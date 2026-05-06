"""
Storage Cleanup Utility
------------------------
A portable Windows utility to scan and delete files older than a configurable
age threshold. Includes protected system-path checks, user exclusions,
dry-run, recycle-bin option, independent file-level selection, and Windows
Task Scheduler integration for automated scan + export.

Features:
- Two interactive modes: Scan & Export, Review & Delete
- Settings & Schedules: user exclusion paths/patterns, scheduled scans
- Headless command-line mode for scheduled execution

Author: Built with Claude (Anthropic)
"""

import argparse
import csv
import fnmatch
import json
import os
import queue
import sys
import threading
from datetime import datetime, timedelta
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from send2trash import send2trash
    SEND2TRASH_AVAILABLE = True
except ImportError:
    SEND2TRASH_AVAILABLE = False

# pywin32 is only available on Windows. We use it for Task Scheduler.
TASK_SCHED_AVAILABLE = False
try:
    if os.name == "nt":
        import win32com.client  # noqa: F401
        import pywintypes  # noqa: F401
        TASK_SCHED_AVAILABLE = True
except ImportError:
    TASK_SCHED_AVAILABLE = False


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

APP_NAME = "Storage Cleanup Utility"
APP_VERSION = "1.1.1"

# Log and state directory in the user's Documents folder
USER_DOCS = Path(os.path.expanduser("~")) / "Documents" / "StorageCleanupUtility"
LOG_DIR = USER_DOCS / "logs"
SCHEDULE_DIR = USER_DOCS / "schedules"
STATE_FILE = USER_DOCS / "state.json"

# Protected system paths — scanning/deleting these is blocked.
# Matched case-insensitively against the user-supplied path.
# Paths are blocked if the user-supplied path EQUALS or is INSIDE any of these.
PROTECTED_PATHS = [
    r"C:\Windows",
    r"C:\Program Files",
    r"C:\Program Files (x86)",
    r"C:\ProgramData",
    r"C:\$Recycle.Bin",
    r"C:\System Volume Information",
    r"C:\Recovery",
    r"C:\Boot",
    r"C:\PerfLogs",
]

# Any path containing \AppData\ is blocked (covers all users' AppData)
PROTECTED_PATH_FRAGMENTS = [
    r"\AppData\Roaming",
    r"\AppData\Local",
    r"\AppData\LocalLow",
]

# The Windows system drive root is also blocked (usually C:\)
SYSTEM_DRIVE = os.environ.get("SystemDrive", "C:").upper()


# ---------------------------------------------------------------------------
# Utility functions
# ---------------------------------------------------------------------------

def ensure_dirs():
    """Create log, schedule, and state directories if they don't exist."""
    LOG_DIR.mkdir(parents=True, exist_ok=True)
    SCHEDULE_DIR.mkdir(parents=True, exist_ok=True)


def format_size(size_bytes):
    """Convert bytes to human-readable string."""
    if size_bytes is None:
        return "—"
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if size_bytes < 1024.0:
            return f"{size_bytes:,.2f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:,.2f} PB"


def normalize_path(path_str):
    """Return a normalized absolute path for comparison."""
    try:
        return os.path.normcase(os.path.abspath(path_str))
    except Exception:
        return os.path.normcase(path_str)


def is_protected_path(path_str):
    """
    Return (is_protected, reason) tuple.
    Blocks the path if it equals or falls inside any protected location.
    """
    if not path_str or not path_str.strip():
        return True, "Path is empty."

    norm = normalize_path(path_str)

    # Block the Windows system drive root exactly (e.g. "c:\")
    system_drive_root = os.path.normcase(SYSTEM_DRIVE + "\\")
    if norm == system_drive_root or norm == os.path.normcase(SYSTEM_DRIVE):
        return True, (
            f"The system drive root ({SYSTEM_DRIVE}\\) is protected. "
            "Please choose a specific subfolder instead."
        )

    # Block the exact protected paths and anything under them
    for p in PROTECTED_PATHS:
        pn = os.path.normcase(p)
        if norm == pn or norm.startswith(pn + os.sep):
            return True, f"This path is protected by the utility: {p}"

    # Block AppData anywhere
    for frag in PROTECTED_PATH_FRAGMENTS:
        if os.path.normcase(frag) in norm:
            return True, (
                f"Paths inside user AppData are protected. "
                f"(Matched fragment: {frag})"
            )

    return False, ""


def to_long_path(path_str):
    r"""
    Convert a Windows path to a long-path form (\\?\...) so that paths longer
    than 260 chars are handled. On non-Windows, returns the path unchanged.
    """
    if os.name != "nt":
        return path_str
    if path_str.startswith("\\\\?\\"):
        return path_str
    abspath = os.path.abspath(path_str)
    if abspath.startswith("\\\\"):
        # UNC path -> \\?\UNC\server\share\...
        return "\\\\?\\UNC\\" + abspath.lstrip("\\")
    return "\\\\?\\" + abspath


# ---------------------------------------------------------------------------
# User exclusions
# ---------------------------------------------------------------------------
# A user exclusion entry is a dict:
#   {
#       "path":    "D:\\Backups\\Critical"  (or "" for pattern-only entry),
#       "pattern": "*.pst"                  (or "" for path-only entry),
#       "note":    "Legal hold — ticket #1234"
#   }
#
# Matching rules:
#   - If path is set and pattern is empty: any file under that path is excluded.
#   - If pattern is set and path is empty: any file matching pattern is excluded
#     (matched against base file name, case-insensitive).
#   - If both are set: file must be under that path AND match the pattern.

def _path_is_under(child, parent):
    """Return True if 'child' equals 'parent' or is inside 'parent'."""
    if not parent:
        return False
    try:
        c = os.path.normcase(os.path.abspath(child))
        p = os.path.normcase(os.path.abspath(parent))
        return c == p or c.startswith(p + os.sep)
    except Exception:
        return False


def file_is_excluded(file_path, file_name, exclusions):
    """
    Return (excluded, matching_entry) — True if any exclusion entry matches.
    """
    if not exclusions:
        return False, None
    fname_lower = (file_name or os.path.basename(file_path)).lower()
    for ex in exclusions:
        path_part = (ex.get("path") or "").strip()
        pat_part = (ex.get("pattern") or "").strip()
        if not path_part and not pat_part:
            continue  # empty entry, ignore
        path_match = True if not path_part else _path_is_under(file_path, path_part)
        pat_match = True if not pat_part else fnmatch.fnmatch(fname_lower, pat_part.lower())
        if path_match and pat_match:
            return True, ex
    return False, None


def dir_is_inside_excluded_path(dir_path, exclusions):
    """
    Return True if dir_path is at or beneath any exclusion entry that has a
    path set (regardless of pattern). This lets us prune entire subtrees
    during the scan walk for efficiency.
    """
    if not exclusions:
        return False
    for ex in exclusions:
        path_part = (ex.get("path") or "").strip()
        if not path_part:
            continue
        if _path_is_under(dir_path, path_part):
            return True
    return False


def validate_exclusion_entry(entry):
    """Return (ok, error_msg). Both path and pattern can be empty individually,
    but at least one must be set."""
    path_part = (entry.get("path") or "").strip()
    pat_part = (entry.get("pattern") or "").strip()
    if not path_part and not pat_part:
        return False, "Provide at least a path, a pattern, or both."
    if path_part:
        # Don't require it to exist — user may add exclusions for paths that
        # might appear later (e.g. on a network share). But warn if obviously bad.
        if any(ch in path_part for ch in '<>"|?*') and not pat_part:
            # < > | are illegal in paths; "*?" suggest the user mistook this for
            # a pattern field
            return False, "Path contains invalid characters. Did you mean to put it in the Pattern field?"
    return True, ""


def load_state():
    """Load last-used settings from state.json (best-effort)."""
    try:
        if STATE_FILE.exists():
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_state(state):
    """Persist last-used settings."""
    try:
        ensure_dirs()
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, indent=2)
    except Exception:
        pass


def timestamp_for_filename():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


# ---------------------------------------------------------------------------
# Scan worker (runs in background thread)
# ---------------------------------------------------------------------------

class ScanWorker:
    """
    Walks the filesystem recursively and reports:
      - per-file info (path, size, mtime, days old)
      - skipped entries (permission / other errors)
      - excluded entries (matched user exclusion rules)
      - progress updates
    Uses a queue to communicate with the GUI thread.
    """

    def __init__(self, root_path, age_days, exclusions, out_queue, cancel_event):
        self.root_path = root_path
        self.age_days = age_days
        self.exclusions = exclusions or []
        self.queue = out_queue
        self.cancel_event = cancel_event

    def run(self):
        try:
            cutoff = datetime.now() - timedelta(days=self.age_days)
            cutoff_ts = cutoff.timestamp()

            files_found = []
            skipped = []
            excluded_paths = []   # excluded files (so user can audit)
            excluded_dirs = []    # pruned subtrees
            files_checked = 0

            walk_root = to_long_path(self.root_path)

            for dirpath, dirnames, filenames in os.walk(walk_root, onerror=self._on_walk_error):
                if self.cancel_event.is_set():
                    break

                # Strip the \\?\ prefix from dirpath for display
                display_dir = dirpath
                if display_dir.startswith("\\\\?\\UNC\\"):
                    display_dir = "\\\\" + display_dir[len("\\\\?\\UNC\\"):]
                elif display_dir.startswith("\\\\?\\"):
                    display_dir = display_dir[len("\\\\?\\"):]

                # Prune subdirs that fall under a path-only exclusion
                kept_dirs = []
                for d in dirnames:
                    sub_display = os.path.join(display_dir, d)
                    if dir_is_inside_excluded_path(sub_display, self.exclusions):
                        # Check that the exclusion has NO pattern (path-only).
                        # If any matching exclusion has a pattern, we still need
                        # to descend to evaluate per-file.
                        only_path_match = False
                        for ex in self.exclusions:
                            ppart = (ex.get("path") or "").strip()
                            patpart = (ex.get("pattern") or "").strip()
                            if ppart and not patpart and _path_is_under(sub_display, ppart):
                                only_path_match = True
                                break
                        if only_path_match:
                            excluded_dirs.append(sub_display)
                            continue
                    kept_dirs.append(d)
                dirnames[:] = kept_dirs

                for name in filenames:
                    if self.cancel_event.is_set():
                        break
                    files_checked += 1
                    full = os.path.join(dirpath, name)
                    display_full = os.path.join(display_dir, name)

                    # Apply user exclusions (path + pattern, or either)
                    ex_hit, ex_entry = file_is_excluded(display_full, name, self.exclusions)
                    if ex_hit:
                        excluded_paths.append({
                            "path": display_full,
                            "rule": _format_excl_rule(ex_entry),
                        })
                        continue

                    try:
                        st = os.stat(full)
                        if st.st_mtime <= cutoff_ts:
                            mtime_dt = datetime.fromtimestamp(st.st_mtime)
                            days_old = (datetime.now() - mtime_dt).days
                            files_found.append({
                                "path": display_full,
                                "name": name,
                                "parent": display_dir,
                                "size": st.st_size,
                                "mtime": mtime_dt,
                                "days_old": days_old,
                            })
                    except Exception as e:
                        skipped.append({
                            "path": display_full,
                            "reason": f"{type(e).__name__}: {e}",
                        })

                    # Progress update every 250 files
                    if files_checked % 250 == 0:
                        self.queue.put(("progress", {
                            "checked": files_checked,
                            "matched": len(files_found),
                            "skipped": len(skipped),
                            "excluded": len(excluded_paths) + len(excluded_dirs),
                            "current": display_dir,
                        }))

            payload = {
                "checked": files_checked,
                "matched": len(files_found),
                "skipped": len(skipped),
                "excluded": len(excluded_paths) + len(excluded_dirs),
                "files": files_found,
                "skipped_list": skipped,
                "excluded_files": excluded_paths,
                "excluded_dirs": excluded_dirs,
            }
            self.queue.put(("cancelled" if self.cancel_event.is_set() else "done", payload))

        except Exception as e:
            self.queue.put(("error", {"message": str(e)}))

    def _on_walk_error(self, err):
        # Errors from os.walk itself (e.g. permission denied on a directory)
        try:
            self.queue.put(("walk_error", {
                "path": getattr(err, "filename", str(err)),
                "reason": str(err),
            }))
        except Exception:
            pass


def _format_excl_rule(entry):
    """Human-readable form of an exclusion entry."""
    if not entry:
        return ""
    parts = []
    if entry.get("path"):
        parts.append(f"path={entry['path']}")
    if entry.get("pattern"):
        parts.append(f"pattern={entry['pattern']}")
    if entry.get("note"):
        parts.append(f"note={entry['note']}")
    return " | ".join(parts)


# ---------------------------------------------------------------------------
# Deletion worker
# ---------------------------------------------------------------------------

class DeleteWorker:
    """
    Deletes the given list of file paths, either to Recycle Bin or permanently.
    Reports progress and results via queue.
    """

    def __init__(self, file_paths, use_recycle_bin, cleanup_empty_dirs,
                 cleanup_roots, out_queue, cancel_event, scan_root=None):
        self.file_paths = file_paths
        self.use_recycle_bin = use_recycle_bin
        self.cleanup_empty_dirs = cleanup_empty_dirs
        self.cleanup_roots = cleanup_roots  # set of parent dirs to check for emptiness
        self.scan_root = scan_root  # bound for ancestor walk; None = no ancestor walk
        self.queue = out_queue
        self.cancel_event = cancel_event

    def run(self):
        results = []
        deleted_count = 0
        deleted_bytes = 0
        failed_count = 0

        for i, path in enumerate(self.file_paths):
            if self.cancel_event.is_set():
                break
            try:
                size = 0
                try:
                    size = os.path.getsize(to_long_path(path))
                except Exception:
                    pass

                if self.use_recycle_bin:
                    if not SEND2TRASH_AVAILABLE:
                        raise RuntimeError(
                            "send2trash library unavailable; cannot send to Recycle Bin."
                        )
                    # send2trash wants the plain path
                    send2trash(os.path.abspath(path))
                else:
                    os.remove(to_long_path(path))

                results.append({
                    "path": path,
                    "status": "DELETED",
                    "reason": "Recycle Bin" if self.use_recycle_bin else "Permanent",
                    "size": size,
                })
                deleted_count += 1
                deleted_bytes += size
            except Exception as e:
                results.append({
                    "path": path,
                    "status": "FAILED",
                    "reason": f"{type(e).__name__}: {e}",
                    "size": 0,
                })
                failed_count += 1

            if (i + 1) % 25 == 0 or i == len(self.file_paths) - 1:
                self.queue.put(("del_progress", {
                    "done": i + 1,
                    "total": len(self.file_paths),
                    "deleted": deleted_count,
                    "failed": failed_count,
                }))

        # Remove empty directories (bottom-up) within the cleanup roots
        removed_dirs = []
        failed_dirs = []
        if self.cleanup_empty_dirs and not self.cancel_event.is_set():
            # Sort directories deepest-first so children are removed before parents
            candidates = sorted(
                set(self.cleanup_roots),
                key=lambda p: len(p),
                reverse=True,
            )
            for d in candidates:
                self._remove_empty_recursive(d, removed_dirs, failed_dirs)

        self.queue.put(("del_done", {
            "results": results,
            "deleted": deleted_count,
            "deleted_bytes": deleted_bytes,
            "failed": failed_count,
            "removed_dirs": removed_dirs,
            "failed_dirs": failed_dirs,
            "cancelled": self.cancel_event.is_set(),
        }))

    def _remove_empty_recursive(self, dirpath, removed, failed):
        """Bottom-up: remove dirpath and any of its subdirectories that are
        now empty. Then walk up to ancestors (bounded by scan_root) and
        remove any that became empty as a result."""
        try:
            long_dir = to_long_path(dirpath)
            if not os.path.isdir(long_dir):
                return
            # Walk bottom-up inside this dir
            for root, dirs, files in os.walk(long_dir, topdown=False):
                if files:
                    continue
                try:
                    if not os.listdir(root):
                        display = root
                        if display.startswith("\\\\?\\UNC\\"):
                            display = "\\\\" + display[len("\\\\?\\UNC\\"):]
                        elif display.startswith("\\\\?\\"):
                            display = display[len("\\\\?\\"):]
                        os.rmdir(root)
                        removed.append(display)
                except Exception as e:
                    failed.append({"path": root, "reason": str(e)})

            # Walk up ancestors but only as long as we stay strictly inside
            # the scan_root. We never delete the scan_root itself or anything
            # above it. This is the safest behavior.
            if self.scan_root:
                self._remove_empty_ancestors(dirpath, removed, failed)
        except Exception as e:
            failed.append({"path": dirpath, "reason": str(e)})

    def _remove_empty_ancestors(self, dirpath, removed, failed):
        """Walk from dirpath upward, removing each folder that is empty.
        Stops at the scan_root boundary or at the first non-empty ancestor."""
        try:
            scan_root_norm = os.path.normcase(os.path.abspath(self.scan_root))
            current = os.path.abspath(dirpath)
            while True:
                parent = os.path.dirname(current)
                if not parent or parent == current:
                    break
                parent_norm = os.path.normcase(os.path.abspath(parent))
                # Critical safety: never go at or above the scan root
                if parent_norm == scan_root_norm:
                    break
                if not parent_norm.startswith(scan_root_norm + os.sep):
                    break
                long_p = to_long_path(parent)
                try:
                    if not os.path.isdir(long_p):
                        break
                    if os.listdir(long_p):
                        break  # has siblings, stop here
                    os.rmdir(long_p)
                    display = parent
                    if display.startswith("\\\\?\\UNC\\"):
                        display = "\\\\" + display[len("\\\\?\\UNC\\"):]
                    elif display.startswith("\\\\?\\"):
                        display = display[len("\\\\?\\"):]
                    removed.append(display)
                    current = parent
                except Exception:
                    break
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Windows Task Scheduler integration
# ---------------------------------------------------------------------------

# Folder name used inside Task Scheduler to group all tasks created by this app
TASK_FOLDER = "\\StorageCleanupUtility"
TASK_PREFIX = "SCU_"   # all task names start with this so we can list them

# Trigger frequency constants from Task Scheduler COM API
TASK_TRIGGER_DAILY   = 2
TASK_TRIGGER_WEEKLY  = 3
TASK_TRIGGER_MONTHLY = 4

# Action type constants
TASK_ACTION_EXEC = 0

# Logon type
TASK_LOGON_INTERACTIVE_TOKEN = 3   # current-user, runs only when logged in
TASK_LOGON_PASSWORD = 1            # store password, runs whether logged in or not

# Task creation flags
TASK_CREATE_OR_UPDATE = 6


def _ts_get_root_folder():
    """Return Task Scheduler root folder, creating our subfolder if missing."""
    if not TASK_SCHED_AVAILABLE:
        raise RuntimeError("Task Scheduler integration requires pywin32 and Windows.")
    scheduler = win32com.client.Dispatch("Schedule.Service")
    scheduler.Connect()
    root = scheduler.GetFolder("\\")
    try:
        return scheduler.GetFolder(TASK_FOLDER)
    except Exception:
        root.CreateFolder(TASK_FOLDER)
        return scheduler.GetFolder(TASK_FOLDER)


def _ts_scheduler():
    scheduler = win32com.client.Dispatch("Schedule.Service")
    scheduler.Connect()
    return scheduler


def _exe_invocation():
    """
    Return (program_path, base_args) needed to launch this utility headlessly.
    When running as a frozen exe, sys.executable IS the exe.
    When running as a script, we run pythonw with the script path.
    """
    if getattr(sys, "frozen", False):
        return sys.executable, ""
    # Script mode (developer / unfrozen): rare in production, but support it
    py = sys.executable.replace("python.exe", "pythonw.exe")
    return py, f'"{os.path.abspath(__file__)}"'


def schedule_create(name, scan_path, age_days, export_dir, exclusions,
                    frequency, time_hhmm, day_of_week=None, day_of_month=None,
                    run_when_not_logged_in=False, password=None):
    """
    Create or update a Windows scheduled task that runs this utility in headless
    scan+export mode.

    Args:
      name: friendly task name (will be prefixed with TASK_PREFIX)
      scan_path: folder to scan
      age_days: age threshold
      export_dir: where the CSV will be written
      exclusions: list of exclusion entries (will be passed via a temp file)
      frequency: 'daily' | 'weekly' | 'monthly'
      time_hhmm: 'HH:MM' (24-hour)
      day_of_week: for weekly — int 1..7 (Sun..Sat) bitmask supported but we use single
      day_of_month: for monthly — int 1..31
      run_when_not_logged_in: bool
      password: optional, required if run_when_not_logged_in is True
    """
    if not TASK_SCHED_AVAILABLE:
        raise RuntimeError("Task Scheduler integration is not available on this system.")

    folder = _ts_get_root_folder()
    scheduler = _ts_scheduler()
    task_def = scheduler.NewTask(0)

    reg_info = task_def.RegistrationInfo
    reg_info.Description = (
        f"Storage Cleanup Utility scheduled scan of {scan_path} "
        f"(files older than {age_days} days)"
    )
    reg_info.Author = "Storage Cleanup Utility"

    settings = task_def.Settings
    settings.Enabled = True
    settings.StartWhenAvailable = True   # catch up if PC was off at scheduled time
    settings.AllowDemandStart = True      # supports "Run Now"
    settings.ExecutionTimeLimit = "PT4H"  # cap a runaway scan at 4h
    settings.MultipleInstances = 2        # IgnoreNew — skip if previous still running
    settings.DisallowStartIfOnBatteries = False
    settings.StopIfGoingOnBatteries = False

    # Trigger
    triggers = task_def.Triggers
    hh, mm = [int(x) for x in time_hhmm.split(":")]
    start_dt = datetime.now().replace(hour=hh, minute=mm, second=0, microsecond=0)
    if start_dt < datetime.now():
        start_dt += timedelta(days=1)
    start_iso = start_dt.strftime("%Y-%m-%dT%H:%M:%S")

    if frequency == "daily":
        t = triggers.Create(TASK_TRIGGER_DAILY)
        t.DaysInterval = 1
    elif frequency == "weekly":
        t = triggers.Create(TASK_TRIGGER_WEEKLY)
        t.WeeksInterval = 1
        # day_of_week stored as a bitmask: 1=Sun, 2=Mon, 4=Tue, 8=Wed, 16=Thu, 32=Fri, 64=Sat
        dow_map = {1: 1, 2: 2, 3: 4, 4: 8, 5: 16, 6: 32, 7: 64}
        t.DaysOfWeek = dow_map.get(day_of_week or 2, 2)
    elif frequency == "monthly":
        t = triggers.Create(TASK_TRIGGER_MONTHLY)
        t.MonthsOfYear = 4095   # all 12 months
        t.DaysOfMonth = 1 << ((day_of_month or 1) - 1)
    else:
        raise ValueError(f"Unsupported frequency: {frequency}")

    t.StartBoundary = start_iso
    t.Enabled = True

    # Action — invoke our exe in headless mode. We persist exclusions to a
    # JSON sidecar so we don't need a long command-line.
    program, base_args = _exe_invocation()
    sidecar = SCHEDULE_DIR / f"{TASK_PREFIX}{_safe_task_name(name)}.json"
    ensure_dirs()
    SCHEDULE_DIR.mkdir(parents=True, exist_ok=True)
    payload = {
        "scan_path": scan_path,
        "age_days": age_days,
        "export_dir": export_dir,
        "exclusions": exclusions or [],
        "task_name": name,
    }
    with open(sidecar, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)

    args = f'{base_args} --headless-from "{sidecar}"'.strip()
    action = task_def.Actions.Create(TASK_ACTION_EXEC)
    action.Path = program
    action.Arguments = args
    action.WorkingDirectory = str(USER_DOCS)

    # Principal (security context)
    principal = task_def.Principal
    principal.RunLevel = 0  # least privilege
    if run_when_not_logged_in and password:
        principal.LogonType = TASK_LOGON_PASSWORD
        username = os.environ.get("USERNAME", "")
        full_name = f"{os.environ.get('USERDOMAIN', '')}\\{username}" if os.environ.get('USERDOMAIN') else username
        principal.UserId = full_name
        folder.RegisterTaskDefinition(
            TASK_PREFIX + _safe_task_name(name),
            task_def,
            TASK_CREATE_OR_UPDATE,
            full_name, password, TASK_LOGON_PASSWORD, ""
        )
    else:
        principal.LogonType = TASK_LOGON_INTERACTIVE_TOKEN
        folder.RegisterTaskDefinition(
            TASK_PREFIX + _safe_task_name(name),
            task_def,
            TASK_CREATE_OR_UPDATE,
            "", "", TASK_LOGON_INTERACTIVE_TOKEN, ""
        )

    return TASK_PREFIX + _safe_task_name(name)


def _safe_task_name(name):
    """Strip characters not allowed in Task Scheduler names."""
    safe = "".join(c for c in name if c.isalnum() or c in " _-")
    safe = safe.strip().replace("  ", " ")
    return safe or "Schedule"


def schedule_list():
    """Return a list of dicts describing all schedules created by this app."""
    if not TASK_SCHED_AVAILABLE:
        return []
    try:
        folder = _ts_get_root_folder()
        out = []
        for task in folder.GetTasks(0):
            try:
                xml = task.Xml
            except Exception:
                xml = ""
            out.append({
                "name": task.Name,
                "enabled": bool(task.Enabled),
                "state": task.State,   # 1=disabled, 3=ready, 4=running
                "last_run": str(task.LastRunTime) if task.LastRunTime else "",
                "next_run": str(task.NextRunTime) if task.NextRunTime else "",
                "last_result": task.LastTaskResult,
                "_task": task,  # keep a handle for delete/run-now
            })
        return out
    except Exception:
        return []


def schedule_delete(task_name):
    if not TASK_SCHED_AVAILABLE:
        raise RuntimeError("Task Scheduler integration not available.")
    folder = _ts_get_root_folder()
    folder.DeleteTask(task_name, 0)
    # remove sidecar
    sidecar = SCHEDULE_DIR / f"{task_name}.json"
    try:
        sidecar.unlink(missing_ok=True)
    except Exception:
        pass


def schedule_run_now(task_name):
    if not TASK_SCHED_AVAILABLE:
        raise RuntimeError("Task Scheduler integration not available.")
    folder = _ts_get_root_folder()
    task = folder.GetTask(task_name)
    task.Run("")


def schedule_set_enabled(task_name, enabled):
    if not TASK_SCHED_AVAILABLE:
        raise RuntimeError("Task Scheduler integration not available.")
    folder = _ts_get_root_folder()
    task = folder.GetTask(task_name)
    task.Enabled = bool(enabled)


# ---------------------------------------------------------------------------
# Headless mode (launched by Task Scheduler)
# ---------------------------------------------------------------------------

def run_headless(sidecar_path):
    """
    Headless scan + export, driven by a JSON sidecar produced when the task
    was created. Runs without showing any GUI. Logs to the standard logs folder.
    """
    ensure_dirs()
    log_path = LOG_DIR / f"scheduled_run_{timestamp_for_filename()}.log"

    def log(msg):
        try:
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
        except Exception:
            pass

    try:
        with open(sidecar_path, "r", encoding="utf-8") as f:
            spec = json.load(f)
    except Exception as e:
        log(f"FATAL: could not read sidecar {sidecar_path}: {e}")
        return 2

    scan_path = spec.get("scan_path", "")
    age_days = int(spec.get("age_days", 30))
    export_dir = spec.get("export_dir", "") or str(USER_DOCS / "scheduled_exports")
    exclusions = spec.get("exclusions", []) or []
    task_name = spec.get("task_name", "Schedule")

    log(f"Starting scheduled scan task='{task_name}' path='{scan_path}' age={age_days}")

    blocked, reason = is_protected_path(scan_path)
    if blocked:
        log(f"BLOCKED: {reason}")
        return 3

    if not os.path.isdir(scan_path):
        log(f"FAILED: scan path is not a folder: {scan_path}")
        return 4

    Path(export_dir).mkdir(parents=True, exist_ok=True)

    # Run scan synchronously (we're in a background process anyway)
    q = queue.Queue()
    cancel = threading.Event()
    worker = ScanWorker(scan_path, age_days, exclusions, q, cancel)
    worker.run()

    # Drain queue, find the final 'done' or 'cancelled' message
    final = None
    while True:
        try:
            msg, data = q.get_nowait()
        except queue.Empty:
            break
        if msg in ("done", "cancelled", "error"):
            final = (msg, data)

    if not final:
        log("FAILED: scan produced no final result")
        return 5

    msg, data = final
    if msg == "error":
        log(f"FAILED: scan error: {data.get('message')}")
        return 6

    files = data["files"]
    skipped = data["skipped_list"]

    # Export CSV
    csv_name = f"scheduled_{_safe_task_name(task_name)}_{timestamp_for_filename()}.csv"
    csv_path = Path(export_dir) / csv_name
    try:
        with open(csv_path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["FullPath", "FileName", "ParentFolder", "SizeBytes", "SizeReadable",
                        "LastModified", "DaysOld", "Selected"])
            for r in files:
                w.writerow([
                    r["path"], r["name"], r["parent"], r["size"],
                    format_size(r["size"]),
                    r["mtime"].strftime("%Y-%m-%d %H:%M:%S"),
                    r["days_old"], "Yes",
                ])
        log(f"Exported {len(files):,} files to {csv_path}")
    except Exception as e:
        log(f"FAILED: could not write CSV: {e}")
        return 7

    if skipped:
        skipped_path = csv_path.with_name(csv_path.stem + "_skipped.csv")
        try:
            with open(skipped_path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["Path", "Reason"])
                for s in skipped:
                    w.writerow([s["path"], s["reason"]])
            log(f"Skipped entries: {len(skipped)} -> {skipped_path}")
        except Exception as e:
            log(f"WARN: could not write skipped CSV: {e}")

    log(f"DONE. Matched {len(files):,} files, total size {format_size(sum(f['size'] for f in files))}")
    return 0


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------

class StorageCleanupApp(tk.Tk):

    def __init__(self):
        super().__init__()
        ensure_dirs()

        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry("1180x760")
        self.minsize(960, 640)

        # State
        self.scan_results = []       # list of dicts from ScanWorker
        self.skipped_results = []
        self.scan_queue = queue.Queue()
        self.scan_cancel = threading.Event()
        self.scan_thread = None

        self.del_queue = queue.Queue()
        self.del_cancel = threading.Event()
        self.del_thread = None

        # Tree item -> file info map for the Review tab
        self.tree_items = {}   # item_id -> {"type": "folder"|"file", "path": ..., "size": ..., "checked": bool, "file_ref": dict|None}
        # Path-keyed checked state, persists across tree rebuilds (e.g. when
        # the user types in the search box, the tree is rebuilt with only the
        # visible/matching files, but their previous checkbox state is restored
        # from this dict).
        self.checked_paths = {}   # full file path -> bool

        self.settings = load_state()
        # User-defined exclusion entries — list of dicts {path, pattern, note}
        self.exclusions = list(self.settings.get("exclusions", []))

        self._build_ui()
        self._poll_scan_queue()
        self._poll_del_queue()

    # ----------------------- UI construction -----------------------

    def _build_ui(self):
        style = ttk.Style(self)
        try:
            style.theme_use("vista" if os.name == "nt" else "clam")
        except Exception:
            pass

        # Top frame: app title
        header = ttk.Frame(self, padding=(10, 6))
        header.pack(fill=tk.X)
        ttk.Label(header, text=APP_NAME, font=("Segoe UI", 14, "bold")).pack(side=tk.LEFT)
        ttk.Label(header, text=f"  v{APP_VERSION}", foreground="#666").pack(side=tk.LEFT)

        # Notebook with two tabs
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 4))

        self.scan_tab = ttk.Frame(self.nb)
        self.review_tab = ttk.Frame(self.nb)
        self.settings_tab = ttk.Frame(self.nb)
        self.nb.add(self.scan_tab, text="  1.  Scan & Export  ")
        self.nb.add(self.review_tab, text="  2.  Review & Delete  ")
        self.nb.add(self.settings_tab, text="  3.  Settings & Schedules  ")

        self._build_scan_tab()
        self._build_review_tab()
        self._build_settings_tab()

        # Status bar
        self.status_var = tk.StringVar(value="Ready.")
        status = ttk.Frame(self, padding=(8, 4), relief=tk.SUNKEN)
        status.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Label(status, textvariable=self.status_var).pack(side=tk.LEFT)
        ttk.Label(status, text=f"Logs: {LOG_DIR}", foreground="#666").pack(side=tk.RIGHT)

    # ----------------------- SCAN TAB -----------------------

    def _build_scan_tab(self):
        f = self.scan_tab

        # Top input row
        top = ttk.LabelFrame(f, text="Scan Settings", padding=10)
        top.pack(fill=tk.X, padx=6, pady=6)

        ttk.Label(top, text="Folder to scan:").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.path_var = tk.StringVar(value=self.settings.get("last_path", ""))
        self.path_entry = ttk.Entry(top, textvariable=self.path_var, width=80)
        self.path_entry.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        ttk.Button(top, text="Browse…", command=self._browse_folder).grid(row=0, column=2)

        ttk.Label(top, text="Delete files older than (days):").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.age_var = tk.StringVar(value=str(self.settings.get("last_age", 30)))
        age_entry = ttk.Entry(top, textvariable=self.age_var, width=10)
        age_entry.grid(row=1, column=1, sticky="w", pady=(8, 0))

        btns = ttk.Frame(top)
        btns.grid(row=1, column=2, sticky="e", pady=(8, 0))
        self.scan_btn = ttk.Button(btns, text="Start Scan", command=self._start_scan)
        self.scan_btn.pack(side=tk.LEFT, padx=(0, 4))
        self.scan_cancel_btn = ttk.Button(btns, text="Cancel", command=self._cancel_scan, state=tk.DISABLED)
        self.scan_cancel_btn.pack(side=tk.LEFT)

        top.columnconfigure(1, weight=1)

        # Progress
        prog_frame = ttk.Frame(f, padding=(6, 0))
        prog_frame.pack(fill=tk.X)
        self.scan_progress = ttk.Progressbar(prog_frame, mode="indeterminate")
        self.scan_progress.pack(fill=tk.X, pady=(2, 4))

        # Filter + summary
        mid = ttk.Frame(f, padding=(6, 0))
        mid.pack(fill=tk.X)
        ttk.Label(mid, text="Filter:").pack(side=tk.LEFT)
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", lambda *a: self._apply_scan_filter())
        ttk.Entry(mid, textvariable=self.filter_var, width=40).pack(side=tk.LEFT, padx=(4, 12))

        self.scan_summary_var = tk.StringVar(value="Total files: 0   |   Total size: 0 B   |   Skipped: 0   |   Excluded: 0")
        ttk.Label(mid, textvariable=self.scan_summary_var, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)

        ttk.Button(mid, text="Export to CSV…", command=self._export_csv).pack(side=tk.RIGHT)
        ttk.Button(mid, text="Proceed to Review →", command=self._goto_review).pack(side=tk.RIGHT, padx=(0, 6))

        # Results table
        tbl_frame = ttk.Frame(f, padding=6)
        tbl_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("name", "path", "size", "mtime", "days")
        self.scan_tree = ttk.Treeview(tbl_frame, columns=columns, show="headings", selectmode="extended")
        self.scan_tree.heading("name", text="File Name", command=lambda: self._sort_scan_by("name"))
        self.scan_tree.heading("path", text="Path", command=lambda: self._sort_scan_by("path"))
        self.scan_tree.heading("size", text="Size", command=lambda: self._sort_scan_by("size"))
        self.scan_tree.heading("mtime", text="Last Modified", command=lambda: self._sort_scan_by("mtime"))
        self.scan_tree.heading("days", text="Days Old", command=lambda: self._sort_scan_by("days"))

        self.scan_tree.column("name", width=220, anchor="w")
        self.scan_tree.column("path", width=480, anchor="w")
        self.scan_tree.column("size", width=110, anchor="e")
        self.scan_tree.column("mtime", width=150, anchor="w")
        self.scan_tree.column("days", width=80, anchor="e")

        ysb = ttk.Scrollbar(tbl_frame, orient="vertical", command=self.scan_tree.yview)
        xsb = ttk.Scrollbar(tbl_frame, orient="horizontal", command=self.scan_tree.xview)
        self.scan_tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)

        self.scan_tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        tbl_frame.rowconfigure(0, weight=1)
        tbl_frame.columnconfigure(0, weight=1)

        self._scan_sort_state = {"col": "days", "reverse": True}

    # ----------------------- REVIEW TAB -----------------------

    def _build_review_tab(self):
        f = self.review_tab

        top = ttk.LabelFrame(f, text="Source", padding=10)
        top.pack(fill=tk.X, padx=6, pady=6)

        ttk.Button(top, text="Use Current Scan Results", command=self._load_from_scan).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(top, text="Load from CSV…", command=self._load_from_csv).grid(row=0, column=1)

        self.review_source_var = tk.StringVar(value="No data loaded.")
        ttk.Label(top, textvariable=self.review_source_var, foreground="#555").grid(row=0, column=2, padx=12, sticky="w")

        # Options
        opts = ttk.LabelFrame(f, text="Deletion Options", padding=10)
        opts.pack(fill=tk.X, padx=6, pady=(0, 6))

        self.delete_mode_var = tk.StringVar(value="recycle")
        ttk.Radiobutton(opts, text="Send to Recycle Bin (safer)", variable=self.delete_mode_var, value="recycle").grid(row=0, column=0, sticky="w", padx=(0, 20))
        ttk.Radiobutton(opts, text="Permanent Delete (cannot undo)", variable=self.delete_mode_var, value="permanent").grid(row=0, column=1, sticky="w")

        self.cleanup_empty_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(opts, text="Remove empty folders after deletion", variable=self.cleanup_empty_var).grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

        # Controls row
        ctrl = ttk.Frame(f, padding=(6, 0))
        ctrl.pack(fill=tk.X)
        ttk.Button(ctrl, text="Select All", command=lambda: self._bulk_check(True)).pack(side=tk.LEFT)
        ttk.Button(ctrl, text="Deselect All", command=lambda: self._bulk_check(False)).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Button(ctrl, text="Expand All", command=lambda: self._expand_all(True)).pack(side=tk.LEFT)
        ttk.Button(ctrl, text="Collapse All", command=lambda: self._expand_all(False)).pack(side=tk.LEFT, padx=(4, 12))

        ttk.Label(ctrl, text="Sort files by:").pack(side=tk.LEFT)
        self.review_sort_var = tk.StringVar(value="size_desc")
        sort_cb = ttk.Combobox(ctrl, textvariable=self.review_sort_var, state="readonly", width=22,
                               values=[
                                   "size_desc", "size_asc",
                                   "days_desc", "days_asc",
                                   "name_asc", "name_desc",
                                   "mtime_desc", "mtime_asc",
                               ])
        sort_cb.pack(side=tk.LEFT, padx=(4, 0))
        sort_cb.bind("<<ComboboxSelected>>", lambda e: self._rebuild_tree())

        # Summary (live)
        self.review_summary_var = tk.StringVar(
            value="Total: 0 files, 0 B   |   Selected: 0 files, 0 B"
        )
        ttk.Label(ctrl, textvariable=self.review_summary_var, font=("Segoe UI", 9, "bold")).pack(side=tk.RIGHT)

        # Search row — filters the tree to show only matching files (and their
        # parent folders for context). Checkbox state is preserved across
        # filter changes.
        search = ttk.Frame(f, padding=(6, 4))
        search.pack(fill=tk.X)
        ttk.Label(search, text="Search:").pack(side=tk.LEFT)
        self.review_search_var = tk.StringVar()
        self.review_search_var.trace_add("write", lambda *a: self._on_review_search_change())
        search_entry = ttk.Entry(search, textvariable=self.review_search_var, width=50)
        search_entry.pack(side=tk.LEFT, padx=(4, 6))

        ttk.Button(search, text="Clear", command=lambda: self.review_search_var.set("")).pack(side=tk.LEFT)
        ttk.Separator(search, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=8)

        # Bulk actions limited to currently-visible (filtered) items.
        # These are the most useful actions when reviewing a search result.
        ttk.Label(search, text="On visible:").pack(side=tk.LEFT)
        ttk.Button(search, text="Deselect (keep these)",
                   command=lambda: self._bulk_check_visible(False)).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(search, text="Select",
                   command=lambda: self._bulk_check_visible(True)).pack(side=tk.LEFT, padx=(4, 0))

        self.review_search_status_var = tk.StringVar(value="")
        ttk.Label(search, textvariable=self.review_search_status_var,
                  foreground="#0066aa").pack(side=tk.LEFT, padx=(12, 0))

        # Tree
        tree_frame = ttk.Frame(f, padding=6)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("size", "mtime", "days")
        self.review_tree = ttk.Treeview(tree_frame, columns=columns, show="tree headings", selectmode="none")
        self.review_tree.heading("#0", text="  Folder / File (click checkbox to toggle)")
        self.review_tree.heading("size", text="Size")
        self.review_tree.heading("mtime", text="Last Modified")
        self.review_tree.heading("days", text="Days Old")
        self.review_tree.column("#0", width=620, anchor="w")
        self.review_tree.column("size", width=110, anchor="e")
        self.review_tree.column("mtime", width=160, anchor="w")
        self.review_tree.column("days", width=80, anchor="e")

        ysb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.review_tree.yview)
        xsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.review_tree.xview)
        self.review_tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)

        self.review_tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.review_tree.bind("<Button-1>", self._on_tree_click)

        # Action row
        act = ttk.Frame(f, padding=6)
        act.pack(fill=tk.X)
        ttk.Button(act, text="Dry Run (Preview)", command=self._dry_run).pack(side=tk.LEFT)
        self.delete_btn = ttk.Button(act, text="DELETE Selected", command=self._start_delete)
        self.delete_btn.pack(side=tk.LEFT, padx=(6, 0))
        self.del_cancel_btn = ttk.Button(act, text="Cancel", command=self._cancel_delete, state=tk.DISABLED)
        self.del_cancel_btn.pack(side=tk.LEFT, padx=(6, 0))

        self.del_progress = ttk.Progressbar(act, mode="determinate")
        self.del_progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=12)

    # ----------------------- SETTINGS & SCHEDULES TAB -----------------------

    def _build_settings_tab(self):
        f = self.settings_tab

        # Two side-by-side sections: Exclusions on the left, Schedules on the right
        paned = ttk.PanedWindow(f, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        # ---------- Exclusions section ----------
        excl_frame = ttk.LabelFrame(paned, text="Exclusion Rules", padding=8)
        paned.add(excl_frame, weight=1)

        ttk.Label(
            excl_frame,
            text=("Files matching any rule below will be skipped during scan and delete.\n"
                  "Provide a path, a filename pattern (e.g. *.pst, *important*), or both."),
            foreground="#444",
            justify="left",
        ).pack(anchor="w", pady=(0, 6))

        # Listbox of exclusions with columns
        list_frame = ttk.Frame(excl_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("path", "pattern", "note")
        self.excl_tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse", height=12)
        self.excl_tree.heading("path", text="Path")
        self.excl_tree.heading("pattern", text="Pattern")
        self.excl_tree.heading("note", text="Note")
        self.excl_tree.column("path", width=240, anchor="w")
        self.excl_tree.column("pattern", width=110, anchor="w")
        self.excl_tree.column("note", width=160, anchor="w")

        ysb = ttk.Scrollbar(list_frame, orient="vertical", command=self.excl_tree.yview)
        self.excl_tree.configure(yscrollcommand=ysb.set)
        self.excl_tree.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        btns = ttk.Frame(excl_frame)
        btns.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(btns, text="Add…", command=self._excl_add).pack(side=tk.LEFT)
        ttk.Button(btns, text="Edit…", command=self._excl_edit).pack(side=tk.LEFT, padx=(4, 0))
        ttk.Button(btns, text="Remove", command=self._excl_remove).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Button(btns, text="Import CSV…", command=self._excl_import).pack(side=tk.LEFT)
        ttk.Button(btns, text="Export CSV…", command=self._excl_export).pack(side=tk.LEFT, padx=(4, 0))

        self._refresh_excl_tree()

        # ---------- Schedules section ----------
        sched_frame = ttk.LabelFrame(paned, text="Scheduled Scans (Windows Task Scheduler)", padding=8)
        paned.add(sched_frame, weight=1)

        if not TASK_SCHED_AVAILABLE:
            ttk.Label(
                sched_frame,
                text=("Task Scheduler integration is not available.\n"
                      "Make sure you are on Windows and the build includes pywin32."),
                foreground="#a00",
                justify="left",
            ).pack(anchor="w", pady=8)
        else:
            ttk.Label(
                sched_frame,
                text=("Schedules run scan + export automatically. They never delete.\n"
                      "Each scan saves a CSV in the chosen export folder."),
                foreground="#444",
                justify="left",
            ).pack(anchor="w", pady=(0, 6))

            list_frame2 = ttk.Frame(sched_frame)
            list_frame2.pack(fill=tk.BOTH, expand=True)
            cols = ("name", "next_run", "last_run", "state")
            self.sched_tree = ttk.Treeview(list_frame2, columns=cols, show="headings", selectmode="browse", height=10)
            self.sched_tree.heading("name", text="Task Name")
            self.sched_tree.heading("next_run", text="Next Run")
            self.sched_tree.heading("last_run", text="Last Run")
            self.sched_tree.heading("state", text="State")
            self.sched_tree.column("name", width=180, anchor="w")
            self.sched_tree.column("next_run", width=140, anchor="w")
            self.sched_tree.column("last_run", width=140, anchor="w")
            self.sched_tree.column("state", width=80, anchor="w")
            ysb2 = ttk.Scrollbar(list_frame2, orient="vertical", command=self.sched_tree.yview)
            self.sched_tree.configure(yscrollcommand=ysb2.set)
            self.sched_tree.grid(row=0, column=0, sticky="nsew")
            ysb2.grid(row=0, column=1, sticky="ns")
            list_frame2.rowconfigure(0, weight=1)
            list_frame2.columnconfigure(0, weight=1)

            btns2 = ttk.Frame(sched_frame)
            btns2.pack(fill=tk.X, pady=(6, 0))
            ttk.Button(btns2, text="New Schedule…", command=self._sched_new).pack(side=tk.LEFT)
            ttk.Button(btns2, text="Run Now", command=self._sched_run_now).pack(side=tk.LEFT, padx=(4, 0))
            ttk.Button(btns2, text="Enable / Disable", command=self._sched_toggle).pack(side=tk.LEFT, padx=(4, 0))
            ttk.Button(btns2, text="Delete", command=self._sched_delete).pack(side=tk.LEFT, padx=(4, 12))
            ttk.Button(btns2, text="Refresh", command=self._refresh_sched_tree).pack(side=tk.LEFT)

            self._refresh_sched_tree()

    # ----------------------- Exclusions logic -----------------------

    def _refresh_excl_tree(self):
        for i in self.excl_tree.get_children():
            self.excl_tree.delete(i)
        for ex in self.exclusions:
            self.excl_tree.insert(
                "", tk.END,
                values=(ex.get("path", ""), ex.get("pattern", ""), ex.get("note", "")),
            )

    def _save_exclusions(self):
        self.settings["exclusions"] = self.exclusions
        save_state(self.settings)

    def _excl_add(self):
        result = ExclusionDialog(self, "Add Exclusion").result
        if not result:
            return
        ok, err = validate_exclusion_entry(result)
        if not ok:
            messagebox.showerror("Invalid entry", err)
            return
        self.exclusions.append(result)
        self._save_exclusions()
        self._refresh_excl_tree()

    def _excl_edit(self):
        sel = self.excl_tree.selection()
        if not sel:
            messagebox.showinfo("No selection", "Select an exclusion to edit.")
            return
        idx = self.excl_tree.index(sel[0])
        current = self.exclusions[idx]
        result = ExclusionDialog(self, "Edit Exclusion", current).result
        if not result:
            return
        ok, err = validate_exclusion_entry(result)
        if not ok:
            messagebox.showerror("Invalid entry", err)
            return
        self.exclusions[idx] = result
        self._save_exclusions()
        self._refresh_excl_tree()

    def _excl_remove(self):
        sel = self.excl_tree.selection()
        if not sel:
            messagebox.showinfo("No selection", "Select an exclusion to remove.")
            return
        idx = self.excl_tree.index(sel[0])
        if messagebox.askyesno("Remove exclusion", "Remove the selected exclusion?"):
            self.exclusions.pop(idx)
            self._save_exclusions()
            self._refresh_excl_tree()

    def _excl_export(self):
        if not self.exclusions:
            messagebox.showinfo("Nothing to export", "No exclusion rules defined.")
            return
        path = filedialog.asksaveasfilename(
            title="Export exclusion rules",
            defaultextension=".csv",
            initialfile=f"exclusions_{timestamp_for_filename()}.csv",
            filetypes=[("CSV files", "*.csv")],
        )
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["Path", "Pattern", "Note"])
                for ex in self.exclusions:
                    w.writerow([ex.get("path", ""), ex.get("pattern", ""), ex.get("note", "")])
            messagebox.showinfo("Exported", f"Exclusion rules exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Export failed", str(e))

    def _excl_import(self):
        path = filedialog.askopenfilename(
            title="Import exclusion rules CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            new_entries = []
            with open(path, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    entry = {
                        "path": (row.get("Path") or row.get("path") or "").strip(),
                        "pattern": (row.get("Pattern") or row.get("pattern") or "").strip(),
                        "note": (row.get("Note") or row.get("note") or "").strip(),
                    }
                    ok, _ = validate_exclusion_entry(entry)
                    if ok:
                        new_entries.append(entry)
            choice = messagebox.askyesnocancel(
                "Import exclusions",
                f"Imported {len(new_entries)} valid entries.\n\n"
                f"Yes  = Replace existing rules\n"
                f"No   = Append to existing rules\n"
                f"Cancel = Abort"
            )
            if choice is None:
                return
            if choice:
                self.exclusions = new_entries
            else:
                self.exclusions.extend(new_entries)
            self._save_exclusions()
            self._refresh_excl_tree()
        except Exception as e:
            messagebox.showerror("Import failed", str(e))

    # ----------------------- Schedules logic -----------------------

    def _refresh_sched_tree(self):
        if not TASK_SCHED_AVAILABLE:
            return
        for i in self.sched_tree.get_children():
            self.sched_tree.delete(i)
        try:
            tasks = schedule_list()
        except Exception as e:
            messagebox.showerror("Cannot list schedules", str(e))
            return
        state_map = {0: "Unknown", 1: "Disabled", 2: "Queued", 3: "Ready", 4: "Running"}
        for t in tasks:
            self.sched_tree.insert(
                "", tk.END,
                values=(t["name"], t.get("next_run") or "—", t.get("last_run") or "—",
                        state_map.get(t.get("state"), "—"))
            )

    def _sched_new(self):
        if not TASK_SCHED_AVAILABLE:
            messagebox.showerror("Unavailable", "Task Scheduler integration not available.")
            return
        dlg = ScheduleDialog(self, exclusions=self.exclusions)
        if not dlg.result:
            return
        spec = dlg.result
        try:
            task_name = schedule_create(
                name=spec["name"],
                scan_path=spec["scan_path"],
                age_days=spec["age_days"],
                export_dir=spec["export_dir"],
                exclusions=spec["exclusions"],
                frequency=spec["frequency"],
                time_hhmm=spec["time"],
                day_of_week=spec.get("day_of_week"),
                day_of_month=spec.get("day_of_month"),
                run_when_not_logged_in=spec.get("run_when_not_logged_in", False),
                password=spec.get("password"),
            )
            messagebox.showinfo(
                "Schedule created",
                f"Created scheduled task:\n  {task_name}\n\n"
                f"You can verify or edit it via the Windows Task Scheduler if needed.\n"
                f"Use 'Run Now' to test it immediately."
            )
            self._refresh_sched_tree()
        except Exception as e:
            messagebox.showerror("Could not create schedule", str(e))

    def _sched_selected_name(self):
        sel = self.sched_tree.selection()
        if not sel:
            messagebox.showinfo("No selection", "Select a schedule from the list.")
            return None
        vals = self.sched_tree.item(sel[0], "values")
        return vals[0] if vals else None

    def _sched_run_now(self):
        name = self._sched_selected_name()
        if not name:
            return
        try:
            schedule_run_now(name)
            messagebox.showinfo("Triggered", f"Run-now triggered for:\n{name}\n\n"
                                f"Check the export folder and the logs folder shortly.")
            self._refresh_sched_tree()
        except Exception as e:
            messagebox.showerror("Run failed", str(e))

    def _sched_toggle(self):
        name = self._sched_selected_name()
        if not name:
            return
        try:
            tasks = schedule_list()
            current = next((t for t in tasks if t["name"] == name), None)
            new_state = not (current and current.get("enabled"))
            schedule_set_enabled(name, new_state)
            self._refresh_sched_tree()
        except Exception as e:
            messagebox.showerror("Toggle failed", str(e))

    def _sched_delete(self):
        name = self._sched_selected_name()
        if not name:
            return
        if not messagebox.askyesno("Delete schedule", f"Delete this schedule?\n\n  {name}"):
            return
        try:
            schedule_delete(name)
            self._refresh_sched_tree()
        except Exception as e:
            messagebox.showerror("Delete failed", str(e))

    # ----------------------- Scan tab logic -----------------------

    def _browse_folder(self):
        start = self.path_var.get() or os.path.expanduser("~")
        d = filedialog.askdirectory(title="Choose folder to scan", initialdir=start)
        if d:
            self.path_var.set(os.path.normpath(d))

    def _validate_scan_inputs(self):
        path = self.path_var.get().strip().strip('"')
        if not path:
            messagebox.showerror("Path required", "Please select a folder to scan.")
            return None

        if not os.path.isdir(path):
            messagebox.showerror("Not a folder", f"This path is not a valid folder:\n\n{path}")
            return None

        blocked, reason = is_protected_path(path)
        if blocked:
            messagebox.showerror(
                "Protected path blocked",
                f"Cannot scan this location.\n\n{reason}"
            )
            return None

        # If the user is trying to scan inside one of their own exclusions, warn
        # (but allow with confirmation — user might want to scan a parent and
        # let pattern exclusions filter out specific files within it).
        for ex in self.exclusions:
            ppart = (ex.get("path") or "").strip()
            patpart = (ex.get("pattern") or "").strip()
            if ppart and not patpart and _path_is_under(path, ppart):
                if not messagebox.askyesno(
                    "Path is in your exclusion list",
                    f"The folder you are about to scan is inside an excluded path:\n\n"
                    f"  {ppart}\n\n"
                    f"With this exclusion in place, the entire scan would skip everything.\n\n"
                    f"Continue anyway?"
                ):
                    return None
                break

        try:
            age = int(self.age_var.get().strip())
            if age < 0:
                raise ValueError()
        except ValueError:
            messagebox.showerror("Invalid age", "Please enter a non-negative integer for days.")
            return None

        return path, age

    def _start_scan(self):
        validated = self._validate_scan_inputs()
        if not validated:
            return
        path, age = validated

        # Persist
        self.settings["last_path"] = path
        self.settings["last_age"] = age
        save_state(self.settings)

        # Clear previous
        self.scan_results = []
        self.skipped_results = []
        for i in self.scan_tree.get_children():
            self.scan_tree.delete(i)

        self.scan_cancel.clear()
        self.scan_btn.configure(state=tk.DISABLED)
        self.scan_cancel_btn.configure(state=tk.NORMAL)
        self.scan_progress.start(10)
        self.status_var.set(f"Scanning {path} …")
        self._update_scan_summary(0, 0, 0, 0)

        worker = ScanWorker(path, age, self.exclusions, self.scan_queue, self.scan_cancel)
        self.scan_thread = threading.Thread(target=worker.run, daemon=True)
        self.scan_thread.start()

    def _cancel_scan(self):
        self.scan_cancel.set()
        self.status_var.set("Cancelling scan…")

    def _poll_scan_queue(self):
        try:
            while True:
                msg, data = self.scan_queue.get_nowait()
                if msg == "progress":
                    self.status_var.set(
                        f"Scanning… checked {data['checked']:,} | matched {data['matched']:,} "
                        f"| skipped {data['skipped']:,} | excluded {data.get('excluded', 0):,}"
                    )
                elif msg == "walk_error":
                    self.skipped_results.append({"path": data["path"], "reason": data["reason"]})
                elif msg in ("done", "cancelled"):
                    self.scan_progress.stop()
                    self.scan_btn.configure(state=tk.NORMAL)
                    self.scan_cancel_btn.configure(state=tk.DISABLED)
                    self.scan_results = data["files"]
                    self.skipped_results.extend(data["skipped_list"])
                    # Stash excluded info so we can include it in the export
                    self._last_excluded_files = data.get("excluded_files", [])
                    self._last_excluded_dirs = data.get("excluded_dirs", [])
                    self._populate_scan_tree(self.scan_results)
                    total_size = sum(f["size"] for f in self.scan_results)
                    self._update_scan_summary(
                        len(self.scan_results), total_size,
                        len(self.skipped_results), data.get("excluded", 0)
                    )
                    verb = "Cancelled" if msg == "cancelled" else "Scan complete"
                    extra = ""
                    if data.get("excluded", 0):
                        extra = f", {data['excluded']:,} excluded by your rules"
                    self.status_var.set(
                        f"{verb}. {len(self.scan_results):,} files matched, "
                        f"{format_size(total_size)} total{extra}."
                    )
                elif msg == "error":
                    self.scan_progress.stop()
                    self.scan_btn.configure(state=tk.NORMAL)
                    self.scan_cancel_btn.configure(state=tk.DISABLED)
                    self.status_var.set("Scan failed.")
                    messagebox.showerror("Scan error", data["message"])
        except queue.Empty:
            pass
        self.after(150, self._poll_scan_queue)

    def _populate_scan_tree(self, files):
        for i in self.scan_tree.get_children():
            self.scan_tree.delete(i)
        self._apply_scan_filter(files=files)

    def _apply_scan_filter(self, files=None):
        if files is None:
            files = self.scan_results
        q = self.filter_var.get().strip().lower()
        for i in self.scan_tree.get_children():
            self.scan_tree.delete(i)
        for f in files:
            if q and q not in f["name"].lower() and q not in f["path"].lower():
                continue
            self.scan_tree.insert("", tk.END, values=(
                f["name"],
                f["path"],
                format_size(f["size"]),
                f["mtime"].strftime("%Y-%m-%d %H:%M:%S"),
                f["days_old"],
            ))

    def _sort_scan_by(self, col):
        rev = (self._scan_sort_state["col"] == col and not self._scan_sort_state["reverse"])
        self._scan_sort_state = {"col": col, "reverse": rev}
        key_map = {
            "name": lambda f: f["name"].lower(),
            "path": lambda f: f["path"].lower(),
            "size": lambda f: f["size"],
            "mtime": lambda f: f["mtime"],
            "days": lambda f: f["days_old"],
        }
        self.scan_results.sort(key=key_map[col], reverse=rev)
        self._apply_scan_filter()

    def _update_scan_summary(self, count, total_size, skipped_count, excluded_count=0):
        self.scan_summary_var.set(
            f"Total files: {count:,}   |   Total size: {format_size(total_size)}   "
            f"|   Skipped: {skipped_count:,}   |   Excluded: {excluded_count:,}"
        )

    def _export_csv(self):
        if not self.scan_results:
            messagebox.showinfo("Nothing to export", "No scan results to export. Run a scan first.")
            return
        default = f"scan_results_{timestamp_for_filename()}.csv"
        path = filedialog.asksaveasfilename(
            title="Export scan results",
            defaultextension=".csv",
            initialfile=default,
            filetypes=[("CSV files", "*.csv")],
        )
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["FullPath", "FileName", "ParentFolder", "SizeBytes", "SizeReadable",
                            "LastModified", "DaysOld", "Selected"])
                for r in self.scan_results:
                    w.writerow([
                        r["path"], r["name"], r["parent"], r["size"],
                        format_size(r["size"]),
                        r["mtime"].strftime("%Y-%m-%d %H:%M:%S"),
                        r["days_old"], "Yes",
                    ])

            # Also write skipped list next to it
            if self.skipped_results:
                skipped_path = os.path.splitext(path)[0] + "_skipped.csv"
                with open(skipped_path, "w", newline="", encoding="utf-8") as f:
                    w = csv.writer(f)
                    w.writerow(["Path", "Reason"])
                    for s in self.skipped_results:
                        w.writerow([s["path"], s["reason"]])
                messagebox.showinfo(
                    "Export complete",
                    f"Exported {len(self.scan_results):,} files to:\n{path}\n\n"
                    f"Skipped entries ({len(self.skipped_results)}) saved to:\n{skipped_path}"
                )
            else:
                messagebox.showinfo("Export complete", f"Exported {len(self.scan_results):,} files to:\n{path}")
            self.status_var.set(f"Exported to {path}")
        except Exception as e:
            messagebox.showerror("Export failed", str(e))

    def _goto_review(self):
        if not self.scan_results:
            if not messagebox.askyesno(
                "No scan results",
                "You have no scan results. Switch to the Review tab anyway?\n"
                "(You can load a previously exported CSV there.)"
            ):
                return
        else:
            self._load_from_scan()
        self.nb.select(self.review_tab)

    # ----------------------- Review tab logic -----------------------

    def _load_from_scan(self):
        if not self.scan_results:
            messagebox.showinfo("No data", "No scan results available. Run a scan first.")
            return
        # Fresh load — start with everything checked. Clear search.
        self.checked_paths = {f["path"]: True for f in self.scan_results}
        if hasattr(self, "review_search_var"):
            self.review_search_var.set("")
        self.review_source_var.set(
            f"Loaded {len(self.scan_results):,} files from current scan."
        )
        self._rebuild_tree()
        self.nb.select(self.review_tab)

    def _load_from_csv(self):
        path = filedialog.askopenfilename(
            title="Load scan results CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            loaded = []
            with open(path, "r", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                for row in reader:
                    try:
                        mtime = datetime.strptime(row["LastModified"], "%Y-%m-%d %H:%M:%S")
                    except Exception:
                        mtime = datetime.now()
                    try:
                        size = int(row.get("SizeBytes", 0))
                    except Exception:
                        size = 0
                    try:
                        days_old = int(row.get("DaysOld", 0))
                    except Exception:
                        days_old = 0
                    loaded.append({
                        "path": row["FullPath"],
                        "name": row.get("FileName") or os.path.basename(row["FullPath"]),
                        "parent": row.get("ParentFolder") or os.path.dirname(row["FullPath"]),
                        "size": size,
                        "mtime": mtime,
                        "days_old": days_old,
                    })
            self.scan_results = loaded
            # Fresh load — start with everything checked. Clear search.
            self.checked_paths = {f["path"]: True for f in self.scan_results}
            if hasattr(self, "review_search_var"):
                self.review_search_var.set("")
            self.review_source_var.set(
                f"Loaded {len(loaded):,} files from CSV: {os.path.basename(path)}"
            )
            self._rebuild_tree()
        except Exception as e:
            messagebox.showerror("Load failed", f"Could not read CSV:\n{e}")

    def _rebuild_tree(self):
        for i in self.review_tree.get_children():
            self.review_tree.delete(i)
        self.tree_items.clear()

        # Initialize checkbox state for any newly-seen files (default: checked).
        # Existing entries in self.checked_paths are preserved across rebuilds.
        for f in self.scan_results:
            if f["path"] not in self.checked_paths:
                self.checked_paths[f["path"]] = True

        # Apply search filter (if any)
        search_q = (self.review_search_var.get() or "").strip().lower() if hasattr(self, "review_search_var") else ""
        if search_q:
            visible_files = [f for f in self.scan_results
                              if search_q in f["name"].lower() or search_q in f["path"].lower()]
        else:
            visible_files = list(self.scan_results)

        # Update search status label
        if hasattr(self, "review_search_status_var"):
            if search_q:
                self.review_search_status_var.set(
                    f"Showing {len(visible_files):,} of {len(self.scan_results):,} files matching '{search_q}'"
                )
            else:
                self.review_search_status_var.set("")

        # Group by parent folder (only visible files)
        folder_map = {}
        for f in visible_files:
            folder_map.setdefault(f["parent"], []).append(f)

        # Sort files within each folder according to user's choice
        sort_choice = self.review_sort_var.get()
        key_fn, reverse = self._sort_key_for_choice(sort_choice)
        for parent in folder_map:
            folder_map[parent].sort(key=key_fn, reverse=reverse)

        # Sort folders by total size desc (most impactful first) — always
        folder_totals = {p: sum(f["size"] for f in files) for p, files in folder_map.items()}
        sorted_folders = sorted(folder_map.keys(), key=lambda p: folder_totals[p], reverse=True)

        for parent in sorted_folders:
            files = folder_map[parent]
            total_size = folder_totals[parent]
            # Folder visual state: ☑ if all visible children checked,
            # ☐ if none, ◪ if some
            child_states = [self.checked_paths.get(f["path"], True) for f in files]
            if all(child_states):
                fglyph = "☑"
            elif not any(child_states):
                fglyph = "☐"
            else:
                fglyph = "◪"
            folder_label = f"{fglyph}  📁  {parent}    ({len(files):,} files — {format_size(total_size)})"
            folder_id = self.review_tree.insert(
                "", tk.END,
                text=folder_label,
                values=("", "", ""),
                open=bool(search_q),  # auto-expand when searching
            )
            self.tree_items[folder_id] = {
                "type": "folder",
                "path": parent,
                "size": total_size,
                "checked": any(child_states),
                "file_ref": None,
            }
            for fobj in files:
                checked = self.checked_paths.get(fobj["path"], True)
                glyph = "☑" if checked else "☐"
                file_label = f"{glyph}  📄  {fobj['name']}"
                file_id = self.review_tree.insert(
                    folder_id, tk.END,
                    text=file_label,
                    values=(
                        format_size(fobj["size"]),
                        fobj["mtime"].strftime("%Y-%m-%d %H:%M:%S"),
                        fobj["days_old"],
                    ),
                )
                self.tree_items[file_id] = {
                    "type": "file",
                    "path": fobj["path"],
                    "size": fobj["size"],
                    "checked": checked,
                    "file_ref": fobj,
                }
        self._update_review_summary()

    def _sort_key_for_choice(self, choice):
        mapping = {
            "size_desc": (lambda f: f["size"], True),
            "size_asc": (lambda f: f["size"], False),
            "days_desc": (lambda f: f["days_old"], True),
            "days_asc": (lambda f: f["days_old"], False),
            "name_asc": (lambda f: f["name"].lower(), False),
            "name_desc": (lambda f: f["name"].lower(), True),
            "mtime_desc": (lambda f: f["mtime"], True),
            "mtime_asc": (lambda f: f["mtime"], False),
        }
        return mapping.get(choice, (lambda f: f["size"], True))

    def _on_tree_click(self, event):
        # Only toggle when the user clicks in the "#0" cell area (where the checkbox glyph is)
        region = self.review_tree.identify("region", event.x, event.y)
        col = self.review_tree.identify_column(event.x)
        if col != "#0":
            return
        if region not in ("tree", "cell"):
            return
        item = self.review_tree.identify_row(event.y)
        if not item:
            return

        # Determine if click was on the checkbox glyph (first ~20 px of label area)
        # Using a simple heuristic: identify_element returns 'text' for label clicks.
        # We'll toggle on any click in the tree column for simplicity, EXCEPT when
        # clicking the expand indicator (element 'Treeitem.indicator').
        elem = self.review_tree.identify_element(event.x, event.y)
        if "indicator" in (elem or ""):
            return  # let the expand/collapse happen naturally

        self._toggle_item(item)
        return "break"  # prevent default selection behavior

    def _toggle_item(self, item_id):
        info = self.tree_items.get(item_id)
        if not info:
            return
        new_state = not info["checked"]
        info["checked"] = new_state

        if info["type"] == "file":
            # Update persistent state and refresh visuals
            self.checked_paths[info["path"]] = new_state
            self._refresh_item_label(item_id)
            # Update parent folder glyph based on its (now-changed) children
            parent = self.review_tree.parent(item_id)
            if parent:
                self._refresh_folder_visual(parent)

        elif info["type"] == "folder":
            # Bulk-toggle all currently-VISIBLE children of this folder.
            # Note: this only affects files visible under this folder right now
            # (which respects any active search filter — searched-out files are
            # untouched, intentionally).
            self._refresh_item_label(item_id)
            for child in self.review_tree.get_children(item_id):
                child_info = self.tree_items.get(child)
                if child_info and child_info["type"] == "file":
                    child_info["checked"] = new_state
                    self.checked_paths[child_info["path"]] = new_state
                    self._refresh_item_label(child)

        self._update_review_summary()

    def _refresh_item_label(self, item_id):
        info = self.tree_items[item_id]
        current = self.review_tree.item(item_id, "text")
        # Strip any existing checkbox glyph prefix we might have set
        stripped = current
        for glyph in ("☑  ", "☐  ", "◪  "):
            if stripped.startswith(glyph):
                stripped = stripped[len(glyph):]
                break
        new_prefix = "☑  " if info["checked"] else "☐  "
        self.review_tree.item(item_id, text=new_prefix + stripped)

    def _refresh_folder_visual(self, folder_id):
        """
        Update folder checkbox glyph based on its currently-visible children:
        all checked -> ☑, none -> ☐, some -> ◪ (partial)
        """
        info = self.tree_items.get(folder_id)
        if not info or info["type"] != "folder":
            return
        children = self.review_tree.get_children(folder_id)
        if not children:
            return
        states = [self.tree_items[c]["checked"] for c in children if c in self.tree_items]
        all_on = all(states)
        none_on = not any(states)

        # Folder's own "checked" indicator is true if at least one child is checked
        info["checked"] = any(states)

        current = self.review_tree.item(folder_id, "text")
        stripped = current
        for glyph in ("☑  ", "☐  ", "◪  "):
            if stripped.startswith(glyph):
                stripped = stripped[len(glyph):]
                break
        if all_on:
            prefix = "☑  "
        elif none_on:
            prefix = "☐  "
        else:
            prefix = "◪  "
        self.review_tree.item(folder_id, text=prefix + stripped)

    def _bulk_check(self, state):
        """Select All / Deselect All — affects EVERY file (visible or not),
        because that's what the user expects from a top-level button."""
        # Update persistent state for every known file
        for f in self.scan_results:
            self.checked_paths[f["path"]] = state
        # Update currently-visible tree items to reflect new state
        for item_id, info in self.tree_items.items():
            info["checked"] = state
            self._refresh_item_label(item_id)
        self._update_review_summary()

    def _bulk_check_visible(self, state):
        """Select / Deselect only the files currently VISIBLE in the tree
        (i.e. those matching the current search filter)."""
        affected = 0
        for item_id, info in self.tree_items.items():
            if info["type"] == "file":
                self.checked_paths[info["path"]] = state
                info["checked"] = state
                self._refresh_item_label(item_id)
                affected += 1
        # Update folder glyphs to reflect partial states
        for item_id, info in self.tree_items.items():
            if info["type"] == "folder":
                self._refresh_folder_visual(item_id)
        self._update_review_summary()
        verb = "selected" if state else "deselected (will NOT be deleted)"
        self.status_var.set(f"{affected:,} visible files {verb}.")

    def _on_review_search_change(self):
        """Called every time the user types in the search box. Rebuilds the
        tree to show only matching files. Checkbox state is preserved across
        rebuilds via self.checked_paths."""
        self._rebuild_tree()

    def _expand_all(self, expand):
        for item_id, info in self.tree_items.items():
            if info["type"] == "folder":
                self.review_tree.item(item_id, open=expand)

    def _get_selected_files(self):
        """Return list of file dicts whose checkbox is checked.
        IMPORTANT: this reads from self.checked_paths (the source of truth),
        NOT from currently-visible tree items. Hidden-by-search files that
        remain checked WILL be included."""
        selected = []
        for f in self.scan_results:
            if self.checked_paths.get(f["path"], True):
                selected.append(f)
        return selected

    def _update_review_summary(self):
        total_files = len(self.scan_results)
        total_size = sum(f["size"] for f in self.scan_results)
        sel_files = 0
        sel_size = 0
        for f in self.scan_results:
            if self.checked_paths.get(f["path"], True):
                sel_files += 1
                sel_size += f["size"]
        # If a search is active, also show how many of the visible items are selected
        visible_count = sum(1 for info in self.tree_items.values() if info["type"] == "file")
        visible_sel = sum(1 for info in self.tree_items.values()
                          if info["type"] == "file" and info["checked"])
        if visible_count != total_files:
            self.review_summary_var.set(
                f"Total: {total_files:,} files, {format_size(total_size)}   |   "
                f"Selected: {sel_files:,} files, {format_size(sel_size)}   |   "
                f"Visible: {visible_count:,} ({visible_sel:,} selected)"
            )
        else:
            self.review_summary_var.set(
                f"Total: {total_files:,} files, {format_size(total_size)}   |   "
                f"Selected: {sel_files:,} files, {format_size(sel_size)}"
            )

    # ----------------------- Delete actions -----------------------

    def _dry_run(self):
        sel = self._get_selected_files()
        if not sel:
            messagebox.showinfo("Nothing selected", "No files are selected for deletion.")
            return

        mode = "Recycle Bin" if self.delete_mode_var.get() == "recycle" else "Permanent Delete"
        total_size = sum(f["size"] for f in sel)
        preview_lines = [
            f"DRY RUN PREVIEW",
            f"Mode: {mode}",
            f"Files to be deleted: {len(sel):,}",
            f"Total size to be freed: {format_size(total_size)}",
            f"Remove empty folders after: {'Yes' if self.cleanup_empty_var.get() else 'No'}",
            "",
            "First 30 files that would be deleted:",
        ]
        for f in sel[:30]:
            preview_lines.append(f"  • {f['path']}  ({format_size(f['size'])})")
        if len(sel) > 30:
            preview_lines.append(f"  … and {len(sel) - 30:,} more.")
        preview_lines.append("")
        preview_lines.append("No files have been deleted.")

        self._show_scrollable_dialog("Dry Run Preview", "\n".join(preview_lines))

    def _start_delete(self):
        sel = self._get_selected_files()
        if not sel:
            messagebox.showinfo("Nothing selected", "No files are selected for deletion.")
            return

        # Block deletion if any selected file falls under a protected path
        for f in sel:
            blocked, reason = is_protected_path(f["path"])
            if blocked:
                messagebox.showerror(
                    "Protected file blocked",
                    f"Cannot delete: {f['path']}\n\n{reason}\n\nAborting entire deletion."
                )
                return

        # Apply user exclusions as a final safety net (e.g., if user loaded an
        # old CSV scanned without current exclusions). Excluded files are filtered
        # out silently with a notification.
        if self.exclusions:
            kept = []
            removed_by_excl = []
            for f in sel:
                hit, ex = file_is_excluded(f["path"], f["name"], self.exclusions)
                if hit:
                    removed_by_excl.append((f, ex))
                else:
                    kept.append(f)
            if removed_by_excl:
                if not messagebox.askyesno(
                    "Files matched your exclusions",
                    f"{len(removed_by_excl):,} of the {len(sel):,} selected files match your "
                    f"current exclusion rules and will be SKIPPED.\n\n"
                    f"Continue with deletion of the remaining {len(kept):,} files?"
                ):
                    return
                sel = kept
            if not sel:
                messagebox.showinfo(
                    "All selections excluded",
                    "All selected files matched your exclusion rules. Nothing to delete."
                )
                return

        mode_label = "Recycle Bin" if self.delete_mode_var.get() == "recycle" else "Permanent Delete"
        total_size = sum(f["size"] for f in sel)

        if self.delete_mode_var.get() == "permanent":
            confirm = messagebox.askyesno(
                "Confirm PERMANENT deletion",
                f"You are about to PERMANENTLY DELETE {len(sel):,} files "
                f"totaling {format_size(total_size)}.\n\n"
                f"This cannot be undone.\n\nContinue?",
                icon="warning",
            )
        else:
            confirm = messagebox.askyesno(
                "Confirm deletion",
                f"Send {len(sel):,} files ({format_size(total_size)}) to the Recycle Bin?",
            )
        if not confirm:
            return

        if self.delete_mode_var.get() == "recycle" and not SEND2TRASH_AVAILABLE:
            messagebox.showerror(
                "Recycle Bin unavailable",
                "The send2trash library is not available in this build. "
                "Please use Permanent Delete, or rebuild including send2trash."
            )
            return

        # Gather parent directories of selected files for empty-dir cleanup
        cleanup_roots = set()
        for f in sel:
            cleanup_roots.add(f["parent"])

        self.del_cancel.clear()
        self.delete_btn.configure(state=tk.DISABLED)
        self.del_cancel_btn.configure(state=tk.NORMAL)
        self.del_progress.configure(maximum=len(sel), value=0)
        self.status_var.set(f"Deleting {len(sel):,} files ({mode_label})…")

        paths = [f["path"] for f in sel]
        # Scan root is the path the user originally scanned. We use it to bound
        # the empty-folder cleanup so we never delete folders above what was
        # scanned. If the user loaded from CSV without scanning, we conservatively
        # don't enable ancestor cleanup (still cleans up immediate empty parents).
        scan_root = self.path_var.get().strip().strip('"') if self.path_var.get() else None
        if scan_root and not os.path.isdir(scan_root):
            scan_root = None

        worker = DeleteWorker(
            paths,
            use_recycle_bin=(self.delete_mode_var.get() == "recycle"),
            cleanup_empty_dirs=self.cleanup_empty_var.get(),
            cleanup_roots=cleanup_roots,
            out_queue=self.del_queue,
            cancel_event=self.del_cancel,
            scan_root=scan_root,
        )
        self.del_thread = threading.Thread(target=worker.run, daemon=True)
        self.del_thread.start()

    def _cancel_delete(self):
        self.del_cancel.set()
        self.status_var.set("Cancelling deletion…")

    def _poll_del_queue(self):
        try:
            while True:
                msg, data = self.del_queue.get_nowait()
                if msg == "del_progress":
                    self.del_progress.configure(value=data["done"])
                    self.status_var.set(
                        f"Deleting… {data['done']}/{data['total']} "
                        f"(OK: {data['deleted']}, Failed: {data['failed']})"
                    )
                elif msg == "del_done":
                    self.delete_btn.configure(state=tk.NORMAL)
                    self.del_cancel_btn.configure(state=tk.DISABLED)
                    self.del_progress.configure(value=0)
                    self._handle_delete_done(data)
        except queue.Empty:
            pass
        self.after(150, self._poll_del_queue)

    def _handle_delete_done(self, data):
        # Write log CSV
        log_path = LOG_DIR / f"deletion_log_{timestamp_for_filename()}.csv"
        try:
            ensure_dirs()
            with open(log_path, "w", newline="", encoding="utf-8") as f:
                w = csv.writer(f)
                w.writerow(["Path", "Status", "Reason", "SizeBytes"])
                for r in data["results"]:
                    w.writerow([r["path"], r["status"], r["reason"], r["size"]])
                if data["removed_dirs"]:
                    w.writerow([])
                    w.writerow(["--- Empty folders removed ---"])
                    for d in data["removed_dirs"]:
                        w.writerow([d, "DIR_REMOVED", "", ""])
                if data["failed_dirs"]:
                    w.writerow([])
                    w.writerow(["--- Empty-folder removals failed ---"])
                    for d in data["failed_dirs"]:
                        w.writerow([d.get("path", ""), "DIR_FAILED", d.get("reason", ""), ""])
        except Exception as e:
            log_path = None
            messagebox.showwarning("Log write failed", f"Could not write deletion log:\n{e}")

        verb = "Cancelled" if data.get("cancelled") else "Complete"
        summary = (
            f"Deletion {verb}.\n\n"
            f"Deleted: {data['deleted']:,}\n"
            f"Size freed: {format_size(data['deleted_bytes'])}\n"
            f"Failed: {data['failed']:,}\n"
            f"Empty folders removed: {len(data['removed_dirs'])}\n"
        )
        if log_path:
            summary += f"\nFull log saved to:\n{log_path}"

        self.status_var.set(
            f"Deletion {verb.lower()}: {data['deleted']:,} deleted, "
            f"{format_size(data['deleted_bytes'])} freed, {data['failed']:,} failed."
        )
        messagebox.showinfo("Deletion Summary", summary)

        # Remove successfully deleted files from the in-memory list and rebuild tree
        deleted_paths = {r["path"] for r in data["results"] if r["status"] == "DELETED"}
        self.scan_results = [f for f in self.scan_results if f["path"] not in deleted_paths]
        # Also clean up the checkbox-state dict
        for p in deleted_paths:
            self.checked_paths.pop(p, None)
        self._rebuild_tree()

    # ----------------------- Misc -----------------------

    def _show_scrollable_dialog(self, title, text):
        win = tk.Toplevel(self)
        win.title(title)
        win.geometry("720x480")
        frame = ttk.Frame(win, padding=8)
        frame.pack(fill=tk.BOTH, expand=True)
        txt = tk.Text(frame, wrap="none", font=("Consolas", 10))
        ysb = ttk.Scrollbar(frame, orient="vertical", command=txt.yview)
        xsb = ttk.Scrollbar(frame, orient="horizontal", command=txt.xview)
        txt.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        txt.grid(row=0, column=0, sticky="nsew")
        ysb.grid(row=0, column=1, sticky="ns")
        xsb.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        txt.insert("1.0", text)
        txt.configure(state="disabled")
        ttk.Button(win, text="Close", command=win.destroy).pack(pady=6)


# ---------------------------------------------------------------------------
# Modal dialogs for adding exclusions and creating schedules
# ---------------------------------------------------------------------------

class ExclusionDialog(tk.Toplevel):
    """Modal dialog to add or edit an exclusion entry."""

    def __init__(self, parent, title, current=None):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.resizable(False, False)
        self.result = None
        current = current or {}

        body = ttk.Frame(self, padding=12)
        body.pack(fill=tk.BOTH, expand=True)

        ttk.Label(body, text="Path (folder or file). Leave blank for pattern-only rule:").grid(row=0, column=0, columnspan=3, sticky="w")
        self.path_var = tk.StringVar(value=current.get("path", ""))
        ttk.Entry(body, textvariable=self.path_var, width=60).grid(row=1, column=0, columnspan=2, sticky="ew", pady=(2, 0))
        ttk.Button(body, text="Browse…", command=self._browse).grid(row=1, column=2, padx=(6, 0))

        ttk.Label(body, text="Filename pattern (e.g. *.pst, *important*). Leave blank for path-only rule:").grid(row=2, column=0, columnspan=3, sticky="w", pady=(10, 0))
        self.pattern_var = tk.StringVar(value=current.get("pattern", ""))
        ttk.Entry(body, textvariable=self.pattern_var, width=60).grid(row=3, column=0, columnspan=3, sticky="ew", pady=(2, 0))

        ttk.Label(body, text="Note (optional):").grid(row=4, column=0, columnspan=3, sticky="w", pady=(10, 0))
        self.note_var = tk.StringVar(value=current.get("note", ""))
        ttk.Entry(body, textvariable=self.note_var, width=60).grid(row=5, column=0, columnspan=3, sticky="ew", pady=(2, 0))

        info = ttk.Label(
            body,
            text=("If both Path and Pattern are provided, a file must match BOTH to be excluded.\n"
                  "If only one is provided, that one alone is enough to exclude a file."),
            foreground="#555", justify="left",
        )
        info.grid(row=6, column=0, columnspan=3, sticky="w", pady=(12, 0))

        btns = ttk.Frame(body)
        btns.grid(row=7, column=0, columnspan=3, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="OK", command=self._ok).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(btns, text="Cancel", command=self._cancel).pack(side=tk.RIGHT)

        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self._cancel())

        self.update_idletasks()
        # Center over parent
        try:
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            w, h = self.winfo_width(), self.winfo_height()
            self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")
        except Exception:
            pass

        self.grab_set()
        self.wait_window()

    def _browse(self):
        d = filedialog.askdirectory(title="Choose folder to exclude", parent=self)
        if d:
            self.path_var.set(os.path.normpath(d))

    def _ok(self):
        self.result = {
            "path": self.path_var.get().strip(),
            "pattern": self.pattern_var.get().strip(),
            "note": self.note_var.get().strip(),
        }
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


class ScheduleDialog(tk.Toplevel):
    """Modal dialog to create a new scheduled scan."""

    def __init__(self, parent, exclusions):
        super().__init__(parent)
        self.title("Create Scheduled Scan")
        self.transient(parent)
        self.resizable(False, False)
        self.result = None
        self._parent_exclusions = exclusions or []

        body = ttk.Frame(self, padding=12)
        body.pack(fill=tk.BOTH, expand=True)

        # Friendly name
        ttk.Label(body, text="Schedule name:").grid(row=0, column=0, sticky="w", pady=(0, 2))
        self.name_var = tk.StringVar(value=f"Scan_{datetime.now().strftime('%Y%m%d')}")
        ttk.Entry(body, textvariable=self.name_var, width=50).grid(row=0, column=1, columnspan=2, sticky="ew")

        # Path
        ttk.Label(body, text="Folder to scan:").grid(row=1, column=0, sticky="w", pady=(8, 2))
        self.path_var = tk.StringVar()
        ttk.Entry(body, textvariable=self.path_var, width=50).grid(row=1, column=1, sticky="ew")
        ttk.Button(body, text="Browse…", command=self._browse_scan).grid(row=1, column=2, padx=(6, 0))

        # Age
        ttk.Label(body, text="Age threshold (days):").grid(row=2, column=0, sticky="w", pady=(8, 2))
        self.age_var = tk.StringVar(value="30")
        ttk.Entry(body, textvariable=self.age_var, width=10).grid(row=2, column=1, sticky="w")

        # Export dir
        ttk.Label(body, text="Export folder:").grid(row=3, column=0, sticky="w", pady=(8, 2))
        default_export = str(USER_DOCS / "scheduled_exports")
        self.export_var = tk.StringVar(value=default_export)
        ttk.Entry(body, textvariable=self.export_var, width=50).grid(row=3, column=1, sticky="ew")
        ttk.Button(body, text="Browse…", command=self._browse_export).grid(row=3, column=2, padx=(6, 0))

        # Frequency
        ttk.Label(body, text="Frequency:").grid(row=4, column=0, sticky="w", pady=(8, 2))
        self.freq_var = tk.StringVar(value="weekly")
        freq_cb = ttk.Combobox(body, textvariable=self.freq_var, state="readonly",
                               values=["daily", "weekly", "monthly"], width=14)
        freq_cb.grid(row=4, column=1, sticky="w")
        freq_cb.bind("<<ComboboxSelected>>", lambda e: self._update_freq_widgets())

        # Day-of-week (weekly)
        self.dow_label = ttk.Label(body, text="Day of week:")
        self.dow_label.grid(row=5, column=0, sticky="w", pady=(8, 2))
        self.dow_var = tk.StringVar(value="Monday")
        self.dow_cb = ttk.Combobox(body, textvariable=self.dow_var, state="readonly",
                                   values=["Sunday", "Monday", "Tuesday", "Wednesday",
                                           "Thursday", "Friday", "Saturday"], width=14)
        self.dow_cb.grid(row=5, column=1, sticky="w")

        # Day-of-month (monthly)
        self.dom_label = ttk.Label(body, text="Day of month:")
        self.dom_var = tk.StringVar(value="1")
        self.dom_cb = ttk.Combobox(body, textvariable=self.dom_var, state="readonly",
                                   values=[str(i) for i in range(1, 29)], width=14)

        # Time
        ttk.Label(body, text="Time of day (HH:MM, 24-hour):").grid(row=6, column=0, sticky="w", pady=(8, 2))
        self.time_var = tk.StringVar(value="03:00")
        ttk.Entry(body, textvariable=self.time_var, width=10).grid(row=6, column=1, sticky="w")

        # Run-mode
        ttk.Label(body, text="Run mode:").grid(row=7, column=0, sticky="w", pady=(8, 2))
        self.runmode_var = tk.StringVar(value="logged_in")
        rm_frame = ttk.Frame(body)
        rm_frame.grid(row=7, column=1, columnspan=2, sticky="w")
        ttk.Radiobutton(rm_frame, text="Only when I'm logged in (no password needed)",
                        variable=self.runmode_var, value="logged_in",
                        command=self._update_pwd_state).pack(anchor="w")
        ttk.Radiobutton(rm_frame, text="Whether I'm logged in or not (requires my Windows password)",
                        variable=self.runmode_var, value="not_logged_in",
                        command=self._update_pwd_state).pack(anchor="w")

        # Password
        self.pwd_label = ttk.Label(body, text="Windows password:")
        self.pwd_label.grid(row=8, column=0, sticky="w", pady=(8, 2))
        self.pwd_var = tk.StringVar(value="")
        self.pwd_entry = ttk.Entry(body, textvariable=self.pwd_var, width=30, show="*")
        self.pwd_entry.grid(row=8, column=1, sticky="w")

        # Use exclusions checkbox
        self.use_excl_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(body, text=f"Apply current exclusion rules ({len(self._parent_exclusions)} entries)",
                        variable=self.use_excl_var).grid(row=9, column=0, columnspan=3, sticky="w", pady=(10, 0))

        # Buttons
        btns = ttk.Frame(body)
        btns.grid(row=10, column=0, columnspan=3, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Create Schedule", command=self._ok).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(btns, text="Cancel", command=self._cancel).pack(side=tk.RIGHT)

        body.columnconfigure(1, weight=1)

        self._update_freq_widgets()
        self._update_pwd_state()

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.bind("<Escape>", lambda e: self._cancel())

        self.update_idletasks()
        try:
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            w, h = self.winfo_width(), self.winfo_height()
            self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")
        except Exception:
            pass

        self.grab_set()
        self.wait_window()

    def _browse_scan(self):
        d = filedialog.askdirectory(title="Choose folder to scan", parent=self)
        if d:
            self.path_var.set(os.path.normpath(d))

    def _browse_export(self):
        d = filedialog.askdirectory(title="Choose export folder", parent=self)
        if d:
            self.export_var.set(os.path.normpath(d))

    def _update_freq_widgets(self):
        freq = self.freq_var.get()
        # Hide both, then show the relevant one
        self.dow_label.grid_remove()
        self.dow_cb.grid_remove()
        self.dom_label.grid_remove()
        self.dom_cb.grid_remove()
        if freq == "weekly":
            self.dow_label.grid(row=5, column=0, sticky="w", pady=(8, 2))
            self.dow_cb.grid(row=5, column=1, sticky="w")
        elif freq == "monthly":
            self.dom_label.grid(row=5, column=0, sticky="w", pady=(8, 2))
            self.dom_cb.grid(row=5, column=1, sticky="w")

    def _update_pwd_state(self):
        if self.runmode_var.get() == "not_logged_in":
            self.pwd_entry.configure(state="normal")
            self.pwd_label.configure(foreground="#000")
        else:
            self.pwd_entry.configure(state="disabled")
            self.pwd_label.configure(foreground="#999")
            self.pwd_var.set("")

    def _ok(self):
        # Validate
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror("Missing", "Schedule name is required.", parent=self)
            return
        scan_path = self.path_var.get().strip()
        if not scan_path or not os.path.isdir(scan_path):
            messagebox.showerror("Missing", "Please select a valid folder to scan.", parent=self)
            return
        blocked, reason = is_protected_path(scan_path)
        if blocked:
            messagebox.showerror("Protected path", reason, parent=self)
            return
        try:
            age = int(self.age_var.get().strip())
            if age < 0:
                raise ValueError()
        except Exception:
            messagebox.showerror("Invalid", "Age threshold must be a non-negative integer.", parent=self)
            return
        export_dir = self.export_var.get().strip()
        if not export_dir:
            messagebox.showerror("Missing", "Export folder is required.", parent=self)
            return
        try:
            Path(export_dir).mkdir(parents=True, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Cannot create export folder", str(e), parent=self)
            return
        time_s = self.time_var.get().strip()
        try:
            hh, mm = time_s.split(":")
            hh, mm = int(hh), int(mm)
            if not (0 <= hh < 24 and 0 <= mm < 60):
                raise ValueError()
        except Exception:
            messagebox.showerror("Invalid time", "Time must be in HH:MM format (24-hour).", parent=self)
            return

        run_when_not_logged_in = (self.runmode_var.get() == "not_logged_in")
        password = None
        if run_when_not_logged_in:
            password = self.pwd_var.get()
            if not password:
                messagebox.showerror("Password required",
                                     "Enter your Windows password (it is stored securely by Windows, "
                                     "never by this utility).", parent=self)
                return

        dow_map = {"Sunday": 1, "Monday": 2, "Tuesday": 3, "Wednesday": 4,
                   "Thursday": 5, "Friday": 6, "Saturday": 7}

        self.result = {
            "name": name,
            "scan_path": scan_path,
            "age_days": age,
            "export_dir": export_dir,
            "exclusions": list(self._parent_exclusions) if self.use_excl_var.get() else [],
            "frequency": self.freq_var.get(),
            "time": f"{hh:02d}:{mm:02d}",
            "day_of_week": dow_map.get(self.dow_var.get(), 2),
            "day_of_month": int(self.dom_var.get() or 1),
            "run_when_not_logged_in": run_when_not_logged_in,
            "password": password,
        }
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def parse_args(argv):
    p = argparse.ArgumentParser(description=APP_NAME, add_help=True)
    p.add_argument("--headless-from", metavar="JSON_FILE",
                   help="Run a scheduled scan defined in the given JSON sidecar (no GUI).")
    return p.parse_known_args(argv)[0]


def main():
    args = parse_args(sys.argv[1:])

    if args.headless_from:
        # Headless mode triggered by Task Scheduler — no GUI, just scan + export.
        sys.exit(run_headless(args.headless_from))

    app = StorageCleanupApp()
    app.mainloop()


if __name__ == "__main__":
    main()
