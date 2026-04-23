"""
Storage Cleanup Utility
------------------------
A portable Windows utility to scan and delete files older than a configurable
age threshold. Includes protected system-path checks, dry-run, recycle-bin
option, and independent file-level selection.

Author: Built with Claude (Anthropic)
"""

import os
import sys
import csv
import json
import shutil
import threading
import queue
from datetime import datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from send2trash import send2trash
    SEND2TRASH_AVAILABLE = True
except ImportError:
    SEND2TRASH_AVAILABLE = False


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

APP_NAME = "Storage Cleanup Utility"
APP_VERSION = "1.0.0"

# Log and state directory in the user's Documents folder
USER_DOCS = Path(os.path.expanduser("~")) / "Documents" / "StorageCleanupUtility"
LOG_DIR = USER_DOCS / "logs"
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
    """Create log and state directories if they don't exist."""
    LOG_DIR.mkdir(parents=True, exist_ok=True)


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
      - progress updates
    Uses a queue to communicate with the GUI thread.
    """

    def __init__(self, root_path, age_days, out_queue, cancel_event):
        self.root_path = root_path
        self.age_days = age_days
        self.queue = out_queue
        self.cancel_event = cancel_event

    def run(self):
        try:
            cutoff = datetime.now() - timedelta(days=self.age_days)
            cutoff_ts = cutoff.timestamp()

            files_found = []
            skipped = []
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

                for name in filenames:
                    if self.cancel_event.is_set():
                        break
                    files_checked += 1
                    full = os.path.join(dirpath, name)
                    display_full = os.path.join(display_dir, name)
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
                            "current": display_dir,
                        }))

            if self.cancel_event.is_set():
                self.queue.put(("cancelled", {
                    "checked": files_checked,
                    "matched": len(files_found),
                    "skipped": len(skipped),
                    "files": files_found,
                    "skipped_list": skipped,
                }))
            else:
                self.queue.put(("done", {
                    "checked": files_checked,
                    "matched": len(files_found),
                    "skipped": len(skipped),
                    "files": files_found,
                    "skipped_list": skipped,
                }))

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


# ---------------------------------------------------------------------------
# Deletion worker
# ---------------------------------------------------------------------------

class DeleteWorker:
    """
    Deletes the given list of file paths, either to Recycle Bin or permanently.
    Reports progress and results via queue.
    """

    def __init__(self, file_paths, use_recycle_bin, cleanup_empty_dirs,
                 cleanup_roots, out_queue, cancel_event):
        self.file_paths = file_paths
        self.use_recycle_bin = use_recycle_bin
        self.cleanup_empty_dirs = cleanup_empty_dirs
        self.cleanup_roots = cleanup_roots  # set of parent dirs to check for emptiness
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
        """Bottom-up: remove dirpath and ancestors if they become empty."""
        try:
            long_dir = to_long_path(dirpath)
            if not os.path.isdir(long_dir):
                return
            # Walk bottom-up inside this dir
            for root, dirs, files in os.walk(long_dir, topdown=False):
                if files:
                    continue
                # Check each subdirectory
                try:
                    if not os.listdir(root):
                        display = root
                        if display.startswith("\\\\?\\UNC\\"):
                            display = "\\\\" + display[len("\\\\?\\UNC\\"):]
                        elif display.startswith("\\\\?\\"):
                            display = display[len("\\\\?\\"):]
                        # Don't remove the scan root itself — only its subfolders
                        # (safer default)
                        os.rmdir(root)
                        removed.append(display)
                except Exception as e:
                    failed.append({"path": root, "reason": str(e)})
        except Exception as e:
            failed.append({"path": dirpath, "reason": str(e)})


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------

class StorageCleanupApp(tk.Tk):

    def __init__(self):
        super().__init__()
        ensure_dirs()

        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry("1100x720")
        self.minsize(900, 600)

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

        self.settings = load_state()

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
        self.nb.add(self.scan_tab, text="  1.  Scan & Export  ")
        self.nb.add(self.review_tab, text="  2.  Review & Delete  ")

        self._build_scan_tab()
        self._build_review_tab()

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

        self.scan_summary_var = tk.StringVar(value="Total files: 0   |   Total size: 0 B   |   Skipped: 0")
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
        self._update_scan_summary(0, 0, 0)

        worker = ScanWorker(path, age, self.scan_queue, self.scan_cancel)
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
                        f"Scanning… checked {data['checked']:,} | matched {data['matched']:,} | skipped {data['skipped']:,}"
                    )
                elif msg == "walk_error":
                    self.skipped_results.append({"path": data["path"], "reason": data["reason"]})
                elif msg in ("done", "cancelled"):
                    self.scan_progress.stop()
                    self.scan_btn.configure(state=tk.NORMAL)
                    self.scan_cancel_btn.configure(state=tk.DISABLED)
                    self.scan_results = data["files"]
                    self.skipped_results.extend(data["skipped_list"])
                    self._populate_scan_tree(self.scan_results)
                    total_size = sum(f["size"] for f in self.scan_results)
                    self._update_scan_summary(len(self.scan_results), total_size, len(self.skipped_results))
                    verb = "Cancelled" if msg == "cancelled" else "Scan complete"
                    self.status_var.set(
                        f"{verb}. {len(self.scan_results):,} files matched, {format_size(total_size)} total."
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

    def _update_scan_summary(self, count, total_size, skipped_count):
        self.scan_summary_var.set(
            f"Total files: {count:,}   |   Total size: {format_size(total_size)}   |   Skipped: {skipped_count:,}"
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

        # Group by parent folder
        folder_map = {}  # parent -> list of file dicts
        for f in self.scan_results:
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
            folder_label = f"☑  📁  {parent}    ({len(files):,} files — {format_size(total_size)})"
            folder_id = self.review_tree.insert(
                "", tk.END,
                text=folder_label,
                values=("", "", ""),
                open=False,
            )
            self.tree_items[folder_id] = {
                "type": "folder",
                "path": parent,
                "size": total_size,
                "checked": True,
                "file_ref": None,
            }
            for fobj in files:
                file_label = f"☑  📄  {fobj['name']}"
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
                    "checked": True,
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
        self._refresh_item_label(item_id)

        # If this is a folder, bulk-toggle all children to the same state (Option B)
        if info["type"] == "folder":
            for child in self.review_tree.get_children(item_id):
                child_info = self.tree_items.get(child)
                if child_info:
                    child_info["checked"] = new_state
                    self._refresh_item_label(child)

        # If this is a file, update parent folder's visual state based on children
        if info["type"] == "file":
            parent = self.review_tree.parent(item_id)
            if parent:
                self._refresh_folder_visual(parent)

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
        Update folder checkbox glyph based on children state:
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

        # Folder's own "checked" is true if at least one child is checked
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
        for item_id, info in self.tree_items.items():
            info["checked"] = state
            self._refresh_item_label(item_id)
        self._update_review_summary()

    def _expand_all(self, expand):
        for item_id, info in self.tree_items.items():
            if info["type"] == "folder":
                self.review_tree.item(item_id, open=expand)

    def _get_selected_files(self):
        """Return list of file dicts whose file-level checkbox is checked."""
        selected = []
        for item_id, info in self.tree_items.items():
            if info["type"] == "file" and info["checked"]:
                selected.append(info["file_ref"])
        return selected

    def _update_review_summary(self):
        total_files = 0
        total_size = 0
        sel_files = 0
        sel_size = 0
        for info in self.tree_items.values():
            if info["type"] == "file":
                total_files += 1
                total_size += info["size"]
                if info["checked"]:
                    sel_files += 1
                    sel_size += info["size"]
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
        worker = DeleteWorker(
            paths,
            use_recycle_bin=(self.delete_mode_var.get() == "recycle"),
            cleanup_empty_dirs=self.cleanup_empty_var.get(),
            cleanup_roots=cleanup_roots,
            out_queue=self.del_queue,
            cancel_event=self.del_cancel,
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
# Entry point
# ---------------------------------------------------------------------------

def main():
    app = StorageCleanupApp()
    app.mainloop()


if __name__ == "__main__":
    main()
