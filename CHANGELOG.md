# Changelog

## Version 1.1.0 — Exclusions & Schedules

### New features

**User exclusion rules**
- Add custom paths and/or filename patterns that scanning and deletion will skip.
- Three rule types supported:
  - **Path-only:** any file under the given folder is skipped (entire subtrees are pruned during scan for efficiency).
  - **Pattern-only:** any file matching a wildcard pattern (e.g. `*.pst`, `*important*`) is skipped anywhere.
  - **Path + pattern:** both must match (e.g. only `.pst` files inside `D:\Mail`).
- Multiple entries supported.
- Exclusions persist across app sessions in `state.json`.
- Import / export rules as CSV for sharing.
- Each rule can carry an optional **note** (for "why" documentation).
- Excluded counts shown in scan summary and status bar.
- Exclusions enforced **at scan time AND at delete time** — even loading a stale CSV won't bypass them.

**Scheduled scans (Windows Task Scheduler integration)**
- Create automated scans that run on a schedule, with CSV export.
- Frequency: **daily**, **weekly** (with day-of-week), or **monthly** (with day-of-month).
- 24-hour time picker.
- Two run modes:
  - "Only when I'm logged in" — no password needed.
  - "Whether I'm logged in or not" — survives restarts/logoffs, requires Windows password (handed directly to Task Scheduler, never stored by this utility).
- **Scheduled tasks only do scan + export** — never deletion. Deletion remains a deliberate human action.
- Manage schedules from the new **Settings & Schedules** tab:
  - List all schedules with state, next run, last run.
  - Run-Now button for testing.
  - Enable/Disable without deleting.
  - Delete.
- Optional: apply current exclusion rules to scheduled runs.
- Each scheduled run produces a timestamped CSV in your chosen export folder, and a log entry in `Documents\StorageCleanupUtility\logs\`.
- Tasks are visible in Windows Task Scheduler under folder `\StorageCleanupUtility`.

### Improvements

- **Empty folder cleanup** is now smarter. Previously only the immediate parent of deleted files was checked. Now, after deletion:
  - Empty subdirectories are removed (as before).
  - Empty *ancestor* folders are also removed, **bounded by the original scan root** (the scan root itself is never removed).
  - Sibling folders with files are always preserved.

- **Three-tab UI:** Scan & Export | Review & Delete | Settings & Schedules. Default window size enlarged to fit the new tab.

- Scan summary now includes excluded count.

### Internal

- New `--headless-from <json>` command-line flag used by scheduled tasks to run scan + export without showing the GUI.
- New module functions for Task Scheduler: `schedule_create`, `schedule_list`, `schedule_delete`, `schedule_run_now`, `schedule_set_enabled`.
- New helpers: `file_is_excluded`, `dir_is_inside_excluded_path`, `validate_exclusion_entry`.
- New build dependency: `pywin32==308` for Task Scheduler COM access.

### Compatibility

- The `state.json` from v1.0 is forward-compatible. Exclusions just won't be present until you add them.
- CSVs from v1.0 still import correctly into v1.1.

### Build

- The .exe grew from ~20 MB (v1.0) to ~30–35 MB (v1.1) due to bundling `pywin32`.
- GitHub Actions build time increased by ~1 minute (still under 5 minutes total).
