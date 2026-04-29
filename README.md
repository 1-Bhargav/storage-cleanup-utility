# Storage Cleanup Utility

A portable Windows utility to scan attached or shared storage and delete files older than a configurable age, with a safe, selective workflow. Now with user-defined exclusions and scheduled scans via Windows Task Scheduler.

**Current version: 1.1.0**

## Features

### Core (v1.0)
- **Three-tab workflow:** Scan & Export, Review & Delete, Settings & Schedules
- **Minimal GUI** built with Tkinter — no install, no dependencies once built
- **Configurable age threshold** (default: 30 days, based on Last Modified date)
- **Recursive scanning** of any folder, including attached and shared (UNC) storage
- **Protected system paths** — scans and deletes are blocked for Windows, Program Files, ProgramData, AppData, etc.
- **CSV export** of scan results (and a separate CSV of skipped items)
- **CSV re-import** — scan today, delete later
- **Expandable folder tree** with independent file-level checkboxes
- **Live summary** — total files, total size, selected counts, excluded counts
- **Sort by size / name / date / age**
- **Filter box** to find specific files quickly
- **Dry-run preview** before any deletion
- **Recycle Bin or Permanent Delete** — chosen at runtime
- **Empty-folder cleanup** after deletion (optional, on by default)
- **Deletion logs** saved to `%USERPROFILE%\Documents\StorageCleanupUtility\logs\`
- **Single portable 64-bit .exe** — copy anywhere, no installer

### New in v1.1

- **User exclusion rules** — define paths, filename patterns, or both that scanning and deletion will skip. Multiple entries supported. Persistent across sessions. Import/Export as CSV.
- **Scheduled scans** — create Windows Task Scheduler entries from inside the app for automated daily/weekly/monthly scan + export. Fully managed: list, run-now, enable/disable, delete.
- **Headless mode** — the .exe accepts a `--headless-from <json>` flag used internally by scheduled tasks; runs scan + export without showing the GUI.

## How to Build the .exe

See **[BUILD_INSTRUCTIONS.md](BUILD_INSTRUCTIONS.md)** for the step-by-step GitHub Actions walkthrough (recommended — no local setup needed).

If you have a Windows machine with Python, you can also just double-click `build.bat`.

## How to Use

### Tab 1 — Scan & Export

1. **Browse…** to pick a folder, or paste a path. UNC paths (`\\server\share\folder`) are supported.
2. Enter how many days old files must be (default: 30).
3. Click **Start Scan**.
4. Use the filter box or click column headers to sort by name, size, date, or age.
5. Click **Export to CSV…** to save the list. A second CSV with skipped items is saved alongside.

### Tab 2 — Review & Delete

1. Click **Use Current Scan Results** (or **Load from CSV…** to use an earlier scan).
2. Expand folders, uncheck anything you want to keep.
   - Unchecking a folder unchecks all files inside it (bulk helper).
   - You can re-check specific files inside an unchecked folder to keep them in the delete list.
   - **Rule: only files whose own checkbox is checked get deleted.**
3. Choose **Recycle Bin** (safer) or **Permanent Delete**.
4. Click **Dry Run (Preview)** to see exactly what would happen — no deletion.
5. Click **DELETE Selected**, confirm the warning, and the deletion runs.
6. Empty folders left behind are cleaned up automatically (if checkbox enabled).

### Tab 3 — Settings & Schedules

#### Exclusion rules (left side)

A file is skipped during scan and delete if it matches any rule.

- **Path-only rule:** any file under the given folder is excluded. Useful for entire archives like `D:\Backups\Critical`.
- **Pattern-only rule:** any file with a matching name is excluded anywhere. Examples: `*.pst`, `*important*`, `archive_2023_??.zip`.
- **Path + pattern:** both must match — e.g., only `.pst` files inside `D:\Mail`.

Buttons: **Add / Edit / Remove**, plus **Import CSV / Export CSV** for sharing exclusion lists across machines or team members.

Rules persist across app sessions (stored in `state.json`).

#### Scheduled scans (right side)

Schedules use **Windows Task Scheduler** under the hood. They run only the scan + export step — never deletion. The exported CSV lands in your chosen export folder, ready for you to review later.

Click **New Schedule…** to create one:

- **Frequency:** daily, weekly (with day-of-week), or monthly (with day-of-month).
- **Time:** 24-hour HH:MM.
- **Run mode:**
  - *Only when I'm logged in* — simple, no password needed, runs only on a logged-in interactive session.
  - *Whether I'm logged in or not* — survives logoffs and restarts, requires your Windows password (handed directly to Task Scheduler; never stored by this utility).
- **Apply current exclusion rules** — if checked, the scheduled run uses the same exclusions visible in the GUI.

After creation, the schedule appears in the list. Buttons:
- **Run Now** — manually trigger the schedule for testing.
- **Enable / Disable** — toggle without deleting.
- **Delete** — permanently remove the scheduled task.
- **Refresh** — re-read the list from Task Scheduler.

Each scheduled run produces:
- A CSV export in your chosen folder (`scheduled_<name>_<timestamp>.csv`)
- A log file in `%USERPROFILE%\Documents\StorageCleanupUtility\logs\scheduled_run_<timestamp>.log`

You can also see and edit these tasks in the Windows Task Scheduler GUI under the folder `\StorageCleanupUtility`.

## Protected Paths

These locations are blocked — the utility refuses to scan or delete here:

- `C:\` (system drive root)
- `C:\Windows`, `C:\Program Files`, `C:\Program Files (x86)`, `C:\ProgramData`
- `C:\$Recycle.Bin`, `C:\System Volume Information`, `C:\Recovery`, `C:\Boot`, `C:\PerfLogs`
- Any path containing `\AppData\Roaming`, `\AppData\Local`, or `\AppData\LocalLow`

You can scan anywhere else, including non-system drive roots and shared/UNC storage.

## File Layout

```
Storage_Cleanup_Utility.exe                <- the executable you ship

%USERPROFILE%\Documents\
    StorageCleanupUtility\
        state.json                          <- last path, age, exclusion rules
        logs\
            deletion_log_*.csv              <- after each deletion run
            scheduled_run_*.log             <- after each scheduled scan
        schedules\
            SCU_<name>.json                 <- sidecar for each schedule
        scheduled_exports\                  <- default location for scheduled CSVs
            scheduled_<name>_<timestamp>.csv
```

## Requirements at Runtime

- Windows 10 or 11, 64-bit
- For scheduled scans: standard Task Scheduler service running (default on Windows)
- Nothing else — Python, libraries, send2trash, pywin32 are all bundled into the .exe

## License

Use freely within your organization.
