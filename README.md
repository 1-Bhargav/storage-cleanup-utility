# Storage Cleanup Utility

A portable Windows utility to scan attached or shared storage and delete files older than a configurable age, with a safe, selective workflow.

## Features

- **Two-mode workflow:** Scan & Export, then Review & Delete
- **Minimal GUI** built with Tkinter — no install, no dependencies once built
- **Configurable age threshold** (default: 30 days, based on Last Modified date)
- **Recursive scanning** of any folder, including attached and shared (UNC) storage
- **Protected system paths** — scans and deletes are blocked for Windows, Program Files, ProgramData, AppData, etc.
- **CSV export** of scan results (and a separate CSV of skipped/inaccessible items)
- **CSV re-import** — scan today, delete later
- **Expandable folder tree** with independent file-level checkboxes
- **Live summary** — total files, total size, and selected counts update as you check/uncheck
- **Sort by size / name / date / age**
- **Filter box** to find specific files quickly
- **Dry-run preview** before any deletion
- **Recycle Bin or Permanent Delete** — your choice at runtime
- **Empty-folder cleanup** after deletion (optional, on by default)
- **Deletion logs** saved to `%USERPROFILE%\Documents\StorageCleanupUtility\logs\`
- **Single portable 64-bit .exe** — copy anywhere, no installer

## How to Build the .exe

See **[BUILD_INSTRUCTIONS.md](BUILD_INSTRUCTIONS.md)** for the step-by-step GitHub Actions walkthrough (recommended — no local setup needed).

If you later have access to a Windows machine with Python, you can also just double-click `build.bat`.

## How to Use

1. **Launch** `Storage_Cleanup_Utility.exe` (double-click).
2. **Scan & Export tab:**
   - Click **Browse…** to pick a folder, or paste a path (UNC paths like `\\server\share\folder` work too).
   - Enter how many days old files must be (default: 30).
   - Click **Start Scan**.
   - Use the filter box or click column headers to sort.
   - Click **Export to CSV…** to save the list for later review.
3. **Review & Delete tab:**
   - Click **Use Current Scan Results** (or **Load from CSV…** to reuse an earlier scan).
   - Expand folders, uncheck anything you want to keep.
     - Unchecking a folder unchecks all files inside it (bulk helper).
     - You can then individually re-check specific files to keep them in the delete list even if their folder is unchecked.
     - **Rule: only files whose own checkbox is checked get deleted. Everything else is left untouched.**
   - Choose **Recycle Bin** (safer) or **Permanent Delete**.
   - Click **Dry Run (Preview)** first to see exactly what will happen.
   - Click **DELETE Selected**. Confirm the warning. Done.

## Protected Paths

These locations are blocked — the utility will refuse to scan or delete here:

- `C:\` (system drive root)
- `C:\Windows`, `C:\Program Files`, `C:\Program Files (x86)`, `C:\ProgramData`
- `C:\$Recycle.Bin`, `C:\System Volume Information`, `C:\Recovery`, `C:\Boot`, `C:\PerfLogs`
- Any path containing `\AppData\Roaming`, `\AppData\Local`, or `\AppData\LocalLow`

You can scan anywhere else, including:
- Specific subfolders on C: that aren't in the blocked list
- Non-system drive roots (`D:\`, `E:\`, etc.) and any subfolders
- Shared / network storage via UNC paths (`\\server\share\…`)

## File Layout

```
Storage_Cleanup_Utility.exe         <- the only file you need to ship
%USERPROFILE%\Documents\
    StorageCleanupUtility\
        state.json                   <- remembers your last path & age
        logs\
            deletion_log_*.csv       <- created after each deletion run
```

## Requirements at Runtime

- Windows 10 or 11, 64-bit
- Nothing else — Python, libraries, and all dependencies are bundled into the .exe

## License

Use freely within your organization.
