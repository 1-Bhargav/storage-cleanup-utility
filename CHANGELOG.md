# Changelog

## Version 1.1.1 — Review-tab search

### New
- **Search box in Review & Delete tab.** Type any text to filter the tree to files whose name or path matches. Auto-expands matching folders so you can see results immediately.
- **"Deselect (keep these)" button** — when you've filtered to a subset (e.g. searched for "important"), one click unchecks every visible file so they're protected from deletion. Then clear the search and continue.
- **"Select" button** — opposite action: select every visible file at once.
- **Visible/Selected counters in the summary** — when a search is active, see total + selected + visible (selected) counts.
- Checkbox state is now **path-keyed and persistent across filter changes**. Filtering does not lose your selections; unchecked files stay unchecked when you clear the search.

### Behavior notes
- Top-level "Select All" / "Deselect All" buttons affect every file (visible or not), as expected.
- The "On visible:" buttons (next to the search box) only affect currently-visible (filtered) files.
- Hidden-but-checked files **will still be deleted** when you click DELETE Selected. The summary always shows "Selected: N" reflecting all checked files regardless of visibility.

## Version 1.1.0 — Exclusions & Schedules

### New features
- User exclusion rules (path / pattern / both), persistent in `state.json`, importable/exportable as CSV.
- Scheduled scans via Windows Task Scheduler (daily / weekly / monthly).
- Headless `--headless-from <json>` CLI mode for scheduled runs.
- Three-tab UI (added "Settings & Schedules").
- Smarter empty-folder cleanup (walks ancestors, bounded by scan root).

### Internal
- New build dependency: `pywin32==308`.
- New helpers: `file_is_excluded`, `dir_is_inside_excluded_path`, `validate_exclusion_entry`.
- New module: Task Scheduler create/list/delete/run-now/enable functions.

## Version 1.0.0 — Initial release

- Two-tab GUI (Scan & Export, Review & Delete).
- Configurable age threshold.
- Recursive scan, including UNC paths.
- Hard-coded protected system paths.
- CSV export & re-import.
- Expandable tree with checkboxes (Option-B folder bulk-toggle).
- Dry-run preview.
- Recycle Bin or Permanent Delete.
- Empty-folder cleanup.
- Deletion logs in `Documents\StorageCleanupUtility\logs\`.
- Single-file portable .exe via PyInstaller.
