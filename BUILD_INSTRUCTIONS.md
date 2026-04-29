# Build Instructions — Using GitHub Actions

This guide walks you through building `Storage_Cleanup_Utility.exe` using GitHub's free Windows build servers. **No software needs to be installed on your computer.** You'll only need a web browser.

**Total time:** ~10 minutes for the one-time setup. After that, every rebuild is automatic.

---

## If you already built v1.0 and want to update to v1.1

If you already have a GitHub repository from the previous build:

1. Go to your repository on GitHub.
2. For each of these files in your repo, click the file → click the pencil (✏️) icon → replace the entire content with the new version from the v1.1 source:
   - `storage_cleanup_utility.py`
   - `.github/workflows/build.yml`
   - `requirements.txt`
   - `build.bat`
   - `README.md`
   - `BUILD_INSTRUCTIONS.md`
3. After saving each file, the build runs automatically. Check the **Actions** tab for the green checkmark, then download the new artifact.

A faster alternative: delete the old files in GitHub, then upload the new ones (drag-and-drop). The build triggers automatically on commit.

---

## Fresh setup (from scratch)

### Step 1 — Create a free GitHub account

1. Go to **https://github.com/signup**
2. Enter your email, pick a username and password.
3. Verify your email.

Skip if you already have one.

### Step 2 — Create a new repository

1. Log in to GitHub.
2. Click the **"+"** icon in the top-right corner → **"New repository"**.
3. Fill in:
   - **Repository name:** `storage-cleanup-utility` (or any name)
   - **Visibility:** **Private** is recommended.
   - Do NOT check "Add a README file" — you'll upload your own.
4. Click **"Create repository"**.

### Step 3 — Upload the project files

You have two reliable options. **Option A is recommended** because the `.github` folder is "hidden" on most systems and may not upload correctly via drag-and-drop.

#### Option A — Upload normal files first, then create the workflow file directly in GitHub

1. On the new repository page, click **"Add file"** → **"Upload files"**.
2. Drag and drop these files (and only these — skip the `.github` folder for now):
   - `storage_cleanup_utility.py`
   - `requirements.txt`
   - `build.bat`
   - `README.md`
   - `BUILD_INSTRUCTIONS.md`
   - `.gitignore` *(optional — if it doesn't show up due to "hidden file" issues, skip it; nothing breaks)*
3. Scroll down, type a commit message like `Initial upload`, click **"Commit changes"**.
4. Now create the workflow file directly in GitHub:
   - Click **"Add file"** → **"Create new file"**.
   - In the filename box at the top, type **exactly** (with the slashes): `.github/workflows/build.yml`
   - Each `/` automatically creates a folder. So `.github` becomes a folder, `workflows` becomes a subfolder, and `build.yml` is the file.
   - In the large content box, paste the **entire content of build.yml** from the source zip.
   - Scroll down, commit message like `Add build workflow`, click **"Commit changes"**.

#### Option B — Drag-and-drop with hidden files visible

1. **On Windows:** Open File Explorer → View tab → check "Hidden items".
   **On Mac:** Finder → press `Cmd + Shift + .` (period) to show hidden files.
2. Drag and drop everything (including the `.github` folder) onto the GitHub upload page.
3. Confirm `.github/workflows/build.yml` is listed before committing.
4. Commit.

If the `.github` folder isn't visible after dragging, fall back to **Option A** above.

### Step 4 — Watch the build run

1. Click the **"Actions"** tab.
2. The latest workflow run will appear. Click it to see live progress.
3. Building takes about **3–5 minutes** (slightly longer than v1.0 because pywin32 is added).
4. When finished, the dot turns into a **green checkmark** ✅.

### Step 5 — Download your .exe

1. On the **Actions** tab, click the successful run.
2. Scroll to the bottom of the page.
3. Under **"Artifacts"** click **`Storage_Cleanup_Utility`** to download a ZIP.
4. Extract the ZIP — inside is **`Storage_Cleanup_Utility.exe`**.

That's your finished utility. The .exe is ~30–35 MB (larger than v1.0 because pywin32 is bundled for Task Scheduler integration).

---

## Rebuilding later

**To update the code:** edit any file in GitHub → commit → build runs automatically → download new artifact.

**To rebuild without code changes:** Go to **Actions → "Build Storage Cleanup Utility"** (left sidebar) → **Run workflow** → **Run workflow** (green button).

---

## Troubleshooting

**"Artifact retention expired"** — GitHub deletes artifacts after 30 days on free accounts. Trigger a rebuild and a fresh artifact appears.

**"Workflow file not found"** — The `.github/workflows/build.yml` file must be at exactly that path. If you accidentally created it at the wrong path, delete and recreate using Option A above.

**Build fails at "Install build dependencies"** — Usually a temporary GitHub or PyPI network glitch. Wait 5 minutes and click "Re-run all jobs" on the failed run.

**Windows Defender flags the .exe** — Known false positive with PyInstaller-built executables that include pywin32. Click **"More info"** → **"Run anyway"**, or add an exclusion in Windows Security.

**Task Scheduler tab says "not available"** — This means pywin32 didn't get bundled. Check the Actions log to confirm the install step succeeded; if it didn't, re-run the workflow.

**Scheduled tasks don't run when machine sleeps** — Task Scheduler honors Windows sleep/hibernate settings. The schedule has `StartWhenAvailable = True` so missed runs catch up after wake. For 24/7 reliability, configure your machine to stay awake at the scheduled time.

**"Run when not logged in" needs a password** — This is by design. Windows requires the password to be cached securely so the task can run without an interactive session. The password is given directly to Task Scheduler — this utility never stores it.

---

## Privacy note

If your repository is **Private**, only you can see the code and the built .exe. GitHub's build servers run in isolated containers and are destroyed after each build. Your built .exe is only accessible to you (or anyone you give repo access to).
