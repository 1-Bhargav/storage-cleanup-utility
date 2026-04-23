# Build Instructions — Using GitHub Actions

This guide walks you through building `Storage_Cleanup_Utility.exe` using GitHub's free Windows build servers. **No software needs to be installed on your computer.** You'll only need a web browser.

**Total time:** ~10 minutes for the one-time setup. After that, every rebuild is automatic.

---

## Step 1 — Create a free GitHub account

1. Go to **https://github.com/signup**
2. Enter your email, pick a username and password.
3. Verify your email.

**Skip this step if you already have a GitHub account.**

---

## Step 2 — Create a new repository

1. Log in to GitHub.
2. Click the **"+"** icon in the top-right corner → **"New repository"**.
3. Fill in:
   - **Repository name:** `storage-cleanup-utility` (or any name you want)
   - **Visibility:** Choose **Private** (recommended) — only you can see it. **Public** also works and is free.
   - Leave everything else at defaults. **Do NOT** check "Add a README file" — you'll upload your own.
4. Click **"Create repository"**.

You'll land on a page that says *"Quick setup — if you've done this kind of thing before"*. Keep this tab open.

---

## Step 3 — Upload the project files

1. On the repository page, click the link that says **"uploading an existing file"** (in the blue setup box), or go to the **"Add file"** dropdown → **"Upload files"**.
2. Drag-and-drop **all** of the following files and folders from the project into the browser:
   - `storage_cleanup_utility.py`
   - `requirements.txt`
   - `build.bat`
   - `.gitignore`
   - `README.md`
   - `BUILD_INSTRUCTIONS.md` (this file)
   - The **`.github`** folder (containing `workflows/build.yml`) — **this is critical; the build won't start without it**

   > **Note:** When dragging folders, GitHub's uploader preserves the folder structure automatically. If you only see individual files after dragging, check that `.github/workflows/build.yml` appears with its path intact.

3. Scroll down to the **"Commit changes"** section.
4. In the first text box type: `Initial upload`
5. Click **"Commit changes"** (green button).

---

## Step 4 — Watch the build run automatically

1. Click the **"Actions"** tab at the top of your repository page.
2. You'll see a workflow run called **"Initial upload"** (or similar), with a yellow dot indicating it's running.
3. Click on it to see live progress. Building takes about **2–4 minutes**.
4. When finished, the dot turns into a **green checkmark** ✅.

### If you see a red X instead

Click the failed run to see the error. The most common cause is a missing file — usually the `.github/workflows/build.yml` file didn't upload with its folder path. Re-upload it and the build will retry automatically.

---

## Step 5 — Download your .exe

1. Still on the **Actions** tab, click on the successful green-checkmark run.
2. Scroll to the bottom of the page.
3. Under the **"Artifacts"** section you'll see **`Storage_Cleanup_Utility`**.
4. Click it to download a ZIP file.
5. Extract the ZIP — inside you'll find **`Storage_Cleanup_Utility.exe`**.

**That's your finished utility.** Copy it to any Windows 10/11 machine and double-click to run. No installation needed.

---

## Rebuilding later (if you ever change the code)

**Option A — Edit in the browser:**
1. Open the file in GitHub.
2. Click the pencil (✏️) icon to edit.
3. Make your changes. Click **"Commit changes"**.
4. The build runs automatically. Download the new .exe from Actions.

**Option B — Trigger a rebuild without code changes:**
1. Go to **Actions** tab → **"Build Storage Cleanup Utility"** (left sidebar).
2. Click **"Run workflow"** → **"Run workflow"** (green button).
3. New build starts. Download when done.

---

## Troubleshooting

**"Artifact retention expired"** — GitHub automatically deletes artifacts after 30 days on free accounts. Just trigger a rebuild (Option B above) and a fresh artifact will be produced.

**"Workflow file not found"** — The `.github/workflows/build.yml` file must be at exactly that path. If you accidentally uploaded just `build.yml` to the root, delete it, then upload the full `.github` folder again via drag-and-drop.

**Windows Defender flags the .exe** — This is a known false positive with PyInstaller-built executables. Click **"More info"** → **"Run anyway"**, or add an exclusion in Windows Security. The source code in this repo is what got built — nothing malicious is in it.

**The .exe is ~15–25 MB** — Normal. Python and all libraries are bundled inside.

---

## Privacy note

If your repository is **Private**, only you can see the code and the built .exe. GitHub's build servers run in isolated containers and are destroyed after each build; they don't retain your code or the exe beyond the artifact retention window.
