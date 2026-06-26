# Patent Reference Checker — Installation Guide

## What this is

A Microsoft Word add-in that:
1. Scans patent specification text for reference number conflicts (e.g., number 102 used for both "interface" and "module", or "interface" assigned two different numbers)
2. Marks conflicts with Word comments for review
3. Generates a sorted reference numerals table at the end of the document

---

## Requirements

- Microsoft Word 2016 (build 7870 or later), or Microsoft 365
- Internet connection (loads add-in files from GitHub Pages)

No Python, no Node.js, no installation of any software required.

---

## One-time setup per machine

### Step 1 — Save manifest.xml somewhere accessible

Copy `manifest.xml` from this repository to a folder on your computer (or a shared network drive if multiple people will use the add-in).

Example: `C:\Users\YourName\Documents\PatentAddin\manifest.xml`

### Step 2 — Add the folder as a Trusted Catalog in Word

1. Open Microsoft Word.
2. Go to **File → Options → Trust Center → Trust Center Settings**.
3. Click **Trusted Add-in Catalogs** in the left panel.
4. In the **Catalog Url** box, type the folder path where you saved `manifest.xml`:
   `C:\Users\YourName\Documents\PatentAddin`
5. Click **Add catalog**.
6. Check the **Show in Menu** checkbox next to the catalog you just added.
7. Click **OK** twice.
8. **Restart Word.**

> **Shared network drive:** If `manifest.xml` is on a shared drive (e.g. `\\server\shared\PatentAddin`), enter that UNC path instead. Everyone who maps that drive can install the add-in the same way — no individual setup needed beyond adding the catalog once.

---

## Installing the add-in

After completing the one-time setup above:

1. Open a patent specification document.
2. Go to **Insert → My Add-ins**.
3. Click the **SHARED FOLDER** tab.
4. Select **Patent Reference Checker** and click **Add**.
5. A **"Check References"** button appears in the **Home** tab ribbon.

The add-in loads from GitHub Pages — as long as you have internet, it just works. Nothing to run or install.

---

## Using the add-in

### Scan
Click **Check References** in the Home ribbon → task pane opens on the right.
Click **Scan Document** → the pane displays:
- **Conflicts** (red badge): reference number inconsistencies detected
- **Reference Dictionary**: all detected component–number pairs, sorted by number

### Mark conflicts in the document
Click **Mark Conflicts in Doc** → Word comments are added at each conflict location.
Review them in the Comments pane (**Review → Show Comments**).

### Insert reference table
Click **Insert Reference Table** → a "Reference Numerals" heading and two-column table are appended to the end of the document, sorted by reference number.

---

## Troubleshooting

| Symptom | Fix |
|---------|-----|
| "Check References" button missing after restart | Go to Insert → My Add-ins → Shared Folder tab and re-add it |
| Task pane shows blank or error | Check internet connection — add-in files load from GitHub Pages |
| "InsertComment is not a function" | Upgrade to Office 2016 build 7870+ or Microsoft 365 |
| Too many false positives in scan | The scanner may flag figure references or page numbers; known v1 limitation — use the dictionary to identify and ignore them |
| Want to use offline | Not supported in v1 — requires internet to load Office.js and add-in files |

---

## How it works (technical summary)

- Add-in files are hosted at `https://rxrm27.github.io/ms-word-addon/`
- `manifest.xml` tells Word where to find the files — it's the only file you need to distribute
- No server to run, no software to install, no ongoing maintenance
