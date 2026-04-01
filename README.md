# File Sorter

Windows desktop app for flattening a folder and sorting its files into root-level folders.

## Features

- Sorts a selected folder into root-level folders
- Supports `By Extension` mode
- Supports `By Category` mode
- Pulls files up from nested subfolders
- Sends the original subfolders to the Recycle Bin after sorting
- Renames duplicates safely, for example `photo (1).jpg`
- Skips protected files and folders it cannot access
- Relaunches as administrator automatically when needed

## Grouping Modes

### By Extension

Creates one folder per extension, for example:

- `JPG`
- `PNG`
- `PDF`
- `TXT`
- `No Extension`

### By Category

Groups files into:

- `Images`
- `Video`
- `Project Files`
- `Documents`
- `Misc`

`Project Files` includes common saved project formats such as:

- Adobe Photoshop: `.psd`, `.psb`
- VEGAS Pro / Sony Vegas: `.veg`, `.vegtemplate`
- Adobe Premiere Pro: `.prproj`
- Adobe After Effects: `.aep`
- Blender: `.blend`
- Affinity: `.afphoto`, `.afdesign`, `.afpub`

## Important Behavior

- The selected folder stays in place.
- Files from nested subfolders are moved into root-level output folders.
- Original subfolders are sent to the Recycle Bin after sorting when possible.
- The Windows system drive root such as `C:\` is blocked on purpose.
- Non-system drive roots such as external drives are allowed.
- The app folder itself is protected so it cannot sort and move its own source files.

If you need to preserve the original folder structure, do not run this tool on that folder.

## Requirements

- Windows
- PowerShell
- Permission to approve the admin prompt when the app starts

## How To Run

### Option 1

Double-click `Run File Sorter.bat`

### Option 2

Run from PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File ".\FileSorter.ps1"
```

## How To Use

1. Launch the app.
2. Accept the admin prompt if Windows asks.
3. Choose a grouping mode.
4. Browse to the folder you want to sort.
5. Click `Sort Folder`.
6. Wait for the completion message.

## Files

- `FileSorter.ps1` - WinForms app UI and launch flow
- `Sorter.psm1` - sorting logic
- `Run File Sorter.bat` - double-click launcher

## Notes

- Protected files and folders are skipped instead of stopping the whole sort.
- Cleanup goes through the Windows Recycle Bin instead of permanent deletion.
- Category mode is meant for broad cleanup.
- Extension mode is better when you want exact file-type buckets.
