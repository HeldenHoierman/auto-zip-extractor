# Auto Zip Extractor

Double-click a `.zip` file and it extracts automatically — no dialogs, no manual steps. The contents appear in a new folder, the zip goes to the Recycle Bin, and File Explorer opens the result.

## What it does

**Before:** Double-clicking a zip opens it as a browseable folder in Explorer.

**After:** Double-clicking a zip:
1. Extracts contents into a new folder alongside the zip
2. Sends the zip to the Recycle Bin
3. Opens the extracted folder in File Explorer

Everything happens silently — no terminal window, no progress bar.

**Example:**
```
C:\Downloads\mod-pack.zip
→ C:\Downloads\mod-pack\   (extracted, folder opens)
→ mod-pack.zip             (moved to Recycle Bin)
```

## Setup

1. Move the folder somewhere permanent — the registry will point to this location
2. Run `setup.bat`

That's it. Double-clicking any `.zip` file will now trigger auto-extraction.

## Uninstall

Run `uninstall.bat` to remove the file association and restore Windows default behavior.

## Files

| File | Purpose |
|---|---|
| `setup.bat` | Registers the custom `.zip` handler in the registry |
| `uninstall.bat` | Removes it and restores Windows default |
| `extract-zip.vbs` | Silent launcher — avoids the console window flash |
| `extract-zip.ps1` | Does the actual extraction, deletion, and navigation |

## Requirements

- Windows 10 or 11
- PowerShell 5.1+ (included with Windows)

## Notes

- The zip goes to the **Recycle Bin**, not permanent deletion — you can recover it if needed
- If a folder with the same name already exists, files are extracted into it (conflicts are overwritten)
- Does not support password-protected zips or non-standard compression formats (use 7-Zip for those)
- Only affects your user account — no admin rights required

For a detailed technical breakdown of how each file works, see [HOW-IT-WORKS.md](HOW-IT-WORKS.md).

## License

This is free and unencumbered software released into the public domain. See [LICENSE](LICENSE).
