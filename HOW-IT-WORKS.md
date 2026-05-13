# Zip Extractor — How It Works (Deep Documentation)

This document explains every file in the project in plain English. It is written for someone who understands what the tool does but doesn't write code — the goal is that if something breaks, you can understand exactly what went wrong and why, without being left stranded.

---

## Table of Contents

1. [The Big Picture](#1-the-big-picture)
2. [File Map](#2-file-map)
3. [How It Works End-to-End (Step-by-Step)](#3-how-it-works-end-to-end-step-by-step)
4. [File-by-File Breakdown](#4-file-by-file-breakdown)
   - [setup.bat](#setupbat)
   - [uninstall.bat](#uninstallbat)
   - [extract-zip.vbs](#extract-zipvbs)
   - [extract-zip.ps1](#extract-zipps1)
5. [What Happens to Your Files](#5-what-happens-to-your-files)
6. [Why There Are Four Files for One Thing](#6-why-there-are-four-files-for-one-thing)
7. [Common Problems and What Causes Them](#7-common-problems-and-what-causes-them)
8. [Glossary of Terms](#8-glossary-of-terms)

---

## 1. The Big Picture

Zip Extractor replaces the default Windows behavior when you double-click a `.zip` file.

**Default Windows behavior:** Opens the zip as a browseable folder in File Explorer. You still have to manually extract the contents somewhere.

**After running setup:** Double-clicking a `.zip` file automatically:
1. Extracts the contents into a new folder in the same location as the zip
2. Sends the original `.zip` file to the Recycle Bin
3. Opens that new folder in File Explorer

The entire operation is invisible — no terminal window, no dialog boxes, nothing. It just happens, and the resulting folder opens up in front of you.

**Example:**
- You double-click `C:\Downloads\mod-pack.zip`
- A folder `C:\Downloads\mod-pack\` is created containing all the contents
- `mod-pack.zip` goes to the Recycle Bin
- File Explorer navigates to `C:\Downloads\mod-pack\`

To undo this and go back to Windows default behavior, you run `uninstall.bat`.

---

## 2. File Map

```
zip-extractor/
│
├── setup.bat          ← Run once to register the custom .zip handler
├── uninstall.bat      ← Run once to remove it and restore Windows default
│
├── extract-zip.vbs    ← The "launcher" — Windows calls this when you open a .zip
└── extract-zip.ps1    ← The "worker" — does the actual extraction and cleanup
```

The split between `.vbs` and `.ps1` exists because of a Windows limitation. More on that in [Section 6](#6-why-there-are-four-files-for-one-thing).

---

## 3. How It Works End-to-End (Step-by-Step)

Here is the exact sequence of events from double-clicking a zip file to the folder opening:

1. **You double-click a `.zip` file** in File Explorer.

2. **Windows looks up the file association for `.zip`** in the registry. Normally it finds its own built-in handler. After `setup.bat` has been run, it finds `ZipAutoExtract` instead — a custom handler that points to `extract-zip.vbs`.

3. **Windows launches `wscript.exe extract-zip.vbs "C:\path\to\file.zip"`**. `wscript.exe` is the Windows Script Host — the engine that runs `.vbs` files. It passes the path of the zip you double-clicked as an argument.

4. **`extract-zip.vbs` runs.** Its only job is to figure out where `extract-zip.ps1` lives (it's always in the same folder as the `.vbs`) and launch it silently via PowerShell, passing along the zip file path. The VBS window never appears.

5. **`extract-zip.ps1` runs with the zip file path as input.** It:
   - Validates the path (does the file actually exist?)
   - Determines the destination folder name (same name as the zip, minus `.zip`, in the same directory)
   - Extracts all contents into that folder
   - Sends the original zip to the Recycle Bin
   - Finds any open File Explorer window that's showing the folder the zip was in, and navigates it to the new extracted folder
   - If no Explorer window was open there, opens a new one pointing to the new folder

6. **You see the extracted folder open in File Explorer.** The zip is gone (in the Recycle Bin). Done.

---

## 4. File-by-File Breakdown

---

### `setup.bat`

```bat
@echo off
setlocal

set "HANDLER=%~dp0extract-zip.vbs"

echo Registering .zip file association...
reg add "HKCU\Software\Classes\.zip"                              /ve /d "ZipAutoExtract" /f >nul
reg add "HKCU\Software\Classes\ZipAutoExtract"                    /ve /d "Zip Auto Extract" /f >nul
reg add "HKCU\Software\Classes\ZipAutoExtract\shell\open\command" /ve /d "wscript.exe \"%HANDLER%\" \"%%1\"" /f >nul
reg add "HKCU\Software\Classes\ZipAutoExtract\DefaultIcon"        /ve /d "%SystemRoot%\system32\zipfldr.dll,0" /f >nul
```

This script runs once and writes four registry entries. That's all it does.

**Line 1 — `@echo off`:** Prevents the command prompt from printing each command on screen before it runs.

**Line 2 — `setlocal`:** Creates an isolated environment for variables defined in this script. Any `set` commands made here are discarded when the script ends, so they don't contaminate your system environment.

**Line 4 — `set "HANDLER=%~dp0extract-zip.vbs"`:** Builds the full path to the `.vbs` file and stores it in a variable called `HANDLER`.
- `%~dp0` is a special .bat variable meaning "the drive letter and folder path of the currently running .bat file." If the `.bat` file is at `C:\Tools\zip-extractor\setup.bat`, then `%~dp0` is `C:\Tools\zip-extractor\`.
- Appending `extract-zip.vbs` gives the full absolute path to the VBS file.
- This is important: the registry entry will permanently store this path. If you move the folder later, the path becomes stale and double-clicking zips will fail. That's exactly why the script reminds you: *"Do not move this folder."*

**The four `reg add` commands:**

Registry entries come in a hierarchy — keys (like folders) that contain values (like files). This script creates a branch under `HKCU\Software\Classes` (the current user's file associations). Breaking down each entry:

---

**Entry 1:**
```
reg add "HKCU\Software\Classes\.zip" /ve /d "ZipAutoExtract" /f >nul
```
- `HKCU\Software\Classes\.zip` — this key is where Windows looks up what handler to use for `.zip` files.
- `/ve` — sets the *default value* of this key (the one with no name, displayed as `(Default)` in Registry Editor).
- `/d "ZipAutoExtract"` — the data to write: the name of the handler class to use. This is just a name; the actual handler definition comes in the next entries.
- `/f` — force overwrite without prompting.
- `>nul` — discard the "Operation completed successfully" output text so it doesn't clutter the screen.

**What this does:** Tells Windows "when the user opens a `.zip` file, look for the handler named `ZipAutoExtract`."

---

**Entry 2:**
```
reg add "HKCU\Software\Classes\ZipAutoExtract" /ve /d "Zip Auto Extract" /f >nul
```
- Creates the handler class key `ZipAutoExtract`.
- The default value `"Zip Auto Extract"` is the human-readable display name for this handler. It would appear in "Open with" dialogs.

---

**Entry 3:**
```
reg add "HKCU\Software\Classes\ZipAutoExtract\shell\open\command" /ve /d "wscript.exe \"%HANDLER%\" \"%%1\"" /f >nul
```
This is the most important entry. The `shell\open\command` subkey tells Windows exactly what program to run when the file is opened (double-clicked).

The command written to the registry is:
```
wscript.exe "C:\path\to\extract-zip.vbs" "%1"
```

- `wscript.exe` — the Windows Script Host engine that runs `.vbs` files.
- `"C:\path\to\extract-zip.vbs"` — the full path to the VBS launcher (stored in `%HANDLER%`).
- `"%1"` — a placeholder that Windows replaces with the full path of the file being opened. So when you double-click `C:\Downloads\mod.zip`, Windows substitutes `%1` with `C:\Downloads\mod.zip`.

**Why the backslashes before the quotes (`\"`):** Inside a `reg add` command, double quotes that should appear literally in the stored value need to be escaped with `\`. So `\"` in the bat file becomes `"` in the registry.

**Why `%%1` instead of `%1`:** Inside a `.bat` file, `%` is a special character used for variables (`%HANDLER%`, `%~dp0`). To write a literal `%` into the registry, you double it: `%%1` → writes `%1` to the registry.

---

**Entry 4:**
```
reg add "HKCU\Software\Classes\ZipAutoExtract\DefaultIcon" /ve /d "%SystemRoot%\system32\zipfldr.dll,0" /f >nul
```
- `DefaultIcon` is a registry subkey that tells Windows what icon to show for files handled by this class.
- `%SystemRoot%\system32\zipfldr.dll,0` points to the standard Windows zip icon (the first icon, index `0`, from Windows' own zip folder library). This keeps `.zip` files showing the familiar zip icon in File Explorer instead of a blank generic icon.

---

### `uninstall.bat`

```bat
@echo off
echo Removing custom .zip file association...
reg delete "HKCU\Software\Classes\.zip"            /f >nul 2>&1
reg delete "HKCU\Software\Classes\ZipAutoExtract"  /f >nul 2>&1
echo Done. .zip files will use the Windows default again.
pause
```

This script deletes the two top-level registry keys that `setup.bat` created. Deleting a key in the registry also deletes all its subkeys, so two deletions clean up all four entries that setup created.

**`/f`** — force delete without a confirmation prompt.

**`>nul 2>&1`** — this suppresses all output from the `reg delete` command, including error messages.
- `>nul` discards standard output (the success message).
- `2>&1` redirects the error output stream (`2`) to the same place as standard output (`1`), which is already going to `nul`. This means if the key doesn't exist (already deleted or never created), the error is silently swallowed rather than showing a scary red message.

**After running uninstall:** Windows no longer finds `ZipAutoExtract` in the user registry for `.zip` files. It falls back to the system-level default (the built-in Windows zip folder viewer). No reboot or log-off required — the change takes effect immediately.

---

### `extract-zip.vbs`

```vbs
Dim WshShell, zipPath, psScript, cmd
Set WshShell = CreateObject("WScript.Shell")
zipPath  = WScript.Arguments(0)
psScript = Replace(WScript.ScriptFullName, ".vbs", ".ps1")

cmd = "powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File """ & psScript & """ -ZipPath """ & zipPath & """"
WshShell.Run cmd, 0, False
```

This file is written in VBScript — an old Windows scripting language that's been part of Windows since the late 1990s. It's used here as a thin launcher, not as the main logic.

**Line 1 — `Dim WshShell, zipPath, psScript, cmd`:** Declares four variables. `Dim` is VBScript for "declare variable."

**Line 2 — `Set WshShell = CreateObject("WScript.Shell")`:** Creates a `WScript.Shell` object — a built-in COM object that provides access to shell functionality, most importantly the ability to run external programs.

**Line 3 — `zipPath = WScript.Arguments(0)`:** `WScript.Arguments` is the list of arguments that were passed to this script. Index `0` is the first argument — the path to the zip file that Windows passed in (from the `%1` in the registry command entry). This is the full path, e.g., `C:\Downloads\mod.zip`.

**Line 4 — `psScript = Replace(WScript.ScriptFullName, ".vbs", ".ps1")`:**
- `WScript.ScriptFullName` is a built-in VBScript property that gives you the full path of the currently running script file, e.g., `C:\Tools\zip-extractor\extract-zip.vbs`.
- `Replace(..., ".vbs", ".ps1")` swaps the `.vbs` extension for `.ps1`, giving `C:\Tools\zip-extractor\extract-zip.ps1`.
- This is how the VBS always finds the PS1 file regardless of where the folder is located — it finds itself, then looks for a file with the same name and `.ps1` extension next to it.

**Lines 6–7 — Building and running the command:**

```vbs
cmd = "powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File """ & psScript & """ -ZipPath """ & zipPath & """"
WshShell.Run cmd, 0, False
```

This builds a command string and runs it. The `&` operator in VBScript concatenates strings.

The final command looks like:
```
powershell.exe -ExecutionPolicy Bypass -NoProfile -WindowStyle Hidden -File "C:\Tools\zip-extractor\extract-zip.ps1" -ZipPath "C:\Downloads\mod.zip"
```

Breaking down the PowerShell flags:
- `-ExecutionPolicy Bypass` — overrides Windows' restriction on running unsigned scripts. Without this, PowerShell might refuse to run the `.ps1` file.
- `-NoProfile` — don't load the user's PowerShell profile (personal customizations). Keeps the environment clean and predictable.
- `-WindowStyle Hidden` — this is the critical one. It tells PowerShell to run completely invisibly — no terminal window ever appears. The extraction happens in total silence.
- `-File "..."` — the script to run.
- `-ZipPath "..."` — a named argument passed to the script (corresponds to `param([string]$ZipPath)` at the top of the `.ps1`).

**`WshShell.Run cmd, 0, False`:**
- First argument: the command to run.
- Second argument `0`: the window style — `0` means hidden. Belt-and-suspenders alongside `-WindowStyle Hidden`.
- Third argument `False`: don't wait for the command to finish before returning. The VBS script exits immediately after launching PowerShell, rather than hanging until the extraction is complete.

---

### `extract-zip.ps1`

This is where the real work happens. It receives the zip file path and performs three operations: extract, delete, navigate.

#### Parameter Declaration (Line 1)

```powershell
param([string]$ZipPath)
```

`param(...)` at the top of a PowerShell script declares the script's input parameters. This means when the script is called with `-ZipPath "C:\Downloads\mod.zip"`, that value is available inside the script as `$ZipPath`.

#### Input Validation (Line 3)

```powershell
if (-not $ZipPath -or -not (Test-Path -LiteralPath $ZipPath)) { exit 1 }
```

This is a safety check before doing anything. It exits immediately with error code `1` (failure) if either condition is true:
- `-not $ZipPath` — the `$ZipPath` variable is empty or was never set.
- `-not (Test-Path -LiteralPath $ZipPath)` — the file path was given but doesn't actually exist on disk.

**Why `-LiteralPath` instead of `-Path`:** PowerShell's `-Path` parameter interprets certain characters as wildcards (like `[`, `]`, `*`, `?`). A file named `mod[v2].zip` would cause `-Path` to fail. `-LiteralPath` treats the path as a plain string with no special characters, making it safe for any valid filename.

#### Calculating the Destination (Lines 5–6)

```powershell
$parentDir = [IO.Path]::GetDirectoryName($ZipPath)
$dest      = Join-Path $parentDir ([IO.Path]::GetFileNameWithoutExtension($ZipPath))
```

These two lines figure out where to extract the files.

`[IO.Path]::GetDirectoryName($ZipPath)` — a .NET method that extracts just the folder part of a full file path. Given `C:\Downloads\mod-pack.zip`, it returns `C:\Downloads`.

`[IO.Path]::GetFileNameWithoutExtension($ZipPath)` — another .NET method that extracts the filename without its extension. Given `C:\Downloads\mod-pack.zip`, it returns `mod-pack`.

`Join-Path $parentDir (...)` — combines a folder path and a name using the correct path separator for the OS. Given `C:\Downloads` and `mod-pack`, it returns `C:\Downloads\mod-pack`.

**Result:** `$dest = "C:\Downloads\mod-pack"` — the folder that will be created to hold the extracted contents.

#### The try/catch Block (Lines 8–33)

Everything from here is wrapped in a `try/catch`. If anything fails, the `catch` block at the bottom shows an error dialog instead of crashing silently.

---

#### Step 1 — Extract (Line 9)

```powershell
Expand-Archive -LiteralPath $ZipPath -DestinationPath $dest -Force
```

`Expand-Archive` is a built-in PowerShell cmdlet that extracts zip files. Parameters:
- `-LiteralPath $ZipPath` — the source zip file (using literal path, same reason as `Test-Path` above).
- `-DestinationPath $dest` — the folder to extract into. If the folder doesn't exist, PowerShell creates it.
- `-Force` — if files already exist at the destination, overwrite them. Without `-Force`, the cmdlet would throw an error if the destination folder already contains anything.

---

#### Step 2 — Send Zip to Recycle Bin (Lines 11–14)

```powershell
Add-Type -AssemblyName Microsoft.VisualBasic
[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
    $ZipPath, 'OnlyErrorDialogs', 'SendToRecycleBin'
)
```

**Why not just `Remove-Item`?** PowerShell's `Remove-Item` permanently deletes a file — it bypasses the Recycle Bin entirely. For a tool that runs automatically on double-click, permanent silent deletion is dangerous. The Recycle Bin approach means you can recover the original zip if something goes wrong.

**`Add-Type -AssemblyName Microsoft.VisualBasic`:** This loads the `Microsoft.VisualBasic` .NET library into the current PowerShell session. PowerShell is built on .NET and can use any .NET library, but they have to be explicitly loaded first. This library is included with Windows and provides file system utilities, including the Recycle Bin delete function.

**`[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(...)`:** Calls the static `DeleteFile` method from this library. Three arguments:
1. `$ZipPath` — the file to delete.
2. `'OnlyErrorDialogs'` — the UI option. `OnlyErrorDialogs` means: don't show any confirmation prompts or progress bars, but do show a dialog if an actual error occurs (like the file is locked). The other option would be `AllDialogs` which shows the animated "sending to recycle bin" window.
3. `'SendToRecycleBin'` — the recycle option. This sends the file to the Recycle Bin instead of permanently deleting it. The alternative `DeletePermanently` would bypass the bin.

---

#### Step 3 — Navigate to the Extracted Folder (Lines 16–29)

This section's goal: make File Explorer show the newly created folder. It tries to reuse an existing open window first, and only opens a new one if needed.

```powershell
$shell     = New-Object -ComObject Shell.Application
$navigated = $false
foreach ($w in $shell.Windows()) {
    try {
        if ($w.Document.Folder.Self.Path -eq $parentDir) {
            $w.Navigate2($dest)
            $navigated = $true
            break
        }
    } catch {}
}
if (-not $navigated) {
    Start-Process explorer.exe -ArgumentList "`"$dest`""
}
```

**`New-Object -ComObject Shell.Application`:** Creates a `Shell.Application` COM object — a Windows interface that gives scripts programmatic access to File Explorer windows that are currently open. COM objects are a Windows technology for inter-process communication; this one is built into Windows and is the official way to interact with the running Explorer shell.

**`$shell.Windows()`:** Returns a collection of all currently open Explorer/Internet Explorer windows. Each item in this collection represents one open window.

**The `foreach` loop:** Iterates over every open window. For each one, it tries to read `$w.Document.Folder.Self.Path` — the full path of the folder that window is currently showing.

**The inner `try/catch {}`:** Some windows in the `Shell.Windows()` collection might not be File Explorer windows (e.g., Internet Explorer instances, or windows in a partially-initialized state). Accessing `.Document.Folder.Self.Path` on those would throw an error. The empty `catch {}` silently swallows those errors and moves on to the next window — it's intentional defensive coding.

**The check `if ($w.Document.Folder.Self.Path -eq $parentDir)`:** If this window is showing the same folder that contains the zip file, it's the right window to reuse.

**`$w.Navigate2($dest)`:** Tells that Explorer window to navigate to the destination folder. This is the same as clicking a folder in the navigation pane — the window's contents change to show `$dest`. `Navigate2` is the modern version of `Navigate` and handles a wider range of path formats.

**`$navigated = $true; break`:** Marks that we found and navigated a window, then exits the loop — no need to keep checking other windows.

**The fallback `if (-not $navigated)`:** If no open Explorer window was showing `$parentDir`, open a brand new window:
```powershell
Start-Process explorer.exe -ArgumentList "`"$dest`""
```
This launches a new `explorer.exe` process, passing the destination folder path as an argument. The backtick-escaped quotes (`` ` ``) around `$dest` ensure the path is quoted correctly even if it contains spaces.

---

#### Error Handling (Lines 30–33)

```powershell
} catch {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("Failed to extract:`n$_", "Zip Extractor")
}
```

If anything in the `try` block fails — extraction error, file locked, disk full, permissions issue, anything — execution jumps to this `catch` block.

`Add-Type -AssemblyName System.Windows.Forms` loads the Windows Forms .NET library, which provides UI components including dialog boxes. It has to be loaded before use, just like `Microsoft.VisualBasic` above.

`[System.Windows.Forms.MessageBox]::Show(...)` shows a standard Windows message box popup.
- First argument: the message. `"Failed to extract:\n$_"` — the text "Failed to extract:" followed by a newline (`\n`) and the actual error message. In PowerShell, `$_` inside a `catch` block refers to the error that was caught.
- Second argument: the title bar text of the dialog ("Zip Extractor").

This means if something goes wrong, you'll see a popup dialog explaining what failed, rather than the operation silently not working.

---

## 5. What Happens to Your Files

Here is the exact file system state before and after double-clicking a zip:

**Before:**
```
C:\Downloads\
    mod-pack.zip          ← exists, 45 MB
```

**After:**
```
C:\Downloads\
    mod-pack\             ← new folder, created by the script
        readme.txt
        data\
        config.ini
                          ← mod-pack.zip is gone from here (in Recycle Bin)
```

**Key behaviors:**

- **The destination folder name** is always the zip filename with the extension removed. `archive.zip` → `archive\`. `my file (v2).zip` → `my file (v2)\`.

- **The destination folder is always in the same directory as the zip.** If the zip is on your Desktop, the folder appears on your Desktop. If it's in Downloads, the folder appears in Downloads.

- **If a folder with that name already exists,** `-Force` on `Expand-Archive` means the files are extracted into it anyway, overwriting any conflicting files. Files in the existing folder that don't conflict are left untouched.

- **The original zip goes to the Recycle Bin, not permanent deletion.** You can open the Recycle Bin and restore it if needed.

- **No confirmation, no progress bar.** The entire operation is silent. The only visible effect is File Explorer showing the new folder when it's done.

---

## 6. Why There Are Four Files for One Thing

You might wonder why this needs a `.bat`, a `.vbs`, and a `.ps1` — why not just one script? Each file solves a specific problem that the others can't.

### Why `setup.bat` exists

Registry editing via `reg add` is a simple, built-in command that works reliably from a `.bat` file without needing PowerShell or elevation. It could have been a PowerShell script, but a `.bat` file is something anyone can double-click without worrying about execution policies.

### Why `extract-zip.vbs` exists (and not just `extract-zip.ps1` directly)

Windows file associations can point to an executable (`.exe`) or a script engine with a script file, but they have a limitation: **they can't launch a PowerShell window in hidden/invisible mode directly from a registry association.**

If the registry `shell\open\command` pointed directly to `powershell.exe -WindowStyle Hidden -File "extract-zip.ps1" "%1"`, Windows would still flash a console window briefly on screen before PowerShell had a chance to hide it.

VBScript doesn't have this problem. When launched via `wscript.exe` (as opposed to `cscript.exe`), VBS runs completely without a window from the very start. And `WshShell.Run` with window style `0` launches its child process (PowerShell) in a truly hidden state from the moment it starts.

So the chain is:
- **Registry** calls `wscript.exe extract-zip.vbs` → no window ever appears
- **VBScript** calls `powershell.exe -WindowStyle Hidden -File extract-zip.ps1` → starts hidden
- **PowerShell** does the actual work invisibly

### Why `extract-zip.ps1` exists (and not just VBScript)

VBScript is old and limited. It has no built-in zip extraction, no Recycle Bin API (without external libraries), and no easy way to interface with File Explorer's open windows. PowerShell has `Expand-Archive`, the .NET `Microsoft.VisualBasic.FileIO.FileSystem` Recycle Bin API, and the `Shell.Application` COM interface for window navigation. The PS1 is the right tool for the actual work.

### Summary

| File | Role | Why this type |
|---|---|---|
| `setup.bat` | One-time registry setup | Simple, no elevation tricks needed, anyone can run it |
| `uninstall.bat` | One-time registry cleanup | Same |
| `extract-zip.vbs` | Silent launcher | Only way to start PowerShell with truly no window flash |
| `extract-zip.ps1` | The actual extraction logic | Has the APIs needed: zip, recycle bin, explorer navigation |

---

## 7. Common Problems and What Causes Them

### Double-clicking a zip still opens it as a folder (Windows default behavior)

**Cause 1:** `setup.bat` was never run, or was run but failed.  
**Fix:** Run `setup.bat` again. Check that no error messages appeared.

**Cause 2:** Another application (7-Zip, WinRAR, etc.) has claimed the `.zip` association at the system level (`HKLM`), which overrides the user-level setting (`HKCU`) that this script writes to.  
**Fix:** In Windows Settings → Apps → Default Apps, find `.zip` and set it back to "Zip Auto Extract" or check if the other application is overriding it.

**Cause 3:** The registry was written correctly but Windows cached the old association. Windows sometimes caches file associations.  
**Fix:** Sign out and sign back in, or restart.

### Nothing happens when I double-click a zip

**Cause 1:** The `.vbs` file was moved or deleted. The registry still has the old path.  
**Fix:** Run `uninstall.bat` then `setup.bat` again from the current folder location.

**Cause 2:** The folder was moved after `setup.bat` was run. The registry path is now stale.  
**Fix:** Same — run `uninstall.bat` then `setup.bat` again from the new location.

**Cause 3:** `wscript.exe` (Windows Script Host) has been disabled on the machine. Some corporate IT policies disable it.  
**Fix:** This would need to be re-enabled via Group Policy — may require IT involvement.

### I see "Failed to extract" error dialog

The error message in the dialog will tell you the specific reason. Common causes:

- **"Access to the path is denied"** — You don't have write permission to the folder where the zip is located. This happens with zips in system folders, on read-only network drives, or on locked external drives. Move the zip somewhere you have write access (like Downloads or Desktop) and try again.

- **"The archive entry was compressed using an unsupported compression method"** — The zip uses a compression format that `Expand-Archive` doesn't support (e.g., zip files compressed with Deflate64, some 7-Zip methods, or password-protected zips). `Expand-Archive` only handles standard zip/Deflate. Use 7-Zip or WinRAR to extract these manually.

- **"The file is in use by another process"** — Something else has the zip file open (antivirus scan, another extraction attempt, etc.). Wait a moment and try again.

### The extraction works but no File Explorer window opens

**Cause:** The `explorer.exe` launch at the end of the script failed silently, or the `Shell.Application` navigation failed.  
**Fix:** The files are still there — open File Explorer manually and navigate to the folder where the zip was. The extracted folder should be there.

### The zip was extracted but NOT deleted

**Cause:** The extraction succeeded but the Recycle Bin deletion failed (e.g., the file became locked after extraction, or you're on a network share that doesn't support the Recycle Bin).  
**Fix:** Delete the zip manually. This is safe — the contents are already extracted.

### After running `uninstall.bat`, zips open strangely or show an error

**Cause:** Windows may fall back to a stale or broken built-in association before it fully resets to the default. This can happen if another app had previously set a user-level association that `uninstall.bat` didn't clean up.  
**Fix:** In Windows Settings → Apps → Default Apps → scroll down to "Choose defaults by file type" → find `.zip` → set it to Windows Explorer (the Windows built-in).

---

## 8. Glossary of Terms

**`%~dp0`:** A special variable available inside `.bat` files that expands to the drive letter and directory path of the batch file itself. Always ends with a backslash. Used to build absolute paths relative to the script's location, so the script works regardless of where the folder is placed.

**`%1`:** In a Windows registry file association command, `%1` is a placeholder that Windows replaces with the full path of the file that was double-clicked. It becomes the argument passed to the handler program.

**Assembly / AssemblyName:** In .NET, an "assembly" is a compiled library of code — roughly equivalent to a `.dll` file. `Add-Type -AssemblyName` in PowerShell loads one of these libraries so its functions can be called.

**COM Object:** Component Object Model — a Windows technology that allows programs to expose functionality to other programs. `Shell.Application` and `WScript.Shell` are COM objects. PowerShell can create and use COM objects with `New-Object -ComObject`.

**`cscript.exe` vs `wscript.exe`:** Both run VBScript files. `cscript.exe` runs in a terminal window (command line output). `wscript.exe` runs without any window — it's the "Windows" (GUI-suppressed) script host. This project uses `wscript.exe` specifically because it's windowless.

**Default Value (Registry):** Every registry key can have a nameless "default" value, displayed as `(Default)` in Registry Editor. The `/ve` flag in `reg add` targets this default value. For file association keys, this default value is typically the display name or handler class name.

**Execution Policy:** A PowerShell safety feature that controls whether scripts can run. `Bypass` overrides all restrictions for a single invocation without changing system settings. This is used because the `.ps1` file is not signed with a code-signing certificate.

**`HKCU` vs `HKLM`:** Two top-level branches of the Windows Registry.
- `HKCU` (`HKEY_CURRENT_USER`) — settings that apply only to the currently logged-in user. No admin rights needed to write here.
- `HKLM` (`HKEY_LOCAL_MACHINE`) — settings that apply to all users on the machine. Requires Administrator rights to write.

This project writes to `HKCU`, which means it only affects your own account and doesn't need admin elevation.

**`Join-Path`:** A PowerShell cmdlet that combines path components correctly. It handles slashes automatically: `Join-Path "C:\Downloads" "myfile"` → `C:\Downloads\myfile`, regardless of whether the first argument ends with a slash or not.

**`-LiteralPath`:** A PowerShell parameter that treats the path exactly as given, with no wildcard expansion. Use it when a filename might contain characters like `[`, `]`, or `*` that PowerShell would otherwise interpret as pattern-matching wildcards.

**Navigate2:** A method on the Shell COM `IWebBrowserApp` interface (which File Explorer windows implement). It tells an open Explorer window to change its current location to a new folder — the programmatic equivalent of clicking a folder.

**`param(...)`:** PowerShell syntax for declaring a script's input parameters. Placed at the very top of a `.ps1` file. Values are passed from the command line using `-ParameterName value` syntax.

**Recycle Bin:** Windows' temporary holding area for deleted files. Files sent here can be restored. The Microsoft.VisualBasic `DeleteFile` method with `SendToRecycleBin` is the standard programmatic way to send a file there.

**`reg add` / `reg delete`:** Built-in Windows command-line utilities for reading and writing registry keys and values. Available in all versions of Windows without any additional installation.

**`Shell.Application`:** A COM object provided by Windows that gives scripts access to the running File Explorer shell. Can enumerate open windows, navigate them, and perform shell operations. Accessed via `New-Object -ComObject Shell.Application` in PowerShell.

**`try/catch`:** An error-handling structure. Code in the `try` block runs normally. If an error occurs, execution immediately jumps to the `catch` block. Code after the error in `try` is skipped. This prevents the script from crashing silently — instead, the error is caught and shown in a dialog.

**VBScript (`.vbs`):** Visual Basic Script — a scripting language built into Windows since the late 1990s, run by `wscript.exe` or `cscript.exe`. In this project it serves only as a windowless launcher to avoid the console window flash that occurs when launching PowerShell directly from a registry file association.

**`WshShell.Run cmd, 0, False`:** A VBScript method to launch an external program. The three arguments are: (1) the command to run, (2) the window style (`0` = hidden), (3) whether to wait for it to finish (`False` = don't wait, return immediately).
