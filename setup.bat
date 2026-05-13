@echo off
setlocal

set "HANDLER=%~dp0extract-zip.vbs"

echo.
echo === Zip Extractor Setup ===
echo.

echo Registering .zip file association...
reg add "HKCU\Software\Classes\.zip"                              /ve /d "ZipAutoExtract" /f >nul
reg add "HKCU\Software\Classes\ZipAutoExtract"                    /ve /d "Zip Auto Extract" /f >nul
reg add "HKCU\Software\Classes\ZipAutoExtract\shell\open\command" /ve /d "wscript.exe \"%HANDLER%\" \"%%1\"" /f >nul
reg add "HKCU\Software\Classes\ZipAutoExtract\DefaultIcon"        /ve /d "%SystemRoot%\system32\zipfldr.dll,0" /f >nul

echo Done!
echo.
echo NOTE: Do not move this folder. Run uninstall.bat first if you need to move it,
echo       then re-run setup.bat to re-register.
pause
