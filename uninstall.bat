@echo off
echo Removing custom .zip file association...
reg delete "HKCU\Software\Classes\.zip"            /f >nul 2>&1
reg delete "HKCU\Software\Classes\ZipAutoExtract"  /f >nul 2>&1
echo Done. .zip files will use the Windows default again.
pause
