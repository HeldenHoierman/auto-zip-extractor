param([string]$ZipPath)

if (-not $ZipPath -or -not (Test-Path -LiteralPath $ZipPath)) { exit 1 }

$parentDir = [IO.Path]::GetDirectoryName($ZipPath)
$dest      = Join-Path $parentDir ([IO.Path]::GetFileNameWithoutExtension($ZipPath))

try {
    Expand-Archive -LiteralPath $ZipPath -DestinationPath $dest -Force

    Add-Type -AssemblyName Microsoft.VisualBasic
    [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
        $ZipPath, 'OnlyErrorDialogs', 'SendToRecycleBin'
    )

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
} catch {
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.MessageBox]::Show("Failed to extract:`n$_", "Zip Extractor")
}
