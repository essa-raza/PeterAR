$ErrorActionPreference = "Stop"

$python = "C:\Users\essar\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$distPath = Join-Path $projectRoot "dist"
$buildPath = Join-Path $projectRoot "build"
$venvPath = Join-Path $projectRoot ".venv"

if (-not (Test-Path $venvPath)) {
    & $python -m venv $venvPath
}

$venvPython = Join-Path $venvPath "Scripts\python.exe"

& $venvPython -m pip install --upgrade pip
& $venvPython -m pip install -r (Join-Path $projectRoot "requirements.txt")

& $venvPython (Join-Path $projectRoot "scripts\make_icon.py")

if (Test-Path $distPath) {
    Remove-Item -Recurse -Force $distPath
}

if (Test-Path $buildPath) {
    Remove-Item -Recurse -Force $buildPath
}

& $venvPython -m PyInstaller `
    --noconfirm `
    --clean `
    --windowed `
    --onefile `
    --name "CustomerNotesMerger" `
    --icon (Join-Path $projectRoot "assets\customer_notes_merger.ico") `
    --add-data (Join-Path $projectRoot "Docs;Docs") `
    (Join-Path $projectRoot "app.py")

Write-Host ""
Write-Host "Build complete:"
Write-Host (Join-Path $projectRoot "dist\CustomerNotesMerger.exe")
