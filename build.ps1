$ErrorActionPreference = "Stop"
Push-Location $PSScriptRoot
try {
    if (-not (Test-Path '.venv-build\Scripts\python.exe')) {
        python -m venv .venv-build
        if ($LASTEXITCODE -ne 0) { throw 'Could not create build environment' }
    }
    & .\.venv-build\Scripts\python.exe -m pip install -r requirements-build.txt
    if ($LASTEXITCODE -ne 0) { throw 'Could not install build dependencies' }
    & .\.venv-build\Scripts\python.exe -m unittest -v test_automation
    if ($LASTEXITCODE -ne 0) { throw 'Tests failed; build cancelled' }
    & .\.venv-build\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --windowed --name UnityHelper main.py
    if ($LASTEXITCODE -ne 0) { throw 'Build failed' }
    Write-Host "Built $PSScriptRoot\dist\UnityHelper.exe"
} finally {
    Pop-Location
}
