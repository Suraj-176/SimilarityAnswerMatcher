$ErrorActionPreference = "Stop"

$pythonExe = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"

if (-not (Test-Path $pythonExe)) {
    Write-Error "Project virtual environment not found at $pythonExe"
}

& $pythonExe -m streamlit run (Join-Path $PSScriptRoot "similarity_app.py")
