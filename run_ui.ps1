$ErrorActionPreference = "Stop"

# One-command launcher for the Streamlit UI - no manual venv activation needed.
#   .\run_ui.ps1
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$venvPython = Join-Path $root ".venv\Scripts\python.exe"
if (-not (Test-Path $venvPython)) {
    Write-Error "Virtual env not found at .venv. Run .\setup.ps1 first."
    exit 1
}

# Install Streamlit into the venv on first run (find_spec avoids importing it).
& $venvPython -c "import importlib.util, sys; sys.exit(0 if importlib.util.find_spec('streamlit') else 1)"
if ($LASTEXITCODE -ne 0) {
    Write-Host "Streamlit not installed - installing it into the venv..."
    & $venvPython -m pip install streamlit
}

Write-Host "Starting the UI at http://localhost:8501  (press Ctrl+C to stop)"
& $venvPython -m streamlit run (Join-Path $root "ui\app.py")
