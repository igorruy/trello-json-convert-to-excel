param([switch]$NoVenv)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

if (-not $NoVenv) {
    if (-not (Test-Path ".venv")) {
        python -m venv .venv
    }
}

$venvPy = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
$py = if (Test-Path $venvPy) { $venvPy } else { "python" }

& $py -m pip install --upgrade pip
if (Test-Path "requirements.txt") {
    & $py -m pip install -r requirements.txt
}

& $py -m streamlit run app.py --server.headless true --browser.gatherUsageStats false
