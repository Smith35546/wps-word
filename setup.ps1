$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$venvPath = Join-Path $projectRoot '.venv'

if (-not (Test-Path (Join-Path $venvPath 'Scripts\\python.exe'))) {
    $python = Get-Command python -ErrorAction SilentlyContinue
    if (-not $python) {
        throw 'Python 3.11 or later was not found. Install Python and run this script again.'
    }
    & $python.Source -m venv $venvPath
}

$venvPython = Join-Path $venvPath 'Scripts\\python.exe'
& $venvPython -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) {
    throw 'Unable to upgrade pip in the project virtual environment.'
}
& $venvPython -m pip install -r (Join-Path $projectRoot 'requirements.txt')
if ($LASTEXITCODE -ne 0) {
    throw 'Unable to install the project dependencies.'
}
Write-Host 'Environment is ready. Run .\\run.ps1 to start the application.' -ForegroundColor Green
