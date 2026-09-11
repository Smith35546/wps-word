$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$python = Join-Path $projectRoot '.venv\\Scripts\\python.exe'
if (-not (Test-Path $python)) {
    throw 'Dependencies are not installed. Run .\\setup.ps1 first.'
}
& $python (Join-Path $projectRoot 'app\\main.py')
