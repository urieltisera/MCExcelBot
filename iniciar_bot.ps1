$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $projectRoot

$activateScript = Join-Path $projectRoot ".venv\Scripts\Activate.ps1"
if (-not (Test-Path -LiteralPath $activateScript)) {
    throw "No se encontró .venv. Creá el entorno virtual antes de iniciar el bot."
}

& $activateScript
python scripts\run_bot.py
