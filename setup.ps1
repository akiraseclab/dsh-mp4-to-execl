$ErrorActionPreference = 'Stop'
$skill = $PSScriptRoot
if (-not (Get-Command python -ErrorAction SilentlyContinue)) { throw 'Install Python 3.10+ and add it to PATH.' }
if (-not (Get-Command ffmpeg -ErrorAction SilentlyContinue)) { throw 'Install FFmpeg and add it to PATH.' }
if (-not (Get-Command ffprobe -ErrorAction SilentlyContinue)) { throw 'Install ffprobe and add it to PATH.' }
python -m pip install -r (Join-Path $skill 'requirements.txt')
if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed.' }
python (Join-Path $skill 'scripts\build_template.py')
if ($LASTEXITCODE -ne 0) { throw 'Template generation failed.' }
Write-Host "Ready. Confirm this directory is in your DSH skills configuration: $skill"
