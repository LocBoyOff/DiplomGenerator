$ErrorActionPreference = 'Stop'
Set-Location -LiteralPath $PSScriptRoot
$buildPython = Join-Path $PSScriptRoot '.venv-build\Scripts\python.exe'
if (-not (Test-Path -LiteralPath $buildPython)) {
    python -m venv .venv-build
    if ($LASTEXITCODE -ne 0) { throw 'Не удалось создать окружение сборки.' }
}
& $buildPython -m pip install -r requirements-build.txt
if ($LASTEXITCODE -ne 0) { throw 'Не удалось установить зависимости сборки.' }
& $buildPython -m PyInstaller --noconfirm DiplomaGenerator.spec
if ($LASTEXITCODE -ne 0) { throw 'Сборка EXE завершилась ошибкой.' }
Write-Host 'Готово: dist\DiplomaGenerator.exe'
