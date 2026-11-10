@echo off
setlocal
chcp 65001 >nul
echo Criando atalho na Area de Trabalho...

set "CURRENT_DIR=%~dp0"
set "BAT_PATH=%CURRENT_DIR%Gerar_PDF.bat"
set "ICO_PATH=%CURRENT_DIR%app.ico"

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
  "$wsh = New-Object -ComObject WScript.Shell; " ^
  "$desktop = [System.Environment]::GetFolderPath('Desktop'); " ^
  "$shortcutPath = Join-Path $desktop 'Gerar PDF de Resumos.lnk'; " ^
  "$s = $wsh.CreateShortcut($shortcutPath); " ^
  "$s.TargetPath = $env:BAT_PATH; " ^
  "$s.WorkingDirectory = (Split-Path -Parent $env:BAT_PATH); " ^
  "$s.Description = 'Arraste uma pasta ou imagens aqui para gerar o PDF'; " ^
  "if (Test-Path -LiteralPath $env:ICO_PATH) { $s.IconLocation = $env:ICO_PATH } else { $s.IconLocation = 'shell32.dll,265' }; " ^
  "$s.WindowStyle = 7; " ^
  "$s.Save(); " ^
  "Write-Host 'Pronto! O atalho foi criado com sucesso em:' $shortcutPath"

echo.
pause
