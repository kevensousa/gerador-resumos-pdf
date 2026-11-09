@echo off
setlocal
echo Criando atalho na Area de Trabalho...
powershell -NoProfile -ExecutionPolicy Bypass -Command "$wsh = New-Object -ComObject WScript.Shell; $desktop = [System.Environment]::GetFolderPath('Desktop'); $s = $wsh.CreateShortcut(\"$desktop\Gerar PDF de Resumos.lnk\"); $s.TargetPath = \"%~dp0Gerar_PDF.bat\"; $s.WorkingDirectory = \"%~dp0\"; $s.Description = 'Arraste uma pasta ou imagens aqui para gerar o PDF'; $s.IconLocation = 'shell32.dll,265'; $s.WindowStyle = 7; $s.Save()"
echo Pronto! O atalho 'Gerar PDF de Resumos' foi criado na sua Area de Trabalho.
pause
