@echo off
setlocal
powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0gerar_pdf.ps1" %*
