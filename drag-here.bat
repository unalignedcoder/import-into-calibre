@echo off
powershell.exe -ExecutionPolicy Bypass -File "%~dp0ImportToCalibre.ps1" %*
pause
