@echo off
chcp 65001 >nul
title EasyPDFConverter
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0EasyPDFConverter.ps1" -NoPause %*
echo.
pause
