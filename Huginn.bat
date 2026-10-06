@echo off
chcp 65001 >nul
setlocal
title HUGINN - Gerador de Massa de Testes

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Huginn_Gerador.ps1"

pause
endlocal
