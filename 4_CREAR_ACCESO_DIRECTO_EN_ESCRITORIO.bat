@echo off
chcp 65001 >nul
title Crear acceso directo - Agente Ejecucion MINAGRI
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0programa\crear_acceso.ps1"
echo.
pause
