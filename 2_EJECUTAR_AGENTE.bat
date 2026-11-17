@echo off
chcp 65001 >nul
title Agente Ejecucion Presupuestaria MINAGRI
cd /d "%~dp0programa"
where py >nul 2>nul
if %errorlevel%==0 (py -3 ejecutar.py %*) else (python ejecutar.py %*)
