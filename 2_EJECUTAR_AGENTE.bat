@echo off
chcp 65001 >nul
title Agente Ejecucion Presupuestaria MINAGRI
set "PYDIR=%~dp0python"
cd /d "%~dp0programa"
if exist "%PYDIR%\python.exe" goto incluido
where py >nul 2>nul
if %errorlevel%==0 (py -3 ejecutar.py %*) else (python ejecutar.py %*)
goto :eof
:incluido
"%PYDIR%\python.exe" ejecutar.py %*
