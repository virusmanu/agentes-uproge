@echo off
chcp 65001 >nul
title Actualizar Agente Ejecucion Presupuestaria MINAGRI
rem Puede arrastrar el .zip de la actualizacion sobre este archivo.
set "PYDIR=%~dp0python"
cd /d "%~dp0programa"
if exist "%PYDIR%\python.exe" goto incluido
where py >nul 2>nul
if %errorlevel%==0 (py -3 actualizar.py %*) else (python actualizar.py %*)
goto :eof
:incluido
"%PYDIR%\python.exe" actualizar.py %*
