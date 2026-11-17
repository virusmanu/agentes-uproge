@echo off
chcp 65001 >nul
title Actualizar Agente Ejecucion Presupuestaria MINAGRI
rem Puede arrastrar el .zip de la actualizacion sobre este archivo.
cd /d "%~dp0programa"
where py >nul 2>nul
if %errorlevel%==0 (py -3 actualizar.py %*) else (python actualizar.py %*)
