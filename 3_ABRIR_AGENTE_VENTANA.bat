@echo off
rem Abre la version grafica del agente (ventana), sin consola.
set "PYDIR=%~dp0python"
cd /d "%~dp0programa"
if exist "%PYDIR%\pythonw.exe" goto incluido
where pyw >nul 2>nul
if %errorlevel%==0 (start "" pyw -3 ventana.pyw) else (start "" pythonw ventana.pyw)
goto :eof
:incluido
start "" "%PYDIR%\pythonw.exe" ventana.pyw
