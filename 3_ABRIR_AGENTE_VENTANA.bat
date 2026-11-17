@echo off
rem Abre la version grafica del agente (ventana), sin consola.
cd /d "%~dp0programa"
where pyw >nul 2>nul
if %errorlevel%==0 (start "" pyw -3 ventana.pyw) else (start "" pythonw ventana.pyw)
