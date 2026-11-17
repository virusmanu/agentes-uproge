@echo off
chcp 65001 >nul
title Instalar - Agente Ejecucion MINAGRI
echo Instalando componentes de Python (pandas, openpyxl, xlrd)...
py -3 -m pip install --user --upgrade pandas openpyxl xlrd
if errorlevel 1 python -m pip install --user --upgrade pandas openpyxl xlrd
echo.
echo Listo. Ya puede usar EJECUTAR_AGENTE.bat
pause
