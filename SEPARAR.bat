@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PY=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
if not exist "%PY%" set "PY=python"
echo Usando: %PY%
"%PY%" -m pip install --quiet --disable-pip-version-check pypdf openpyxl
"%PY%" separar_faturas_boletos.py
echo.
pause
