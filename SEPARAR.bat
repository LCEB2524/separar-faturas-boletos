@echo off
chcp 65001 >nul
cd /d "%~dp0"
python separar_faturas_boletos.py
echo.
pause
