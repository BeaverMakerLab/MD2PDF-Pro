@echo off
chcp 65001 >nul
cd /d "%~dp0"
start "" pythonw "%~dp0md2pdf_pro.py"
exit
