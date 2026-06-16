@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt
python -m PyInstaller --noconsole --onedir --name ExcelPPTGenerator --clean -y app.py
pause
