@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo First run: installing libraries, please wait...
  py -3 -m venv .venv 2>nul || python -m venv .venv
  if not exist ".venv\Scripts\python.exe" (
    echo Python is not installed. Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH".
    pause
    exit /b 1
  )
  ".venv\Scripts\python.exe" -m pip install --upgrade pip -q
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt -q
)
".venv\Scripts\python.exe" pdf_workshop.py
pause
