@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" goto local
python app.py %*
goto done
:local
".venv\Scripts\python.exe" app.py %*
:done
if errorlevel 1 pause
