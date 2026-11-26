@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" goto missing
".venv\Scripts\python.exe" -m unittest discover -s tests -v
goto done
:missing
echo 请先安装 Python 并创建 .venv，详见 README.md。
:done
pause
