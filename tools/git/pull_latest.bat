@echo off
setlocal
cd /d "%~dp0"
python "%~dp0pull_latest.py" %*
if errorlevel 1 exit /b 1
