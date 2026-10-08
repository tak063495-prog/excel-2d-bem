@echo off
cd /d "%~dp0"
python verify.py
set "analysis_exit=%errorlevel%"
pause
exit /b %analysis_exit%
