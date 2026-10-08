@echo off
cd /d "%~dp0"
if "%~1"=="" goto missing
python run_bem.py "%~1" --out output_model
set "analysis_exit=%errorlevel%"
pause
exit /b %analysis_exit%
:missing
echo Usage: run_model.bat model.json
pause
exit /b 1
