@echo off
cd /d "%~dp0"
set "analysis_state=plane_strain"
if not "%~1"=="" set "analysis_state=%~1"
python run_bem.py examples/two_materials_2d_order2_%analysis_state%.json --out output_2d
set "analysis_exit=%errorlevel%"
if "%analysis_exit%"=="2" echo Accuracy unverified. Available results saved.
pause
exit /b %analysis_exit%
