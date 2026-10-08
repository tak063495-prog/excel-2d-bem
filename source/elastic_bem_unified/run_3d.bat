@echo off
cd /d "%~dp0"
python run_bem.py examples/two_materials_3d_order2_three_dimensional.json --out output_3d
set "analysis_exit=%errorlevel%"
if "%analysis_exit%"=="2" echo Accuracy unverified. Available results saved.
pause
exit /b %analysis_exit%
