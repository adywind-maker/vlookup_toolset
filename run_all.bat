@echo off
cd /d "%~dp0"

echo ============================================
echo  Excel Tools - Run All
echo ============================================
echo.

echo [1/3] txt2xlsx - convert text to Excel...
python excel_tools.py txt2xlsx
if errorlevel 1 ( echo [ERROR] txt2xlsx failed. & pause & exit /b 1 )
echo.

echo [2/3] vlookup - merge data...
python excel_tools.py vlookup --src1 test1_converted.xlsx --src2 test2.ods --output output_vlookup.xlsx
if errorlevel 1 ( echo [ERROR] vlookup failed. & pause & exit /b 1 )
echo.

echo [3/3] extract - filter columns...
python excel_tools.py extract --src test1_converted.xlsx --columns "ID,Net Pay" --output my_output.xlsx
if errorlevel 1 ( echo [ERROR] extract failed. & pause & exit /b 1 )
echo.

echo ============================================
echo  All done!
echo ============================================
pause
