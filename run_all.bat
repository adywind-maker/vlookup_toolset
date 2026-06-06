@echo off
cd /d "%~dp0"

:: ============================================================
::  設定區 - 只需修改這裡
:: ============================================================

:: [1] txt2xlsx：要轉換的文字檔資料夾（預設為腳本所在位置）
set TXT_FOLDER=C:\Users\ASUS\repo\vlookup

:: [2] vlookup：來源檔案與輸出路徑
set VLOOKUP_SRC1=test1_converted.xlsx
set VLOOKUP_SRC2=test2.ods
set VLOOKUP_OUTPUT=output_vlookup.xlsx

:: [3] extract：來源檔案、要篩選的欄位（逗號分隔）、輸出路徑
set EXTRACT_SRC=test1_converted.xlsx
set EXTRACT_COLUMNS=ID,Net Pay
set EXTRACT_OUTPUT=my_output.xlsx

:: ============================================================
::  執行區 - 以下不需修改
:: ============================================================

echo ============================================
echo  Excel Tools - Run All
echo ============================================
echo.

echo [1/3] txt2xlsx - convert text to Excel...
python excel_tools.py txt2xlsx --folder "%TXT_FOLDER%"
if errorlevel 1 ( echo [ERROR] txt2xlsx failed. & pause & exit /b 1 )
echo.

echo [2/3] vlookup - merge data...
python excel_tools.py vlookup --src1 "%VLOOKUP_SRC1%" --src2 "%VLOOKUP_SRC2%" --output "%VLOOKUP_OUTPUT%"
if errorlevel 1 ( echo [ERROR] vlookup failed. & pause & exit /b 1 )
echo.

echo [3/3] extract - filter columns...
python excel_tools.py extract --src "%EXTRACT_SRC%" --columns "%EXTRACT_COLUMNS%" --output "%EXTRACT_OUTPUT%"
if errorlevel 1 ( echo [ERROR] extract failed. & pause & exit /b 1 )
echo.

echo ============================================
echo  All done!
echo ============================================
pause
