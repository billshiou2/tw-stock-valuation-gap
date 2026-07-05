@echo off
setlocal

cd /d "%~dp0"

if not exist output mkdir output

echo Running full Taiwan valuation-gap scan.
echo Cnyes crawl settings: delay=1.0s, retries=2, backoff=2.0s, progress every 25 stocks.
echo Excel output defaults to both full and _lite files unless .env or CLI overrides TW_STOCK_EXCEL_OUTPUT.
echo Output will be written under output\
echo.

py -3.11 src\tw_target_scan.py ^
  --cnyes-delay 1.0 ^
  --cnyes-retries 2 ^
  --cnyes-backoff 2.0 ^
  --cnyes-progress-every 25 ^
  --cnyes-error-stop-after 30 ^
  --cnyes-error-stop-rate 0.5 ^
  --excel-output both

set EXIT_CODE=%ERRORLEVEL%
echo.
if "%EXIT_CODE%"=="0" (
  echo Scan finished. Check the newest output\tw_valuation_gap_*.xlsx and *_lite.xlsx files.
) else (
  echo Scan failed with exit code %EXIT_CODE%.
)

echo.
if /I not "%~1"=="--no-pause" pause
exit /b %EXIT_CODE%
