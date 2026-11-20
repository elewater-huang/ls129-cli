@echo off
setlocal
cd /d "%~dp0"
echo [1/3] Protocol self-test (no LS129 required)
LS129CLI.exe selftest
if errorlevel 1 goto failed
echo.
echo [2/3] Read LS129: %CD%\LS129.ini
LS129CLI.exe status
if errorlevel 1 goto failed
echo.
echo [3/3] Success. You can now try: LS129CLI.exe record --count 100 --out data.csv
goto end
:failed
echo.
echo CHECK FAILED. See the error above and README.md.
:end
pause
