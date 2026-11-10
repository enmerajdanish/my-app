@echo off
cd /d "%~dp0"
if exist node_modules goto launch
call npm ci
if errorlevel 1 goto fail
:launch
call npm start
pause
exit /b
:fail
echo Setup failed. Check Node.js 24 and internet access, then retry.
pause
exit /b 1
