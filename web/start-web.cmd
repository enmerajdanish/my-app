@echo off
cd /d "%~dp0"
if exist node_modules goto database
call npm ci
if errorlevel 1 goto fail
:database
docker compose up -d --wait
if errorlevel 1 goto fail
call npm run setup
if errorlevel 1 goto fail
start "" http://127.0.0.1:3000
call npm start
pause
exit /b
:fail
echo Setup failed. Check Node.js 24 and Docker Desktop, then retry.
pause
exit /b 1
