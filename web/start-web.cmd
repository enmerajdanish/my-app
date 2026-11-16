@echo off
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 goto fail
if exist node_modules\nodemon\bin\nodemon.js goto run
call npm ci
if errorlevel 1 goto fail
:run
call npm run dev
if errorlevel 1 goto fail
pause
exit /b
:fail
echo Start failed. Check Node.js 24, MongoDB Community Server, and the error above.
pause
exit /b 1
