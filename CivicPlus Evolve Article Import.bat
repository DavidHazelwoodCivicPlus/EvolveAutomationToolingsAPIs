@echo off
cd /d "%~dp0"
taskkill /F /IM node.exe /T >nul 2>&1
start "Node Server" node.exe index.js
timeout /t 2 >nul
start http://localhost:4000
pause