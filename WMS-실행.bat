@echo off
cd /d "%~dp0"
title WMS
echo Open in browser: http://localhost:3000
echo Keep this window open. Close it to stop the server.
node server.js
pause
