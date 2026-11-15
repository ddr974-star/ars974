@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist node_modules (
  echo Installation des outils, une seule fois...
  call npm install --no-audit --no-fund
)
node exporter.js
echo.
pause
