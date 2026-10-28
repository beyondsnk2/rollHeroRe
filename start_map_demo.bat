@echo off
rem Launch the demo server (serves web/index.html at http://127.0.0.1:8765/,
rem and /api/refresh re-exports the latest config from config_v1.xlsx)
set PY=C:\Python312\python.exe
if not exist "%PY%" set PY=python
start "rollHeroRe demo server" "%PY%" "%~dp0work\demo_server.py"
timeout /t 1 /nobreak >nul
start "" "http://127.0.0.1:8765/"
