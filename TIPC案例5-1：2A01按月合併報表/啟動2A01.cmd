@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "TIPC_PY=%USERPROFILE%\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
if exist "%TIPC_PY%" (
  "%TIPC_PY%" -X utf8 server.py
) else (
  py -3 -X utf8 server.py
)
if errorlevel 1 (
  echo 啟動失敗。需要 Python 3 與 openpyxl。可執行：py -3 -m pip install openpyxl
  pause
)
