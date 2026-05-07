@echo off
REM ============================================================
REM Day 03 -- Budget vs Actual Tracker, local launcher
REM
REM Day-N port convention:  port = 1000 + N
REM   Day 03 -> 1003
REM
REM Double-click this file to start the tracker at
REM http://127.0.0.1:1003/ in your default browser.
REM
REM   start.bat            -- default port 1003
REM   start.bat 1103       -- custom port (override)
REM
REM Stops when you close this window or press Ctrl+C.
REM ============================================================

setlocal
cd /d "%~dp0"

REM --- Sanity checks ---------------------------------------------------------
if not exist "..\.venv\Scripts\python.exe" (
    echo.
    echo ERROR: virtual environment not found at ..\.venv\
    echo.
    echo First-time setup ^(run these once from the project root^):
    echo   python -m venv .venv
    echo   .venv\Scripts\python.exe -m pip install -r requirements.txt
    echo   copy .env.example .env
    echo   notepad .env             ^(then paste your ANTHROPIC_API_KEY^)
    echo.
    pause
    exit /b 1
)

if not exist "server.py" (
    echo.
    echo ERROR: server.py not found in this folder.
    echo Make sure start.bat is sitting next to server.py inside day-03-budget-tracker\.
    echo.
    pause
    exit /b 1
)

if not exist "..\.env" (
    echo NOTICE: ..\.env not found. AI commentary will fail until you create it.
    echo Run from project root:  copy .env.example .env  ^&^&  notepad .env
    echo.
)

REM --- Resolve port (arg 1 or default 1003 for Day 03) ---------------------
set "PORT=%~1"
if "%PORT%"=="" set "PORT=1003"

REM --- Detect port-in-use and pick an alternate -----------------------------
REM Fallback ports for Day 03 are 1103, 1203, 1303 (keeps day-distinct).
powershell -NoProfile -Command "if (Get-NetTCPConnection -LocalPort %PORT% -State Listen -ErrorAction SilentlyContinue) { exit 1 } else { exit 0 }"
if errorlevel 1 (
    echo NOTICE: port %PORT% is already in use. Trying 1103, 1203, 1303 in turn...
    for %%P in (1103 1203 1303) do (
        powershell -NoProfile -Command "if (Get-NetTCPConnection -LocalPort %%P -State Listen -ErrorAction SilentlyContinue) { exit 1 } else { exit 0 }"
        if not errorlevel 1 (
            set "PORT=%%P"
            goto :got_port
        )
    )
    echo ERROR: no fallback port free. Pass an explicit port:  start.bat 1234
    pause
    exit /b 1
)
:got_port

REM --- Open the browser after the server has had time to bind --------------
start "" /b powershell -NoProfile -WindowStyle Hidden -Command ^
    "Start-Sleep -Seconds 2; Start-Process 'http://127.0.0.1:%PORT%/'"

REM --- Run the Flask server in the foreground (Ctrl+C to stop) -------------
echo.
echo Starting Day 03 - Budget vs Actual Tracker on port %PORT% ...
echo Local URL:  http://127.0.0.1:%PORT%/
echo Press Ctrl+C or close this window to stop.
echo.

set "PYTHONIOENCODING=utf-8"
"..\.venv\Scripts\python.exe" "server.py" --port %PORT%

endlocal
