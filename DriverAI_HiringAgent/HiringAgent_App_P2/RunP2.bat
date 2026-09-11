@echo off
rem ============================================================================
rem  Interactive Command Prompt launcher for a full live P2 scoring pass.
rem
rem  - Works from any current directory.
rem  - Repairs the local environment when needed.
rem  - Processes up to 500 eligible rows so a desktop/CMD pass drains the normal
rem    queue instead of silently leaving candidates behind at a small batch cap.
rem  - Extra CLI arguments are forwarded; a later --batch-size argument overrides 500.
rem ============================================================================
setlocal
cd /d "%~dp0"
title DriverAI Hiring Agent P2 - Live Scoring

if not exist ".venv\Scripts\python.exe" goto :repair
".venv\Scripts\python.exe" -c "import sys" >nul 2>&1
if not errorlevel 1 goto :run

:repair
echo [setup] P2 environment is missing or unusable. Repairing it now...
call "%~dp0Launch.bat" --setup-only
if errorlevel 1 goto :fail_setup
if not exist ".venv\Scripts\python.exe" goto :fail_setup
".venv\Scripts\python.exe" -c "import sys" >nul 2>&1
if errorlevel 1 goto :fail_setup

:run
echo [run] Scoring all eligible SharePoint candidates (up to 500 rows)...
".venv\Scripts\python.exe" bot.py --score-sharepoint --batch-size 500 %*
set "P2_EXIT=%ERRORLEVEL%"
if not "%P2_EXIT%"=="0" (
    echo [error] P2 exited with code %P2_EXIT%. Review the messages above and P2_Logs.
)
exit /b %P2_EXIT%

:fail_setup
echo [error] P2 could not create a usable Python environment. Run Launch.bat once and retry.
exit /b 1
