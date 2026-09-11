@echo off
rem ============================================================================
rem  Daily SharePoint scoring run - scheduled at 9:00 AM by the Windows Task
rem  Scheduler task "HiringAgent P2 Daily 9AM" (safe to double-click manually
rem  too; it is the exact same run the P2 app's button starts).
rem
rem  Crash/shutdown safe: the queue lives in SharePoint (Status column), and
rem  decline emails are driven by the 'Decline Sent' marker column - so an
rem  interrupted run simply continues where it left off on the next 9 AM run
rem  or the next manual run. Nothing is processed or emailed twice.
rem ============================================================================
cd /d "%~dp0"
rem  The venv must actually RUN, not merely exist: a .venv copied from another machine
rem  keeps python.exe on disk while pyvenv.cfg still points at the original interpreter,
rem  so every call dies with "No Python at ...". Unattended, that failure lands only in
rem  task_console.log and the 9 AM run silently stops happening. Validate, and let
rem  Launch.bat rebuild the environment before giving up.
if not exist ".venv\Scripts\python.exe" goto :repair
".venv\Scripts\python.exe" -c "import sys" >nul 2>&1
if not errorlevel 1 goto :run
:repair
echo [%DATE% %TIME%] venv unusable - rebuilding via Launch.bat >> "P2_Logs\task_console.log"
call "%~dp0Launch.bat" --setup-only >> "P2_Logs\task_console.log" 2>&1
if not exist ".venv\Scripts\python.exe" (
    echo [%DATE% %TIME%] FATAL: could not rebuild the environment; run aborted >> "P2_Logs\task_console.log"
    exit /b 1
)
:run
".venv\Scripts\python.exe" bot.py --score-sharepoint >> "P2_Logs\task_console.log" 2>&1
