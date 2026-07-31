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
".venv\Scripts\python.exe" bot.py --score-sharepoint >> "P2_Logs\task_console.log" 2>&1
