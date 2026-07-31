@echo off
REM DriverAI Hiring Agent - one-click launcher.
REM First run installs everything automatically: Python if missing, then the
REM app environment and its dependencies. After that it just starts the app.
cd /d "%~dp0"
title DriverAI Hiring Agent

if exist ".venv\Scripts\python.exe" goto :deps

REM ---- find a working Python by actually running it, so the fake
REM ---- Microsoft Store alias can never be picked by mistake
set "PY="
for /f "delims=" %%p in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do set "PY=%%p"
if defined PY goto :makevenv
for /f "delims=" %%p in ('python -c "import sys; print(sys.executable)" 2^>nul') do set "PY=%%p"
if defined PY goto :makevenv

echo [setup] Python not found. Installing it automatically - one-time step...
winget install -e --id Python.Python.3.12 --accept-source-agreements --accept-package-agreements
for /f "delims=" %%p in ('py -3 -c "import sys; print(sys.executable)" 2^>nul') do set "PY=%%p"
if defined PY goto :makevenv
if exist "%LocalAppData%\Programs\Python\Python312\python.exe" set "PY=%LocalAppData%\Programs\Python\Python312\python.exe"
if defined PY goto :makevenv
goto :fail_python

:makevenv
echo [setup] First run - creating the app environment...
"%PY%" -m venv .venv
if errorlevel 1 goto :fail_venv

:deps
set "VPY=%~dp0.venv\Scripts\python.exe"
"%VPY%" -c "import yaml, pandas, openpyxl, pypdf, requests, dotenv, customtkinter" >nul 2>&1
if not errorlevel 1 goto :ocrlibs
echo [setup] Installing dependencies - about 1 minute, one time only...
"%VPY%" -m pip install --upgrade pip -q
"%VPY%" -m pip install -r requirements-desktop.txt -q
if errorlevel 1 goto :fail_deps

:ocrlibs
REM OCR python libs (pymupdf/pytesseract/pillow). Best-effort: never block the app if a
REM wheel is unavailable for this Python - OCR just stays off until it can be installed.
"%VPY%" -c "import fitz, pytesseract, PIL" >nul 2>&1
if not errorlevel 1 goto :tesseract
echo [setup] Adding OCR libraries for scanned PDFs...
"%VPY%" -m pip install pymupdf pytesseract pillow -q

:tesseract
REM Tesseract BINARY - needed to actually read scanned / image-only PDFs. Idempotent:
REM skip if already on PATH or in the standard install folder.
where tesseract >nul 2>&1 && goto :ollama
if exist "%ProgramFiles%\Tesseract-OCR\tesseract.exe" goto :ollama
echo [setup] Installing Tesseract OCR engine (scanned-PDF support) - one time...
winget install -e --id UB-Mannheim.TesseractOCR --accept-source-agreements --accept-package-agreements

:ollama
REM Local AI is optional - the app scores fully without it (built-in keyword engine).
REM Auto-install Ollama + pull its model on first run too, so after this ONE setup pass
REM (which needs internet) the app's AI layer also works fully offline. Best-effort at every
REM step - a failure here just logs and falls through; the app never fails to start over it.
set "OLLAMA_EXE="
where ollama >nul 2>&1 && set "OLLAMA_EXE=ollama"
if not defined OLLAMA_EXE if exist "%LocalAppData%\Programs\Ollama\ollama.exe" set "OLLAMA_EXE=%LocalAppData%\Programs\Ollama\ollama.exe"
if defined OLLAMA_EXE goto :ollama_model
echo [setup] Installing Ollama (local AI engine) - one time, needs internet...
winget install -e --id Ollama.Ollama --accept-source-agreements --accept-package-agreements
where ollama >nul 2>&1 && set "OLLAMA_EXE=ollama"
if not defined OLLAMA_EXE if exist "%LocalAppData%\Programs\Ollama\ollama.exe" set "OLLAMA_EXE=%LocalAppData%\Programs\Ollama\ollama.exe"
if not defined OLLAMA_EXE (echo [info] Ollama install skipped/failed - app scores fully with its built-in engine. Optional local AI: https://ollama.com & goto :run)

:ollama_model
for /f "delims=" %%m in ('"%VPY%" -c "from hiring_agent.config import OLLAMA_MODEL; print(OLLAMA_MODEL)" 2^>nul') do set "OLLAMA_MODEL_NAME=%%m"
if not defined OLLAMA_MODEL_NAME set "OLLAMA_MODEL_NAME=llama3.2"
"%OLLAMA_EXE%" list 2>nul | findstr /i /c:"%OLLAMA_MODEL_NAME%" >nul
if not errorlevel 1 goto :ollama_ready
echo [setup] Downloading local AI model '%OLLAMA_MODEL_NAME%' - one time, a few GB, needs internet...
"%OLLAMA_EXE%" pull %OLLAMA_MODEL_NAME%
if errorlevel 1 (echo [info] Model download failed/skipped - app scores fully with its built-in engine. Retry later with: ollama pull %OLLAMA_MODEL_NAME% & goto :run)

:ollama_ready
echo [ok] Ollama ready with '%OLLAMA_MODEL_NAME%' - local AI extraction/scoring enabled, works offline from here on.

:run
echo [ok] Starting DriverAI Hiring Agent...
"%VPY%" app.py %*
if errorlevel 1 goto :fail_app
exit /b 0

:fail_python
echo.
echo [ERROR] Could not install Python automatically.
echo         Install it from https://www.python.org/downloads/
echo         Tick "Add python.exe to PATH" during install, then run this again.
start https://www.python.org/downloads/
pause
exit /b 1

:fail_venv
echo.
echo [ERROR] Could not create the app environment. Delete the .venv folder and retry.
pause
exit /b 1

:fail_deps
echo.
echo [ERROR] Dependency install failed - check your internet connection and retry.
pause
exit /b 1

:fail_app
echo.
echo [ERROR] The app closed with an error - see the messages above.
pause
exit /b 1
