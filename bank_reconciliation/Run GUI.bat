@echo off
setlocal

REM Locate this script directory
set "HERE=%~dp0"

REM Activate your virtual environment
call "C:\Users\TP2507088\Downloads\Automation\venv\Scripts\activate.bat"
if errorlevel 1 (
  echo Failed to activate venv. Falling back to system Python...
) else (
  echo Activated venv: %VIRTUAL_ENV%
)

REM Run the GUI with whichever Python is active now
python "%HERE%run_gui.py"

endlocal
