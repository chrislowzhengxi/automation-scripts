@echo off
setlocal ENABLEDELAYEDEXPANSION

REM --- Resolve project root (folder of this .bat) ---
set "SCRIPT_DIR=%~dp0"
pushd "%SCRIPT_DIR%"

REM --- Prefer repo venv if present ---
set "VENV_PY=%SCRIPT_DIR%venv\Scripts\pythonw.exe"
set "SYS_PY_W=pythonw.exe"
set "SYS_PY=python.exe"

REM --- Choose interpreter (no console) ---
if exist "%VENV_PY%" (
  set "PYEXE=%VENV_PY%"
) else (
  where /q %SYS_PY_W%
  if %ERRORLEVEL%==0 (
    set "PYEXE=%SYS_PY_W%"
  ) else (
    REM Fallback to console python (will open a console if pythonw isn't installed)
    where /q %SYS_PY%
    if %ERRORLEVEL%==0 (
      set "PYEXE=%SYS_PY%"
    ) else (
      echo [ERROR] Python not found. Please install Python 3 or use the repo venv.
      echo Tried: "%VENV_PY%", "%SYS_PY_W%", "%SYS_PY%"
      pause
      popd
      exit /b 1
    )
  )
)

REM --- Run GUI (cwd = repo root so -m imports work) ---
REM If your file lives elsewhere, update the relative path below.
"%PYEXE%" -B "%SCRIPT_DIR%run_gui_fill_updated.py"
popd
exit /b 0
