@echo off
setlocal EnableExtensions

set "HERE=%~dp0"
pushd "%HERE%"

set "MERGER=%HERE%merge_excels.py"
set "VENV_PYW=%HERE%..\venv\Scripts\pythonw.exe"
set "VENV_PY=%HERE%..\venv\Scripts\python.exe"

if not exist "%MERGER%" (
  echo merge_excels.py not found at: %MERGER%
  pause
  exit /b 1
)

REM Prefer pythonw (no console), else python.exe
if exist "%VENV_PYW%" (
  "%VENV_PYW%" "%MERGER%" --gui
  goto :done
)

if exist "%VENV_PY%" (
  "%VENV_PY%" "%MERGER%" --gui
  goto :done
)

where py >nul 2>nul && ( py -3 "%MERGER%" --gui & goto :done )
where python >nul 2>nul && ( python "%MERGER%" --gui & goto :done )

echo Could not find Python. Please set up ..\venv or install Python.
pause

:done
popd
endlocal
