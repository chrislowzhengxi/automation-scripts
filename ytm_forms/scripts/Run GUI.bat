@echo off
set PYTHON=%~dp0..\..\venv\Scripts\python.exe
if not exist "%PYTHON%" set PYTHON=py
"%PYTHON%" "%~dp0run_gui.py"
pause
