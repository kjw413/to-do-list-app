@echo off
cd /d "%~dp0"
rem Run without a console window (pythonw)
start "" pythonw "%~dp0todo_app.py"