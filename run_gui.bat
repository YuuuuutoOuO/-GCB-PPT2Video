@echo off
python gui.py
if errorlevel 1 (
    echo.
    echo [ERROR] Failed to start GUI.
    pause
)
