@echo off
title Nexus AI Launcher
cd /d "%~dp0"

:: Activate the virtual environment
if exist venv\Scripts\activate.bat (
    call venv\Scripts\activate.bat
) else if exist .venv\Scripts\activate.bat (
    call .venv\Scripts\activate.bat
)

echo Starting Nexus AI Local Core...
streamlit run app.py
pause