@echo off
setlocal
if not exist venv (
    echo Virtual environment not found. Run create_env_and_install.bat first.
    pause
    exit /b 1
)
call venv\Scripts\activate.bat
python ui_app.py
