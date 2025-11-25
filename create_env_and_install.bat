@echo off
setlocal
REM Create venv
if not exist venv (
    echo Creating virtual environment...
    python -m venv venv
)
echo Activating virtual environment...
call venv\Scripts\activate.bat
echo Installing Python packages...
pip install --upgrade pip
pip install -r requirements.txt
echo Installing Playwright browser (Chromium)...
python -m playwright install chromium
echo Done. You can now run run_app.bat to start the Customer Letter Generator.
pause
