@echo off
chcp 65001 >nul

if exist "venv" goto :start

echo ============================================
echo   First-time setup...
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 goto :nopython

echo [1/3] Python OK
python --version
echo [2/3] Creating virtual environment...
python -m venv venv
echo [3/3] Installing packages...
call venv\Scripts\activate.bat
pip install -r requirements.txt --quiet
echo.
echo   Setup complete!
echo.

:start
call venv\Scripts\activate.bat 2>nul
echo ============================================
echo   PPT Auto Generator - Starting...
echo   Browser will open automatically.
echo   Do NOT close this window.
echo ============================================
echo.
python app.py
pause
exit /b 0

:nopython
echo [Error] Python is not installed.
echo   Please install Python 3.10+ from python.org
echo   Check "Add Python to PATH" during install.
pause
exit /b 1
