@echo off
chcp 65001 >nul
echo ============================================
echo   PPT Auto Generator - Setup
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 goto :nopython

echo [1/3] Python OK
python --version
goto :makevenv

:nopython
echo [Error] Python is not installed.
echo   Please install Python 3.10+ from python.org
echo   Check "Add Python to PATH" during install.
pause
exit /b 1

:makevenv
echo [2/3] Creating virtual environment...
if not exist "venv" python -m venv venv

echo [3/3] Installing packages...
call venv\Scripts\activate.bat
pip install -r requirements.txt --quiet

echo.
echo ============================================
echo   Setup complete!
echo   Drag your .docx file onto run.bat
echo ============================================
pause
