@echo off
chcp 65001 >nul

if not exist "venv" goto :novenv
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

:novenv
echo [Error] Virtual environment not found.
echo   Run setup.bat first.
pause
exit /b 1
