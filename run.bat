@echo off
chcp 65001 >nul
echo ============================================
echo   PPT Auto Generator
echo ============================================
echo.

if not exist "venv" goto :novenv
call venv\Scripts\activate.bat 2>nul
goto :checkargs

:novenv
echo [Error] Virtual environment not found.
echo   Run setup.bat first.
pause
exit /b 1

:checkargs
if "%~1"=="" goto :noinput
set INPUT_FILE=%~1
goto :run

:noinput
set /p INPUT_FILE="Enter .docx file path: "
if "%INPUT_FILE%"=="" goto :error

:run
python main.py "%INPUT_FILE%" %2 %3 %4
echo.
pause
exit /b 0

:error
echo [Error] No input file specified.
pause
exit /b 1
