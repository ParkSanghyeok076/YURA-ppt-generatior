@echo off
chcp 65001 >nul
echo ============================================
echo   사내강사용 PPT 자동 생성 도구 - 설치
echo ============================================
echo.

:: Python 설치 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo [오류] Python이 설치되어 있지 않습니다.
    echo   https://www.python.org/downloads/ 에서 Python 3.10 이상을 설치하세요.
    echo   설치 시 "Add Python to PATH" 체크를 반드시 해주세요.
    pause
    exit /b 1
)

echo [1/3] Python 확인 완료
python --version

:: 가상환경 생성
echo [2/3] 가상환경 생성 중...
if not exist "venv" (
    python -m venv venv
)

:: 패키지 설치
echo [3/3] 필수 패키지 설치 중...
call venv\Scripts\activate.bat
pip install -r requirements.txt --quiet

echo.
echo ============================================
echo   설치 완료!
echo   run.bat 을 더블클릭하여 실행하세요.
echo ============================================
pause
