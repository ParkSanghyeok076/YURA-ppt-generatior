@echo off
chcp 65001 >nul
echo ============================================
echo   사내강사용 PPT 자동 생성 도구
echo ============================================
echo.

:: 가상환경 활성화
call venv\Scripts\activate.bat 2>nul

:: 입력 파일 확인
if "%~1"=="" (
    echo 사용법: Word 양식 파일을 이 파일 위에 드래그 앤 드롭 하세요.
    echo.
    echo   또는 명령어: run.bat 강의안.docx
    echo   출력 지정:   run.bat 강의안.docx -o 결과.pptx
    echo.
    set /p INPUT_FILE="Word 파일 경로를 입력하세요: "
) else (
    set INPUT_FILE=%~1
)

if "%INPUT_FILE%"=="" (
    echo [오류] 입력 파일이 지정되지 않았습니다.
    pause
    exit /b 1
)

:: 실행
python main.py "%INPUT_FILE%" %2 %3 %4

echo.
pause
