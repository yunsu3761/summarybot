@echo off
chcp 65001 >nul
echo ==========================================================
echo   후보기술 개요서 자동 요약 봇 - 시작 스크립트 (Windows용)
echo ==========================================================
echo.

:: 1. 파이썬 설치 확인
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [오류] Python이 설치되어 있지 않거나 환경 변수(PATH)에 등록되어 있지 않습니다.
    echo 파이썬 공식 홈페이지(https://www.python.org/)에서 설치해주세요.
    pause
    exit /b
)

:: 2. 가상환경 세팅
if not exist ".venv\Scripts\activate.bat" (
    echo [안내] 파이썬 가상환경(.venv)을 생성합니다. 잠시만 기다려주세요...
    python -m venv .venv
)

:: 3. 가상환경 활성화
call .venv\Scripts\activate.bat

:: 4. 필요한 패키지 설치
echo [안내] 필수 패키지를 설치합니다. 이 작업은 처음 실행 시 시간이 걸릴 수 있습니다...
pip install -r backend\requirements.txt

:: 5. 서버 실행 (브라우저 자동 오픈 포함)
echo.
echo ==========================================================
echo   준비가 완료되었습니다! 서버를 실행합니다...
echo   (이 창을 닫으면 요약 봇이 종료됩니다.)
echo ==========================================================
echo.
python backend\server.py

pause
