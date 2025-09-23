@echo off
chcp 65001 > nul
title KOSHA KRAS 위험성평가 시스템

echo.
echo ================================================================
echo 🛡️  KOSHA KRAS 위험성평가 시스템
echo ================================================================
echo.

REM Python 설치 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python이 설치되지 않았습니다.
    echo.
    echo 📥 Python 설치 방법:
    echo 1. https://python.org 에서 Python 3.8 이상 다운로드
    echo 2. 설치 시 "Add Python to PATH" 체크박스 선택
    echo 3. 설치 완료 후 컴퓨터 재시작
    echo.
    pause
    exit /b 1
)

echo ✅ Python 설치 확인됨
python --version

REM 필요한 패키지 설치
echo.
echo 📦 필요한 패키지를 설치하고 있습니다...
pip install flask flask-cors openpyxl

if errorlevel 1 (
    echo.
    echo ❌ 패키지 설치에 실패했습니다.
    echo 💡 인터넷 연결을 확인하고 다시 시도해주세요.
    pause
    exit /b 1
)

echo.
echo ✅ 패키지 설치 완료
echo.

REM 서버 시작
echo 🚀 KOSHA KRAS 서버를 시작합니다...
echo.
echo 📌 브라우저에서 다음 주소로 접속하세요:
echo    http://localhost:5000
echo.
echo 💡 서버를 종료하려면 Ctrl+C를 누르세요.
echo.

python app.py

pause
