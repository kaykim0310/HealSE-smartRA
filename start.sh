#!/bin/bash

echo "================================================================"
echo "🛡️  KOSHA KRAS 위험성평가 시스템"
echo "================================================================"
echo

# Python 설치 확인
if ! command -v python3 &> /dev/null; then
    echo "❌ Python3가 설치되지 않았습니다."
    echo
    echo "📥 Python 설치 방법:"
    echo "Ubuntu/Debian: sudo apt update && sudo apt install python3 python3-pip"
    echo "CentOS/RHEL: sudo yum install python3 python3-pip"
    echo "macOS: brew install python3"
    echo
    exit 1
fi

echo "✅ Python 설치 확인됨"
python3 --version

# pip 확인
if ! command -v pip3 &> /dev/null; then
    echo "❌ pip3가 설치되지 않았습니다."
    echo "📥 pip 설치: sudo apt install python3-pip"
    exit 1
fi

# 필요한 패키지 설치
echo
echo "📦 필요한 패키지를 설치하고 있습니다..."
pip3 install flask flask-cors openpyxl

if [ $? -ne 0 ]; then
    echo
    echo "❌ 패키지 설치에 실패했습니다."
    echo "💡 다음 명령어를 수동으로 실행해보세요:"
    echo "pip3 install --user flask flask-cors openpyxl"
    exit 1
fi

echo
echo "✅ 패키지 설치 완료"
echo

# 서버 시작
echo "🚀 KOSHA KRAS 서버를 시작합니다..."
echo
echo "📌 브라우저에서 다음 주소로 접속하세요:"
echo "   http://localhost:5000"
echo
echo "💡 서버를 종료하려면 Ctrl+C를 누르세요."
echo

python3 app.py
