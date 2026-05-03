#!/bin/bash
cd "$(dirname "$0")"

echo "=========================================================="
echo "  후보기술 개요서 자동 요약 봇 - 시작 스크립트 (Mac용)"
echo "=========================================================="
echo ""

# 1. 파이썬 설치 확인
if ! command -v python3 &> /dev/null; then
    echo "[오류] Python3가 설치되어 있지 않습니다."
    exit 1
fi

# 2. 가상환경 세팅
if [ ! -f ".venv/bin/activate" ]; then
    echo "[안내] 파이썬 가상환경(.venv)을 생성합니다. 잠시만 기다려주세요..."
    python3 -m venv .venv
fi

# 3. 가상환경 활성화
source .venv/bin/activate

# 4. 필요한 패키지 설치
echo "[안내] 필수 패키지를 설치합니다. 이 작업은 처음 실행 시 시간이 걸릴 수 있습니다..."
pip install -r backend/requirements.txt

# 5. 서버 실행
echo ""
echo "=========================================================="
echo "  준비가 완료되었습니다! 서버를 실행합니다..."
echo "  (터미널 창을 닫으면 요약 봇이 종료됩니다.)"
echo "=========================================================="
echo ""
python backend/server.py
