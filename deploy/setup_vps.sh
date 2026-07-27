#!/usr/bin/env bash
# 이카운트 대조 시스템 VPS 세팅 스크립트 (Ubuntu 기준)
# 서버에 접속한 뒤 이 스크립트를 실행하면 코드·의존성·인증키 템플릿까지 준비된다.
#
# 사용법:
#   bash deploy/setup_vps.sh
# 또는 아직 코드가 없다면 먼저 clone 후 실행 (아래 README_vps.md 참고).

set -euo pipefail

REPO_DIR="${REPO_DIR:-$HOME/pov-erp-stock-alert}"

echo "==> 1) 시스템 패키지 설치 (python3, pip, venv, git)"
sudo apt-get update -y
sudo apt-get install -y python3 python3-pip python3-venv git

echo "==> 2) 서버 시간대를 한국(Asia/Seoul)으로"
sudo timedatectl set-timezone Asia/Seoul || echo "  (timedatectl 미지원 — 수동 확인 필요)"

echo "==> 3) 코드 위치 확인: $REPO_DIR"
if [ ! -d "$REPO_DIR/.git" ]; then
  echo "  !! $REPO_DIR 에 레포가 없습니다."
  echo "     먼저 git clone 하세요 (비공개 레포라 토큰/배포키 필요 — README_vps.md 참고)."
  exit 1
fi
cd "$REPO_DIR"

echo "==> 4) 파이썬 가상환경 + 의존성"
python3 -m venv .venv
# shellcheck disable=SC1091
. .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt

echo "==> 5) 인증키 템플릿 생성 (.secrets/ecount.json)"
mkdir -p .secrets
if [ ! -f .secrets/ecount.json ]; then
  cat > .secrets/ecount.json <<'JSON'
{
  "COM_CODE": "여기에_회사코드",
  "USER_ID": "여기에_API용_사용자ID",
  "API_CERT_KEY": "여기에_인증키",
  "MODE": "test"
}
JSON
  echo "  → .secrets/ecount.json 을 만들었습니다. 실제 값으로 채우세요."
else
  echo "  → .secrets/ecount.json 이미 있음 (건드리지 않음)."
fi

echo ""
echo "==> 완료. 다음 순서:"
echo "  1. nano .secrets/ecount.json  으로 인증키 채우기"
echo "  2. 이 서버의 고정 IP를 이카운트 [IP등록]에 등록"
echo "  3. .venv/bin/python ecount_probe.py --check   # 인증 검증"
echo "  4. crontab 등록 (deploy/README_vps.md 참고)"
