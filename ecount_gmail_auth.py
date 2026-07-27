#!/usr/bin/env python3
"""Gmail OAuth 최초 인증 — 브라우저가 있는 로컬 PC에서 딱 1회만 실행.

서버(VPS)는 헤드리스라 브라우저 로그인을 못 하므로, 이 스크립트는 로컬(맥)에서
실행해서 token.json을 발급받은 뒤 서버 .secrets/ 로 옮기는 용도다.

사용법 (로컬 PC):
  pip3 install google-auth-oauthlib google-api-python-client
  python3 ecount_gmail_auth.py --client-secret client_secret.json --out token.json

실행하면 브라우저가 열리고 povbotpovbot@gmail.com으로 로그인 후 권한 동의하면
--out 경로에 토큰이 저장된다. 이후 그 token.json을 서버의 .secrets/gmail_token.json 으로
scp 등으로 옮기면 된다.
"""
from __future__ import annotations

import argparse

from google_auth_oauthlib.flow import InstalledAppFlow

# 읽기 전용 권한만 요청 (메일 삭제/발송 불가능 — 최소 권한 원칙)
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--client-secret", default="client_secret.json", help="구글 클라우드에서 다운받은 OAuth 클라이언트 JSON")
    ap.add_argument("--out", default="token.json", help="발급된 토큰 저장 경로")
    args = ap.parse_args()

    flow = InstalledAppFlow.from_client_secrets_file(args.client_secret, SCOPES)
    creds = flow.run_local_server(port=0)

    with open(args.out, "w") as f:
        f.write(creds.to_json())
    print(f"[auth] 토큰 저장 완료: {args.out}")
    print("[auth] 이 파일을 서버의 .secrets/gmail_token.json 으로 옮기세요.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
