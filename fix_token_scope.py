#!/usr/bin/env python3
"""손상된 gmail_token.json의 권한(scope)을 복구하는 1회성 스크립트.

배경(2026-08-02): ecount_gmail_fetch.py가 자기만의 좁은 SCOPES(gmail.readonly만)로
공유 토큰 파일을 새로고침(refresh)해서, 스프레드시트 권한이 필요할 때 403
insufficient scope로 실패하는 버그가 있었다(ecount_gmail_fetch.py 수정으로 재발은
막았지만, 이미 손상된 기존 토큰 파일 자체는 이 스크립트로 한 번 복구해야 한다).

Credentials.from_authorized_user_file()은 파일에 저장된 "scopes" 값을 우선시해서
코드에서 아무리 넓은 scopes를 넘겨도 무시되므로, 여기서는 그 헬퍼를 안 쓰고
token/refresh_token만 파일에서 읽어 Credentials를 직접 만들고 원하는 전체 권한으로
강제 새로고침한다.

사용법 (서버에서):
  .venv/bin/python fix_token_scope.py
"""
from __future__ import annotations

import json
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

from ecount_sheets_setup import SCOPES, _TOKEN_FILE


def main() -> int:
    if not _TOKEN_FILE.exists():
        print(f"토큰 파일이 없습니다: {_TOKEN_FILE}")
        return 2

    info = json.loads(_TOKEN_FILE.read_text())
    print(f"[fix] 복구 전 scopes: {info.get('scopes')}")

    creds = Credentials(
        token=None,  # 무조건 새로고침하도록 비워둠
        refresh_token=info["refresh_token"],
        token_uri=info.get("token_uri", "https://oauth2.googleapis.com/token"),
        client_id=info["client_id"],
        client_secret=info["client_secret"],
        scopes=SCOPES,
    )
    creds.refresh(Request())
    _TOKEN_FILE.write_text(creds.to_json())

    new_info = json.loads(_TOKEN_FILE.read_text())
    print(f"[fix] 복구 후 scopes: {new_info.get('scopes')}")
    if set(new_info.get("scopes", [])) >= set(SCOPES):
        print("[fix] 정상 복구됨 — 필요한 권한이 모두 포함되어 있습니다.")
        return 0
    print("[fix] 여전히 권한이 부족합니다 — 브라우저로 재인증(ecount_gmail_auth.py)이 필요할 수 있습니다.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
