#!/usr/bin/env python3
"""이카운트 "판매현황" 자동알림 이메일에서 첨부 엑셀을 받아오는 서버용 스크립트.

로컬(맥)에서 ecount_gmail_auth.py로 발급받은 token.json을 서버의
.secrets/gmail_token.json 에 둔 뒤 이 스크립트를 실행한다. 서버는 브라우저가
없어도 refresh_token으로 계속 갱신되므로 이후엔 완전자동으로 동작한다.

아직 이카운트에서 실제 자동알림 메일을 한 번도 안 받아봤다면, 정확한 발신자/제목을
모르니 --query를 넓게 잡아 목록만 출력해서 확인하는 용도로 먼저 쓴다.

사용법:
  # 1) 최근 메일 중 첨부파일 있는 것 목록만 확인 (아직 뭐가 오는지 모를 때)
  python3 ecount_gmail_fetch.py --list --query "has:attachment newer_than:2d"

  # 2) 실제 발신자/제목 확인되면 정확히 좁혀서 첨부 다운로드
  python3 ecount_gmail_fetch.py --query "from:noreply@ecount.com newer_than:1d has:attachment" \
      --download-to cron_tracking/ecount/sales_status_latest.xlsx
"""
from __future__ import annotations

import argparse
import base64
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

_TOKEN_FILE = Path(__file__).parent / ".secrets" / "gmail_token.json"
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


def _load_creds() -> Credentials:
    if not _TOKEN_FILE.exists():
        raise SystemExit(
            f"{_TOKEN_FILE} 이 없습니다. 로컬에서 ecount_gmail_auth.py로 발급받은 "
            "token.json을 이 경로로 옮기세요."
        )
    creds = Credentials.from_authorized_user_file(str(_TOKEN_FILE), SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        _TOKEN_FILE.write_text(creds.to_json())
    return creds


def list_messages(service, query: str) -> list[dict]:
    resp = service.users().messages().list(userId="me", q=query, maxResults=20).execute()
    return resp.get("messages", [])


def describe(service, msg_id: str) -> dict:
    msg = service.users().messages().get(userId="me", id=msg_id, format="metadata",
                                          metadataHeaders=["From", "Subject", "Date"]).execute()
    headers = {h["name"]: h["value"] for h in msg["payload"].get("headers", [])}
    return {"id": msg_id, **headers}


def download_first_attachment(service, msg_id: str, out_path: Path) -> Path | None:
    msg = service.users().messages().get(userId="me", id=msg_id, format="full").execute()
    parts = msg["payload"].get("parts", []) or []
    for part in parts:
        filename = part.get("filename") or ""
        if not filename.lower().endswith((".xlsx", ".xls")):
            continue
        body = part.get("body", {})
        att_id = body.get("attachmentId")
        if not att_id:
            continue
        att = service.users().messages().attachments().get(
            userId="me", messageId=msg_id, id=att_id
        ).execute()
        data = base64.urlsafe_b64decode(att["data"])
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_bytes(data)
        return out_path
    return None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--query", default="has:attachment newer_than:2d", help="Gmail 검색 쿼리")
    ap.add_argument("--list", action="store_true", help="목록만 출력하고 다운로드는 안 함")
    ap.add_argument("--download-to", help="가장 최근 매치 메일의 첫 엑셀 첨부를 이 경로로 저장")
    args = ap.parse_args()

    creds = _load_creds()
    service = build("gmail", "v1", credentials=creds)

    messages = list_messages(service, args.query)
    if not messages:
        print("[gmail] 쿼리에 매치되는 메일이 없습니다.")
        return 1

    print(f"[gmail] {len(messages)}건 매치:")
    for m in messages:
        info = describe(service, m["id"])
        print(f"  - id={info['id']}  From={info.get('From')}  Subject={info.get('Subject')}  Date={info.get('Date')}")

    if args.list or not args.download_to:
        return 0

    latest_id = messages[0]["id"]
    out_path = Path(args.download_to)
    saved = download_first_attachment(service, latest_id, out_path)
    if saved:
        print(f"[gmail] 첨부 저장 완료: {saved}")
    else:
        print("[gmail] 해당 메일에서 엑셀 첨부를 못 찾았습니다.")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
