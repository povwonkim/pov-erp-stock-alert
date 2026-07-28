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
import re
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


def get_html_body(service, msg_id: str) -> str:
    """메일 본문(HTML) 전체를 합쳐서 반환. 자동알림 메일은 첨부 없이 '수신문서보기' 링크만
    본문에 들어있는 경우가 있어(2026-07-28 실메일로 확인), 이 링크를 뽑는 데 쓴다."""
    msg = service.users().messages().get(userId="me", id=msg_id, format="full").execute()

    def _walk(payload):
        mime = payload.get("mimeType", "")
        body = payload.get("body", {})
        if mime.startswith("text/html") and body.get("data"):
            yield body["data"]
        for part in payload.get("parts", []) or []:
            yield from _walk(part)

    html = ""
    for data in _walk(msg["payload"]):
        html += base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")
    return html


def find_view_link(html: str) -> str | None:
    """본문 HTML에서 '수신문서보기' 버튼 링크를 찾는다.

    실제로는 ViewMailContents URL이 본문에 직접 있지 않고, l.ecount.com 단축링크로
    감싸져 있다(2026-07-28 실메일로 확인). 그 단축링크가 SEND_CM_ID(그날 알림 고유 ID)를
    포함한 실제 URL로 리다이렉트하므로, 브라우저(Playwright)가 이 단축링크를 그대로 열면
    자동으로 따라간다 — 매일 이 단축링크 자체도 새로 발급되므로 매번 새로 추출해야 한다."""
    patterns = [
        r'href=["\'](https?://l\.ecount\.com/[^"\']*)["\']',
        r'href=["\'](https?://[^"\']*ViewMailContents[^"\']*)["\']',
        r'href=["\'](https?://[^"\']*AutomationBridge[^"\']*)["\']',
        r'(https?://l\.ecount\.com/\S+)',
        r'(https?://[^\s"\'<>]*ViewMailContents[^\s"\'<>]*)',
        r'(https?://[^\s"\'<>]*AutomationBridge[^\s"\'<>]*)',
    ]
    for pattern in patterns:
        m = re.search(pattern, html, re.IGNORECASE)
        if m:
            return m.group(1).replace("&amp;", "&")
    return None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--query", default="has:attachment newer_than:2d", help="Gmail 검색 쿼리")
    ap.add_argument("--list", action="store_true", help="목록만 출력하고 다운로드는 안 함")
    ap.add_argument("--download-to", help="가장 최근 매치 메일의 첫 엑셀 첨부를 이 경로로 저장")
    ap.add_argument("--find-link", action="store_true",
                     help="가장 최근 매치 메일 본문에서 '수신문서보기' 링크를 찾아 출력(첨부 없는 알림 메일용)")
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

    if args.find_link:
        latest_id = messages[0]["id"]
        html = get_html_body(service, latest_id)
        link = find_view_link(html)
        if link:
            print(f"[gmail] 수신문서보기 링크: {link}")
        else:
            print("[gmail] 본문에서 링크를 못 찾았습니다 — 본문 HTML 구조가 다를 수 있습니다.")
            return 1
        return 0

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
