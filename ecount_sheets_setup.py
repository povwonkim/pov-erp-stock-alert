#!/usr/bin/env python3
"""이 프로젝트 전용 구글 스프레드시트에 탭/헤더를 세팅한다.

스프레드시트 자체는 admin@pointofview.kr이 직접 만들고 povbotpovbot@gmail.com을
편집자로 공유해둔 상태여야 한다 (회사 실제 계정이 문서를 소유, 봇은 쓰기 권한만).
이 스크립트는 파일을 새로 만들지 않고, 이미 있는 스프레드시트 ID에 필요한 탭과
헤더 행만 채운다 (이미 있는 탭은 건드리지 않음).

사용법 (서버에서, gmail_token.json이 Sheets 권한 포함해서 재발급된 뒤):
  .venv/bin/python ecount_sheets_setup.py --spreadsheet-id <admin이 공유한 시트의 ID>

한 번 실행하고 나면 ID를 .secrets/sheet_id.json 에 저장해서 이후 스크립트들이
계속 그 ID로 데이터를 쓴다.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

_TOKEN_FILE = Path(__file__).parent / ".secrets" / "gmail_token.json"
_SHEET_ID_FILE = Path(__file__).parent / ".secrets" / "sheet_id.json"

SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

# 탭 이름 + 헤더 행. 순서가 곧 시트 안 탭 순서.
TABS: dict[str, list[str]] = {
    "RAW_재고현황": ["창고코드", "창고명", "품목코드", "품목명", "사이즈", "재고수량", "수집일시"],
    "RAW_판매현황": ["일자", "브랜드", "품목코드", "품명", "수량", "단가", "공급가액", "부가세", "합계", "적요", "창고명", "담당자", "수집일시"],
    "품목마스터": ["품목코드", "품목명", "브랜드", "브랜드코드", "갱신일"],
    "일별변동계산": ["날짜", "품목코드", "브랜드", "전일재고", "재고수량", "출고수량", "입고수량(계산)"],
    "디자인팀_발주필요": ["브랜드", "품목코드", "품목명", "재고수량", "저재고기준", "업데이트일시"],
    "관리팀_전체재고": ["브랜드", "품목코드", "품목명", "창고명", "재고수량", "전일재고", "입고수량(계산)", "출고수량", "업데이트일시"],
}


def _load_creds() -> Credentials:
    if not _TOKEN_FILE.exists():
        raise SystemExit(f"{_TOKEN_FILE} 이 없습니다. ecount_gmail_auth.py로 발급받은 토큰을 먼저 옮기세요.")
    creds = Credentials.from_authorized_user_file(str(_TOKEN_FILE), SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        _TOKEN_FILE.write_text(creds.to_json())
    missing = set(SCOPES) - set(creds.scopes or [])
    if missing:
        raise SystemExit(
            f"토큰에 이 권한이 없습니다: {missing}\n"
            "맥에서 ecount_gmail_auth.py를 다시 실행해서 (Sheets/Drive 권한 포함) 토큰을 재발급하세요."
        )
    return creds


def existing_tab_names(sheets_service, spreadsheet_id: str) -> set[str]:
    meta = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id, fields="sheets.properties.title").execute()
    return {s["properties"]["title"] for s in meta.get("sheets", [])}


def add_missing_tabs(sheets_service, spreadsheet_id: str, have: set[str]) -> list[str]:
    to_add = [name for name in TABS if name not in have]
    if not to_add:
        return []
    requests = [{"addSheet": {"properties": {"title": name}}} for name in to_add]
    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={"requests": requests}
    ).execute()
    return to_add


def write_headers(sheets_service, spreadsheet_id: str, tab_names: list[str]) -> None:
    if not tab_names:
        return
    data = [{"range": f"'{name}'!A1", "values": [TABS[name]]} for name in tab_names]
    sheets_service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"valueInputOption": "RAW", "data": data},
    ).execute()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", help="admin@pointofview.kr이 만들고 봇 계정에 편집자로 공유한 시트 ID (또는 URL)")
    args = ap.parse_args()

    if args.spreadsheet_id:
        # URL로 붙여넣었을 수도 있으니 ID만 추출
        spreadsheet_id = args.spreadsheet_id
        if "/d/" in spreadsheet_id:
            spreadsheet_id = spreadsheet_id.split("/d/")[1].split("/")[0]
        _SHEET_ID_FILE.write_text(json.dumps({"spreadsheet_id": spreadsheet_id}, ensure_ascii=False, indent=2))
    elif _SHEET_ID_FILE.exists():
        spreadsheet_id = json.loads(_SHEET_ID_FILE.read_text())["spreadsheet_id"]
    else:
        raise SystemExit("--spreadsheet-id 를 지정하세요 (admin이 만들고 봇에 공유한 시트의 ID 또는 URL).")

    creds = _load_creds()
    sheets_service = build("sheets", "v4", credentials=creds)

    have = existing_tab_names(sheets_service, spreadsheet_id)
    added = add_missing_tabs(sheets_service, spreadsheet_id, have)
    write_headers(sheets_service, spreadsheet_id, added)

    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
    print(f"[sheets] 대상: {url}")
    if added:
        print(f"[sheets] 새로 추가된 탭: {', '.join(added)}")
    else:
        print("[sheets] 이미 모든 탭이 존재해서 추가한 것 없음")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
