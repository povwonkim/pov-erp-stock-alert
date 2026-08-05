#!/usr/bin/env python3
"""서식(색상/폰트/조건부서식) 확인용 샘플 데이터를 넣거나 지운다.

실제 자동화 데이터가 아니라 디자인 검토용 예시 몇 줄이다. --clear로 지우면
5행부터(헤더 아래) 전부 삭제된다.

사용법:
  .venv/bin/python ecount_sheets_sample_data.py --spreadsheet-id <ID>          # 샘플 넣기
  .venv/bin/python ecount_sheets_sample_data.py --spreadsheet-id <ID> --clear  # 지우기
"""
from __future__ import annotations

import argparse

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

from ecount_sheets_setup import _TOKEN_FILE, SCOPES, DATA_START_IDX

SAMPLE_ROWS = {
    "관리팀_전체재고": [
        # 브랜드, 품목코드, 품목명, 조달유형, 상태, 리드타임(일), 4개 창고, 총재고, 전일재고, 입고, 출고, 최근7일, DOI, 조치방안, 메모
        ["POV_original", "POV-1001", "오브제 노트 A5 · 크림", "자체제작", "🔴 위험", 35, 1, 0, 0, 0, 1, 2, 0, 0, 1, 7, "긴급 제작 필요", ""],
        ["POV Atelier", "POV-1002", "레더 펜슬캡 · 네이비", "자체제작", "🟠 주의", 35, 3, 2, 1, 1, 7, 8, 0, 1, 1, 49, "제작 검토", ""],
        ["국내OEM", "DM-4001", "라벨 스티커 A4", "국내사입", "🔴 마이너스재고", 7, -2, 0, 0, 0, -2, 0, 0, 0, 12, "", "재고 데이터 즉시 확인", "이카운트 전표 확인 필요"],
        ["KAWECO", "IM-3001", "카웨코 스포츠 만년필 · 블랙", "해외수입", "🔵 과잉", 21, 60, 40, 30, 25, 155, 158, 0, 0, 2, 230, "프로모션/할인 검토", ""],
        ["PAULA SKENE", "0031-20002-B", "card_tree with dove&candles", "해외수입", "🟢 정상", 21, 30, 25, 20, 10, 85, 85, 0, 0, 0, "", "-", ""],
        ["POV x Hello Kitty", "POV-HK01", "keep sack pouch S", "자체제작", "🟠 품절-지속", 35, 0, 0, 0, 0, 0, 0, 0, 0, 45, "", "재입고 골든타임(90일 45개)", "품절 18일째"],
        ["POV Atelier", "POV-SLG09", "round leather case_L", "자체제작", "🟡 재고소량", 35, 1, 1, 0, 0, 2, 2, 0, 0, 0, "", "마케팅 부스트 또는 정리 검토", "(사용x/구버전)"],
    ],
    "디자인팀_발주필요": [
        # 브랜드, 품목코드, 품목명, 조달유형, 상태, 우선순위, 리드타임(일), 재고, DOI(소진일), 7일 판매, 90일 판매, 품절(일), 조치, 메모
        ["POV_original", "POV-1001", "오브제 노트 A5 · 크림", "자체제작", "🔴 위험", 1, 35, 1, 7, 1, 9, "", "긴급 제작 필요", ""],
        ["POV Atelier", "POV-1002", "레더 펜슬캡 · 네이비", "자체제작", "🟠 주의", 2, 35, 7, 49, 1, 12, "", "제작 검토", ""],
        ["POV x Hello Kitty", "POV-HK01", "keep sack pouch S", "자체제작", "🟠 품절-지속", 3, 35, 0, "", 0, 45, 18, "제작 검토", "재입고 골든타임"],
    ],
}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", required=True)
    ap.add_argument("--clear", action="store_true", help="샘플 데이터를 넣는 대신 지운다")
    args = ap.parse_args()

    creds = Credentials.from_authorized_user_file(str(_TOKEN_FILE), SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
    service = build("sheets", "v4", credentials=creds)

    if args.clear:
        ranges = [f"'{name}'!A{DATA_START_IDX + 1}:Z1000" for name in SAMPLE_ROWS]
        service.spreadsheets().values().batchClear(
            spreadsheetId=args.spreadsheet_id, body={"ranges": ranges}
        ).execute()
        print("[sample] 샘플 데이터 삭제 완료:", ", ".join(SAMPLE_ROWS))
        return 0

    data = [
        {"range": f"'{name}'!A{DATA_START_IDX + 1}", "values": rows}
        for name, rows in SAMPLE_ROWS.items()
    ]
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=args.spreadsheet_id,
        body={"valueInputOption": "RAW", "data": data},
    ).execute()
    print("[sample] 샘플 데이터 입력 완료:", ", ".join(SAMPLE_ROWS))
    print("[sample] 확인 끝나면 --clear 옵션으로 지우세요.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
