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

# 오프라인 재고 관리 대상 창고 — 확정 4개 (2026-07-28). POP-UP/orer.archive/OTS/OY/POV는
# 이 시스템의 범위에서 완전히 제외(집계에도 포함 안 함). 온라인(MXN 계열) 창고는 별도 시스템.
OFFLINE_WAREHOUSES = ["POINT OF VIEW(법인)", "THE HYUNDAI SEOUL", "시시호시-수원점", "신세계 강남-피숀"]

# ---- 서식(색/줄무늬/얼림/폰트) — 2026-07-28 디자인 검토(별도 세션) 스펙 반영.
# 배너는 어둡게 위로, 헤더는 밝게 아래로(둘 다 어두우면 얼린 두 줄이 두꺼운 띠가 됨). ----
BANNER1_BG = {"red": 0.102, "green": 0.102, "blue": 0.102}   # #1A1A1A — 1행(탭 설명)
BANNER2_BG = {"red": 0.200, "green": 0.200, "blue": 0.200}   # #333333 — 2행(할 일)
BANNER3_BG = {"red": 0.945, "green": 0.945, "blue": 0.945}   # #F1F1F1 — 3행(정렬근거, 헤더와 같은 톤)
BANNER_FG_LIGHT = {"red": 1, "green": 1, "blue": 1}
BANNER3_FG = {"red": 0.4, "green": 0.4, "blue": 0.4}

HEADER_BG = {"red": 0.945, "green": 0.945, "blue": 0.945}    # #F1F1F1 — 4행(진짜 헤더)
HEADER_FG = {"red": 0.133, "green": 0.133, "blue": 0.133}    # #222222
HEADER_BORDER = {"red": 0.2, "green": 0.2, "blue": 0.2}      # #333333 하단 굵은 테두리

STRIPE_BG = {"red": 0.980, "green": 0.973, "blue": 0.965}    # #FAF8F6 — 데이터 줄무늬(웜 오프화이트)
BODY_FG = {"red": 0.2, "green": 0.2, "blue": 0.2}            # #333333 — 순수 검정 금지

FONT_BODY = "Noto Sans KR"
FONT_MONO = "Roboto Mono"

SEMANTIC_DANGER_BG = {"red": 0.976, "green": 0.816, "blue": 0.816}   # F9D0D0
SEMANTIC_DANGER_FG = {"red": 0.800, "green": 0.000, "blue": 0.000}   # CC0000
SEMANTIC_WARNING_BG = {"red": 1.000, "green": 0.976, "blue": 0.769}  # FFF9C4
SEMANTIC_WARNING_FG = {"red": 0.702, "green": 0.420, "blue": 0.000}  # B36B00
SEMANTIC_INFO_BG = {"red": 0.839, "green": 0.918, "blue": 0.973}     # D6EAF8 (과잉/과다재고)
SEMANTIC_INFO_FG = {"red": 0.084, "green": 0.396, "blue": 0.753}     # 1565C0
SEMANTIC_DIM_FG = {"red": 0.600, "green": 0.600, "blue": 0.600}      # 999999 (텍스트만, 배경 없음)

# "상태" 컬럼에 부분일치(SEARCH)로 색 입히는 규칙 — (부분문자열, 배경 or None, 글자색, 굵게).
# 문자열들이 서로 겹치지 않아 순서 무관. "정상"은 의도적으로 무색(색은 예외를 알리는 신호).
# 품절이 오래될수록 신호가 약해진다: 신규(빨강,지금 채우면 매출 삶) → 지속(주황) → 장기(회색,이미 죽음).
STATUS_COLOR_RULES = [
    ("초과주문", SEMANTIC_DANGER_BG, SEMANTIC_DANGER_FG, True),
    ("위험", SEMANTIC_DANGER_BG, SEMANTIC_DANGER_FG, True),
    ("품절-신규", SEMANTIC_DANGER_BG, SEMANTIC_DANGER_FG, True),
    ("주의", SEMANTIC_WARNING_BG, SEMANTIC_WARNING_FG, True),
    ("품절-지속", SEMANTIC_WARNING_BG, SEMANTIC_WARNING_FG, True),
    ("과잉", SEMANTIC_INFO_BG, SEMANTIC_INFO_FG, True),
    ("재고소량", None, SEMANTIC_DIM_FG, False),
    ("품절-장기", None, SEMANTIC_DIM_FG, False),
]

# 스크롤해도 식별자 컬럼(브랜드/품목코드/품목명 등)이 계속 보이게 얼릴 컬럼 수.
FREEZE_COLS = {
    "RAW_재고현황": 4,
    "RAW_판매현황": 3,
    "품목마스터": 3,
    "일별재고이력": 4,
    "디자인팀_발주필요": 3,
    "관리팀_전체재고": 3,
    "악성재고": 3,
    "악성품절": 3,
}

# 컬럼명 패턴 → 숫자서식. 0을 "–"로(0이 빽빽하면 "재고 없음"이 안 보임), 숫자는 Roboto Mono
# 고정폭 + 우측정렬로 자릿수를 맞춘다(재고량 스캔 속도가 가장 크게 체감되는 변경).
QTY_COLS = {"재고수량", "총재고", "전일재고", "입고수량(계산)", "출고수량",
            "최근7일판매량", "최근90일판매량", "수량", "공급가액", "부가세", "합계", "단가"}
DAY_COLS = {"DOI", "품절경과일", "미판매경과일", "미입고경과일", "리드타임(일)"}
INT_COLS = {"우선순위"}
MONEY_COLS = {"재고평가금액", "입고단가"}
DATE_COLS = {"최근판매일", "최근입고일", "갱신일", "날짜"}

NUMBER_FORMAT_QTY = {"type": "NUMBER", "pattern": '#,##0;-#,##0;"–"'}
NUMBER_FORMAT_DAY = {"type": "NUMBER", "pattern": '0"일";;"–"'}
NUMBER_FORMAT_INT = {"type": "NUMBER", "pattern": "#,##0"}
NUMBER_FORMAT_MONEY = {"type": "NUMBER", "pattern": '#,##0"원";;"–"'}
NUMBER_FORMAT_DATE = {"type": "DATE", "pattern": "yyyy-mm-dd"}

# 탭 이름 → {note: 3행 배너([① 탭 설명, ② 할 일, ③ 정렬·갱신 근거]), headers: 헤더 행}.
# 컬럼 순서 규칙(2026-07-28 통일): 식별자(브랜드/품목코드/품목명 등) → 상태(+우선순위) →
# 나머지 지표 → 조치방안/메모(맨 끝). 상태가 있는 탭은 전부 조치방안도 같이 둔다(둘은 항상 짝).
# "수집일시/업데이트일시/일자"처럼 매 행 반복되는 배치 기준시각은 컬럼에 넣지 않고 배너 3행에
# 넣는다(단, 일별재고이력의 "날짜"는 매 행이 실제로 다른 값이라 예외 — 그대로 컬럼 유지).
# 3층 구조 (카페24 온라인 대조 시스템과 동일한 패턴을 오프라인에 적용, DOI 기반 우선순위 포함):
#   1층 RAW(원본 수집) → 2층 일별재고이력(DOI·상태판정 엔진) → 3층 결과물(용도별 뷰)
TABS: dict[str, dict] = {
    # 1층 RAW — 사람이 직접 보는 탭이 아니라, 아래 계산의 원재료.
    "RAW_재고현황": {
        "note": [
            "📋 이카운트 재고API 원본 스냅샷 (사람이 보는 탭 아님, 계산용 원본)",
            "직접 볼 일 없음 — 데이터 이상 있을 때만 원본 확인용",
            "정렬 없음(API 응답 순서 그대로) · 🕒 기준일시: 매일 자동 갱신",
        ],
        "headers": ["창고코드", "창고명", "품목코드", "품목명", "사이즈", "재고수량"],
    },
    "RAW_판매현황": {
        "note": [
            "📋 이카운트 판매현황 이메일 원본 (사람이 보는 탭 아님, 계산용 원본)",
            "직접 볼 일 없음 — 데이터 이상 있을 때만 원본 확인용",
            "정렬 없음 · 🕒 기준일(전일): 매일 자동 갱신",
        ],
        "headers": ["브랜드", "품목코드", "품명", "수량", "단가", "공급가액", "부가세", "합계", "적요", "창고명", "담당자"],
    },
    "품목마스터": {
        "note": [
            "📋 품목코드↔브랜드↔조달유형 매핑표",
            "신상품 브랜드가 '미분류'로 뜨면 조달유형만 채워주세요 — 그 뒤로 같은 브랜드는 자동 적용",
            "정렬: 품목코드순 · 🕒 신상품 추가 시에만 사람이 가끔 수동 갱신",
        ],
        "headers": ["품목코드", "품목명", "브랜드", "브랜드코드", "조달유형", "리드타임(일)", "갱신일"],
    },
    # 2층 일별재고이력 — 3층의 모든 결과물이 여기서 계산돼 나오는 엔진. 사람이 매일 볼 필요는
    # 없지만, 왜 그런 상태/조치가 나왔는지 근거를 확인하고 싶을 때 여기를 본다.
    # 상태값: 위험/주의(조달유형별 리드타임 기준, 아래 참고)/재고소량·판매없음(재고1~5&7일판매0)/
    #        품절-신규(0~9일째)/품절-지속(10~29일째)/품절-장기(30일+)/과잉(DOI>180일)/초과주문(재고<0)/정상
    # 위험/주의 임계값: 자체제작 DOI≤35일/35~49일 · 국내사입 DOI≤7일/7~14일 · 해외수입 DOI≤21일/21~35일
    "일별재고이력": {
        "note": [
            "📋 품목별 일별 재고·입출고·DOI·상태 계산 이력 — 3층 결과물 탭들이 전부 여기서 계산됨",
            "직접 볼 일 없음 — 왜 그런 상태/조치가 나왔는지 근거 확인용",
            "정렬: 날짜 오름차순 → 품목코드순 · 🕒 매일 품목당 한 행씩 자동 누적",
        ],
        "headers": [
            "날짜", "브랜드", "품목코드", "품목명", "상태", "우선순위",
            "전일재고", "재고수량", "출고수량", "입고수량(계산)",
            "최근7일판매량", "최근90일판매량", "DOI", "품절경과일", "조치방안",
        ],
    },
    # 3층 결과물 (용도별 뷰) — 온라인의 "위험/주의/재고소량·판매없음/신규품절/지속품절" 리스트에 대응.
    # 정렬: 위험·주의는 DOI 오름차순, 품절-지속/장기는 최근90일판매량 내림차순(부활가치 큰 것 우선).
    # 재고수량 = 오프라인 4개 창고 합계 (창고별 상세는 관리팀_전체재고에서 확인 — 디자인팀은
    # "만들어야 하는가"만 판단하면 되므로 창고별 상세 불필요).
    "디자인팀_발주필요": {
        "note": [
            "📋 디자인팀용 — 지금 발주(제작)해야 하는 품목만 골라서 보여줌",
            "우선순위 1번부터 순서대로 제작 발주 검토 (전체 재고는 관리팀_전체재고 참고)",
            "정렬: 우선순위 오름차순 → 재고수량 오름차순 · 🕒 마지막 갱신: 매일 자동",
        ],
        "headers": [
            "브랜드", "품목코드", "품목명", "조달유형", "상태", "우선순위",
            "리드타임(일)", "재고수량", "DOI", "최근7일판매량", "최근90일판매량", "품절경과일",
            "조치방안", "메모",
        ],
    },
    # 이 팀이 국내발주/해외발주/창고이동을 모두 판단하는 실제 운영 마스터 시트 — 가장 중요.
    # 창고별 수량을 컬럼으로 나란히 둬서 창고이동 필요 여부(A매장 과잉·B매장 품절)를 한 행에서
    # 바로 확인 가능하게 한다. 조달유형별로 DOI 임계값이 다르므로 리드타임을 컬럼으로 명시.
    "관리팀_전체재고": {
        "note": [
            "📋 관리팀용 마스터 시트 — 오프라인 전체 재고를 창고별로 보고 국내발주/해외발주/창고이동을 판단",
            "상태가 위험/주의/과잉인 행부터 확인 · 창고별 컬럼에 0인 칸 있으면 창고이동 검토",
            "정렬: 상태 우선순위 → 품목코드순 · 🕒 마지막 갱신: 매일 자동",
        ],
        "headers": [
            "브랜드", "품목코드", "품목명", "조달유형", "상태",
            "리드타임(일)", *OFFLINE_WAREHOUSES, "총재고",
            "전일재고", "입고수량(계산)", "출고수량", "최근7일판매량", "DOI",
            "조치방안", "메모",
        ],
    },
    # 재고 과잉 + 장기 미출고 (판매 없음, 재고만 쌓임 — 디자인팀_발주필요와 반대 축)
    "악성재고": {
        "note": [
            "📋 재고 과잉 + 오래 안 팔린 품목 — 프로모션/폐기 판단용 (반대 축: 디자인팀_발주필요)",
            "재고평가금액 큰 것부터 프로모션 또는 폐기 결정",
            "정렬: 재고평가금액 내림차순 · 🕒 마지막 갱신: 매일 자동",
        ],
        "headers": ["브랜드", "품목코드", "품목명", "재고수량", "최근판매일", "미판매경과일", "입고단가", "재고평가금액", "조치방안", "메모"],
    },
    # 품절-장기 중에서도 최근90일판매량=0(부활가치 없음)인 것만. 판매량 있는 장기품절은
    # 디자인팀_발주필요의 "재입고 골든타임"으로 남긴다. 재발주/제작/단종 최종 판단용.
    # 컬럼은 "방치 근거"(품절경과일/조달유형/리드타임)와 "부활가치 근거"(최근입고일/미입고경과일/
    # 최근90일판매량)를 모두 담아서 왜 여기 있는지 판단할 정보를 다 준다.
    "악성품절": {
        "note": [
            "📋 재고 없이 오래 방치된 품목(부활가치 없음) — 재발주/단종 최종 판단용",
            "품절경과일 오래된 것부터 재발주할지 단종할지 결정",
            "정렬: 품절경과일 내림차순 · 🕒 마지막 갱신: 매일 자동",
        ],
        "headers": [
            "브랜드", "품목코드", "품목명", "조달유형", "리드타임(일)",
            "재고수량", "최근판매일", "미판매경과일", "품절경과일",
            "최근입고일", "미입고경과일", "최근90일판매량", "조치방안", "메모",
        ],
    },
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


# 과거 이름 -> 현재 이름. 실행 시 존재하면 새로 만들지 않고 이름만 바꿔서 데이터 유지.
RENAMED_FROM = {
    "일별변동계산": "일별재고이력",
}


def existing_tab_names(sheets_service, spreadsheet_id: str) -> set[str]:
    meta = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id, fields="sheets.properties.title").execute()
    return {s["properties"]["title"] for s in meta.get("sheets", [])}


def rename_legacy_tabs(sheets_service, spreadsheet_id: str, have: set[str]) -> list[tuple[str, str]]:
    meta = sheets_service.spreadsheets().get(
        spreadsheetId=spreadsheet_id, fields="sheets(properties(sheetId,title))"
    ).execute()
    id_by_title = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}
    requests = []
    renamed = []
    for old, new in RENAMED_FROM.items():
        if old in have and new not in have:
            requests.append({
                "updateSheetProperties": {
                    "properties": {"sheetId": id_by_title[old], "title": new},
                    "fields": "title",
                }
            })
            renamed.append((old, new))
    if requests:
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": requests}
        ).execute()
    return renamed


def add_missing_tabs(sheets_service, spreadsheet_id: str, have: set[str]) -> list[str]:
    to_add = [name for name in TABS if name not in have]
    if not to_add:
        return []
    requests = [{"addSheet": {"properties": {"title": name}}} for name in to_add]
    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, body={"requests": requests}
    ).execute()
    return to_add


# 행 레이아웃(고정): 1~3행 = 배너(탭설명/할일/정렬근거), 4행 = 헤더, 5행부터 = 데이터.
BANNER_ROWS = 3
HEADER_ROW_IDX = BANNER_ROWS       # 0-indexed = 3 (=4행)
DATA_START_IDX = BANNER_ROWS + 1   # 0-indexed = 4 (=5행)


def write_headers(sheets_service, spreadsheet_id: str, tab_names: list[str]) -> None:
    """1~3행에 note 배너(탭설명/할일/정렬근거), 4행에 헤더. 데이터는 5행부터 시작하는 게 전제.

    과거 실행에서 컬럼 수·행 수가 달랐던 적이 있으면 그 잔여 셀이 안 지워지고 남아 새 내용과
    뒤섞일 수 있어서, 쓰기 전에 1~4행을 넉넉히(Z열까지) 지워둔다.
    """
    if not tab_names:
        return
    clear_ranges = [f"'{name}'!A1:Z{HEADER_ROW_IDX + 1}" for name in tab_names]
    sheets_service.spreadsheets().values().batchClear(
        spreadsheetId=spreadsheet_id, body={"ranges": clear_ranges}
    ).execute()

    data = []
    for name in tab_names:
        spec = TABS[name]
        for i, line in enumerate(spec["note"]):
            data.append({"range": f"'{name}'!A{i + 1}", "values": [[line]]})
        data.append({"range": f"'{name}'!A{HEADER_ROW_IDX + 1}", "values": [spec["headers"]]})
    sheets_service.spreadsheets().values().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"valueInputOption": "RAW", "data": data},
    ).execute()

    # 기존 서식(병합/조건부서식/줄무늬)을 먼저 깨끗이 지운 뒤에 새로 입힌다 — 재실행할 때마다
    # 쌓이지 않게. 삭제 대상은 실제 메타데이터에서 조회해서 정확한 ID/인덱스로 지운다.
    meta = sheets_service.spreadsheets().get(
        spreadsheetId=spreadsheet_id,
        fields="sheets(properties(sheetId,title),conditionalFormats,bandedRanges(bandedRangeId))",
    ).execute()
    by_title = {s["properties"]["title"]: s for s in meta["sheets"]}

    delete_requests = []
    for name in tab_names:
        sheet = by_title[name]
        sheet_id = sheet["properties"]["sheetId"]
        n_rules = len(sheet.get("conditionalFormats", []))
        # 인덱스가 삭제할 때마다 당겨지므로 뒤에서부터 지운다.
        for idx in range(n_rules - 1, -1, -1):
            delete_requests.append({"deleteConditionalFormatRule": {"sheetId": sheet_id, "index": idx}})
        for banded in sheet.get("bandedRanges", []):
            delete_requests.append({"deleteBanding": {"bandedRangeId": banded["bandedRangeId"]}})
        delete_requests.append({
            "unmergeCells": {
                "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": BANNER_ROWS, "startColumnIndex": 0, "endColumnIndex": 26},
            }
        })
    if delete_requests:
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": delete_requests}
        ).execute()

    requests = []
    for name in tab_names:
        headers = TABS[name]["headers"]
        ncols = max(len(headers), 1)
        sheet_id = by_title[name]["properties"]["sheetId"]

        # 배너 3행 — 병합은 안 함(컬럼 고정과 충돌하므로). 대신 행 전체(전체 컬럼)에 배경색을
        # 칠해서 같은 "띠" 모양을 낸다. 텍스트는 A열에만 있고 나머지 칸은 빈 채로 배경만 깔림.
        banner_rows = [
            (0, BANNER1_BG, BANNER_FG_LIGHT, True),
            (1, BANNER2_BG, BANNER_FG_LIGHT, True),
            (2, BANNER3_BG, BANNER3_FG, False),
        ]
        for row_idx, bg, fg, bold in banner_rows:
            requests.append({
                "repeatCell": {
                    "range": {"sheetId": sheet_id, "startRowIndex": row_idx, "endRowIndex": row_idx + 1},
                    "cell": {"userEnteredFormat": {
                        "backgroundColor": bg,
                        "textFormat": {"bold": bold, "italic": not bold, "foregroundColor": fg, "fontFamily": FONT_BODY, "fontSize": 9},
                        "verticalAlignment": "MIDDLE",
                    }},
                    "fields": "userEnteredFormat(textFormat,backgroundColor,verticalAlignment)",
                }
            })

        # 헤더 행(4행) — 밝은 배경 + 진한 글씨 + 하단 굵은 테두리.
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": HEADER_ROW_IDX, "endRowIndex": HEADER_ROW_IDX + 1},
                "cell": {"userEnteredFormat": {
                    "backgroundColor": HEADER_BG,
                    "textFormat": {"bold": True, "foregroundColor": HEADER_FG, "fontFamily": FONT_BODY, "fontSize": 10},
                    "horizontalAlignment": "CENTER",
                    "verticalAlignment": "MIDDLE",
                    "borders": {"bottom": {"style": "SOLID_THICK", "color": HEADER_BORDER}},
                }},
                "fields": "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment,verticalAlignment,borders)",
            }
        })

        requests.append({
            "updateSheetProperties": {
                "properties": {"sheetId": sheet_id, "gridProperties": {
                    "frozenRowCount": HEADER_ROW_IDX + 1, "frozenColumnCount": FREEZE_COLS.get(name, 0),
                }},
                "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
            }
        })
        requests.append({
            "addBanding": {
                "bandedRange": {
                    "range": {"sheetId": sheet_id, "startRowIndex": DATA_START_IDX, "endRowIndex": 1000,
                              "startColumnIndex": 0, "endColumnIndex": ncols},
                    "rowProperties": {
                        "headerColor": {"red": 1, "green": 1, "blue": 1},
                        "firstBandColor": {"red": 1, "green": 1, "blue": 1},
                        "secondBandColor": STRIPE_BG,
                    },
                }
            }
        })

        # 데이터 영역 기본 폰트/글자색부터 깔고(Noto Sans KR, #333333), 숫자류 컬럼을
        # Roboto Mono + 우측정렬 + 전용 숫자서식으로 덮어쓴다(나중 요청이 우선 적용됨).
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": DATA_START_IDX, "endRowIndex": 1000,
                           "startColumnIndex": 0, "endColumnIndex": ncols},
                "cell": {"userEnteredFormat": {"textFormat": {"fontFamily": FONT_BODY, "foregroundColor": BODY_FG}}},
                "fields": "userEnteredFormat.textFormat",
            }
        })
        numeric_col_formats = (
            [(c, NUMBER_FORMAT_QTY) for c in QTY_COLS]
            + [(c, NUMBER_FORMAT_DAY) for c in DAY_COLS]
            + [(c, NUMBER_FORMAT_INT) for c in INT_COLS]
            + [(c, NUMBER_FORMAT_MONEY) for c in MONEY_COLS]
        )
        for col_name, fmt in numeric_col_formats:
            if col_name in headers:
                col = headers.index(col_name)
                requests.append({
                    "repeatCell": {
                        "range": {"sheetId": sheet_id, "startRowIndex": DATA_START_IDX, "endRowIndex": 1000,
                                   "startColumnIndex": col, "endColumnIndex": col + 1},
                        "cell": {"userEnteredFormat": {
                            "numberFormat": fmt,
                            "horizontalAlignment": "RIGHT",
                            "textFormat": {"fontFamily": FONT_MONO},
                        }},
                        "fields": "userEnteredFormat(numberFormat,horizontalAlignment,textFormat)",
                    }
                })
        for col_name in DATE_COLS:
            if col_name in headers:
                col = headers.index(col_name)
                requests.append({
                    "repeatCell": {
                        "range": {"sheetId": sheet_id, "startRowIndex": DATA_START_IDX, "endRowIndex": 1000,
                                   "startColumnIndex": col, "endColumnIndex": col + 1},
                        "cell": {"userEnteredFormat": {"numberFormat": NUMBER_FORMAT_DATE}},
                        "fields": "userEnteredFormat.numberFormat",
                    }
                })

        if "상태" in headers:
            status_col = headers.index("상태")
            for substr, bg, fg, bold in STATUS_COLOR_RULES:
                fmt = {"textFormat": {"bold": bold, "foregroundColor": fg}}
                if bg is not None:
                    fmt["backgroundColor"] = bg
                requests.append({
                    "addConditionalFormatRule": {
                        "rule": {
                            "ranges": [{"sheetId": sheet_id, "startRowIndex": DATA_START_IDX, "endRowIndex": 1000,
                                        "startColumnIndex": status_col, "endColumnIndex": status_col + 1}],
                            "booleanRule": {
                                "condition": {"type": "TEXT_CONTAINS", "values": [{"userEnteredValue": substr}]},
                                "format": fmt,
                            },
                        },
                        "index": 0,
                    }
                })

        for qty_col_name in ("재고수량", "총재고"):
            if qty_col_name in headers:
                qty_col = headers.index(qty_col_name)
                requests.append({
                    "addConditionalFormatRule": {
                        "rule": {
                            "ranges": [{"sheetId": sheet_id, "startRowIndex": DATA_START_IDX, "endRowIndex": 1000,
                                        "startColumnIndex": qty_col, "endColumnIndex": qty_col + 1}],
                            "booleanRule": {
                                "condition": {"type": "NUMBER_LESS", "values": [{"userEnteredValue": "0"}]},
                                "format": {"backgroundColor": SEMANTIC_DANGER_BG, "textFormat": {"bold": True, "foregroundColor": SEMANTIC_DANGER_FG}},
                            },
                        },
                        "index": 0,
                    }
                })

        # 관리팀_전체재고: 창고이동 신호 — 총재고>0인데 특정 창고만 0이면 그 칸을 흐리게(회색)
        # 표시해서, 가로로 훑으면 "어디로 보내야 하는지"가 바로 보이게 한다.
        if name == "관리팀_전체재고" and "총재고" in headers:
            total_col = headers.index("총재고")
            total_letter = _col_letter(total_col)
            for wh in OFFLINE_WAREHOUSES:
                if wh not in headers:
                    continue
                wh_col = headers.index(wh)
                wh_letter = _col_letter(wh_col)
                requests.append({
                    "addConditionalFormatRule": {
                        "rule": {
                            "ranges": [{"sheetId": sheet_id, "startRowIndex": DATA_START_IDX, "endRowIndex": 1000,
                                        "startColumnIndex": wh_col, "endColumnIndex": wh_col + 1}],
                            "booleanRule": {
                                "condition": {
                                    "type": "CUSTOM_FORMULA",
                                    "values": [{"userEnteredValue": f"=AND(${wh_letter}{DATA_START_IDX + 1}=0,${total_letter}{DATA_START_IDX + 1}>0)"}],
                                },
                                "format": {"textFormat": {"foregroundColor": SEMANTIC_DIM_FG}},
                            },
                        },
                        "index": 0,
                    }
                })

    if requests:
        try:
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body={"requests": requests}
            ).execute()
        except Exception as e:  # 병합이 기존 데이터와 충돌하는 등 — 헤더 텍스트 자체는 이미 써졌으니 계속 진행
            print(f"[sheets] 서식 적용 중 일부 실패(무시하고 진행): {e}")


def _col_letter(idx: int) -> str:
    """0-indexed 컬럼 번호 -> A1 표기 알파벳 (A, B, ..., Z, AA, ...)."""
    letters = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


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
    renamed = rename_legacy_tabs(sheets_service, spreadsheet_id, have)
    if renamed:
        have = existing_tab_names(sheets_service, spreadsheet_id)  # 이름 바뀐 뒤 다시 조회
    added = add_missing_tabs(sheets_service, spreadsheet_id, have)
    # note/헤더는 매번 전체 탭 기준으로 동기화한다(1~2행만 덮어씀, 3행부터의 데이터는 안 건드림) —
    # TABS의 컬럼 정의가 바뀌었을 때(예: 상태/우선순위 컬럼 추가) 기존 탭도 따라가게.
    write_headers(sheets_service, spreadsheet_id, list(TABS.keys()))

    url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}"
    print(f"[sheets] 대상: {url}")
    if renamed:
        print(f"[sheets] 이름 변경된 탭: {', '.join(f'{o} → {n}' for o, n in renamed)}")
    if added:
        print(f"[sheets] 새로 추가된 탭: {', '.join(added)}")
    else:
        print("[sheets] 새로 추가된 탭 없음 (기존 탭 헤더만 최신화)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
