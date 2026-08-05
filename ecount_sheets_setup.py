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
# 이름은 이카운트 API가 실제로 내려주는 WH_DES 문자열과 정확히 일치해야 필터링이 된다
# ("POINT OF VIEW (법인)" — 괄호 앞에 띄어쓰기 있음, 2026-07-28 실제 API 응답으로 확인).
# 2026-07-28: 시시호시-수원점/신세계 강남-피숀 매장을 빼고 MXN/MXN(온라인)을 추가(사용자 확정) —
# 이 시스템은 이제 "오프라인 매장만"이 아니라 "회사 전체 재고 중 관리 대상 4개 창고"로 범위가
# 넓어짐. 온라인(MXN 계열) 재고를 봐야 오프라인 품절 시 온라인→오프라인 재고 이동 판단이 가능.
OFFLINE_WAREHOUSES = ["POINT OF VIEW (법인)", "THE HYUNDAI SEOUL", "MXN", "MXN(온라인)"]

# 창고명 -> WH_CD. 재고API가 15개 창고 통합 조회 시 10000건 제한에 걸리는 문제(README '탐사
# 결과' 참고)를 창고별로 나눠 호출해서 우회하는 데 쓴다 (2026-07-28 WH_CD 필터 동작 확인됨:
# WH_CD=00014 단독 조회 시 6676건, 전체 10000건보다 적고 정상 응답).
OFFLINE_WAREHOUSE_CODES = {
    "POINT OF VIEW (법인)": "00014",
    "THE HYUNDAI SEOUL": "00012",
    "MXN": "00030",
    "MXN(온라인)": "00033",
}

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

STRIPE_BG = {"red": 0.961, "green": 0.961, "blue": 0.961}    # #F5F5F5 — 데이터 줄무늬(아주 옅은 그레이,
# 2026-07-28 재조정: #FAF8F6은 너무 옅고, #F0EBE6은 너무 진했다는 피드백 — 무채색 그레이로 확정)
BODY_FG = {"red": 0.2, "green": 0.2, "blue": 0.2}            # #333333 — 순수 검정 금지

FONT_BODY = "Noto Sans KR"
FONT_MONO = "Roboto Mono"

SEMANTIC_DANGER_BG = {"red": 0.976, "green": 0.816, "blue": 0.816}   # F9D0D0 — 🔴
SEMANTIC_DANGER_FG = {"red": 0.800, "green": 0.000, "blue": 0.000}   # CC0000
SEMANTIC_ORANGE_BG = {"red": 0.988, "green": 0.878, "blue": 0.769}   # FDE0C4 — 🟠 (2026-07-28 신규,
SEMANTIC_ORANGE_FG = {"red": 0.902, "green": 0.451, "blue": 0.000}   # E67300  6색 동그라미 도입)
SEMANTIC_WARNING_BG = {"red": 1.000, "green": 0.976, "blue": 0.769}  # FFF9C4 — 🟡
SEMANTIC_WARNING_FG = {"red": 0.702, "green": 0.420, "blue": 0.000}  # B36B00
SEMANTIC_INFO_BG = {"red": 0.839, "green": 0.918, "blue": 0.973}     # D6EAF8 — 🔵 (과잉/과다재고)
SEMANTIC_INFO_FG = {"red": 0.084, "green": 0.396, "blue": 0.753}     # 1565C0
SEMANTIC_DIM_FG = {"red": 0.600, "green": 0.600, "blue": 0.600}      # 999999 — ⚫ (텍스트만, 배경 없음)
SEMANTIC_BROWN_BG = {"red": 0.929, "green": 0.855, "blue": 0.780}    # EDDAC7 — 🟤 (2026-07-28 신규,
SEMANTIC_BROWN_FG = {"red": 0.545, "green": 0.271, "blue": 0.075}   # 8B4513  품절-신규+지속 통합)

# "상태" 컬럼에 부분일치(SEARCH)로 색 입히는 규칙 — (부분문자열, 배경 or None, 글자색, 굵게).
# 문자열들이 서로 겹치지 않아 순서 무관. "정상"은 의도적으로 무색(색은 예외를 알리는 신호).
# 2026-07-28: 상태 이모지를 동그라미로 통일 — ⛔️ 마이너스재고(위험과 헷갈리지 않게 분리) →
# 🔴 위험 → 🟠 주의 → 🟡 재고소량 → 🔵 과잉 → 🟤 품절-지속(0~29일, 원래 품절-신규/지속 2단계였는데
# 표에서 흩어져 헷갈린다는 피드백으로 통합) → ⚫ 품절-장기(30일+, 조용히 죽은 상품).
STATUS_COLOR_RULES = [
    ("마이너스재고", SEMANTIC_DANGER_BG, SEMANTIC_DANGER_FG, True),
    ("위험", SEMANTIC_DANGER_BG, SEMANTIC_DANGER_FG, True),
    ("주의", SEMANTIC_ORANGE_BG, SEMANTIC_ORANGE_FG, True),
    ("품절-지속", SEMANTIC_BROWN_BG, SEMANTIC_BROWN_FG, True),
    ("재고소량", SEMANTIC_WARNING_BG, SEMANTIC_WARNING_FG, True),
    ("과잉", SEMANTIC_INFO_BG, SEMANTIC_INFO_FG, True),
    # 대시보드 악성재고 블록 라벨(2026-08-05 "과잉"→"미판매재고"로 수정)도 같은 파란 톤을
    # 유지하도록 별도 매칭 규칙 추가 — 안 넣으면 위 "과잉" 규칙만 걸리던 게 텍스트가 바뀌면서
    # 색이 안 먹는 회귀가 생긴다.
    ("미판매재고", SEMANTIC_INFO_BG, SEMANTIC_INFO_FG, True),
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
    "샘플의심재고": 3,
}

# 컬럼명 패턴 → 숫자서식. 0을 "–"로(0이 빽빽하면 "재고 없음"이 안 보임), 숫자는 Roboto Mono
# 고정폭 + 우측정렬로 자릿수를 맞춘다(재고량 스캔 속도가 가장 크게 체감되는 변경).
# "재고수량"은 RAW_재고현황(원본 필드명 그대로 유지)에만 남고, 나머지 탭은 "재고"로 줄인
# 이름을 쓴다 — 둘 다 넣어둬야 두 이름 다 숫자서식이 적용된다.
QTY_COLS = {"재고수량", "재고", "총재고", "전일재고", "입고", "출고",
            "7일 판매", "90일 판매", "수량", "공급가액", "부가세", "합계", "단가"}
DAY_COLS = {"DOI(소진일)", "품절(일)", "미판매(일)", "미입고(일)", "리드타임(일)"}
INT_COLS = {"우선순위"}
MONEY_COLS = {"재고금액", "입고단가"}
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
    #        품절-지속(0~29일째)/품절-장기(30일+)/과잉(DOI>180일)/마이너스재고(재고<0)/정상
    # 위험/주의 임계값: 자체제작 DOI≤30일/30~44일 · 국내사입 DOI≤7일/7~14일 · 해외수입 DOI≤21일/21~35일
    "일별재고이력": {
        "note": [
            "📋 품목별 일별 재고·입출고·DOI·상태 계산 이력 — 3층 결과물 탭들이 전부 여기서 계산됨",
            "직접 볼 일 없음 — 왜 그런 상태/조치가 나왔는지 근거 확인용",
            "정렬: 날짜 오름차순 → 품목코드순 · 🕒 매일 품목당 한 행씩 자동 누적",
        ],
        "headers": [
            "날짜", "브랜드", "품목코드", "품목명", "상태", "우선순위",
            "전일재고", "재고", "출고", "입고",
            "7일 판매", "90일 판매", "DOI(소진일)", "품절(일)", "조치",
        ],
    },
    # 3층 결과물 (용도별 뷰) — 온라인의 "위험/주의/재고소량·판매없음/신규품절/지속품절" 리스트에 대응.
    # 정렬: 위험·주의는 DOI(소진일) 오름차순, 품절-지속/장기는 90일 판매 내림차순(부활가치 큰 것 우선).
    # 재고 = 오프라인 4개 창고 합계 (창고별 상세는 관리팀_전체재고에서 확인 — 디자인팀은
    # "만들어야 하는가"만 판단하면 되므로 창고별 상세 불필요).
    "디자인팀_발주필요": {
        "note": [
            "📋 디자인팀용 — 자체제작 브랜드만, 지금 발주(제작)해야 하는 품목만 골라서 보여줌",
            "국내사입/해외수입 브랜드는 여기 안 뜸(디자인팀 소관 아님) — 전체 재고는 관리팀_전체재고 참고",
            "정렬: 우선순위 오름차순 → 재고 오름차순 · 🕒 마지막 갱신: 매일 자동",
        ],
        "headers": [
            "브랜드", "품목코드", "품목명", "조달유형", "상태", "우선순위",
            "리드타임(일)", "재고", "DOI(소진일)", "7일 판매", "90일 판매", "품절(일)",
            "조치", "메모",
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
            "전일재고", "입고", "출고", "7일 판매", "DOI(소진일)",
            "조치", "메모",
        ],
    },
    # 재고 있음(수량 무관) + 90일 이상 미판매 (재고 1개짜리도 포함 — 디자인팀_발주필요와 반대 축)
    "악성재고": {
        "note": [
            "📋 재고 있는데 90일 이상 안 팔린 품목(수량 무관) — 프로모션/폐기 판단용 (반대 축: 디자인팀_발주필요)",
            "재고금액 큰 것부터 프로모션 또는 폐기 결정",
            "정렬: 재고금액 내림차순 · 🕒 마지막 갱신: 매일 자동",
        ],
        "headers": ["브랜드", "품목코드", "품목명", "재고", "최근판매일", "미판매(일)", "입고단가", "재고금액", "조치", "메모"],
    },
    # 품절-장기 중에서도 최근90일판매량=0(부활가치 없음)인 것만. 판매량 있는 장기품절은
    # 디자인팀_발주필요의 "재입고 골든타임"으로 남긴다. 재발주/제작/단종 최종 판단용.
    # 컬럼은 "방치 근거"(품절경과일/조달유형/리드타임)와 "부활가치 근거"(최근입고일/미입고경과일/
    # 최근90일판매량)를 모두 담아서 왜 여기 있는지 판단할 정보를 다 준다.
    "악성품절": {
        "note": [
            "📋 재고 없이 90일 이상 방치된 품목(부활가치 없음) — 재발주/단종 최종 판단용",
            "품절구간(90일+/120일+/150일+/180일+/1년+/2년+)이 오래된 것부터 검토(2026-08-04 확정)",
            "정렬: 품절(일) 내림차순 · 🕒 마지막 갱신: 매일 자동",
        ],
        "headers": [
            "브랜드", "품목코드", "품목명", "조달유형", "리드타임(일)",
            "재고", "최근판매일", "미판매(일)", "품절(일)", "품절구간",
            "최근입고일", "미입고(일)", "90일 판매", "조치", "메모",
        ],
    },
    # 재고 1~2개인데 입고·판매 둘 다 3개월(90일) 이상 없는 품목 — 실재고가 아니라 진열/샘플용일
    # 가능성. 가용재고로 잘못 읽힐 수 있어 사람이 이카운트에서 직접 확인하도록 별도로 뽑는다
    # (2026-07-28 도입, 2026-07-30 기준을 60일→3개월로 상향 확정).
    # 이 시스템이 직접 쌓은 일별재고이력이 90일 이상 있어야 정확해진다 — 도입 초기엔 "미판매(일)/
    # 미입고(일)"을 판단할 근거 자체가 없어서 거의 비어 있는 게 정상(오탐 방지를 위해 근거 없으면
    # 아예 후보에서 제외, 빈 리스트=이상 없음이 아니라 아직 판단 불가일 수 있음에 유의).
    "샘플의심재고": {
        "note": [
            "📋 재고 1~2개 & 입고·판매 둘 다 3개월(90일) 이상 없음 — 진열/샘플용일 가능성, 실재고 아닐 수 있음",
            "이카운트에서 실제 샘플/전시용인지 확인 후 처리 (가용재고로 잘못 잡히지 않게)",
            "정렬: 미판매(일) 내림차순 · 🕒 마지막 갱신: 매일 자동 (3개월치 이력 쌓이기 전엔 대부분 공란)",
            "입고단가/재고금액은 2026-07-30 스냅샷 기준(단가는 SKU당 잘 안 바뀌어 당분간 유효)",
        ],
        "headers": [
            "브랜드", "품목코드", "품목명", "조달유형",
            "재고", "최근판매일", "미판매(일)", "최근입고일", "미입고(일)",
            "입고단가", "재고금액",
            "조치", "메모",
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
        # 맨 위(0행)는 탭 타이틀 — 대시보드와 동일하게 11pt·왼쪽정렬 고정(2026-08-06 사용자가
        # 모든 탭에서 직접 이렇게 맞춰뒀던 걸, 이 함수 재실행 시 9pt·기본정렬로 되돌리던 문제 수정).
        banner_rows = [
            (0, BANNER1_BG, BANNER_FG_LIGHT, True, 11, "LEFT"),
            (1, BANNER2_BG, BANNER_FG_LIGHT, True, 9, None),
            (2, BANNER3_BG, BANNER3_FG, False, 9, None),
        ]
        for row_idx, bg, fg, bold, font_size, halign in banner_rows:
            cell_format = {
                "backgroundColor": bg,
                "textFormat": {"bold": bold, "italic": not bold, "foregroundColor": fg, "fontFamily": FONT_BODY, "fontSize": font_size},
                "verticalAlignment": "MIDDLE",
            }
            fields = "userEnteredFormat(textFormat,backgroundColor,verticalAlignment)"
            if halign:
                cell_format["horizontalAlignment"] = halign
                fields = "userEnteredFormat(textFormat,backgroundColor,verticalAlignment,horizontalAlignment)"
            requests.append({
                "repeatCell": {
                    "range": {"sheetId": sheet_id, "startRowIndex": row_idx, "endRowIndex": row_idx + 1},
                    "cell": {"userEnteredFormat": cell_format},
                    "fields": fields,
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
                "cell": {"userEnteredFormat": {"textFormat": {"fontFamily": FONT_BODY, "foregroundColor": BODY_FG, "fontSize": 9}}},
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
                            "textFormat": {"fontFamily": FONT_MONO, "fontSize": 9},
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
                        "cell": {"userEnteredFormat": {"numberFormat": NUMBER_FORMAT_DATE, "horizontalAlignment": "RIGHT"}},
                        "fields": "userEnteredFormat(numberFormat,horizontalAlignment)",
                    }
                })

        # 악성재고/샘플의심재고는 조치 칸도 오른쪽 정렬(2026-08-06 사용자 요청 — 값 컬럼들과
        # 시각적으로 줄맞춤). 다른 탭의 조치 칸(디자인팀_발주필요 등)은 건드리지 않는다.
        if name in ("악성재고", "샘플의심재고") and "조치" in headers:
            action_col = headers.index("조치")
            requests.append({
                "repeatCell": {
                    "range": {"sheetId": sheet_id, "startRowIndex": DATA_START_IDX, "endRowIndex": 1000,
                               "startColumnIndex": action_col, "endColumnIndex": action_col + 1},
                    "cell": {"userEnteredFormat": {"horizontalAlignment": "RIGHT"}},
                    "fields": "userEnteredFormat.horizontalAlignment",
                }
            })

        # 총재고는 다른 숫자 컬럼(창고별 재고 등)과 구분되는 합계 컬럼이라 항상 볼드체로 강조
        # (2026-07-28 사용자 요청) — numeric_col_formats 다음에 둬야 폰트만 있는 그 서식을
        # 안 덮어쓰이고 볼드가 유지된다.
        if "총재고" in headers:
            total_col = headers.index("총재고")
            requests.append({
                "repeatCell": {
                    "range": {"sheetId": sheet_id, "startRowIndex": DATA_START_IDX, "endRowIndex": 1000,
                               "startColumnIndex": total_col, "endColumnIndex": total_col + 1},
                    "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontFamily": FONT_MONO, "fontSize": 9}}},
                    "fields": "userEnteredFormat.textFormat",
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

        for qty_col_name in ("재고수량", "재고", "총재고"):
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


# ---------------------------------------------------------------------------
# "읽는 법" 탭 — 전사 배포용 온보딩 탭. 처음 보는 사람이 색/이모지/용어를 몰라도
# 읽고 바로 쓸 수 있게 하는 게 목적(2026-07-28, 전사 배포 검토 중 지적 반영).
# TABS의 note+headers 패턴과 다른 자유형 레이아웃이라 별도 함수로 작성/관리한다.
# ---------------------------------------------------------------------------
READ_ME_TAB = "읽는 법"
_README_NCOLS = 5

SECTION_BG = {"red": 0.90, "green": 0.90, "blue": 0.90}
SECTION_FG = {"red": 0.133, "green": 0.133, "blue": 0.133}

# 상태 신호(③ 표) 배경/글자색 — STATUS_COLOR_RULES와 같은 팔레트를 재사용해 시트 본문의
# 실제 색과 "읽는 법"의 범례가 항상 같은 색을 쓰게 한다.
_CIRCLE_STYLE = {
    "⛔️": (SEMANTIC_DANGER_BG, SEMANTIC_DANGER_FG),
    "🔴": (SEMANTIC_DANGER_BG, SEMANTIC_DANGER_FG),
    "🟠": (SEMANTIC_ORANGE_BG, SEMANTIC_ORANGE_FG),
    "🟡": (SEMANTIC_WARNING_BG, SEMANTIC_WARNING_FG),
    "🔵": (SEMANTIC_INFO_BG, SEMANTIC_INFO_FG),
    "🟤": (SEMANTIC_BROWN_BG, SEMANTIC_BROWN_FG),
    "⚫": (None, SEMANTIC_DIM_FG),
    "🟢": (None, BODY_FG),
}


def _readme_content() -> list[tuple[str, list[str]]]:
    """(행 종류, 셀 값 5개) 목록. 행 종류에 따라 write_read_me_tab이 다르게 서식을 입힌다."""
    rows: list[tuple[str, list[str]]] = []
    pad = lambda *vals: list(vals) + [""] * (_README_NCOLS - len(vals))  # noqa: E731

    rows.append(("banner1", pad("처음 보시는 분은 이 탭부터 읽어 주세요")))
    rows.append(("banner2", pad("오프라인(이카운트) 재고를 매일 자동으로 판정해 보여주는 시스템입니다")))
    rows.append(("blank", pad()))

    rows.append(("section", pad("① 하루에 이렇게 씁니다")))
    rows.append(("text", pad("1. 대시보드 — 오늘 처리할 목록이 급한 순서대로 블록에 나뉘어 있습니다")))
    rows.append(("text", pad("2. 블록 제목 오른쪽에 \"무엇을 하라\"가 적혀 있습니다")))
    rows.append(("text", pad("3. 블록 우측 '출처'에 어느 탭에서 왔는지 적혀 있습니다")))
    rows.append(("blank", pad()))

    rows.append(("section", pad("② 어느 탭을 보면 되나")))
    rows.append(("tablehdr", pad("탭", "누가", "언제", "수정", "무엇을 보나")))
    for r in [
        ["대시보드", "전원", "매일", "읽기만", "오늘 처리할 것 전부 · 여기서 시작하세요"],
        ["관리팀_전체재고", "관리팀", "매일", "메모만", "창고별 재고 · 발주와 매장 간 이동 결정"],
        ["디자인팀_발주필요", "디자인팀", "매일", "메모만", "제작 발주가 필요한 품목만"],
        ["악성재고", "관리팀", "주 1회", "메모만", "재고가 쌓여 도는 품목"],
        ["악성품절", "관리팀", "월 1회", "메모만", "되살릴지 접을지 결정할 품목"],
        ["품목마스터", "담당자", "새 브랜드 등장 시만", "직접 수정", "품목 ↔ 브랜드 ↔ 조달유형 ↔ 리드타임 (같은 브랜드 신상품은 자동 등록됨, '미분류'만 확인)"],
        ["일별재고이력", "—", "—", "수정 금지", "모든 판정이 계산되는 곳"],
        ["RAW_재고현황", "—", "—", "수정 금지", "이카운트 재고 원본"],
        ["RAW_판매현황", "—", "—", "수정 금지", "이카운트 판매 원본"],
    ]:
        rows.append(("tablerow", r))
    rows.append(("blank", pad()))

    rows.append(("section", pad("③ 동그라미가 뜻하는 것")))
    rows.append(("tablehdr", pad("상태", "언제까지", "조치")))
    for status_text, deadline, action in [
        ("⛔️ 마이너스재고", "오늘", "재고가 음수예요 — 이카운트 전표부터 확인하세요!"),
        ("🔴 위험", "오늘", "지금 발주해도 늦어요 — 바로 발주하세요!"),
        ("🟠 주의", "이번 주", "아직 여유 있지만 이번 주 안에 발주하세요!"),
        ("🟡 재고소량", "여유 있을 때", "재고는 적은데 7일간 안 팔렸어요 — 노출이 막혔는지 확인하세요!"),
        ("🔵 과잉", "월 1회", "180일치 넘게 쌓였어요 — 프로모션·번들을 검토하세요!"),
        ("🟤 품절-지속(0~29일)", "이번 주", "지금 채우면 매출을 살릴 수 있어요 — 재입고 여부를 결정하세요!"),
        ("⚫ 품절-장기(30일+)", "월 1회", "재발주할지 단종할지 결정하세요!"),
        ("🟢 정상", "—", "지금 할 일 없어요"),
    ]:
        rows.append(("statusrow", pad(status_text, deadline, action)))
    rows.append(("blank", pad()))

    rows.append(("section", pad("④ 표에 나오는 말")))
    rows.append(("tablehdr", pad("용어", "뜻")))
    for term, desc in [
        ("DOI(소진일)", "지금 재고가 며칠 뒤에 바닥나는지. 재고 ÷ 최근 7일 하루평균 판매량. 예: 재고 62개를 하루 2개씩 팔면 31일"),
        ("리드타임(일)", "발주해서 물건이 실제로 들어오기까지 걸리는 날짜"),
        ("조달유형", "자체제작 / 국내사입 / 해외수입. 리드타임이 다르므로 위험 판정 기준도 다릅니다"),
        ("총재고", "창고 4곳의 재고를 더한 값. 한 곳이라도 0이면 다른 곳에 남아있을 수 있습니다"),
        ("7일 판매 / 90일 판매", "최근 7일 · 90일 동안 팔린 수량. 7일은 지금 속도, 90일은 원래 속도"),
        ("재고금액", "남은 재고 × 입고단가. 이 품목에 묶여 있는 돈"),
        ("품절(일) / 미판매(일)", "품절된 지 며칠 / 마지막으로 팔린 지 며칠"),
        ("우선순위", "1이 가장 급합니다. 같은 상태 안에서 처리 순서를 정할 때 씁니다"),
    ]:
        rows.append(("glossaryrow", pad(term, desc)))
    rows.append(("blank", pad()))

    rows.append(("section", pad("⑤ 지켜 주세요")))
    for b in [
        "회색으로 표시된 탭(일별재고이력·RAW 2개)은 스크립트가 매일 다시 씁니다. 직접 고쳐도 다음날 사라집니다.",
        "나머지 탭에서도 값은 자동으로 채워집니다. 사람이 적는 곳은 메모 컬럼뿐입니다.",
        "품목마스터는 반대로 사람이 쓰는 탭입니다. 여기 적은 내용은 지워지지 않습니다.",
        "숫자가 이상해 보이면 고치지 마시고 RAW 탭의 마지막 갱신 시각부터 확인해 주세요.",
        "문의처: (담당자 이름과 슬랙 채널을 알려주시면 이 자리에 채워넣습니다) — 이 시트가 이상하면 여기로 연락하세요.",
    ]:
        rows.append(("text", pad(f"· {b}")))

    return rows


def write_read_me_tab(sheets_service, spreadsheet_id: str, sheet_id: int) -> None:
    content = _readme_content()
    values = [v for _, v in content]
    nrows = len(content)

    sheets_service.spreadsheets().values().batchClear(
        spreadsheetId=spreadsheet_id, body={"ranges": [f"'{READ_ME_TAB}'!A1:Z500"]}
    ).execute()
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=f"'{READ_ME_TAB}'!A1",
        valueInputOption="RAW", body={"values": values},
    ).execute()

    # 재실행 안전 — 기존 병합/줄무늬 지우고 새로 입힌다.
    meta = sheets_service.spreadsheets().get(
        spreadsheetId=spreadsheet_id, fields="sheets(properties(sheetId,title),bandedRanges(bandedRangeId))"
    ).execute()
    delete_requests = []
    for s in meta["sheets"]:
        if s["properties"]["title"] != READ_ME_TAB:
            continue
        delete_requests.append({
            "unmergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": nrows,
                                        "startColumnIndex": 0, "endColumnIndex": _README_NCOLS}},
        })
        for banded in s.get("bandedRanges", []):
            delete_requests.append({"deleteBanding": {"bandedRangeId": banded["bandedRangeId"]}})
    if delete_requests:
        sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": delete_requests}).execute()

    requests = []
    # 기본 폰트/줄바꿈부터 전체에 깔고, 아래서 행별로 덮어쓴다(나중 요청이 우선 적용).
    requests.append({
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": nrows,
                       "startColumnIndex": 0, "endColumnIndex": _README_NCOLS},
            "cell": {"userEnteredFormat": {
                "textFormat": {"fontFamily": FONT_BODY, "fontSize": 9, "foregroundColor": BODY_FG},
                "wrapStrategy": "WRAP", "verticalAlignment": "MIDDLE",
            }},
            "fields": "userEnteredFormat(textFormat,wrapStrategy,verticalAlignment)",
        }
    })

    def merge_row(r0: int, bg: dict, fg: dict, bold: bool, size: int = 10, align: str = "LEFT") -> None:
        requests.append({
            "mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r0 + 1,
                                      "startColumnIndex": 0, "endColumnIndex": _README_NCOLS}, "mergeType": "MERGE_ALL"},
        })
        fmt = {"backgroundColor": bg, "textFormat": {"bold": bold, "foregroundColor": fg, "fontFamily": FONT_BODY, "fontSize": size},
               "verticalAlignment": "MIDDLE", "horizontalAlignment": align}
        requests.append({
            "repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r0 + 1,
                                      "startColumnIndex": 0, "endColumnIndex": _README_NCOLS},
                            "cell": {"userEnteredFormat": fmt},
                            "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment,horizontalAlignment)"},
        })

    for i, (rtype, vals) in enumerate(content):
        if rtype == "banner1":
            merge_row(i, BANNER1_BG, BANNER_FG_LIGHT, True, size=12, align="CENTER")
        elif rtype == "banner2":
            merge_row(i, BANNER3_BG, BANNER3_FG, False, size=9, align="LEFT")
        elif rtype == "section":
            merge_row(i, SECTION_BG, SECTION_FG, True, size=11, align="LEFT")
        elif rtype == "tablehdr":
            requests.append({
                "repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": i, "endRowIndex": i + 1,
                                          "startColumnIndex": 0, "endColumnIndex": _README_NCOLS},
                                "cell": {"userEnteredFormat": {"backgroundColor": HEADER_BG,
                                         "textFormat": {"bold": True, "foregroundColor": HEADER_FG, "fontFamily": FONT_BODY, "fontSize": 9}}},
                                "fields": "userEnteredFormat(backgroundColor,textFormat)"},
            })
        elif rtype == "statusrow":
            # vals[0]은 "⛔️ 마이너스재고"처럼 이모지+상태이름이 합쳐진 문자열 — 앞쪽 이모지로
            # _CIRCLE_STYLE에서 색을 찾는다(신호/상태이름을 별도 컬럼으로 안 나누고 한 셀로 합침,
            # 2026-07-28 사용자 요청: 두 컬럼이 항상 같이 봐야 하는 정보라 분리할 이유가 없었음).
            bg, fg = next((v for k, v in _CIRCLE_STYLE.items() if vals[0].startswith(k)), (None, BODY_FG))
            fmt = {"textFormat": {"foregroundColor": fg, "fontFamily": FONT_BODY, "fontSize": 9, "bold": bg is not None}}
            if bg is not None:
                fmt["backgroundColor"] = bg
            requests.append({
                "repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": i, "endRowIndex": i + 1,
                                          "startColumnIndex": 0, "endColumnIndex": _README_NCOLS},
                                "cell": {"userEnteredFormat": fmt},
                                "fields": "userEnteredFormat(backgroundColor,textFormat)"},
            })
            # 조치(C~E열)는 텍스트가 기니까 병합해서 한 칸처럼 보이게.
            requests.append({
                "mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": i, "endRowIndex": i + 1,
                                          "startColumnIndex": 2, "endColumnIndex": _README_NCOLS}, "mergeType": "MERGE_ALL"},
            })
        elif rtype == "text":
            # 순서 안내(①)처럼 한 줄짜리 문장 행 — A열에만 텍스트가 있고 B~E는 비어서
            # 병합 안 하면 어색해 보인다는 사용자 피드백으로 전체 폭 병합(2026-07-28).
            requests.append({
                "mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": i, "endRowIndex": i + 1,
                                          "startColumnIndex": 0, "endColumnIndex": _README_NCOLS}, "mergeType": "MERGE_ALL"},
            })
        elif rtype == "glossaryrow":
            requests.append({
                "mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": i, "endRowIndex": i + 1,
                                          "startColumnIndex": 1, "endColumnIndex": _README_NCOLS}, "mergeType": "MERGE_ALL"},
            })
            requests.append({
                "repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": i, "endRowIndex": i + 1,
                                          "startColumnIndex": 0, "endColumnIndex": 1},
                                "cell": {"userEnteredFormat": {"textFormat": {"bold": True, "fontFamily": FONT_BODY, "fontSize": 9}}},
                                "fields": "userEnteredFormat.textFormat"},
            })

    # 열 너비 — A(용어/탭이름) 좁게, B~D 중간, E(설명) 넓게.
    requests.append({"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 1},
                                                     "properties": {"pixelSize": 170}, "fields": "pixelSize"}})
    requests.append({"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 1, "endIndex": 4},
                                                     "properties": {"pixelSize": 100}, "fields": "pixelSize"}})
    requests.append({"updateDimensionProperties": {"range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": 4, "endIndex": 5},
                                                     "properties": {"pixelSize": 430}, "fields": "pixelSize"}})

    try:
        sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()
    except Exception as e:
        print(f"[sheets] 읽는 법 탭 서식 적용 중 일부 실패(무시하고 진행): {e}")


# 탭 순서 — 사람이 매일 보는 결과물이 앞, 원본/계산용은 뒤로 (2026-07-28, 전사 배포 검토 반영).
# 존재하지 않는 탭은 reorder_tabs가 건너뛰므로 아직 안 만든 탭(샘플의심재고 등)을 미리 적어둬도 안전.
TAB_ORDER = [
    READ_ME_TAB, "대시보드", "관리팀_전체재고", "디자인팀_발주필요",
    "악성재고", "악성품절", "샘플의심재고", "일별재고이력", "품목마스터",
    "RAW_재고현황", "RAW_판매현황",
]


def reorder_tabs(sheets_service, spreadsheet_id: str) -> None:
    meta = sheets_service.spreadsheets().get(
        spreadsheetId=spreadsheet_id, fields="sheets.properties(sheetId,title)"
    ).execute()
    id_by_title = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}
    requests = []
    idx = 0
    for name in TAB_ORDER:
        if name not in id_by_title:
            continue
        requests.append({"updateSheetProperties": {"properties": {"sheetId": id_by_title[name], "index": idx}, "fields": "index"}})
        idx += 1
    if requests:
        sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()


def _ensure_dashboard_placeholder(sheets_service, spreadsheet_id: str, sheet_id: int) -> None:
    """대시보드는 매일 daily_runner가 통째로 다시 쓰므로, 여기선 최초 1회용 안내 배너만 넣는다."""
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range="'대시보드'!A1",
        valueInputOption="RAW",
        body={"values": [["아직 실행 전 — ecount_daily_runner.py 최초 실행 후 매일 자동으로 채워집니다"]]},
    ).execute()
    requests = [{
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 11},
            "cell": {"userEnteredFormat": {
                "backgroundColor": BANNER1_BG,
                "textFormat": {"bold": True, "foregroundColor": BANNER_FG_LIGHT, "fontFamily": FONT_BODY, "fontSize": 10},
                "verticalAlignment": "MIDDLE",
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat,verticalAlignment)",
        }
    }]
    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()


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

    # 읽는 법 / 대시보드 — TABS의 note+headers 패턴과 다른 특수 탭. 존재만 보장하고 내용은 각자 채운다.
    have = existing_tab_names(sheets_service, spreadsheet_id)
    extra_added = []
    for name in (READ_ME_TAB, "대시보드"):
        if name not in have:
            sheets_service.spreadsheets().batchUpdate(
                spreadsheetId=spreadsheet_id, body={"requests": [{"addSheet": {"properties": {"title": name}}}]}
            ).execute()
            extra_added.append(name)
    if extra_added:
        added = added + extra_added

    meta = sheets_service.spreadsheets().get(
        spreadsheetId=spreadsheet_id, fields="sheets.properties(sheetId,title)"
    ).execute()
    id_by_title = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}
    write_read_me_tab(sheets_service, spreadsheet_id, id_by_title[READ_ME_TAB])
    if "대시보드" in extra_added:
        _ensure_dashboard_placeholder(sheets_service, spreadsheet_id, id_by_title["대시보드"])

    reorder_tabs(sheets_service, spreadsheet_id)

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
