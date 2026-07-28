#!/usr/bin/env python3
"""오프라인 재고 감시 시스템 — 매일 자동 실행되는 통합 스크립트.

흐름:
  1. 이카운트 창고별재고현황 API로 오늘 시점 재고 스냅샷 조회 (전체 창고, RAW 그대로 저장)
  2. Gmail에서 이카운트 "판매현황" 자동발송 이메일의 첨부 엑셀을 받아 파싱 (전체 창고, RAW 그대로 저장)
  3. 두 원본을 오프라인 4개 창고로 필터링 + 품목마스터(브랜드/조달유형/리드타임) 조인
  4. 일별재고이력에서 전일 데이터를 읽어 전일재고·품절경과일 연속성 계산, 입고수량 역산
  5. DOI·상태·우선순위·조치방안 계산 (README "DOI 기반 우선순위 체계" 참고)
  6. 일별재고이력에 오늘자 행 누적, 3층 결과물 탭(디자인팀_발주필요/관리팀_전체재고/악성재고/
     악성품절) 재작성

기준일(TARGET_DATE) 하나로 통일: 이카운트 판매현황 자동알림의 "기준일자=전일" 관례를 그대로
따라, 이 스크립트가 실행되는 날의 "어제"를 재고/판매 양쪽의 기준일로 쓴다.

사용법 (서버에서):
  .venv/bin/python ecount_daily_runner.py --spreadsheet-id <ID>
  .venv/bin/python ecount_daily_runner.py --spreadsheet-id <ID> --sales-xlsx cron_tracking/ecount/sales.xlsx
  .venv/bin/python ecount_daily_runner.py --spreadsheet-id <ID> --dry-run   # 시트에 안 쓰고 요약만 출력
"""
from __future__ import annotations

import argparse
import time
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from openpyxl import load_workbook

from ecount_client import EcountClient
from ecount_item_master import LEADTIME_BY_TYPE
from ecount_sheets_setup import (
    TABS, OFFLINE_WAREHOUSES, OFFLINE_WAREHOUSE_CODES, DATA_START_IDX, _TOKEN_FILE, SCOPES,
    STATUS_COLOR_RULES, BANNER1_BG, BANNER_FG_LIGHT, BANNER3_BG, BANNER3_FG,
    HEADER_BG, HEADER_FG, FONT_BODY, FONT_MONO, BODY_FG,
    SEMANTIC_DANGER_BG, SEMANTIC_INFO_BG, SEMANTIC_WARNING_BG,
)

KST = timezone(timedelta(hours=9))
DUMP_DIR = Path(__file__).parent / "cron_tracking" / "ecount"

# 이카운트 실서버 조회 API(재고현황/창고별재고현황 포함)는 종류당 1회/10분 제한(2026-07-28
# 공식 문서 확인, HTTP 412 = "API 전송 횟수 기준을 넘은 경우"). 오프라인 창고 4곳을 각각
# 조회하려면 그만큼 간격을 둬야 한다 — 10분(600초) + 여유.
INVENTORY_CALL_INTERVAL_SEC = 610

# DOI 위험/주의 임계값 (일) — README "DOI 기반 우선순위 체계" 표와 동일.
RISK_WARN_BY_TYPE = {
    "자체제작": (35, 49),
    "국내사입": (7, 14),
    "해외수입": (21, 35),
}

PRIORITY_BY_STATUS = {
    "🔴 위험": 1, "🔴 품절-신규": 1,
    "🟠 주의": 2, "🟠 품절-지속": 2,
    "🟡 재고소량": 3,
    "⚫ 품절-장기": 3,
}

ACTION_BY_STATUS = {
    "🔴 마이너스재고": "재고 데이터 확인 필요",
    "🔴 위험": "긴급 제작 필요",
    "🔴 품절-신규": "긴급 제작 필요",
    "🟠 주의": "제작 검토",
    "🟠 품절-지속": "제작 검토",
    "🟡 재고소량": "마케팅 부스트 또는 정리 검토",
    "🔵 과잉": "프로모션/할인 검토",
    "🟢 정상": "-",
}


# ---------------------------------------------------------------------------
# 1. 이카운트 재고 API
# ---------------------------------------------------------------------------

def fetch_inventory_raw(base_date: str) -> list[dict]:
    """창고별재고현황 API 원본 — 오프라인 4개 창고만, 창고별로 나눠서 호출.

    전체 창고를 한 번에 조회하면 10000건 근처에서 잘리는 것으로 의심되는 문제가 있었다
    (README '탐사 결과' 참고). WH_CD로 창고를 하나씩 지정해 호출하면 건수가 훨씬 적게
    나오는 것을 실제로 확인(2026-07-28, WH_CD=00014 단독 6676건 < 전체 10000건) —
    이 시스템은 애초에 오프라인 4개 창고 외엔 안 쓰므로, 온라인 창고는 아예 조회하지 않는다.

    단, 실서버 조회 API는 종류당 1회/10분 제한이 있어(공식 문서 확인, 2026-07-28) 창고 4개를
    연달아 부르면 두 번째부터 무조건 막힌다. 그래서 호출 사이에 실제로 대기한다 — 총 소요시간이
    30~40분이지만, 하루 한 번 도는 배치라 문제없다.
    """
    client = EcountClient()
    rows = []
    for i, (wh_name, wh_code) in enumerate(OFFLINE_WAREHOUSE_CODES.items()):
        if i > 0:
            print(f"[runner] 조회 API 1회/10분 제한 — 다음 창고까지 {INVENTORY_CALL_INTERVAL_SEC}초 대기...")
            time.sleep(INVENTORY_CALL_INTERVAL_SEC)
        print(f"[runner] {wh_name}(WH_CD={wh_code}) 조회 중...")
        data = client.inventory_balance_by_location(base_date, WH_CD=wh_code)
        result = data.get("Data", {}).get("Result") or []
        if len(result) >= 10000:
            print(f"[runner] ⚠️ {wh_name}(WH_CD={wh_code}) 응답이 {len(result)}건 — 10000건 "
                  "근처라 이 창고도 페이지네이션으로 잘렸을 가능성이 있습니다. 확인 필요.")
        for r in result:
            try:
                qty = float(r.get("BAL_QTY") or 0)
            except ValueError:
                qty = 0.0
            rows.append({
                "창고코드": r.get("WH_CD", wh_code),
                "창고명": r.get("WH_DES", wh_name),
                "품목코드": r.get("PROD_CD", ""),
                "품목명": r.get("PROD_DES", ""),
                "사이즈": r.get("PROD_SIZE_DES", ""),
                "재고수량": qty,
            })
    return rows


# ---------------------------------------------------------------------------
# 2. Gmail 판매현황 첨부 엑셀
# ---------------------------------------------------------------------------

def fetch_sales_xlsx_from_gmail(query: str, out_path: Path) -> Path:
    from ecount_gmail_fetch import _load_creds, list_messages, download_first_attachment
    from googleapiclient.discovery import build

    creds = _load_creds()
    service = build("gmail", "v1", credentials=creds)
    messages = list_messages(service, query)
    if not messages:
        raise SystemExit(f"[runner] Gmail 쿼리에 매치되는 메일이 없습니다: {query!r}")
    saved = download_first_attachment(service, messages[0]["id"], out_path)
    if not saved:
        raise SystemExit("[runner] 최신 메일에서 엑셀 첨부를 못 찾았습니다.")
    return saved


def _clean(v) -> str:
    return (str(v).strip() if v is not None else "")


def _to_number(v) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        return 0.0


SALES_HEADER_MAP = {
    "품목그룹2명": "브랜드", "브랜드": "브랜드",
    "품목코드": "품목코드",
    "품명 및 규격": "품명", "품명": "품명",
    "수량": "수량", "단가": "단가", "공급가액": "공급가액", "부가세": "부가세",
    "합계": "합계", "적요": "적요", "창고명": "창고명",
    "사원(담당)명": "담당자", "담당자": "담당자",
}


def parse_sales_xlsx(path: Path) -> list[dict]:
    """판매현황 엑셀(RAW_판매현황 원본) 파싱. 헤더 행을 자동으로 찾는다."""
    wb = load_workbook(path, data_only=True)
    ws = wb.active

    header_row = None
    for r in range(1, min(ws.max_row, 10) + 1):
        vals = [_clean(c.value) for c in ws[r]]
        if "품목코드" in vals and "수량" in vals:
            header_row = r
            break
    if header_row is None:
        raise SystemExit("[runner] 판매현황 엑셀에서 헤더 행(품목코드/수량)을 못 찾았습니다.")

    headers = [_clean(c.value) for c in ws[header_row]]
    col = {SALES_HEADER_MAP[h]: i for i, h in enumerate(headers) if h in SALES_HEADER_MAP}
    if "품목코드" not in col:
        raise SystemExit("[runner] 판매현황 엑셀에 품목코드 컬럼이 없습니다.")

    rows = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        def get(field, default=""):
            idx = col.get(field)
            return row[idx] if idx is not None and idx < len(row) else default

        code = _clean(get("품목코드"))
        if not code:
            continue
        rows.append({
            "브랜드": _clean(get("브랜드")),
            "품목코드": code,
            "품명": _clean(get("품명")),
            "수량": _to_number(get("수량")),
            "단가": _to_number(get("단가")),
            "공급가액": _to_number(get("공급가액")),
            "부가세": _to_number(get("부가세")),
            "합계": _to_number(get("합계")),
            "적요": _clean(get("적요")),
            "창고명": _clean(get("창고명")),
            "담당자": _clean(get("담당자")),
        })
    return rows


# ---------------------------------------------------------------------------
# 3. 구글시트 읽기/쓰기 helpers
# ---------------------------------------------------------------------------

def _load_creds():
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials

    creds = Credentials.from_authorized_user_file(str(_TOKEN_FILE), SCOPES)
    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        _TOKEN_FILE.write_text(creds.to_json())
    return creds


def read_tab_rows(service, spreadsheet_id: str, tab_name: str) -> list[dict]:
    headers = TABS[tab_name]["headers"]
    last_col = chr(ord("A") + len(headers) - 1) if len(headers) <= 26 else "Z"
    rng = f"'{tab_name}'!A{DATA_START_IDX + 1}:{last_col}200000"
    # UNFORMATTED_VALUE 필수 — 안 그러면 숫자서식이 적용된 표시문자열("35일", "–")이 그대로
    # 와서 파싱이 깨진다 (예: 품절경과일=0인 셀은 서식상 "–"로 보이지만 실제 값은 0이어야 함).
    resp = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=rng, valueRenderOption="UNFORMATTED_VALUE"
    ).execute()
    values = resp.get("values", [])
    out = []
    for row in values:
        row = row + [""] * (len(headers) - len(row))
        out.append(dict(zip(headers, row)))
    return out


def replace_tab_rows(service, spreadsheet_id: str, tab_name: str, rows: list[list]) -> None:
    service.spreadsheets().values().batchClear(
        spreadsheetId=spreadsheet_id, body={"ranges": [f"'{tab_name}'!A{DATA_START_IDX + 1}:Z200000"]}
    ).execute()
    if not rows:
        return
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=f"'{tab_name}'!A{DATA_START_IDX + 1}",
        valueInputOption="RAW",
        body={"values": rows},
    ).execute()


def append_history_rows(service, spreadsheet_id: str, target_date_str: str, new_rows: list[list]) -> None:
    """일별재고이력은 누적이지만, 같은 날짜로 재실행하면 그 날짜 행만 지우고 다시 쓴다(재실행 안전)."""
    headers = TABS["일별재고이력"]["headers"]
    existing = read_tab_rows(service, spreadsheet_id, "일별재고이력")
    kept = [[r[h] for h in headers] for r in existing if r["날짜"] != target_date_str]
    all_rows = kept + new_rows
    replace_tab_rows(service, spreadsheet_id, "일별재고이력", all_rows)


def auto_register_new_items(item_master: dict, sales_raw: list[dict], target_date: date) -> list[list]:
    """오늘 판매현황에 새로 나타난(품목마스터에 없는) 품목코드를 자동 등록한다.

    조달유형은 SKU가 아니라 브랜드 단위 속성이므로(README 참고), 같은 브랜드의 기존
    품목마스터 항목에서 다수결로 물려받는다 — 사람이 매번 새 상품을 등록할 필요가 없다.
    브랜드 자체가 처음 등장하는 경우만 '미분류'로 남아 사람이 한 번 확인하면 되고, 그 뒤로는
    같은 브랜드의 모든 신상품에 자동 적용된다(이카운트 원본에 브랜드코드가 없어 국가코드
    휴리스틱은 못 쓰지만, 브랜드명 다수결만으로 충분).

    item_master는 in-place로 갱신되어 오늘 계산부터 바로 반영되고, 반환값은 품목마스터
    시트에 추가로 써야 할 행이다(기존 행은 건드리지 않고 append).
    """
    brand_type_votes: dict[str, dict[str, int]] = {}
    for meta in item_master.values():
        if not meta.get("브랜드") or not meta.get("조달유형") or meta["조달유형"] == "미분류":
            continue
        votes = brand_type_votes.setdefault(meta["브랜드"], {})
        votes[meta["조달유형"]] = votes.get(meta["조달유형"], 0) + 1

    new_rows: list[list] = []
    seen_codes: set[str] = set()
    for r in sales_raw:
        code = r["품목코드"]
        if not code or code in item_master or code in seen_codes:
            continue
        seen_codes.add(code)
        brand = r["브랜드"]
        votes = brand_type_votes.get(brand)
        if votes:
            ptype = max(votes, key=votes.get)
            leadtime = LEADTIME_BY_TYPE.get(ptype, "")
        else:
            ptype, leadtime = "미분류", ""
        item_master[code] = {"브랜드": brand, "조달유형": ptype, "리드타임": leadtime, "품목명": r["품명"]}
        new_rows.append([code, r["품명"], brand, "", ptype, leadtime, target_date.isoformat()])
    return new_rows


def append_item_master_rows(service, spreadsheet_id: str, rows: list[list]) -> None:
    if not rows:
        return
    service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=f"'품목마스터'!A{DATA_START_IDX + 1}",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": rows},
    ).execute()


# ---------------------------------------------------------------------------
# 4. 계산 로직
# ---------------------------------------------------------------------------

def determine_status(재고: float, 최근7일: float, 품절경과일: int, 조달유형: str | None) -> str:
    # 상태 이모지는 상태별 개별 아이콘이 아니라 6색 동그라미로 통일한다(2026-07-28, 전사 배포용
    # "읽는 법" 탭 도입과 함께 확정): 🔴 지금/오늘 · 🟠 곧/이번주 · 🟡 확인 · 🔵 과잉 · ⚫ 끝(방치) ·
    # 🟢 정상. STATUS_COLOR_RULES의 조건부서식은 이모지가 아니라 한글 부분일치라 영향 없음.
    if 재고 < 0:
        return "🔴 마이너스재고"
    if 재고 == 0:
        if 품절경과일 <= 9:
            return "🔴 품절-신규"
        if 품절경과일 <= 29:
            return "🟠 품절-지속"
        return "⚫ 품절-장기"
    if 재고 <= 5 and 최근7일 == 0:
        return "🟡 재고소량"
    if 최근7일 > 0:
        doi = 재고 / (최근7일 / 7)
        risk, warn = RISK_WARN_BY_TYPE.get(조달유형 or "", (None, None))
        if risk is not None:
            if doi <= risk:
                return "🔴 위험"
            if doi <= warn:
                return "🟠 주의"
        if doi > 180:
            return "🔵 과잉"
        return "🟢 정상"
    return "🟢 정상"


def compute_doi(재고: float, 최근7일: float) -> float | str:
    # 재고 0 이하(품절/마이너스재고)일 때도 빈칸 — 안 그러면 마이너스재고(음수 재고)에서 DOI가 음수로
    # 나와 의미 없는 값이 찍힌다 (설계 스펙 "POV_재고관리_마스터_설계.md" 명시 사항).
    if 최근7일 <= 0 or 재고 <= 0:
        return ""
    return round(재고 / (최근7일 / 7), 1)


def build_daily_rows(target_date: date, inventory_raw: list[dict], sales_raw: list[dict],
                      item_master: dict[str, dict], history: list[dict]) -> dict:
    """오늘자 일별재고이력 행들과 3층 결과물을 계산해서 반환."""
    target_str = target_date.isoformat()
    prev_str = (target_date - timedelta(days=1)).isoformat()

    # 오프라인 4개 창고로 필터링한 재고 — 품목코드별 총재고 + 창고별 재고.
    offline_inv = [r for r in inventory_raw if r["창고명"] in OFFLINE_WAREHOUSES]
    total_by_item: dict[str, float] = {}
    by_item_wh: dict[str, dict[str, float]] = {}
    name_by_item: dict[str, str] = {}
    for r in offline_inv:
        code = r["품목코드"]
        total_by_item[code] = total_by_item.get(code, 0.0) + r["재고수량"]
        by_item_wh.setdefault(code, {})
        by_item_wh[code][r["창고명"]] = by_item_wh[code].get(r["창고명"], 0.0) + r["재고수량"]
        if r["품목명"]:
            name_by_item[code] = r["품목명"]

    # 오프라인 4개 창고로 필터링한 판매(출고) — 품목코드별 합계.
    offline_sales = [r for r in sales_raw if r["창고명"] in OFFLINE_WAREHOUSES]
    out_by_item: dict[str, float] = {}
    for r in offline_sales:
        code = r["품목코드"]
        out_by_item[code] = out_by_item.get(code, 0.0) + r["수량"]
        if r["품명"] and code not in name_by_item:
            name_by_item[code] = r["품명"]

    # 이력 인덱싱 — 품목코드별 (날짜 -> row).
    hist_by_item: dict[str, dict[str, dict]] = {}
    for r in history:
        hist_by_item.setdefault(r["품목코드"], {})[r["날짜"]] = r

    # 처리 대상 품목 = 품목마스터 전체(관리팀 전체 그림) ∪ 오늘 재고/판매에 등장한 품목(미분류 포함).
    all_codes = set(item_master) | set(total_by_item) | set(out_by_item)

    history_rows: list[list] = []
    design_rows: list[list] = []
    mgmt_rows: list[list] = []
    malstock_rows: list[list] = []
    maldead_rows: list[list] = []

    for code in sorted(all_codes):
        meta = item_master.get(code, {})
        브랜드 = meta.get("브랜드", "미분류")
        조달유형 = meta.get("조달유형", "미분류")
        리드타임 = meta.get("리드타임", "")
        품목명 = name_by_item.get(code, meta.get("품목명", ""))

        재고 = total_by_item.get(code, 0.0)
        출고 = out_by_item.get(code, 0.0)
        prev_row = hist_by_item.get(code, {}).get(prev_str)
        전일재고 = _to_number(prev_row["재고"]) if prev_row else 0.0
        입고계산 = (재고 - 전일재고) + 출고

        prev_stockout_days = int(_to_number(prev_row["품절(일)"])) if prev_row and prev_row.get("품절(일)") not in ("", None) else 0
        if 재고 <= 0:
            # prev_row가 아예 없는(처음 등장하는) 품목은 "어제도 품절이었다"고 이어받을 근거가
            # 없으므로 무조건 0일차(처음 확인)로 시작한다.
            품절경과일 = (prev_stockout_days + 1) if (prev_row is not None and 전일재고 <= 0) else 0
        else:
            품절경과일 = 0

        # 최근 7일/90일 판매량 = 과거 이력(오늘 제외) + 오늘 출고.
        item_hist = hist_by_item.get(code, {})
        최근7일 = 출고
        최근90일 = 출고
        for delta in range(1, 90):
            d = (target_date - timedelta(days=delta)).isoformat()
            row = item_hist.get(d)
            if not row:
                continue
            qty = _to_number(row.get("출고"))
            if delta <= 6:
                최근7일 += qty
            최근90일 += qty

        상태 = determine_status(재고, 최근7일, 품절경과일, 조달유형 if 조달유형 != "미분류" else None)
        doi = compute_doi(재고, 최근7일)
        조치방안 = ACTION_BY_STATUS.get(상태, "")

        history_rows.append([
            target_str, 브랜드, code, 품목명, 상태, PRIORITY_BY_STATUS.get(상태, 99),
            전일재고, 재고, 출고, 입고계산, 최근7일, 최근90일, doi, 품절경과일, 조치방안,
        ])

        # 최근판매일/최근입고일 — 이력 전체(오늘 포함) 스캔. 날짜가 ISO(YYYY-MM-DD) 문자열이라
        # 문자열 비교가 곧 날짜 비교와 같다.
        최근판매일 = target_str if 출고 > 0 else None
        최근입고일 = target_str if 입고계산 > 0 else None
        for d_str, row in item_hist.items():
            if _to_number(row.get("출고")) > 0 and (최근판매일 is None or d_str > 최근판매일):
                최근판매일 = d_str
            if _to_number(row.get("입고")) > 0 and (최근입고일 is None or d_str > 최근입고일):
                최근입고일 = d_str
        미판매경과일 = (target_date - date.fromisoformat(최근판매일)).days if 최근판매일 else ""
        미입고경과일 = (target_date - date.fromisoformat(최근입고일)).days if 최근입고일 else ""

        # ---- 3층 결과물 라우팅 ----
        if 상태 == "⚫ 품절-장기":
            if 최근90일 > 0:
                design_rows.append([
                    브랜드, code, 품목명, 조달유형, 상태, PRIORITY_BY_STATUS.get(상태, 99),
                    리드타임, 재고, doi, 최근7일, 최근90일, 품절경과일,
                    "재입고 검토", f"재입고 골든타임(최근90일 {최근90일:.0f}개)",
                ])
            else:
                maldead_rows.append([
                    브랜드, code, 품목명, 조달유형, 리드타임, 재고,
                    최근판매일 or "", 미판매경과일, 품절경과일, 최근입고일 or "", 미입고경과일, 최근90일,
                    "재발주 검토 또는 단종 검토 (판단 필요)", "",
                ])
        elif 상태 in PRIORITY_BY_STATUS:
            design_rows.append([
                브랜드, code, 품목명, 조달유형, 상태, PRIORITY_BY_STATUS.get(상태, 99),
                리드타임, 재고, doi, 최근7일, 최근90일, 품절경과일, 조치방안, "",
            ])

        # 관리팀_전체재고 — 전체 품목(필터 없음).
        wh_qtys = [by_item_wh.get(code, {}).get(wh, 0.0) for wh in OFFLINE_WAREHOUSES]
        mgmt_rows.append([
            브랜드, code, 품목명, 조달유형, 상태, 리드타임, *wh_qtys, 재고,
            전일재고, 입고계산, 출고, 최근7일, doi, 조치방안, "",
        ])

        # 악성재고 — 재고 있음 + 90일 이상 미판매.
        if 재고 > 0 and isinstance(미판매경과일, int) and 미판매경과일 >= 90:
            malstock_action = "프로모션/할인 검토" if 미판매경과일 < 180 else "땡처리/폐기 검토"
            malstock_rows.append([
                브랜드, code, 품목명, 재고, 최근판매일 or "", 미판매경과일, "", "", malstock_action, "",
            ])

    # 정렬 — 각 탭 note에 명시된 규칙.
    design_rows.sort(key=lambda r: (r[5], r[7]))            # 우선순위 asc, 재고수량 asc
    mgmt_rows.sort(key=lambda r: (PRIORITY_BY_STATUS.get(r[4], 50), r[1]))  # 상태우선순위 → 품목코드
    malstock_rows.sort(key=lambda r: -(r[5] or 0))           # 미판매(일) 내림차순 — 오래 방치된 것부터
                                                                # (재고금액·재고수량 순으로 보고 싶으면 시트에서 직접 정렬)
    maldead_rows.sort(key=lambda r: -(r[8] or 0))            # 품절경과일 내림차순

    return {
        "history": history_rows, "design": design_rows, "mgmt": mgmt_rows,
        "malstock": malstock_rows, "maldead": maldead_rows,
    }


# ---------------------------------------------------------------------------
# 5. 대시보드 — "오늘 처리할 것" 큐. 이미 계산된 4개 결과물 리스트를 잘라서 조립할 뿐,
#    새로 계산하지 않는다(같은 데이터가 두 군데서 다르게 나오는 걸 막기 위해).
# ---------------------------------------------------------------------------
DASHBOARD_HEADERS = ["순위", "브랜드", "품목명", "조달유형", "상태", "재고", "7일 판매", "DOI(소진일)", "경과(일)", "재고금액", "조치"]

# (제목, 배경색 톤) — 목업의 "색" 컬럼과 동일한 배정.
_BLOCK_TONE = {
    "danger": SEMANTIC_DANGER_BG, "info": SEMANTIC_INFO_BG, "warning": SEMANTIC_WARNING_BG, "dim": None,
}


def _dashboard_row(rank: int, brand, name, ptype, status, qty, sales7, doi, elapsed, money, action) -> list:
    return [rank, brand, name, ptype, status, qty, sales7, doi, elapsed, money, action]


def build_dashboard_blocks(result: dict) -> list[dict]:
    design, mgmt, malstock, maldead = result["design"], result["mgmt"], result["malstock"], result["maldead"]

    def from_design(r, rank):
        # r: 브랜드,코드,품목명,조달유형,상태,우선순위,리드타임,재고,DOI,7일,90일,품절(일),조치,메모
        return _dashboard_row(rank, r[0], r[2], r[3], r[4], r[7], r[9], r[8], r[11], "", r[12])

    def from_mgmt(r, rank, elapsed=""):
        # r: 브랜드,코드,품목명,조달유형,상태,리드타임,4창고,총재고,전일재고,입고,출고,7일,DOI,조치,메모
        return _dashboard_row(rank, r[0], r[2], r[3], r[4], r[10], r[14], r[15], elapsed, "", r[16])

    def from_malstock(r, rank):
        # r: 브랜드,코드,품목명,재고,최근판매일,미판매(일),입고단가,재고금액,조치,메모
        return _dashboard_row(rank, r[0], r[2], "", "🔵 과잉", r[3], "", "", r[5], r[7], r[8])

    def from_maldead(r, rank):
        # r: 브랜드,코드,품목명,조달유형,리드타임,재고,최근판매일,미판매(일),품절(일),최근입고일,미입고(일),90일,조치,메모
        return _dashboard_row(rank, r[0], r[2], r[3], "⚫ 품절-장기", r[5], "", "", r[7], "", r[12])

    blk1_rows = [from_design(r, i + 1) for i, r in enumerate(design[:20])]
    blk2_src = sorted([r for r in design if isinstance(r[11], int) and r[11] >= 3], key=lambda r: -r[11])
    blk2_rows = [from_design(r, i + 1) for i, r in enumerate(blk2_src[:10])]

    def out_of_stock_wh_count(r):
        return sum(1 for v in r[6:10] if v == 0)

    blk3_src = sorted(
        [r for r in mgmt if r[10] > 0 and out_of_stock_wh_count(r) > 0],
        key=lambda r: -out_of_stock_wh_count(r),
    )
    blk3_rows = [from_mgmt(r, i + 1) for i, r in enumerate(blk3_src[:20])]

    blk4_src = sorted(mgmt, key=lambda r: -(r[14] or 0))
    blk4_rows = [from_mgmt(r, i + 1) for i, r in enumerate(blk4_src[:10])]

    # 악성재고/악성품절은 "매일 처리할 목록"이 아니라 "주기적으로 훑어보는 감사 리스트"라 —
    # 상위 몇 개로 자르지 않고 전부 보여준다. 정렬 기준(재고금액/재고수량 등)도 강제하지 않고
    # daily_runner가 계산한 기본 정렬(미판매(일)/품절(일) 내림차순)만 유지, 나머지는 시트에서
    # 직접 정렬해서 보라는 사용자 요청 반영(2026-07-28).
    blk5_rows = [from_malstock(r, i + 1) for i, r in enumerate(malstock)]
    blk6_rows = [from_maldead(r, i + 1) for i, r in enumerate(maldead)]

    return [
        {"title": "오늘 조치 필요", "action": "지금 발주해야 리드타임 커버", "tone": "danger",
         "source": "디자인팀_발주필요", "sort": "우선순위 → 재고 오름차순",
         "total": len(design), "show": 20, "rows": blk1_rows},
        {"title": "3일 이상 품절", "action": "즉시 재입고 검토", "tone": "danger",
         "source": "디자인팀_발주필요", "sort": "품절(일) 내림차순",
         "total": len(blk2_src), "show": 10, "rows": blk2_rows},
        {"title": "창고이동 검토", "action": "결품 매장으로 재고 이동", "tone": "info",
         "source": "관리팀_전체재고 · 총재고>0", "sort": "결품 매장 수 내림차순",
         "total": len(blk3_src), "show": 20, "rows": blk3_rows},
        {"title": "판매 속도 상위", "action": "재입고 우선순위 확인", "tone": "warning",
         "source": "관리팀_전체재고", "sort": "최근 7일 판매량 내림차순",
         "total": len(mgmt), "show": 10, "rows": blk4_rows},
        {"title": "악성재고 정리", "action": "프로모션·번들·폐기 판단 (전체 목록, 필요시 시트에서 직접 정렬)", "tone": "info",
         "source": "악성재고", "sort": "미판매(일) 내림차순",
         "total": len(malstock), "show": len(malstock), "rows": blk5_rows},
        {"title": "악성품절 — 단종 판단", "action": "재발주 여부 · 단종 확정 (전체 목록, 필요시 시트에서 직접 정렬)", "tone": "dim",
         "source": "악성품절", "sort": "품절(일) 내림차순",
         "total": len(maldead), "show": len(maldead), "rows": blk6_rows},
    ]


def build_status_distribution_line(mgmt_rows: list[list]) -> str:
    order = ["🔴", "🟠", "🟡", "🔵", "⚫", "🟢"]
    counts = {c: 0 for c in order}
    for r in mgmt_rows:
        status = r[4] or ""
        for c in order:
            if status.startswith(c):
                counts[c] += 1
                break
    parts = " · ".join(f"{c} {counts[c]}" for c in order)
    return f"오늘 재고 현황 — 전체 {len(mgmt_rows)}개 — {parts}"


def write_status_distribution_banner(service, spreadsheet_id: str, mgmt_rows: list[list]) -> None:
    """관리팀_전체재고 2행(할 일 배너)을 상태 분포 한 줄로 매일 갱신한다."""
    line = build_status_distribution_line(mgmt_rows)
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range="'관리팀_전체재고'!A2",
        valueInputOption="RAW", body={"values": [[line]]},
    ).execute()


def _status_style(status: str):
    """STATUS_COLOR_RULES를 재사용해 대시보드 상태 셀 색을 본문 탭과 동일하게 맞춘다."""
    for substr, bg, fg, bold in STATUS_COLOR_RULES:
        if substr in (status or ""):
            return bg, fg, bold
    return None, BODY_FG, False


def write_dashboard_tab(service, spreadsheet_id: str, sheet_id: int, blocks: list[dict], target_date: date) -> None:
    ncols = len(DASHBOARD_HEADERS)
    values: list[list] = [
        ["오늘 처리할 것을 위에서부터 순서대로", "", "", "", "", "", "", "", "", "", ""],
        [f"결과물 탭 4개에서 자동 집계 · 마지막 갱신: {target_date.isoformat()} (기준일)", "", "", "", "", "", "", "", "", "", ""],
        [""] * ncols,
    ]
    row_styles: list[tuple[int, str, dict | None]] = []  # (row_idx0, kind, tone)

    for block in blocks:
        count_label = f"전체 {block['total']}개" if block["show"] >= block["total"] else f"상위 {block['show']} / 전체 {block['total']}개"
        title_row = f"{block['title']} — {count_label} → {block['action']}"
        meta_row = f"출처: {block['source']} · 정렬: {block['sort']}"
        row_styles.append((len(values), "section", _BLOCK_TONE.get(block["tone"])))
        values.append([title_row] + [""] * (ncols - 2) + [meta_row])
        row_styles.append((len(values), "header", None))
        values.append(list(DASHBOARD_HEADERS))
        status_rows = []
        for r in block["rows"]:
            status_rows.append(len(values))
            values.append(r)
        row_styles.append((-1, "statusrows", status_rows))  # 표시용, 실제 서식 루프에서 사용
        values.append([""] * ncols)  # 블록 사이 spacer

    service.spreadsheets().values().batchClear(
        spreadsheetId=spreadsheet_id, body={"ranges": ["'대시보드'!A1:Z2000"]}
    ).execute()
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range="'대시보드'!A1",
        valueInputOption="RAW", body={"values": values},
    ).execute()

    meta = service.spreadsheets().get(
        spreadsheetId=spreadsheet_id, fields="sheets(properties(sheetId,title),bandedRanges(bandedRangeId))"
    ).execute()
    delete_requests = []
    for s in meta["sheets"]:
        if s["properties"]["sheetId"] != sheet_id:
            continue
        delete_requests.append({"unmergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": len(values), "startColumnIndex": 0, "endColumnIndex": ncols}}})
        for banded in s.get("bandedRanges", []):
            delete_requests.append({"deleteBanding": {"bandedRangeId": banded["bandedRangeId"]}})
    if delete_requests:
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": delete_requests}).execute()

    requests = [{
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": len(values), "startColumnIndex": 0, "endColumnIndex": ncols},
            "cell": {"userEnteredFormat": {"textFormat": {"fontFamily": FONT_BODY, "fontSize": 9, "foregroundColor": BODY_FG}}},
            "fields": "userEnteredFormat.textFormat",
        }
    }]
    # 배너 2행.
    requests.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": ncols}, "mergeType": "MERGE_ALL"}})
    requests.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": 0, "endRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": ncols},
                                     "cell": {"userEnteredFormat": {"backgroundColor": BANNER1_BG, "textFormat": {"bold": True, "foregroundColor": BANNER_FG_LIGHT, "fontFamily": FONT_BODY, "fontSize": 11}, "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}},
                                     "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)"}})
    requests.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": ncols}, "mergeType": "MERGE_ALL"}})
    requests.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": 1, "endRowIndex": 2, "startColumnIndex": 0, "endColumnIndex": ncols},
                                     "cell": {"userEnteredFormat": {"backgroundColor": BANNER3_BG, "textFormat": {"foregroundColor": BANNER3_FG, "fontFamily": FONT_BODY, "fontSize": 9}}},
                                     "fields": "userEnteredFormat(backgroundColor,textFormat)"}})

    for row_idx, kind, extra in row_styles:
        if kind == "section":
            tone_bg = extra
            requests.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": row_idx, "endRowIndex": row_idx + 1, "startColumnIndex": 0, "endColumnIndex": ncols - 6}, "mergeType": "MERGE_ALL"}})
            requests.append({"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": row_idx, "endRowIndex": row_idx + 1, "startColumnIndex": ncols - 6, "endColumnIndex": ncols}, "mergeType": "MERGE_ALL"}})
            fmt = {"textFormat": {"bold": True, "fontFamily": FONT_BODY, "fontSize": 10}}
            if tone_bg is not None:
                fmt["backgroundColor"] = tone_bg
            requests.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": row_idx, "endRowIndex": row_idx + 1, "startColumnIndex": 0, "endColumnIndex": ncols}, "cell": {"userEnteredFormat": fmt}, "fields": "userEnteredFormat(backgroundColor,textFormat)"}})
        elif kind == "header":
            requests.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": row_idx, "endRowIndex": row_idx + 1, "startColumnIndex": 0, "endColumnIndex": ncols},
                                             "cell": {"userEnteredFormat": {"backgroundColor": HEADER_BG, "textFormat": {"bold": True, "foregroundColor": HEADER_FG, "fontFamily": FONT_BODY, "fontSize": 9}}},
                                             "fields": "userEnteredFormat(backgroundColor,textFormat)"}})
        elif kind == "statusrows":
            for r_idx in extra:
                status_val = values[r_idx][4]
                bg, fg, bold = _status_style(status_val)
                fmt = {"textFormat": {"foregroundColor": fg, "bold": bold, "fontFamily": FONT_BODY, "fontSize": 9}}
                if bg is not None:
                    fmt["backgroundColor"] = bg
                requests.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r_idx, "endRowIndex": r_idx + 1, "startColumnIndex": 4, "endColumnIndex": 5}, "cell": {"userEnteredFormat": fmt}, "fields": "userEnteredFormat(backgroundColor,textFormat)"}})
                # 숫자 컬럼(재고/7일판매/DOI/경과/재고금액) 고정폭 우측정렬.
                requests.append({"repeatCell": {"range": {"sheetId": sheet_id, "startRowIndex": r_idx, "endRowIndex": r_idx + 1, "startColumnIndex": 5, "endColumnIndex": 10},
                                                 "cell": {"userEnteredFormat": {"horizontalAlignment": "RIGHT", "textFormat": {"fontFamily": FONT_MONO, "fontSize": 9}}},
                                                 "fields": "userEnteredFormat(horizontalAlignment,textFormat)"}})

    requests.append({"updateSheetProperties": {"properties": {"sheetId": sheet_id, "gridProperties": {"frozenRowCount": 2}}, "fields": "gridProperties.frozenRowCount"}})

    try:
        service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()
    except Exception as e:
        print(f"[runner] 대시보드 서식 적용 중 일부 실패(무시하고 진행): {e}")


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------

def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", required=True)
    ap.add_argument("--base-date", help="YYYY-MM-DD (기본: 실행일 기준 어제, KST)")
    ap.add_argument("--sales-xlsx", help="판매현황 엑셀 경로 (미지정 시 Gmail에서 자동 다운로드)")
    ap.add_argument("--gmail-query", default='from:ecountnotice@ecount.com subject:판매현황 has:attachment newer_than:2d',
                     help="Gmail 검색 쿼리 (기본값: 이카운트 판매현황 자동알림 발신자/제목으로 특정, 2026-07-28 실메일로 확인됨)")
    ap.add_argument("--dry-run", action="store_true", help="계산만 하고 시트에 쓰지 않음")
    args = ap.parse_args()

    target_date = date.fromisoformat(args.base_date) if args.base_date else (datetime.now(KST).date() - timedelta(days=1))
    base_date_str = target_date.strftime("%Y%m%d")
    print(f"[runner] 기준일(TARGET_DATE) = {target_date.isoformat()}")

    print("[runner] 이카운트 재고API 조회 중...")
    inventory_raw = fetch_inventory_raw(base_date_str)
    print(f"[runner] 재고 원본 {len(inventory_raw)}건 (전체 창고)")

    if args.sales_xlsx:
        sales_path = Path(args.sales_xlsx)
    else:
        print(f"[runner] Gmail에서 판매현황 첨부 다운로드 중... (query={args.gmail_query!r})")
        sales_path = fetch_sales_xlsx_from_gmail(
            args.gmail_query, DUMP_DIR / f"sales_status_{target_date.isoformat()}.xlsx"
        )
    sales_raw = parse_sales_xlsx(sales_path)
    print(f"[runner] 판매 원본 {len(sales_raw)}건 (전체 창고, 출처={sales_path})")

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    print("[runner] 품목마스터/이력 로드 중...")
    master_headers = TABS["품목마스터"]["headers"]
    master_rows = read_tab_rows(service, args.spreadsheet_id, "품목마스터")
    item_master = {
        r["품목코드"]: {"브랜드": r["브랜드"], "조달유형": r["조달유형"],
                       "리드타임": _to_number(r["리드타임(일)"]) if r["리드타임(일)"] else "",
                       "품목명": r["품목명"]}
        for r in master_rows if r["품목코드"]
    }
    history = read_tab_rows(service, args.spreadsheet_id, "일별재고이력")
    print(f"[runner] 품목마스터 {len(item_master)}건, 이력 {len(history)}행 로드")

    new_master_rows = auto_register_new_items(item_master, sales_raw, target_date)
    if new_master_rows:
        unclassified = sum(1 for r in new_master_rows if r[4] == "미분류")
        print(f"[runner] 품목마스터 신규 자동등록 {len(new_master_rows)}건"
              + (f" (이 중 {unclassified}건은 새 브랜드라 미분류 — 품목마스터에서 조달유형 확인 필요)" if unclassified else ""))

    result = build_daily_rows(target_date, inventory_raw, sales_raw, item_master, history)
    print(f"[runner] 계산 완료 — 일별이력 {len(result['history'])}건 / 디자인팀_발주필요 "
          f"{len(result['design'])}건 / 관리팀_전체재고 {len(result['mgmt'])}건 / "
          f"악성재고 {len(result['malstock'])}건 / 악성품절 {len(result['maldead'])}건")

    if args.dry_run:
        print("[runner] --dry-run — 시트에 쓰지 않음")
        return 0

    print("[runner] RAW 탭 반영 중...")
    raw_inv_rows = [[r[h] for h in TABS["RAW_재고현황"]["headers"]] for r in inventory_raw]
    raw_sales_rows = [[r[h] for h in TABS["RAW_판매현황"]["headers"]] for r in sales_raw]
    replace_tab_rows(service, args.spreadsheet_id, "RAW_재고현황", raw_inv_rows)
    replace_tab_rows(service, args.spreadsheet_id, "RAW_판매현황", raw_sales_rows)

    print("[runner] 일별재고이력 반영 중...")
    append_history_rows(service, args.spreadsheet_id, target_date.isoformat(), result["history"])

    print("[runner] 3층 결과물 탭 반영 중...")
    replace_tab_rows(service, args.spreadsheet_id, "디자인팀_발주필요", result["design"])
    replace_tab_rows(service, args.spreadsheet_id, "관리팀_전체재고", result["mgmt"])
    replace_tab_rows(service, args.spreadsheet_id, "악성재고", result["malstock"])
    replace_tab_rows(service, args.spreadsheet_id, "악성품절", result["maldead"])
    write_status_distribution_banner(service, args.spreadsheet_id, result["mgmt"])

    if new_master_rows:
        print("[runner] 품목마스터 신규 품목 반영 중...")
        append_item_master_rows(service, args.spreadsheet_id, new_master_rows)

    print("[runner] 대시보드 반영 중...")
    dash_meta = service.spreadsheets().get(
        spreadsheetId=args.spreadsheet_id, fields="sheets.properties(sheetId,title)"
    ).execute()
    dash_id_by_title = {s["properties"]["title"]: s["properties"]["sheetId"] for s in dash_meta["sheets"]}
    if "대시보드" in dash_id_by_title:
        blocks = build_dashboard_blocks(result)
        write_dashboard_tab(service, args.spreadsheet_id, dash_id_by_title["대시보드"], blocks, target_date)
    else:
        print("[runner] ⚠️ '대시보드' 탭이 없습니다 — ecount_sheets_setup.py를 먼저 실행하세요.")

    print("[runner] 완료.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
