#!/usr/bin/env python3
"""소계(합계) 행 오파싱으로 잘못 등록된 유령 품목을 품목마스터/일별재고이력에서 제거.

배경(2026-08-06): ecount_sales_scraper.py가 이카운트 판매현황 웹 표의 "소계" 행(라벨
칸이 colspan으로 병합돼 있어 정상 품목 행보다 칸 수가 적음)을 정상 품목 행으로 오인해서,
소계 숫자(수량합계/금액합계 등)가 품목코드/품명/브랜드 자리에 잘못 들어간 채 자동등록되고
있었다(원인은 ecount_sales_scraper.py에서 수정 완료 — 이 스크립트는 이미 잘못 등록된
기존 데이터를 지우는 1회성 정리용).

식별 규칙(2026-08-06 1차 시도 실패 후 수정): 품목코드가 숫자+콤마로만 이루어진 것만으로는
안 된다 — 실제 이카운트 품목코드 중에도 순수 숫자인 것들이 꽤 있다(예: "07101139",
"0001004878306", 바코드형 코드). 1차 시도로 이 조건만 썼더니 1,389개가 걸렸는데, 그중
다수가 item_cost_lookup.json(2026-07-30 실제 재고현황 엑셀에서 뽑은 진짜 품목 목록)에도
있는 진짜 품목이었다 — 잘못하면 진짜 데이터를 지울 뻔했다.

실제 소계 행 오파싱 사례(스크린샷 확인)는 품목코드뿐 아니라 브랜드 칸까지 동시에
숫자로 깨져있었다(예: 품목코드="8,079,087", 브랜드="1,761") — 진짜 품목은 품목코드가
숫자여도 브랜드는 항상 정상 텍스트(Zeroperzero, PAPERIAN 등)라 이 조합이 절대 안 겹친다.
그래서 "품목코드 AND 브랜드가 둘 다 숫자+콤마 패턴"인 것만 대상으로 하고, 혹시 모를
오탐까지 막기 위해 item_cost_lookup.json에 있는 코드는 한 번 더 제외한다(진짜 품목이란
확실한 증거가 있으면 무조건 보존).

사용법 (서버에서):
  .venv/bin/python cleanup_ghost_items.py --spreadsheet-id <ID> --dry-run
  .venv/bin/python cleanup_ghost_items.py --spreadsheet-id <ID>
"""
from __future__ import annotations

import argparse
import json
import re
from pathlib import Path

from ecount_daily_runner import _load_creds, read_tab_rows, replace_tab_rows
from ecount_sheets_setup import TABS

GHOST_PATTERN_RE = re.compile(r"^[\d,]+$")
_COST_LOOKUP_PATH = Path(__file__).parent / "one_time_seed_2026-07-30" / "item_cost_lookup.json"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", required=True)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    try:
        known_real_codes = set(json.loads(_COST_LOOKUP_PATH.read_text()))
    except FileNotFoundError:
        known_real_codes = set()

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    # ---- 품목마스터 ----
    master_headers = TABS["품목마스터"]["headers"]
    master_rows = read_tab_rows(service, args.spreadsheet_id, "품목마스터")

    def is_ghost(r: dict) -> bool:
        code = r["품목코드"] or ""
        brand = r["브랜드"] or ""
        if code in known_real_codes:
            return False
        return bool(GHOST_PATTERN_RE.match(code)) and bool(GHOST_PATTERN_RE.match(brand))

    ghost_rows = [r for r in master_rows if is_ghost(r)]
    ghost_codes = {r["품목코드"] for r in ghost_rows}
    kept_master = [r for r in master_rows if r["품목코드"] not in ghost_codes]

    print(f"[cleanup] 품목마스터 {len(master_rows)}행 중 유령 품목 {len(ghost_rows)}개 발견")
    for r in ghost_rows[:10]:
        print(f"  - 품목코드={r['품목코드']!r} 품명={r['품목명']!r} 브랜드={r['브랜드']!r} 갱신일={r['갱신일']!r}")
    if len(ghost_rows) > 10:
        print(f"  ... 외 {len(ghost_rows) - 10}개")

    # ---- 일별재고이력(같은 유령 코드로 매일 쌓인 행들) ----
    hist_headers = TABS["일별재고이력"]["headers"]
    hist_rows = read_tab_rows(service, args.spreadsheet_id, "일별재고이력")
    ghost_hist = [r for r in hist_rows if r["품목코드"] in ghost_codes]
    kept_hist = [r for r in hist_rows if r["품목코드"] not in ghost_codes]
    print(f"[cleanup] 일별재고이력 {len(hist_rows)}행 중 유령 품목 관련 {len(ghost_hist)}행 발견")

    if args.dry_run:
        print("[cleanup] --dry-run: 시트에 쓰지 않음")
        return 0

    if not ghost_codes:
        print("[cleanup] 지울 게 없음")
        return 0

    replace_tab_rows(service, args.spreadsheet_id, "품목마스터",
                      [[r[h] for h in master_headers] for r in kept_master])
    print("[cleanup] 품목마스터 정리 완료")

    if ghost_hist:
        replace_tab_rows(service, args.spreadsheet_id, "일별재고이력",
                          [[r[h] for h in hist_headers] for r in kept_hist])
        print("[cleanup] 일별재고이력 정리 완료")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
