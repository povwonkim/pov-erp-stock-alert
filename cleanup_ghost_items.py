#!/usr/bin/env python3
"""소계(합계) 행 오파싱으로 잘못 등록된 유령 품목을 품목마스터/일별재고이력에서 제거.

배경(2026-08-06): ecount_sales_scraper.py가 이카운트 판매현황 웹 표의 "소계" 행(라벨
칸이 colspan으로 병합돼 있어 정상 품목 행보다 칸 수가 적음)을 정상 품목 행으로 오인해서,
소계 숫자(수량합계/금액합계 등)가 품목코드/품명/브랜드 자리에 잘못 들어간 채 자동등록되고
있었다(원인은 ecount_sales_scraper.py에서 수정 완료 — 이 스크립트는 이미 잘못 등록된
기존 데이터를 지우는 1회성 정리용).

식별 규칙: 품목코드가 숫자와 콤마로만 이루어진 경우(예: "8,079,087") — 실제 이카운트
품목코드는 항상 문자가 섞여있어(예: "JJPU0755_001", "ZZFJST02") 이 패턴과 안 겹친다.

사용법 (서버에서):
  .venv/bin/python cleanup_ghost_items.py --spreadsheet-id <ID> --dry-run
  .venv/bin/python cleanup_ghost_items.py --spreadsheet-id <ID>
"""
from __future__ import annotations

import argparse
import re

from ecount_daily_runner import _load_creds, read_tab_rows, replace_tab_rows
from ecount_sheets_setup import TABS

GHOST_CODE_RE = re.compile(r"^[\d,]+$")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", required=True)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    # ---- 품목마스터 ----
    master_headers = TABS["품목마스터"]["headers"]
    master_rows = read_tab_rows(service, args.spreadsheet_id, "품목마스터")
    ghost_codes = {r["품목코드"] for r in master_rows if GHOST_CODE_RE.match(r["품목코드"] or "")}
    kept_master = [r for r in master_rows if r["품목코드"] not in ghost_codes]

    print(f"[cleanup] 품목마스터 {len(master_rows)}행 중 유령 품목 {len(ghost_codes)}개 발견")
    for code in sorted(ghost_codes)[:10]:
        print(f"  - {code!r}")
    if len(ghost_codes) > 10:
        print(f"  ... 외 {len(ghost_codes) - 10}개")

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
