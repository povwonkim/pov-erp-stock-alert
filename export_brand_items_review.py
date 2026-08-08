#!/usr/bin/env python3
"""브랜드 단위로는 판단이 안 되는 통 분류(OTHER/POP-UP 등)에 속한 품목을 품목 단위로
검토할 수 있게 엑셀로 뽑는다(2026-08-07 사용자 요청 — "각각의 품목을 알아야돼서 품목명이랑
현재 재고수량을 같이 보여주면 내가 표시해줄게" / "POP-UP도 287개면 다 확인해야돼").

RAW_품목마스터(이미 등록된 것)와 one_time_seed_2026-07-30/full_item_registry_2026-08-07.json
(이카운트 품목등록 전체 조회 — 아직 등록 안 된 것/EXCLUDED_BRANDS라 애초에 등록 안 되는
것도 포함) 양쪽에서 지정한 브랜드의 품목코드를 모으고, RAW_재고현황(오프라인 4개 창고
합계)에서 현재 재고수량을 붙인다.

사용법 (서버에서):
  .venv/bin/python export_brand_items_review.py --spreadsheet-id <ID> --brand OTHER --out other_items_review.xlsx
  .venv/bin/python export_brand_items_review.py --spreadsheet-id <ID> --brand POP-UP --out popup_items_review.xlsx
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from openpyxl import Workbook

from ecount_daily_runner import _load_creds, read_tab_rows, OFFLINE_WAREHOUSES

REGISTRY_PATH = Path(__file__).parent / "one_time_seed_2026-07-30" / "full_item_registry_2026-08-07.json"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", required=True)
    ap.add_argument("--brand", required=True, help='예: OTHER, POP-UP')
    ap.add_argument("--out", default="brand_items_review.xlsx")
    args = ap.parse_args()

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    # 품목코드 -> 품목명, 지정한 브랜드인 것만.
    items: dict[str, str] = {}

    master_rows = read_tab_rows(service, args.spreadsheet_id, "RAW_품목마스터")
    for r in master_rows:
        if (r.get("브랜드") or "").strip() == args.brand:
            items[r["품목코드"]] = r.get("품목명", "")

    registry = json.loads(REGISTRY_PATH.read_text(encoding="utf-8"))
    for item in registry:
        if item.get("브랜드") == args.brand and item["품목코드"] not in items:
            items[item["품목코드"]] = item.get("품목명", "")

    print(f"[export] {args.brand} 브랜드 품목 {len(items)}건 발견 (품목마스터+전체등록 합산)")

    # 현재 재고수량 — 오프라인 4개 창고 합계.
    inv_rows = read_tab_rows(service, args.spreadsheet_id, "RAW_재고현황")
    stock: dict[str, float] = {}
    for r in inv_rows:
        if r.get("창고명") not in OFFLINE_WAREHOUSES:
            continue
        code = r["품목코드"]
        stock[code] = stock.get(code, 0.0) + float(r.get("재고수량") or 0)

    wb = Workbook()
    ws = wb.active
    ws.title = f"{args.brand} 품목 검토"[:31]
    ws.append(["품목코드", "품목명", "현재재고수량", "실제 브랜드/구분(표시)", "메모"])
    for code in sorted(items, key=lambda c: -stock.get(c, 0.0)):
        ws.append([code, items[code], stock.get(code, 0.0), "", ""])

    for col, width in zip("ABCDE", [20, 40, 14, 24, 30]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"
    wb.save(args.out)
    print(f"[export] {args.out} 저장 완료 — 재고수량 큰 순으로 정렬됨")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
