#!/usr/bin/env python3
"""샘플의심재고가 계속 0건인 게 "지금 1~2개짜리가 실제로 없어서"인지, 아니면 다른 버그인지
확인하는 1회성 진단 스크립트. ecount_daily_runner.py와 같은 방식으로 4개 추적 창고만
합산해서 재고 1~2개인 품목이 몇 개인지, 그중 미판매/미입고 근거가 있는 게 몇 개인지 센다.

사용법 (서버에서, 이미 캐시된 재고 파일 사용):
  .venv/bin/python diag_sample_pool.py --spreadsheet-id <ID> --inventory-cache cron_tracking/ecount/inventory_raw_2026-08-03.json
"""
from __future__ import annotations

import argparse
import json
from datetime import date
from pathlib import Path

from ecount_daily_runner import _load_creds, read_tab_rows
from ecount_sheets_setup import TABS, OFFLINE_WAREHOUSES


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", required=True)
    ap.add_argument("--inventory-cache", required=True)
    ap.add_argument("--base-date", required=True, help="YYYY-MM-DD, 재고 캐시가 대표하는 날짜")
    args = ap.parse_args()

    inventory_raw = json.loads(Path(args.inventory_cache).read_text())
    offline_inv = [r for r in inventory_raw if r["창고명"] in OFFLINE_WAREHOUSES]
    total_by_item: dict[str, float] = {}
    for r in offline_inv:
        code = r["품목코드"]
        total_by_item[code] = total_by_item.get(code, 0.0) + r["재고수량"]

    pool_1_2 = {code: qty for code, qty in total_by_item.items() if qty in (1, 2)}
    print(f"[diag] 4개 창고 합산 재고가 정확히 1~2개인 품목: {len(pool_1_2)}개")

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)
    history = read_tab_rows(service, args.spreadsheet_id, "일별재고이력")
    hist_by_item: dict[str, dict[str, dict]] = {}
    for r in history:
        hist_by_item.setdefault(r["품목코드"], {})[r["날짜"]] = r

    target_date = date.fromisoformat(args.base_date)
    has_sale_hist = has_inbound_hist = has_both_hist = both_stale = 0
    no_hist_at_all = 0
    examples = []
    no_hist_examples = []
    both_examples = []
    for code, qty in pool_1_2.items():
        item_hist = hist_by_item.get(code, {})
        if not item_hist:
            no_hist_at_all += 1
            if len(no_hist_examples) < 5:
                no_hist_examples.append(code)
            continue
        최근판매일 = None
        최근입고일 = None
        for d_str, row in item_hist.items():
            출고 = row.get("출고") or 0
            입고 = row.get("입고") or 0
            if isinstance(출고, (int, float)) and 출고 > 0 and (최근판매일 is None or d_str > 최근판매일):
                최근판매일 = d_str
            if isinstance(입고, (int, float)) and 입고 > 0 and (최근입고일 is None or d_str > 최근입고일):
                최근입고일 = d_str
        if 최근판매일:
            has_sale_hist += 1
        if 최근입고일:
            has_inbound_hist += 1
        if 최근판매일 and 최근입고일:
            has_both_hist += 1
            if len(both_examples) < 8:
                both_examples.append((code, qty, 최근판매일, 최근입고일))
        미판매 = (target_date - date.fromisoformat(최근판매일)).days if 최근판매일 else None
        미입고 = (target_date - date.fromisoformat(최근입고일)).days if 최근입고일 else None
        if 미판매 is not None and 미판매 >= 90 and 미입고 is not None and 미입고 >= 90:
            both_stale += 1
            if len(examples) < 5:
                examples.append((code, qty, 미판매, 미입고))

    print(f"[diag] 일별재고이력에 해당 품목코드 행 자체가 아예 없음: {no_hist_at_all}개 (예: {no_hist_examples})")
    print(f"[diag] 최근판매일 근거 있음: {has_sale_hist}개")
    print(f"[diag] 최근입고일 근거 있음: {has_inbound_hist}개")
    print(f"[diag] 둘 다 근거 있음: {has_both_hist}개")
    print(f"[diag] 둘 다 근거 있고 90일 이상: {both_stale}개")
    for ex in examples:
        print("  90일+ 예시:", ex)
    print("[diag] '둘 다 근거 있음' 품목의 실제 날짜값 샘플(품목코드, 재고, 최근판매일, 최근입고일):")
    for ex in both_examples:
        print("  ", ex)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
