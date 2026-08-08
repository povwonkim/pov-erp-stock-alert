#!/usr/bin/env python3
""""(단종)" 표시가 붙은 품목(사용중단 포함) 중 재고가 남아있는 게 있는지 확인.

배경(2026-08-07): "(단종)" 표시가 붙으면 build_daily_rows에서 아예 통째로 제외돼서
관리팀_전체재고 등 어디에도 안 보인다. OTS 사례로 "사용중단인데 재고 남은 것 = 처리
안 된 자산"이라는 게 확인된 뒤, OTS 말고도 이렇게 조용히 숨어있는 품목이 더 있는지
사용자가 확인 요청.

RAW_품목마스터에서 품목명에 "(단종)"이 포함된 것을 모으고, RAW_재고현황(오프라인 4개
창고 합계)에서 실제 재고수량을 붙여서 재고>0인 것만 보여준다.

사용법 (서버에서):
  .venv/bin/python check_discontinued_stock.py --spreadsheet-id <ID>
"""
from __future__ import annotations

import argparse

from ecount_daily_runner import _load_creds, read_tab_rows, OFFLINE_WAREHOUSES


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", required=True)
    args = ap.parse_args()

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    master_rows = read_tab_rows(service, args.spreadsheet_id, "RAW_품목마스터")
    discontinued = {r["품목코드"]: r for r in master_rows if "(단종)" in (r.get("품목명") or "")}
    print(f"[check] '(단종)' 표시된 품목: {len(discontinued)}건")

    inv_rows = read_tab_rows(service, args.spreadsheet_id, "RAW_재고현황")
    stock: dict[str, float] = {}
    for r in inv_rows:
        if r.get("창고명") not in OFFLINE_WAREHOUSES:
            continue
        code = r["품목코드"]
        stock[code] = stock.get(code, 0.0) + float(r.get("재고수량") or 0)

    with_stock = [(code, r, stock.get(code, 0.0)) for code, r in discontinued.items() if stock.get(code, 0.0) > 0]
    with_stock.sort(key=lambda x: -x[2])

    print(f"[check] 그중 오프라인 4개 창고 재고 1개 이상: {len(with_stock)}건")
    for code, r, qty in with_stock[:50]:
        print(f"  브랜드={r.get('브랜드')!r} 품목코드={code!r} 품목명={r.get('품목명')!r} 재고={qty:g}")
    if len(with_stock) > 50:
        print(f"  ... 외 {len(with_stock) - 50}건")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
