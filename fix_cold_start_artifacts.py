#!/usr/bin/env python3
"""RAW_일별재고이력의 "가짜 입고" 아티팩트를 제거하는 1회성 스크립트.

배경(2026-08-04): 크론이 며칠 실패하던 시기(7/28~29, 7/31 등)에 이력이 군데군데
비어서, 그 다음 날 계산할 때 "전일재고"를 찾을 행이 없어 0으로 기본값 처리됐다.
입고계산 = (재고 - 전일재고) + 출고 라서, 전일재고가 실제로는 있었는데 그냥
기록이 없어서 0으로 잘못 취급되면 그 차이가 전부 "입고"로 잡혀버린다(실제 입고가
아님). 이 가짜 입고 날짜가 저희가 심어둔 진짜 미입고 앵커(재고보유월수 기반)보다
최신이라 최근입고일 계산에서 앵커를 가려버려서 샘플의심재고가 계속 0건으로 나왔다.

식별 규칙: 어떤 행(품목코드, 날짜)에서
  - 바로 전날(날짜-1) 행이 그 품목코드로 아예 존재하지 않고(진짜 구멍)
  - 전일재고가 0으로 기록되어 있고
  - 출고가 0이고
  - 입고가 0보다 크면
→ 진짜 입고가 아니라 구멍 때문에 생긴 아티팩트로 보고 입고를 0으로 고친다.
(품목이 그 시스템에 처음 등장한 진짜 첫날은 건드리지 않도록, "이 품목의 더 이전
날짜 행이 하나도 없는 경우"는 제외한다 — 그런 경우 입고가 있어도 이상하지 않음.)

사용법 (서버에서):
  .venv/bin/python fix_cold_start_artifacts.py --spreadsheet-id <ID> --dry-run
  .venv/bin/python fix_cold_start_artifacts.py --spreadsheet-id <ID>
"""
from __future__ import annotations

import argparse
from datetime import date, timedelta

from ecount_daily_runner import _load_creds, read_tab_rows, replace_tab_rows
from ecount_sheets_setup import TABS


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", required=True)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    headers = TABS["RAW_일별재고이력"]["headers"]
    idx_전일재고 = headers.index("전일재고")
    idx_출고 = headers.index("출고")
    idx_입고 = headers.index("입고")

    existing = read_tab_rows(service, args.spreadsheet_id, "RAW_일별재고이력")
    print(f"[fix] 기존 RAW_일별재고이력 {len(existing)}행")

    rows_as_lists = [[r[h] for h in headers] for r in existing]
    dates_by_code: dict[str, set[str]] = {}
    for r in existing:
        dates_by_code.setdefault(r["품목코드"], set()).add(r["날짜"])

    fixed = 0
    for row in rows_as_lists:
        code = row[headers.index("품목코드")]
        d_str = row[headers.index("날짜")]
        전일재고 = row[idx_전일재고]
        출고 = row[idx_출고]
        입고 = row[idx_입고]
        if not (isinstance(입고, (int, float)) and 입고 > 0):
            continue
        if 전일재고 not in (0, "", None) or (출고 not in (0, "", None)):
            continue
        try:
            d = date.fromisoformat(d_str.strip())
        except ValueError:
            continue
        prev_str = (d - timedelta(days=1)).isoformat()
        item_dates = dates_by_code.get(code, set())
        has_earlier = any(x < d_str for x in item_dates)
        if not has_earlier:
            continue  # 이 품목의 진짜 첫 등장 — 건드리지 않음
        if prev_str in item_dates:
            continue  # 어제 기록이 실제로 있었음 — 진짜 구멍 아님, 건드리지 않음
        row[idx_입고] = 0
        fixed += 1

    print(f"[fix] 가짜 입고로 판단해 0으로 고친 행: {fixed}건")

    if args.dry_run:
        print("[fix] --dry-run: 시트에 쓰지 않음")
        return 0

    rows_as_lists.sort(key=lambda r: (r[0], r[2]))
    replace_tab_rows(service, args.spreadsheet_id, "RAW_일별재고이력", rows_as_lists)
    print(f"[fix] RAW_일별재고이력 반영 완료 (총 {len(rows_as_lists)}행)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
