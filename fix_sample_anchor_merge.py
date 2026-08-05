#!/usr/bin/env python3
"""seed_sample_v2.py의 충돌 처리 버그를 바로잡는 1회성 스크립트.

배경(2026-08-04): seed_sample_v2.py는 (품목코드, 날짜)에 이미 행이 있으면 "건너뛰기"
했는데, 그 기존 행은 악성재고용 출고 앵커(출고=1)일 뿐 입고 정보는 없었다 — 그래서
샘플의심재고에 필요한 입고 신호가 1,063건 중 13건만 반영되고 나머지 1,050건은
누락됐다. 건너뛰지 않고 기존 행에 입고=1을 채워 넣는(병합) 방식으로 다시 처리한다.

사용법 (서버에서):
  .venv/bin/python fix_sample_anchor_merge.py --spreadsheet-id <ID> --dry-run
  .venv/bin/python fix_sample_anchor_merge.py --spreadsheet-id <ID>
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from ecount_daily_runner import _load_creds, read_tab_rows, replace_tab_rows
from ecount_sheets_setup import TABS

SEED_DIR = Path(__file__).parent / "one_time_seed_2026-07-30"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", required=True)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    anchor_rows = json.loads((SEED_DIR / "sample_anchor_rows_v2.json").read_text())
    anchor_by_key = {(r[2], r[0]): r for r in anchor_rows}  # (품목코드, 날짜) -> 앵커 행
    print(f"[fix] 입고 앵커 {len(anchor_rows)}건 로드")

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    headers = TABS["일별재고이력"]["headers"]
    existing = read_tab_rows(service, args.spreadsheet_id, "일별재고이력")
    print(f"[fix] 기존 일별재고이력 {len(existing)}행")

    idx_입고 = headers.index("입고")
    merged = 0
    added = 0
    rows_as_lists = [[r[h] for h in headers] for r in existing]
    existing_keys = {(r["품목코드"], r["날짜"]): i for i, r in enumerate(existing)}

    for key, anchor in anchor_by_key.items():
        if key in existing_keys:
            i = existing_keys[key]
            if not rows_as_lists[i][idx_입고]:  # 비어있거나 0일 때만 채움(진짜 입고 데이터는 안 건드림)
                rows_as_lists[i][idx_입고] = 1
                merged += 1
        else:
            rows_as_lists.append(anchor)
            added += 1

    print(f"[fix] 기존 행에 입고=1 병합: {merged}건 / 새로 추가: {added}건")

    if args.dry_run:
        print("[fix] --dry-run: 시트에 쓰지 않음")
        return 0

    rows_as_lists.sort(key=lambda r: (r[0], r[2]))
    replace_tab_rows(service, args.spreadsheet_id, "일별재고이력", rows_as_lists)
    print(f"[fix] 일별재고이력 반영 완료 (총 {len(rows_as_lists)}행)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
