#!/usr/bin/env python3
"""전체 품목 대상으로 다시 만든 미판매/미입고 앵커를 반영하는 1회성 스크립트.

배경(2026-08-04): 앞서 심은 입고 앵커는 "그 시점(7/30)에 재고 1~2개였던 품목"만
대상으로 했는데, 실제 진단(diag_sample_pool.py) 결과 지금(8/3) 1~2개인 품목 1,824개
중 근거 있는 게 0개였다 — 1~2개는 거래 하나에도 바로 벗어나는 값이라 "그때 1~2개"와
"지금 1~2개"가 거의 안 겹치는 게 원인이었다.

해결: 재고 수량과 무관하게 전체 품목 대상으로 앵커를 다시 만든다(anchor_rows_v3.json).
- 출고(미판매) 앵커: 재고 0 이상이면 다 포함(판매 이력 유무만 근거로 삼음)
- 입고 앵커: 재고보유월수만 있으면 현재 수량과 무관하게 포함
이러면 어느 품목이 어느 날 1~2개로 바뀌든 이미 근거가 준비되어 있다.

같은 (품목코드, 날짜)에 이미 행이 있으면 "건너뛰기"가 아니라 "비어있는 칸만 채우기"로
병합한다(실제 값이 있는 칸은 안 건드림) — 이전 버그(seed_sample_v2.py) 재발 방지.

사용법 (서버에서):
  .venv/bin/python seed_anchors_v3.py --spreadsheet-id <ID> --dry-run
  .venv/bin/python seed_anchors_v3.py --spreadsheet-id <ID>
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

    anchor_rows = json.loads((SEED_DIR / "anchor_rows_v3.json").read_text())
    print(f"[seed-v3] 앵커 {len(anchor_rows)}건 로드")

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    headers = TABS["일별재고이력"]["headers"]
    idx_출고 = headers.index("출고")
    idx_입고 = headers.index("입고")

    existing = read_tab_rows(service, args.spreadsheet_id, "일별재고이력")
    print(f"[seed-v3] 기존 일별재고이력 {len(existing)}행")

    rows_as_lists = [[r[h] for h in headers] for r in existing]
    existing_idx = {(r["품목코드"], r["날짜"]): i for i, r in enumerate(existing)}

    merged = added = 0
    for anchor in anchor_rows:
        key = (anchor[2], anchor[0])
        if key in existing_idx:
            i = existing_idx[key]
            changed = False
            if anchor[idx_출고] and not rows_as_lists[i][idx_출고]:
                rows_as_lists[i][idx_출고] = anchor[idx_출고]
                changed = True
            if anchor[idx_입고] and not rows_as_lists[i][idx_입고]:
                rows_as_lists[i][idx_입고] = anchor[idx_입고]
                changed = True
            if changed:
                merged += 1
        else:
            rows_as_lists.append(anchor)
            added += 1

    print(f"[seed-v3] 기존 행 병합: {merged}건 / 새로 추가: {added}건")

    if args.dry_run:
        print("[seed-v3] --dry-run: 시트에 쓰지 않음")
        return 0

    rows_as_lists.sort(key=lambda r: (r[0], r[2]))
    replace_tab_rows(service, args.spreadsheet_id, "일별재고이력", rows_as_lists)
    print(f"[seed-v3] 일별재고이력 반영 완료 (총 {len(rows_as_lists)}행)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
