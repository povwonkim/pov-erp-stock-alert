#!/usr/bin/env python3
"""샘플의심재고를 4개 추적 창고 기준으로 다시 계산한 시드를 반영하는 1회성 스크립트.

배경(2026-08-04): 처음 seed_aging_stock.py로 만든 샘플의심재고 후보는 "재고 1~2개"를
회사 전체(모든 창고 합산) 재고수량 기준으로 골랐는데, 실제 이 시스템은 4개 추적 창고
(POINT OF VIEW(법인)/THE HYUNDAI SEOUL/MXN/MXN(온라인))만 합산한다 — 그래서 실제
자동실행 결과는 계속 0건이었다. 재고현황 엑셀에 이미 있던 4개 창고별 컬럼만 합산해서
다시 후보를 뽑고(one_time_seed_2026-07-30/sample_seed_v2.json), 근거가 되는 입고 앵커도
새로 만들었다(sample_anchor_rows_v2.json, 기존 것과 별개로 추가만 함 — 충돌 시 건너뜀).

사용법 (서버에서):
  .venv/bin/python seed_sample_v2.py --spreadsheet-id <ID> --dry-run
  .venv/bin/python seed_sample_v2.py --spreadsheet-id <ID>
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

    sample_rows = json.loads((SEED_DIR / "sample_seed_v2.json").read_text())
    anchor_rows = json.loads((SEED_DIR / "sample_anchor_rows_v2.json").read_text())
    print(f"[seed-v2] 샘플의심재고 후보 {len(sample_rows)}건, 입고 앵커 {len(anchor_rows)}건 로드")

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    hist_headers = TABS["RAW_일별재고이력"]["headers"]
    existing = read_tab_rows(service, args.spreadsheet_id, "RAW_일별재고이력")
    existing_keys = {(r["품목코드"], r["날짜"]) for r in existing}
    print(f"[seed-v2] 기존 RAW_일별재고이력 {len(existing)}행")

    new_anchor_rows = [r for r in anchor_rows if (r[2], r[0]) not in existing_keys]
    print(f"[seed-v2] 실제 추가할 입고 앵커 {len(new_anchor_rows)}건 "
          f"(기존과 충돌해 건너뜀 {len(anchor_rows) - len(new_anchor_rows)}건)")

    if args.dry_run:
        print("[seed-v2] --dry-run: 시트에 쓰지 않음")
        return 0

    existing_as_lists = [[r[h] for h in hist_headers] for r in existing]
    all_hist_rows = existing_as_lists + new_anchor_rows
    all_hist_rows.sort(key=lambda r: (r[0], r[2]))
    replace_tab_rows(service, args.spreadsheet_id, "RAW_일별재고이력", all_hist_rows)
    print(f"[seed-v2] RAW_일별재고이력 반영 완료 (총 {len(all_hist_rows)}행)")

    replace_tab_rows(service, args.spreadsheet_id, "샘플의심재고", sample_rows)
    print("[seed-v2] 샘플의심재고 탭 반영 완료")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
