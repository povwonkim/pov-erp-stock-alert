#!/usr/bin/env python3
"""악성재고/샘플의심재고 1회성 시드(seed_aging_stock.py)가 다음 날 자동실행 때 통째로
사라지는 문제를 막기 위한 1회성 스크립트.

배경(2026-07-30): ecount_daily_runner.py는 악성재고/샘플의심재고 탭을 "RAW_일별재고이력"에서
직접 계산한다(최근판매일=RAW_일별재고이력에서 출고>0인 가장 최근 날짜). 그런데 RAW_일별재고이력은
이 시스템이 스스로 쌓은 것만 알아서, 오늘 엑셀 기반으로 만든 seed_aging_stock.py의 결과는
내일 자동실행 때 "판매 이력 근거 없음"으로 덮어써져 사라진다.

해결: 결과 탭이 아니라 "RAW_일별재고이력"에 합성 앵커 행을 심는다. 예를 들어 어떤 품목이
"최근 91~179일 사이 마지막 판매"로 추정됐다면, (오늘-미판매추정일) 날짜에 출고=1인
합성 행 하나를 추가한다. 그러면 ecount_daily_runner.py의 기존 계산 로직이 이 행을
그대로 읽어서 최근판매일을 알아내고, 이후 매일 미판매(일)이 자연스럽게 +1씩 늘어난다
(사용자 확인: "판매안된 날짜수에 +1이 되야하는거 아니냐").

합성 행은 날짜/품목코드/출고/입고 외 다른 칸은 비워둔다(전일재고 연속성 계산은 "어제"
날짜만 조회하므로 이 과거 앵커 행과 충돌하지 않음).

데이터: one_time_seed_2026-07-30/history_anchor_rows.json (로컬에서 미리 계산).

사용법 (서버에서):
  .venv/bin/python seed_history_anchors.py --spreadsheet-id <ID> --dry-run
  .venv/bin/python seed_history_anchors.py --spreadsheet-id <ID>
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

    anchor_rows = json.loads((SEED_DIR / "history_anchor_rows.json").read_text())
    print(f"[seed-hist] 합성 앵커 행 {len(anchor_rows)}건 로드")

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    headers = TABS["RAW_일별재고이력"]["headers"]
    existing = read_tab_rows(service, args.spreadsheet_id, "RAW_일별재고이력")
    existing_keys = {(r["품목코드"], r["날짜"]) for r in existing}
    print(f"[seed-hist] 기존 RAW_일별재고이력 {len(existing)}행")

    skipped = 0
    new_rows = []
    for r in anchor_rows:
        code, date_str = r[2], r[0]
        if (code, date_str) in existing_keys:
            skipped += 1
            continue
        new_rows.append(r)
    print(f"[seed-hist] 실제 추가할 행 {len(new_rows)}건 (기존과 충돌해 건너뜀 {skipped}건)")

    if args.dry_run:
        print("[seed-hist] --dry-run: 시트에 쓰지 않음")
        return 0

    existing_as_lists = [[r[h] for h in headers] for r in existing]
    all_rows = existing_as_lists + new_rows
    all_rows.sort(key=lambda r: (r[0], r[2]))  # 날짜 오름차순 → 품목코드순 (탭 정렬 규칙)

    replace_tab_rows(service, args.spreadsheet_id, "RAW_일별재고이력", all_rows)
    print(f"[seed-hist] RAW_일별재고이력 반영 완료 (총 {len(all_rows)}행)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
