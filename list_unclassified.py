#!/usr/bin/env python3
"""품목마스터에서 조달유형="미분류"인 행을 뽑아서 보여주는 1회성 조회 스크립트.

사용법 (서버에서):
  .venv/bin/python list_unclassified.py --spreadsheet-id <ID>
"""
from __future__ import annotations

import argparse

from ecount_daily_runner import _load_creds, read_tab_rows


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", required=True)
    args = ap.parse_args()

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    rows = read_tab_rows(service, args.spreadsheet_id, "품목마스터")
    unclassified = [r for r in rows if r.get("조달유형") == "미분류"]

    print(f"미분류 {len(unclassified)}건")
    for r in unclassified:
        print(f"  브랜드={r['브랜드']!r} 품목코드={r['품목코드']!r} 품목명={r['품목명']!r}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
