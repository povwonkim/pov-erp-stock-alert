#!/usr/bin/env python3
"""악성재고/샘플의심재고 탭에 엑셀 기반 일회성 시드 데이터를 채워넣는 1회성 스크립트.

배경(2026-07-30): RAW_일별재고이력이 아직 하루치도 없어서(이 시스템 도입 초기) 악성재고/
샘플의심재고 탭이 비어있었다. 사용자가 이카운트에서 직접 4개 리포트를 수동 Excel로
뽑아줘서(재고현황 전체 + 판매현황 90일/180일/1년), 그 데이터로 정확한 "판매 있었는가"
사실관계를 기준으로 후보 리스트를 한 번 만들어 기준점(baseline)으로 심는다.

이후 매일 자동 실행되는 ecount_daily_runner.py가 실제 RAW_일별재고이력이 쌓이는 대로 이
탭들을 정밀한 값으로 덮어쓴다 — 이 스크립트는 그 전까지의 공백을 메우는 용도.

데이터 자체는 one_time_seed_2026-07-30/{malstock,sample}_seed.json에 미리 계산되어
있다(로컬에서 openpyxl로 4개 엑셀을 대조해서 만듦 — 이 스크립트는 그 결과를 시트에
쓰기만 한다).

사용법 (서버에서):
  .venv/bin/python seed_aging_stock.py --spreadsheet-id <ID>
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from ecount_daily_runner import _load_creds, replace_tab_rows

SEED_DIR = Path(__file__).parent / "one_time_seed_2026-07-30"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", required=True)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    malstock_rows = json.loads((SEED_DIR / "malstock_seed.json").read_text())
    sample_rows = json.loads((SEED_DIR / "sample_seed.json").read_text())
    print(f"[seed] 악성재고 {len(malstock_rows)}건, 샘플의심재고 {len(sample_rows)}건 로드")

    if args.dry_run:
        print("[seed] --dry-run: 시트에 쓰지 않음")
        return 0

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    replace_tab_rows(service, args.spreadsheet_id, "악성재고", malstock_rows)
    print("[seed] 악성재고 탭 반영 완료")
    replace_tab_rows(service, args.spreadsheet_id, "샘플의심재고", sample_rows)
    print("[seed] 샘플의심재고 탭 반영 완료")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
