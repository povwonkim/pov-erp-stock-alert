#!/usr/bin/env python3
"""brand_type_overrides.json 기준으로 품목마스터의 기존 행들 조달유형/리드타임을 일괄 수정.

ecount_item_master.py의 classify()/BRAND_OVERRIDES는 "새로 등록되는" 품목에만 적용되고,
이미 품목마스터에 있는 기존 행은 자동으로 안 바뀐다 — brand_type_overrides.json을 고친
뒤에는 이 스크립트로 기존 행도 맞춰줘야 한다(2026-08-06, Travelers company/MD paper/
SAILOR/KAWECO/MOLESKINE를 국내사입으로 확정하면서 도입).

사용법 (서버에서):
  .venv/bin/python backfill_brand_types.py --spreadsheet-id <ID> --dry-run
  .venv/bin/python backfill_brand_types.py --spreadsheet-id <ID>
"""
from __future__ import annotations

import argparse
from datetime import date

from ecount_daily_runner import _load_creds, read_tab_rows, replace_tab_rows
from ecount_item_master import BRAND_OVERRIDES, LEADTIME_BY_TYPE
from ecount_sheets_setup import TABS


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", required=True)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    headers = TABS["품목마스터"]["headers"]
    rows = read_tab_rows(service, args.spreadsheet_id, "품목마스터")
    print(f"[backfill] 품목마스터 {len(rows)}행 로드, brand_type_overrides.json {len(BRAND_OVERRIDES)}개 브랜드")

    today = date.today().isoformat()
    changed = 0
    by_brand_changed: dict[str, int] = {}
    out_rows = []
    for r in rows:
        brand = (r.get("브랜드") or "").strip()
        new_type = BRAND_OVERRIDES.get(brand)
        if new_type and new_type != r.get("조달유형"):
            r["조달유형"] = new_type
            r["리드타임(일)"] = LEADTIME_BY_TYPE[new_type]
            r["갱신일"] = today
            changed += 1
            by_brand_changed[brand] = by_brand_changed.get(brand, 0) + 1
        out_rows.append([r[h] for h in headers])

    print(f"[backfill] 조달유형 수정 대상: {changed}행")
    for brand, count in sorted(by_brand_changed.items(), key=lambda x: -x[1]):
        print(f"  - {brand}: {count}개 → {BRAND_OVERRIDES[brand]}")

    if args.dry_run:
        print("[backfill] --dry-run: 시트에 쓰지 않음")
        return 0
    if changed == 0:
        print("[backfill] 바뀔 게 없어서 쓰지 않음")
        return 0

    replace_tab_rows(service, args.spreadsheet_id, "품목마스터", out_rows)
    print("[backfill] 품목마스터 반영 완료 — 다음 자동실행부터 리드타임/위험판정에도 반영됩니다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
