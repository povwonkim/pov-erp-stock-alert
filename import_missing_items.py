#!/usr/bin/env python3
"""이카운트 "품목등록" 전체 조회(사용중단 포함, 창고 무관) 결과를 기준으로,
RAW_품목마스터에 아직 없는 품목을 채워넣는 1회성 스크립트.

배경(2026-08-07): 지금까지 RAW_품목마스터는 (1) 최초 부트스트랩 때 쓴 "재고변동표"
엑셀(그 기간에 입출고가 있었던 품목만) + (2) 매일 판매현황에 처음 등장하는 품목만
자동 등록 — 이 두 경로로만 채워져서, "재고는 있지만 한 번도 안 팔린 품목"은 구조적으로
RAW_품목마스터에 영원히 안 들어오는 문제가 있었다. 사용자가 이카운트 "품목등록"
화면에서 품목유형 전체·사용구분 전체로 뽑은 전체 목록(18,225개, 창고/재고와 무관한
진짜 마스터 목록)을 줘서, 거기 있는데 RAW_품목마스터에 없는 것만 채워넣는다.

- 브랜드코드(국가코드)가 이 원본엔 없어서, 조달유형은 brand_type_overrides.json의
  브랜드명 매칭에만 의존한다(ecount_item_master.classify 재사용). 못 걸리는 새 브랜드는
  "해외수입"(보수적 기본값)로 넣고 확인 필요 목록에 별도로 표로 뽑아준다.
- 사용중단(사용구분=NO) 품목은 품목명에 "(단종)"을 붙여서, 기존 파이프라인의 단종 제외
  로직(ecount_daily_runner.build_daily_rows)이 자동으로 걸러내게 한다 — 마스터에는
  기록으로 남기되 운영 탭에는 안 뜨게.
- 기존 RAW_품목마스터 행은 절대 안 건드린다(append만, replace 아님) — 오늘 애써 고친
  조달유형 값들이 안전하게 보존된다.

사용법 (서버에서):
  .venv/bin/python import_missing_items.py --spreadsheet-id <ID> --dry-run
  .venv/bin/python import_missing_items.py --spreadsheet-id <ID>
"""
from __future__ import annotations

import argparse
import json
from datetime import date
from pathlib import Path

from ecount_daily_runner import _load_creds, read_tab_rows, append_item_master_rows
from ecount_item_master import classify, LEADTIME_BY_TYPE

REGISTRY_PATH = Path(__file__).parent / "one_time_seed_2026-07-30" / "full_item_registry_2026-08-07.json"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", required=True)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    registry = json.loads(REGISTRY_PATH.read_text(encoding="utf-8"))
    print(f"[import] 이카운트 품목등록 전체 {len(registry)}건 로드")

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    existing = read_tab_rows(service, args.spreadsheet_id, "RAW_품목마스터")
    existing_codes = {r["품목코드"] for r in existing}
    print(f"[import] 기존 RAW_품목마스터 {len(existing_codes)}건")

    today = date.today().isoformat()
    new_rows: list[list] = []
    uncertain: list[tuple[str, str]] = []  # (브랜드, 품목코드) — 새 브랜드라 확실치 않음

    for item in registry:
        code = item["품목코드"]
        if code in existing_codes:
            continue
        brand = item["브랜드"]
        name = item["품목명"]
        if item.get("사용구분") == "NO" and "(단종)" not in name:
            name = f"{name} (단종)"
        ptype, certain = classify(brand, "", code)
        if not certain:
            uncertain.append((brand, code))
        leadtime = LEADTIME_BY_TYPE.get(ptype, "")
        new_rows.append([code, name, brand, "", ptype, leadtime, today])

    print(f"[import] 새로 추가할 품목: {len(new_rows)}건")
    if uncertain:
        uniq_brands = sorted({b for b, _ in uncertain})
        print(f"[import] ⚠️ 새 브랜드라 조달유형 확실치 않음(임시 해외수입) — 브랜드 {len(uniq_brands)}개, 품목 {len(uncertain)}건:")
        for b in uniq_brands[:30]:
            print(f"  - {b!r}")
        if len(uniq_brands) > 30:
            print(f"  ... 외 {len(uniq_brands) - 30}개 브랜드")

    if args.dry_run:
        print("[import] --dry-run: 시트에 쓰지 않음")
        return 0
    if not new_rows:
        print("[import] 추가할 게 없음")
        return 0

    append_item_master_rows(service, args.spreadsheet_id, new_rows)
    print("[import] RAW_품목마스터 반영 완료(append) — 기존 행은 안 건드림.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
