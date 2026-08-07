#!/usr/bin/env python3
"""품목마스터(구글시트, 사람이 방금 직접 고친 최신 상태)를 읽어서 brand_type_overrides.json
갱신용 파일을 만든다(2026-08-06 — 사용자가 시트에서 대량으로 조달유형을 직접 고친 뒤,
그걸 코드 쪽 brand_type_overrides.json에도 반영해서 앞으로 신상품에도 계속 적용되게
해달라고 요청).

브랜드별로 다수결(가장 많은 SKU가 가진 조달유형)을 뽑는다 — OTHER처럼 여러 브랜드가
섞인 브랜드는 다수결이 의미 없어서 제외한다(SKU_OVERRIDES로 개별 처리).

사용법 (서버에서):
  .venv/bin/python sync_brand_overrides_from_master.py --spreadsheet-id <ID> \
      --out brand_type_overrides_from_sheet.json

결과 파일을 다운로드해서 Claude에게 다시 전달하면 브랜드별로 확인 후 저장소의
brand_type_overrides.json에 반영합니다.
"""
from __future__ import annotations

import argparse
import json

from ecount_daily_runner import _load_creds, read_tab_rows

MIXED_BRANDS = {"OTHER"}  # 여러 브랜드가 섞인 통 브랜드 — 다수결로 대표값 뽑으면 안 됨


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", required=True)
    ap.add_argument("--out", default="brand_type_overrides_from_sheet.json")
    args = ap.parse_args()

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    rows = read_tab_rows(service, args.spreadsheet_id, "품목마스터")
    print(f"[sync] 품목마스터 {len(rows)}행 로드")

    votes: dict[str, dict[str, int]] = {}
    for r in rows:
        brand = (r.get("브랜드") or "").strip()
        ptype = (r.get("조달유형") or "").strip()
        if not brand or not ptype or ptype == "미분류" or brand in MIXED_BRANDS:
            continue
        votes.setdefault(brand, {})
        votes[brand][ptype] = votes[brand].get(ptype, 0) + 1

    result = {}
    mixed_found = []
    for brand, counts in sorted(votes.items()):
        dominant = max(counts, key=counts.get)
        result[brand] = dominant
        if len(counts) > 1:
            mixed_found.append((brand, counts))

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2, sort_keys=True)

    print(f"[sync] {len(result)}개 브랜드 → {args.out} 저장 완료")
    if mixed_found:
        print(f"[sync] ⚠️ 브랜드 안에서 조달유형이 갈리는 것 {len(mixed_found)}개(다수결로 대표값만 뽑음):")
        for brand, counts in mixed_found:
            detail = ", ".join(f"{t}:{c}개" for t, c in counts.items())
            print(f"  - {brand!r}: {detail}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
