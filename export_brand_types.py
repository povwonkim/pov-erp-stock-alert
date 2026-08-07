#!/usr/bin/env python3
"""RAW_품목마스터 탭의 전체 브랜드 목록 + 현재 조달유형을 엑셀로 뽑아서 사람이 검토·수정할 수 있게
한다(2026-08-06 사용자 요청 — "수정할게 많아 보이는데 엑셀로 쫙 주면 내가 표기를 다 해서
줄게").

사용법 (서버에서):
  .venv/bin/python export_brand_types.py --spreadsheet-id <ID> --out 브랜드_조달유형_검토.xlsx

받은 파일을 다운로드해서(scp) 열어보면 "조달유형(수정)" 칸이 비어있는 엑셀이 나온다.
자체제작/국내사입/국내위탁/해외수입 중 바꿀 브랜드만 그 칸에 채워서(드롭다운으로 고르면 됨)
Claude에게 다시 전달하면, apply_brand_types.py로 brand_type_overrides.json에 반영하고
기존 RAW_품목마스터 행도 일괄 수정한다.
"""
from __future__ import annotations

import argparse

from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

from ecount_daily_runner import _load_creds, read_tab_rows

VALID_TYPES = ["자체제작", "국내사입", "국내위탁", "해외수입"]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", required=True)
    ap.add_argument("--out", default="brand_type_review.xlsx")
    args = ap.parse_args()

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    rows = read_tab_rows(service, args.spreadsheet_id, "RAW_품목마스터")
    print(f"[export] RAW_품목마스터 {len(rows)}행 로드")

    by_brand: dict[str, dict] = {}
    for r in rows:
        brand = (r.get("브랜드") or "").strip()
        if not brand:
            continue
        s = by_brand.setdefault(brand, {"code": r.get("브랜드코드", ""), "types": {}, "count": 0})
        s["count"] += 1
        t = r.get("조달유형") or "미분류"
        s["types"][t] = s["types"].get(t, 0) + 1

    wb = Workbook()
    ws = wb.active
    ws.title = "브랜드별 조달유형"
    ws.append(["브랜드", "브랜드코드", "품목수", "현재 조달유형", "섞여있음(참고)", "조달유형(수정)", "메모"])

    for brand in sorted(by_brand, key=lambda b: -by_brand[b]["count"]):
        s = by_brand[brand]
        dominant = max(s["types"], key=s["types"].get)
        mixed = ", ".join(f"{t}:{c}개" for t, c in s["types"].items()) if len(s["types"]) > 1 else ""
        ws.append([brand, s["code"], s["count"], dominant, mixed, "", ""])

    # F열(조달유형(수정))에 드롭다운 — 오타/오기입 방지.
    dv = DataValidation(type="list", formula1=f'"{",".join(VALID_TYPES)}"', allow_blank=True)
    dv.error = "자체제작/국내사입/국내위탁/해외수입 중 하나만 선택하세요"
    dv.prompt = "바꿀 브랜드만 채우세요. 비워두면 현재 값 그대로 유지됩니다."
    ws.add_data_validation(dv)
    dv.add(f"F2:F{ws.max_row}")

    for col, width in zip("ABCDEFG", [24, 12, 10, 16, 24, 16, 30]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"

    wb.save(args.out)
    print(f"[export] {len(by_brand)}개 브랜드 → {args.out} 저장 완료")
    print("[export] F열(조달유형(수정))에 바꿀 브랜드만 채워서 다시 전달하면 됩니다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
