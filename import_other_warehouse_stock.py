#!/usr/bin/env python3
"""메인 4개 창고(OFFLINE_WAREHOUSES) 밖의 나머지 19개 창고 재고를 "기타창고_재고현황" 탭에,
이카운트 창고 23개 전체 목록을 "창고목록" 탭에 1회성으로 채워넣는다.

배경(2026-08-07): 이 시스템은 처음부터 오프라인 4개 창고만 API로 자동 조회하도록 설계돼서
(온라인 창고는 아예 조회 안 함), 나머지 19개 창고에 재고가 쌓여도 지금까지 어디에도 안
보이고 있었다. 사용자가 이카운트에서 "29CM 외"(=메인 4개 제외) 재고현황을 직접 뽑아준
결과, 47,585개(1,667개 품목)가 이 19개 창고에 흩어져 있는 게 확인됨 — "이게 진짜 방치재고".

19개 창고는 API 자동조회 대상이 아니므로 매일 자동 갱신되지 않는다(수동 갱신 전용 탭).
창고목록 탭은 담당자(수빈실장)가 각 창고의 용도/처리방침을 직접 채워넣는 용도.

사용법 (서버에서):
  .venv/bin/python import_other_warehouse_stock.py --spreadsheet-id <ID>
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path

from ecount_daily_runner import _load_creds, replace_tab_rows
from ecount_sheets_setup import TABS, OFFLINE_WAREHOUSES

DATA_PATH = Path(__file__).parent / "one_time_seed_2026-08-07" / "other_warehouse_stock.json"

# 이카운트 "창고검색" 화면 전체 목록(2026-08-07 확인, 23개).
ALL_WAREHOUSES = [
    ("00022", "29CM"), ("00038", "DDP디자인스토어"), ("00034", "HQ_office"),
    ("00030", "MXN"), ("00031", "MXN(29cm)"), ("00033", "MXN(온라인)"),
    ("00036", "MXN(올리브영)"), ("00039", "MXN(카카오)"), ("00008", "OTS"),
    ("00027", "OY"), ("00014", "POINT OF VIEW (법인)"), ("00018", "POP-UP"),
    ("00004", "POV"), ("00012", "THE HYUNDAI SEOUL"), ("00002", "orer.archive"),
    ("00021", "시시호시-수원점"), ("00032", "신세계 강남-피숀"), ("00017", "온라인"),
    ("00035", "올리브영"), ("00025", "올리브영(성수N)"), ("00029", "올리브영(홍대)"),
    ("00028", "카카오 선물하기(단품 재고)"), ("00020", "카카오 선물하기(세트 재고)"),
]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--spreadsheet-id", required=True)
    args = ap.parse_args()

    creds = _load_creds()
    from googleapiclient.discovery import build
    service = build("sheets", "v4", credentials=creds)

    # ---- 창고목록 ----
    warehouse_headers = TABS["창고목록"]["headers"]
    warehouse_rows = []
    for code, name in ALL_WAREHOUSES:
        tracked = "Y (자동)" if name in OFFLINE_WAREHOUSES else "N (수동)"
        warehouse_rows.append([code, name, tracked, "", ""])
    replace_tab_rows(service, args.spreadsheet_id, "창고목록", warehouse_rows)
    print(f"[import] 창고목록 {len(warehouse_rows)}건 반영 완료")

    # ---- 기타창고_재고현황 ----
    data = json.loads(DATA_PATH.read_text(encoding="utf-8"))
    warehouses = data["warehouses"]
    items = data["items"]
    stock_headers = TABS["기타창고_재고현황"]["headers"]

    rows = []
    for it in items:
        wh_qtys = it["창고별재고"]
        row = [it["브랜드"], it["품목코드"], it["품명"]]
        row += [wh_qtys.get(wh, "") for wh in warehouses]
        row += [it["재고수량"], it["바코드"], it["재고보유월수"], "", ""]
        rows.append(row)
    rows.sort(key=lambda r: -r[len(warehouses) + 3])  # 재고수량 내림차순

    assert len(stock_headers) == len(rows[0]) if rows else True
    replace_tab_rows(service, args.spreadsheet_id, "기타창고_재고현황", rows)
    print(f"[import] 기타창고_재고현황 {len(rows)}건 반영 완료 (재고수량 합계 {sum(it['재고수량'] for it in items):,.0f}개)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
