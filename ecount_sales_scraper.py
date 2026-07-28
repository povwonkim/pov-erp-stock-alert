#!/usr/bin/env python3
"""이카운트 "판매현황" 자동알림 이메일의 "수신문서보기" 웹 리포트를 스크래핑해서
RAW_판매현황과 같은 형태(list[dict])로 반환하는 서버용 스크립트.

배경: 이 알림 이메일에는 첨부파일이 없고 로그인이 필요한 웹 팝업 링크만 있다
(2026-07-28 확인, ecount_web_diag.py로 진단). 그 팝업 안의 표를 그대로 읽어온다.
링크는 매일 SEND_CM_ID가 바뀌어 하드코딩할 수 없으므로 매번 Gmail에서 새로 찾는다
(ecount_gmail_fetch.find_view_link). 로그인/모달 처리는 ecount_web_diag.py의
open_browser_context / login_if_needed / save_storage_state를 그대로 재사용한다.

사용법 (서버에서):
  .venv/bin/python ecount_sales_scraper.py --out cron_tracking/ecount/sales_status_latest.json
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from ecount_client import _load_secrets
from ecount_web_diag import (
    get_today_view_link,
    login_if_needed,
    open_browser_context,
    save_storage_state,
)

# 실제 웹 표 헤더 → RAW_판매현황 필드명. ecount_daily_runner.SALES_HEADER_MAP과
# 동일한 매핑을 쓰되, 엑셀에는 없던 "일자"(웹 표에만 있는 컬럼)는 무시한다.
SALES_HEADER_MAP = {
    "품목그룹2명": "브랜드", "브랜드": "브랜드",
    "품목코드": "품목코드",
    "품명 및 규격": "품명", "품명": "품명",
    "수량": "수량", "단가": "단가", "공급가액": "공급가액", "부가세": "부가세",
    "합계": "합계", "합 계": "합계",  # 웹 표 헤더는 "합 계"로 띄어쓰기가 들어가 있음(2026-07-28 확인)
    "적요": "적요", "창고명": "창고명",
    "담당자명": "담당자", "사원(담당)명": "담당자", "담당자": "담당자",
}


def _to_number(v: str) -> float:
    v = (v or "").strip().replace(",", "")
    if not v:
        return 0.0
    try:
        return float(v)
    except ValueError:
        return 0.0


def parse_table_rows(header_cells: list[str], body_rows: list[list[str]]) -> list[dict]:
    """헤더 행과 데이터 행(각각 셀 텍스트 리스트)을 RAW_판매현황 형태의 dict 리스트로 변환."""
    col = {SALES_HEADER_MAP[h]: i for i, h in enumerate(header_cells) if h in SALES_HEADER_MAP}
    if "품목코드" not in col:
        raise SystemExit(f"[scraper] 표 헤더에 품목코드 컬럼이 없습니다: {header_cells!r}")

    def get(row: list[str], field: str, default: str = "") -> str:
        idx = col.get(field)
        if idx is None or idx >= len(row):
            return default
        return row[idx]

    rows = []
    for row in body_rows:
        code = get(row, "품목코드").strip()
        if not code:
            continue
        rows.append({
            "브랜드": get(row, "브랜드").strip(),
            "품목코드": code,
            "품명": get(row, "품명").strip(),
            "수량": _to_number(get(row, "수량")),
            "단가": _to_number(get(row, "단가")),
            "공급가액": _to_number(get(row, "공급가액")),
            "부가세": _to_number(get(row, "부가세")),
            "합계": _to_number(get(row, "합계")),
            "적요": get(row, "적요").strip(),
            "창고명": get(row, "창고명").strip(),
            "담당자": get(row, "담당자").strip(),
        })
    return rows


def scrape_sales_status() -> list[dict]:
    from playwright.sync_api import sync_playwright

    secrets = _load_secrets()
    web_user_id = secrets.get("WEB_USER_ID", "POV_API")
    web_password = secrets.get("WEB_PASSWORD", "")
    if not web_password:
        raise SystemExit("[scraper] .secrets/ecount.json에 WEB_PASSWORD가 없습니다.")

    print("[scraper] Gmail에서 오늘의 '수신문서보기' 링크 찾는 중...")
    link = get_today_view_link()
    print(f"[scraper] 링크: {link}")

    DUMP_DIR = Path(__file__).parent / "cron_tracking" / "ecount"
    with sync_playwright() as p:
        browser, context = open_browser_context(p)
        try:
            page = context.new_page()
            page.goto(link, wait_until="networkidle", timeout=60000)
            login_if_needed(page, web_user_id, web_password)
            save_storage_state(context)

            # 표는 <table><tr><td>...</td></tr>...</table> 구조(2026-07-28 확인, 1004행 규모).
            # 이 페이지엔 <table>이 2개 있는데(레이아웃용 + 실제 데이터), 실제 데이터 표는
            # networkidle 이후에도 비동기로 늦게 채워질 때가 있어(2026-07-28, 같은 코드로도
            # 성공/실패가 들쭉날쭉하게 재현됨) — 데이터가 어느 정도 찰 때까지 명시적으로 기다린다.
            try:
                page.wait_for_function(
                    "document.querySelectorAll('table tr').length > 10", timeout=20000
                )
            except Exception as e:
                print(f"[scraper]   표 로딩 대기 타임아웃(무시하고 계속): {e}")

            # 큰 표라 페이지 전체를 순회하는 JS 한 번으로 셀 텍스트를 뽑아 CDP 왕복을 최소화한다
            # (locator.nth() 반복 호출은 행이 많을수록 느리고 메모리 부담도 커짐).
            # <table>이 여러 개면 첫 번째가 아니라 행이 가장 많은(=실제 데이터) 표를 고른다 —
            # 레이아웃용 빈 표가 DOM 순서상 먼저 나올 때 잘못 집던 문제(2026-07-28) 수정.
            table_data = page.evaluate(
                """
                () => {
                  const tables = Array.from(document.querySelectorAll('table'));
                  if (!tables.length) return null;
                  const table = tables.reduce((best, t) =>
                    t.querySelectorAll('tr').length > best.querySelectorAll('tr').length ? t : best
                  , tables[0]);
                  const rows = Array.from(table.querySelectorAll('tr'));
                  return rows.map(tr =>
                    Array.from(tr.querySelectorAll('td,th')).map(td => td.innerText.trim())
                  );
                }
                """
            )

            if not table_data or len(table_data) < 2:
                # 표를 못 찾은 원인을 알 수 있게 실패 시점 화면을 남긴다(뷰포트만 — 큰 표
                # 페이지에서 full_page 스크린샷이 메모리를 과도하게 먹은 적이 있어 그것과 동일한
                # 이유로 여기서도 전체 페이지는 찍지 않는다).
                DUMP_DIR.mkdir(parents=True, exist_ok=True)
                page.screenshot(path=str(DUMP_DIR / "scraper_fail.png"), full_page=False)
                (DUMP_DIR / "scraper_fail.html").write_text(page.content())
                raise SystemExit(
                    f"[scraper] 표를 찾지 못했거나 데이터 행이 없습니다. 실패 시점 화면 저장: "
                    f"{DUMP_DIR / 'scraper_fail.png'} / .html (현재 URL: {page.url})"
                )
        finally:
            browser.close()

    header_cells, *body_rows = table_data
    rows = parse_table_rows(header_cells, body_rows)
    print(f"[scraper] {len(rows)}건 파싱 완료")
    return rows


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", help="결과를 JSON으로 저장할 경로 (미지정 시 표준출력에 건수만 출력)")
    args = ap.parse_args()

    rows = scrape_sales_status()

    if args.out:
        out_path = Path(args.out)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(rows, ensure_ascii=False, indent=2))
        print(f"[scraper] 저장 완료: {out_path} ({len(rows)}행)")
    else:
        print(f"[scraper] {len(rows)}행 (미리보기 최대 3행):")
        for r in rows[:3]:
            print(f"  {r}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
