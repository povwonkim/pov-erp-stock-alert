#!/usr/bin/env python3
"""이카운트 웹 "재고현황"(재고Ⅰ > 출력물 > 재고현황 > 재고현황) 페이지 구조를 확인하기
위한 1회성 진단 스크립트.

배경: 처음엔 "재고잔량분석표"(prgId=E040727)를 시도했는데, 거긴 "입고월별" breakdown이라
"기타"에 몇 개월/몇 년 전 재고가 다 뭉뚱그려져서 정확한 정체기간을 알기 어려웠다. 그런데
사용자가 직접 화면에서 "재고현황"(prgId=E040701, 다른 메뉴) 리포트를 열어보니 맨 오른쪽에
**"재고보유월수"** 컬럼이 있고 "3개월"/"20개월"/"24개월초과"처럼 이카운트가 이미 정확한
개월수를 계산해서 보여준다(2026-07-29 사용자 스크린샷으로 확인) — 게다가 창고별 컬럼(29CM,
시시호시, 신세계, MXN 계열, POV 법인, 현대서울 등)도 이미 다 포함되어 있어 우리 목적에
훨씬 잘 맞는다. ecount_stock_aging_diag.py와 로그인/해시라우팅 로직은 동일하게 재사용.

사용법 (서버에서):
  .venv/bin/python ecount_stock_status_diag.py
"""
from __future__ import annotations

import sys
from pathlib import Path

from ecount_client import _load_secrets
from ecount_web_diag import (
    dismiss_new_device_modal,
    login_if_needed,
    open_browser_context,
    save_storage_state,
)

DUMP_DIR = Path(__file__).parent / "cron_tracking" / "ecount"

# 2026-07-29 사용자가 수동으로 이 화면에 들어갔을 때의 URL(재고현황, 재고잔량분석표와는
# 다른 메뉴 — prgId=E040701). ec_req_sid는 그 세션에서만 유효했을 수 있음.
TARGET_URL = (
    "https://loginac.ecount.com/ec5/view/erp?w_flag=1&ec_req_sid=AC-ETh1fjsf8xoDa"
    "#menuType=MENUTREE_000004&menuSeq=MENUTREE_000212&groupSeq=MENUTREE_000035"
    "&prgId=E040701&depth=4"
)
TARGET_PRG_ID = "prgId=E040701"


def main() -> int:
    secrets = _load_secrets()
    web_user_id = secrets.get("WEB_USER_ID", "POV_API")
    web_password = secrets.get("WEB_PASSWORD", "")
    if not web_password:
        print("[diag] .secrets/ecount.json에 WEB_PASSWORD가 없습니다.", file=sys.stderr)
        return 2

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("[diag] playwright가 설치되어 있지 않습니다.", file=sys.stderr)
        return 2

    DUMP_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser, context = open_browser_context(p)
        try:
            page = context.new_page()
            page.set_default_timeout(15000)
            print(f"[diag] 페이지 로드 중... {TARGET_URL}")
            page.goto(TARGET_URL, wait_until="networkidle", timeout=60000)

            page.screenshot(path=str(DUMP_DIR / "status_01_initial.png"), full_page=False)
            (DUMP_DIR / "status_01_initial.html").write_text(page.content())
            print(f"[diag] 초기 화면 저장: {DUMP_DIR / 'status_01_initial.png'} (현재 URL: {page.url})")

            dismiss_new_device_modal(page)
            had_login_form = page.locator("#txtPass").count() > 0 or page.locator("#passwd").count() > 0
            login_if_needed(page, web_user_id, web_password, com_code=secrets.get("COM_CODE", ""))
            save_storage_state(context)

            if had_login_form:
                print(f"[diag] 로그인 폼 발견 — 로그인 시도 후 현재 URL: {page.url}")
                if TARGET_PRG_ID not in page.url:
                    target_hash = TARGET_URL.split("#", 1)[1]
                    print(f"[diag] 로그인 후 목표 화면이 아님 — 현재 URL에 해시만 재설정: #{target_hash}")
                    page.evaluate(f"window.location.hash = {target_hash!r}")
                    page.wait_for_timeout(3000)
            else:
                print("[diag] 로그인 폼 없음 — 이미 인증된 상태이거나 다른 구조")

            page.wait_for_timeout(2000)
            page.screenshot(path=str(DUMP_DIR / "status_02_after_login.png"), full_page=False)
            (DUMP_DIR / "status_02_after_login.html").write_text(page.content())
            print(f"[diag] 로그인 후 화면 저장: {DUMP_DIR / 'status_02_after_login.png'} (현재 URL: {page.url})")

            summary = page.evaluate(
                """
                () => {
                  const tables = document.querySelectorAll('table').length;
                  const rows = document.querySelectorAll('table tr').length;
                  const inputs = Array.from(document.querySelectorAll('input')).map(el => ({
                    id: el.id, name: el.name, type: el.type, placeholder: el.placeholder,
                  })).slice(0, 30);
                  const buttons = Array.from(document.querySelectorAll('button, input[type=button], a[onclick]'))
                    .map(el => (el.innerText || el.value || '').trim()).filter(Boolean).slice(0, 30);
                  return { tables, rows, inputs, buttons };
                }
                """
            )
            print(f"[diag] <table> {summary['tables']}개, <tr> {summary['rows']}개")
            print(f"[diag] 입력창 최대 30개: {summary['inputs']}")
            print(f"[diag] 버튼/링크 텍스트 최대 30개: {summary['buttons']}")

            zero_qty_chk = page.locator('input[name="ZERO_QTY_INCLUDE_YN"]')
            if zero_qty_chk.count() > 0:
                zero_qty_chk.first.check()
                print("[diag] '재고수량0포함' 체크 완료")
            else:
                print("[diag] '재고수량0포함' 체크박스 없음(이 화면엔 없을 수 있음)")

            search_btn = page.get_by_text("검색(F8)", exact=True)
            if search_btn.count() == 0:
                search_btn = page.get_by_text("Search(F3)", exact=True)
            if search_btn.count() > 0:
                print("[diag] 검색 버튼 발견 — 클릭 시도...", flush=True)
                try:
                    search_btn.first.click(timeout=10000)
                    print("[diag] 클릭 완료", flush=True)
                except Exception as e:
                    print(f"[diag]   클릭 타임아웃/실패(무시하고 계속): {e}", flush=True)
                try:
                    page.wait_for_load_state("networkidle", timeout=20000)
                except Exception as e:
                    print(f"[diag]   networkidle 대기 타임아웃(무시): {e}")
                print("[diag] networkidle 대기 종료, 다음 단계로", flush=True)
                try:
                    page.wait_for_function(
                        "document.querySelectorAll('table tr').length > 1", timeout=30000
                    )
                except Exception as e:
                    print(f"[diag]   표 로딩 대기 타임아웃(무시하고 계속): {e}")
                page.wait_for_timeout(1000)

                page.screenshot(path=str(DUMP_DIR / "status_03_search_result.png"), full_page=False)
                (DUMP_DIR / "status_03_search_result.html").write_text(page.content())
                print(f"[diag] 검색 결과 화면 저장: {DUMP_DIR / 'status_03_search_result.png'}")

                result_summary = page.evaluate(
                    """
                    () => {
                      const tables = Array.from(document.querySelectorAll('table'));
                      const best = tables.reduce((a, t) => t.querySelectorAll('tr').length > a.querySelectorAll('tr').length ? t : a, tables[0]);
                      if (!best) return { tables: 0, rows: 0, headerPreview: [], firstRowPreview: [] };
                      const rows = Array.from(best.querySelectorAll('tr'));
                      const cellText = tr => Array.from(tr.querySelectorAll('td,th')).map(td => td.innerText.trim());
                      return {
                        tables: tables.length,
                        rows: rows.length,
                        headerPreview: rows[0] ? cellText(rows[0]) : [],
                        firstRowPreview: rows[1] ? cellText(rows[1]) : [],
                        secondRowPreview: rows[2] ? cellText(rows[2]) : [],
                      };
                    }
                    """
                )
                print(f"[diag] 검색 후 표: {result_summary}")
            else:
                print("[diag] 검색 버튼을 못 찾음")
        finally:
            browser.close()

    print("\n[diag] 완료. 아래 파일들을 확인해서 실제 화면 구조를 파악하세요:")
    for f in sorted(DUMP_DIR.glob("status_*")):
        print(f"  - {f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
