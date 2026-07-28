#!/usr/bin/env python3
"""이카운트 웹 "재고잔량분석표"(재고Ⅰ > 출력물 > 재고현황 > 재고잔량분석표) 페이지 구조를
확인하기 위한 1회성 진단 스크립트.

배경: 이 리포트는 품목별 "입고(취득) 시점"을 이카운트가 이미 계산해서 보여준다 — 이게 있으면
우리가 일별재고이력을 며칠씩 쌓을 때까지 기다리지 않고도 첫날부터 악성재고/샘플의심재고를
정확히 판별할 수 있다(2026-07-28 사용자가 이카운트 도움말에서 찾아옴).

다만 이건 판매현황처럼 이메일로 오는 팝업 링크가 아니라, 로그인 후 메뉴로 들어가는
일반 리포트 화면이다(URL 예: https://loginac.ecount.com/ec5/view/erp?w_flag=1&
ec_req_sid=...#menuType=MENUTREE_000004&menuSeq=MENUTREE_002775&groupSeq=MENUTREE_000035&
prgId=E040727&depth=4). ec_req_sid는 사용자가 수동으로 브라우저에서 접속했을 때 발급된
값이라 우리 자동화 세션에서 그대로 쓰면 만료됐을 가능성이 높다 — 그래도 일단 시도해보고
안 되면(로그인 화면으로 튕기면) 그 상태를 스크린샷/HTML로 남겨서 다음 단계를 판단한다.
검색 조건(날짜 등)을 채우고 "조회" 버튼을 눌러야 표가 채워지는 화면일 가능성이 높아서,
이번엔 표뿐 아니라 입력폼/버튼 구조도 같이 확인한다.

사용법 (서버에서):
  .venv/bin/python ecount_stock_aging_diag.py
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

# 2026-07-28 사용자가 수동으로 이 화면에 들어갔을 때의 URL. ec_req_sid는 그 세션에서만
# 유효했을 수 있음 — 실패하면 로그인 화면으로 튕기는지 확인하고 대안을 찾는다.
TARGET_URL = (
    "https://loginac.ecount.com/ec5/view/erp?w_flag=1&ec_req_sid=AC-ETgmELJTHcSH2"
    "#menuType=MENUTREE_000004&menuSeq=MENUTREE_002775&groupSeq=MENUTREE_000035"
    "&prgId=E040727&depth=4"
)


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
            print(f"[diag] 페이지 로드 중... {TARGET_URL}")
            page.goto(TARGET_URL, wait_until="networkidle", timeout=60000)

            page.screenshot(path=str(DUMP_DIR / "aging_01_initial.png"), full_page=False)
            (DUMP_DIR / "aging_01_initial.html").write_text(page.content())
            print(f"[diag] 초기 화면 저장: {DUMP_DIR / 'aging_01_initial.png'} (현재 URL: {page.url})")

            dismiss_new_device_modal(page)
            had_login_form = page.locator("#txtPass").count() > 0
            login_if_needed(page, web_user_id, web_password)
            save_storage_state(context)

            if had_login_form:
                print(f"[diag] 로그인 폼 발견 — 로그인 시도 후 현재 URL: {page.url}")
                # 로그인 후 원래 목표 URL(해시 포함)로 다시 시도 — SPA 라우팅이라 로그인 성공 후
                # 루트로 리다이렉트됐을 수 있어, 해시를 다시 세팅해서 해당 메뉴로 이동을 시도한다.
                if "prgId=E040727" not in page.url:
                    print("[diag] 로그인 후 목표 화면이 아님 — 원래 URL로 재시도...")
                    page.goto(TARGET_URL, wait_until="networkidle", timeout=60000)
            else:
                print("[diag] 로그인 폼 없음 — 이미 인증된 상태이거나 다른 구조")

            page.wait_for_timeout(2000)  # SPA 렌더링 여유
            page.screenshot(path=str(DUMP_DIR / "aging_02_after_login.png"), full_page=False)
            (DUMP_DIR / "aging_02_after_login.html").write_text(page.content())
            print(f"[diag] 로그인 후 화면 저장: {DUMP_DIR / 'aging_02_after_login.png'} (현재 URL: {page.url})")

            # 표/입력폼/버튼 구조를 한 번에 파악 — 검색조건을 채우고 조회 버튼을 눌러야 하는
            # 화면일 가능성이 높아서 표뿐 아니라 폼 요소도 같이 센다.
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
        finally:
            browser.close()

    print("\n[diag] 완료. 아래 파일들을 확인해서 실제 화면 구조를 파악하세요:")
    for f in sorted(DUMP_DIR.glob("aging_*")):
        print(f"  - {f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
