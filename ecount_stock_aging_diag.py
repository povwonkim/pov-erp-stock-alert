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
            # 이전 실행에서 검색(F8) 클릭 이후 정체불명의 지점에서 몇 분씩 멈추는 문제가
            # 있었다(2026-07-29) — 어느 액션이 원인인지 특정이 안 돼서, 이후 모든 개별
            # 액션(클릭/evaluate 등)에 공통으로 짧은 기본 타임아웃을 걸어 어디서든 최대
            # 15초 안에 실패하고 다음으로 넘어가게 한다.
            page.set_default_timeout(15000)
            print(f"[diag] 페이지 로드 중... {TARGET_URL}")
            page.goto(TARGET_URL, wait_until="networkidle", timeout=60000)

            page.screenshot(path=str(DUMP_DIR / "aging_01_initial.png"), full_page=False)
            (DUMP_DIR / "aging_01_initial.html").write_text(page.content())
            print(f"[diag] 초기 화면 저장: {DUMP_DIR / 'aging_01_initial.png'} (현재 URL: {page.url})")

            dismiss_new_device_modal(page)
            # 메인 ERP 앱 로그인 폼(#passwd, 회사코드 칸 있음)일 수도 있어 둘 다 확인.
            had_login_form = page.locator("#txtPass").count() > 0 or page.locator("#passwd").count() > 0
            login_if_needed(page, web_user_id, web_password, com_code=secrets.get("COM_CODE", ""))
            save_storage_state(context)

            if had_login_form:
                print(f"[diag] 로그인 폼 발견 — 로그인 시도 후 현재 URL: {page.url}")
                # 로그인 성공 후 ec5/view/erp 앱 화면까지는 오는데 해시(#menuType=...)가
                # 날아간다(2026-07-28 확인) — 원래 TARGET_URL로 다시 이동하면 그 URL에 박혀있는
                # 옛(만료된) ec_req_sid 때문에 도로 로그인 화면으로 튕긴다. 새로 발급된 세션은
                # 그대로 두고 해시만 다시 세팅해서 SPA가 클라이언트 라우팅하게 한다.
                if "prgId=E040727" not in page.url:
                    target_hash = TARGET_URL.split("#", 1)[1]
                    print(f"[diag] 로그인 후 목표 화면이 아님 — 현재 URL에 해시만 재설정: #{target_hash}")
                    page.evaluate(f"window.location.hash = {target_hash!r}")
                    page.wait_for_timeout(3000)
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

            # "재고수량0포함" 체크(2026-07-28 사용자 요청) — name="ZERO_QTY_INCLUDE_YN"
            # (요약 로그의 입력창 목록에서 확인). 재고 0인 품목도 포함해서 봐야 품절 판정에
            # 쓸 수 있다.
            zero_qty_chk = page.locator('input[name="ZERO_QTY_INCLUDE_YN"]')
            if zero_qty_chk.count() > 0:
                zero_qty_chk.first.check()
                print("[diag] '재고수량0포함' 체크 완료")
            else:
                print("[diag] '재고수량0포함' 체크박스를 못 찾음")

            # 검색 조건 화면(품목/창고 필터, 기준일자)까지는 왔으니 "검색(F8)"을 눌러 표를
            # 채워본다 — 기준일자가 이미 오늘로 기본 설정돼 있는 걸 화면에서 확인함.
            search_btn = page.get_by_text("검색(F8)", exact=True)
            if search_btn.count() > 0:
                print("[diag] '검색(F8)' 버튼 발견 — 클릭...")
                search_btn.first.click()
                try:
                    page.wait_for_load_state("networkidle", timeout=20000)
                except Exception as e:
                    print(f"[diag]   networkidle 대기 타임아웃(무시): {e}")
                # 판매현황 스크래핑 때와 동일하게, 검색 직후 표가 비동기로 늦게 채워질 수 있어
                # 실제 행이 어느 정도 찰 때까지 명시적으로 기다린다(2026-07-28: 로딩 스플래시
                # 화면만 찍힌 스크린샷으로 확인 — 3초 고정 대기로는 부족했음).
                try:
                    page.wait_for_function(
                        "document.querySelectorAll('table tr').length > 1", timeout=30000
                    )
                except Exception as e:
                    print(f"[diag]   표 로딩 대기 타임아웃(무시하고 계속): {e}")
                page.wait_for_timeout(1000)

                page.screenshot(path=str(DUMP_DIR / "aging_03_search_result.png"), full_page=False)
                (DUMP_DIR / "aging_03_search_result.html").write_text(page.content())
                print(f"[diag] 검색 결과 화면 저장: {DUMP_DIR / 'aging_03_search_result.png'}")

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
                      };
                    }
                    """
                )
                print(f"[diag] 검색 후 표: {result_summary}")

                # 결과가 iframe 안에 있을 수도 있어(로딩 스플래시가 iframe 자체 로딩화면처럼
                # 보였음) 프레임별로도 확인한다 — 메인 프레임에서 못 찾았을 때만.
                if result_summary["rows"] == 0 and len(page.frames) > 1:
                    print(f"[diag] 메인 프레임에 표 없음 — 하위 프레임 {len(page.frames) - 1}개 확인 중...")
                    for fr in page.frames:
                        if fr == page.main_frame:
                            continue
                        try:
                            fr_summary = fr.evaluate(
                                "() => ({ tables: document.querySelectorAll('table').length, "
                                "rows: document.querySelectorAll('table tr').length })"
                            )
                        except Exception as e:
                            fr_summary = {"error": str(e)}
                        print(f"[diag]   프레임 {fr.url!r}: {fr_summary}")
            else:
                print("[diag] '검색(F8)' 버튼을 못 찾음")
        finally:
            browser.close()

    print("\n[diag] 완료. 아래 파일들을 확인해서 실제 화면 구조를 파악하세요:")
    for f in sorted(DUMP_DIR.glob("aging_*")):
        print(f"  - {f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
