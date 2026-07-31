#!/bin/bash
# 이카운트 일일 파이프라인을 재시도 로직과 함께 실행하는 cron용 래퍼.
#
# 배경(2026-07-31): 10:30 크론 실행이 "품목마스터" 구글시트 읽기 단계에서
# 403 insufficient authentication scopes로 실패했다 — 같은 인증 파일/코드로 몇 시간
# 뒤 수동 실행은 바로 성공했고, 코드·환경 차이도 못 찾아서 구글 OAuth 쪽 일시적
# 현상으로 보고 있다(원인 규명 어려움, 재현 안 됨). 재시도로 흡수한다.
#
# 1차 실패 시 60초 후 재시도하되, 이미 캐시된 재고/판매 원본(ecount_daily_runner.py가
# 매 실행마다 cron_tracking/ecount/에 저장함)을 재사용해서 API 재조회 없이 곧바로
# 계산·쓰기 단계로 넘어간다 — 창고별 재고 API가 10분/회 제한이라 처음부터 다시 돌면
# 30분 넘게 걸려 비효율적.
#
# 사용법: run_daily_with_retry.sh <SPREADSHEET_ID>
set -uo pipefail
cd "$(dirname "$0")"

SPREADSHEET_ID="${1:?사용법: run_daily_with_retry.sh <SPREADSHEET_ID>}"

.venv/bin/python ecount_daily_runner.py --spreadsheet-id "$SPREADSHEET_ID"
STATUS=$?

if [ $STATUS -ne 0 ]; then
    echo "[retry] 1차 실행 실패(종료코드 $STATUS) — 60초 후 캐시 데이터로 재시도"
    sleep 60
    .venv/bin/python ecount_daily_runner.py --spreadsheet-id "$SPREADSHEET_ID" \
        --use-cached-inventory --use-cached-sales
    STATUS=$?
    if [ $STATUS -ne 0 ]; then
        echo "[retry] 2차 재시도도 실패(종료코드 $STATUS) — 로그 확인 후 수동 재실행 필요"
    else
        echo "[retry] 재시도 성공"
    fi
fi

exit $STATUS
