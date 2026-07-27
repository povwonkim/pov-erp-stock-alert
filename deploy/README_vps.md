# VPS 배포 가이드 — 이카운트 대조 매일 자동실행

이카운트 API는 **등록된 고정 IP에서만** 호출을 허용한다. GitHub Actions는 IP가 매번 바뀌어
쓸 수 없으므로, 이카운트 대조 작업은 **고정 IP를 가진 작은 VPS**에서 매일 돌린다.

## 왜 VPS인가
- 고정 IP가 **절대 안 바뀜** → 이카운트 [IP등록]에 한 번 넣으면 끝
- 24시간 켜져 있어 매일 예약실행(cron)에 적합
- 월 3~6천 원 수준

---

## 1. 가비아 클라우드 서버(g클라우드) 생성

가비아 → **클라우드(g클라우드)** → 서버 생성:
- OS: **Ubuntu 22.04 LTS** (또는 최신 LTS)
- 사양: 최저(1 vCPU / 1GB RAM)로 충분
- 계정: 가비아 g클라우드는 서버 생성 시 별도 계정을 안 만들므로 **root**가 아이디. 콘솔에서 root 비밀번호 확인/설정.

## 2. 공인 IP 연결 (= 우리 고정 IP)

가비아는 서버에 **공인 IP를 별도로 연결**해야 외부 통신이 된다. 이 공인 IP가 이카운트에 등록할
고정 주소다.
- 가비아 콘솔 → **공인 IP 연결하기** → 서버(서브넷)에 공인 IP 연결
  (참고: https://customer.gabia.com/manual/32/6758/14729 )
- **방화벽**: 방화벽을 만들어 **SSH(22) 허용** → 해당 방화벽을 서버 서브넷에 적용.
  (아웃바운드 HTTPS는 기본 허용이라 이카운트 호출은 문제 없음)
- 이 공인 IP는 서버를 유지하는 한 바뀌지 않는다 → 이카운트 등록에 적합.

## 3. 이카운트에 IP 등록
- 이카운트 → **API인증키 발급 → [IP등록]** → 위 가비아 공인 IP 입력 → 저장
- 개발 중 내 PC에서도 테스트하려면 내 PC의 IP( whatismyip.com )도 같이 등록해두면 편하다.
- 테스트(sboapi)/운영(oapi) 등록이 나뉠 수 있으니 화면 안내 확인.

## 4. 서버 접속 & 코드 설치

서버 접속: **root + 비밀번호**(또는 키페어)로 SSH. 예: `ssh root@<가비아_공인IP>`
(Windows는 PuTTY, Mac은 터미널. 가비아 매뉴얼: 리눅스 서버 원격 접속 참고)

**비공개 레포 clone** — 이 레포는 private이라 인증이 필요하다. 둘 중 하나:
- **A. Personal Access Token(PAT)**: GitHub → Settings → Developer settings → Fine-grained token
  (이 레포 read 권한) 발급 후:
  ```bash
  git clone https://<TOKEN>@github.com/povwonkim/pov-erp-stock-alert.git ~/pov-erp-stock-alert
  ```
- **B. Deploy key(SSH)**: 서버에서 `ssh-keygen` → 공개키를 GitHub 레포 → Settings → Deploy keys 등록 →
  `git clone git@github.com:povwonkim/pov-erp-stock-alert.git`

clone 후 세팅 스크립트 실행:
```bash
cd ~/pov-erp-stock-alert
bash deploy/setup_vps.sh
```
(python·의존성 설치 + 시간대 Asia/Seoul + `.secrets/ecount.json` 템플릿 생성)

## 5. 인증키 넣기
```bash
nano ~/pov-erp-stock-alert/.secrets/ecount.json
```
회사코드·사용자ID·인증키를 채우고 저장. (이 파일은 .gitignore되어 커밋 안 됨)

인증 검증:
```bash
cd ~/pov-erp-stock-alert
.venv/bin/python ecount_probe.py --check       # Zone+로그인 OK 확인
.venv/bin/python ecount_probe.py --inventory    # 재고 응답 구조 덤프
```
여기까지 되면 IP 등록·인증이 정상이라는 뜻.

## 6. 매일 자동실행 (cron)

서버 시간대가 Asia/Seoul인지 확인(`date`). 그다음:
```bash
crontab -e
```
아래 줄 추가 (매일 **10:30 KST** — 10시 주문마감 배치가 MXN에 반영된 후):
```cron
30 10 * * * cd $HOME/pov-erp-stock-alert && .venv/bin/python ecount_daily_runner.py >> $HOME/ecount_cron.log 2>&1
```
> `ecount_daily_runner.py`(대조→구글시트→슬랙)는 인증키로 실제 데이터 구조를 확인한 뒤 추가된다.
> 그 전까지는 위 줄 대신 `ecount_probe.py --check`로 연결만 확인해도 된다.

로그 확인: `tail -f ~/ecount_cron.log`

---

## 보안 체크
- 인증키는 **서버의 `.secrets/`에만**. 절대 git에 커밋하지 않는다(.gitignore 처리됨).
- 서버 방화벽은 SSH(22)만 열어두면 충분(아웃바운드 HTTPS로 API 호출).
- PAT/Deploy key는 이 레포 **읽기 권한만** 부여.
