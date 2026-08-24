# CLAUDE.md — W Concept 마케팅 KPI 대시보드

> 이 파일은 **모든 모델·모든 세션에 자동 로드**되는 단일 규칙 소스입니다.
> 아래 철칙(R1~R8)은 "가능하면"이 아니라 **필수 게이트**입니다. 지키기 전엔 "됐다"고 보고하지 마세요.

## 시스템 개요
- **대시보드**: `KPI/pmkt-kpi.html` — 단일 파일 정적 SPA(Vanilla JS + Chart.js). GitHub Pages(`https://wckmkt.github.io/dashboard/KPI/pmkt-kpi.html`)로 배포. 백엔드 없음.
- **데이터**: `KPI/pmkt_kpi_data.json` — `C:\bang\gen_realtime_dashboard.py`가 OneDrive Excel 마스터파일을 COM으로 읽어 생성 → git 자동 커밋·push. Windows 작업스케줄러가 하루 5회(09/12/14/16/18시) 구동.
- **레포 2개(별개 git)**:
  - `C:\bang\pmkt-dashboard` = 이 레포(대시보드·데이터, 커밋 대상).
  - `C:\bang` = 파이프라인 스크립트(gen_realtime_dashboard.py, teams_notify.ps1 등). **명시적 요청 없이는 커밋하지 말 것.**
- 세션 작업 디렉토리는 보통 `KPI/`. 상세 운영지식은 fable5 로컬 memory에만 있으니, **모델 무관 규칙은 이 CLAUDE.md가 유일 소스**.
- 이 레포에는 **광고 성과 대시보드(`ad.html`)도 함께** 있다. 데이터·파이프라인이 완전히 별개이므로 아래 전용 섹션을 볼 것.

## 철칙 (Golden Rules — 검증 게이트)

**R1. 숫자는 원본에서, 검증 후 말한다.**
데이터가 이상하다고 단정하기 전에 `python C:\bang\gen_realtime_dashboard.py`로 JSON을 재생성하고 Excel 마스터 셀과 직접 대조한다(기억·추측으로 진단 금지). 값이 어긋나면 **가장 raw한 단위(원 단위, 억 아님)**에서 재계산한다 — 억 단위로 반올림한 뒤 비율(YoY)을 계산하면 오차가 누적된 사례가 있다.

**R2. 기존 검증 함수를 재사용한다. 병렬 계산 로직을 새로 만들지 않는다.**
특히 "지원이"(Q&A): LLM은 **의도 파싱·문장화만** 하고 **모든 숫자는 이미 화면 렌더링에 쓰이는 집계 함수**(`aggWeekFull`/`getFunnelDiagnosis`/전역 `ANNUAL` 배열 등)에서 가져온다. 새 수치 계산 함수를 지어내면 할루시네이션이 생기고 화면 표와 값이 어긋난다.

**R3. 프리뷰로 검증하기 전엔 "됐다"고 하지 않는다.**
브라우저에서 관찰 가능한 변경은: preview 서버 리로드 → `preview_console_logs`(에러 0 확인) → **그 탭 자체의 표/패널 숫자와 결과가 일치하는지 대조**까지 마친 뒤 보고한다. 스크린샷은 최종 증빙용, 검증은 텍스트 도구로.

**R4. 배포는 확인한다. 반영됐다고 가정하지 않는다.**
push 후 `gh run list --repo wckmkt/dashboard`로 Pages 빌드 성공을 확인한다. 빌드가 실패/멈추면 `gh run rerun`과 씨름하지 말고 **파이프라인을 한 번 더 돌려 새 데이터 커밋으로 배포를 강제 트리거**한다(검증된 복구법). "알림(Teams)은 오는데 대시보드가 안 바뀜" 증상은 데이터 커밋은 됐는데 **Pages 배포만 실패**한 경우일 수 있으니 이걸 먼저 의심한다.

**R5. 외부·비가역 액션은 승인 먼저.**
Teams 실제 발송, 공개 페이지에 시크릿 커밋, 프로덕션 재배포 등은 AskUserQuestion으로 확인한 뒤 실행한다.

**R6. 시크릿 반사신경.**
API 키를 절대 에코·저장·재사용하지 않는다. 사용자 스크린샷에 `sk-`/`Bearer`/긴 랜덤 토큰이 보이면 즉시 폐기·재발급을 안내한다(과거 반복 사고). 정적 사이트라 클라이언트 JS에 트리거 URL이 노출될 수밖에 없을 때는, 프록시 서버 대신 **Anthropic Console 지출 한도 설정**을 실질적 방어책으로 안내한다.

**R7. 퍼센트 산수.**
이미 퍼센트인 지표(CR, 취반품률, 달성률 등)의 YoY는 상대변화율 `(a-ly)/ly`가 아니라 **두 값의 차 `a-ly` (단위 %p)**다. 라벨과 수식이 일치하는지 항상 확인한다.

**R8. Windows / PowerShell 함정.**
1. PowerShell `ConvertFrom-Json`은 **대소문자만 다른 형제 키를 충돌로 처리**한다 → JSON 키(특히 채널ID)를 만들 때 원본 데이터의 표기 불일치(대소문자·공백)를 canonicalize한다.
2. `pwsh`(PowerShell 7) ≠ `powershell`(5.1) — 기능·구문 다름.
3. Excel COM 작업 전 마스터 파일이 **닫혀 있어야** 한다(`tasklist | grep EXCEL`로 확인).
4. native curl에 UTF-8(한글) 본문을 넘길 땐 셸 인코딩이 깨지니 **파일로 저장 후 `--data-binary @file`** 사용.

## 핵심 파일 지도
- `KPI/pmkt-kpi.html` — 대시보드 전체. 탭 시스템(`ALL_VIEWS`/`showView`), 지원이 팝업(FAB + `jw*` 함수들), Chart.js 렌더러.
- `KPI/pmkt_kpi_data.json` — `generated` + `RAW`(일→시간×[orders,revenue,net,dau,appdau,adcost]) + `DAILY` `DAILY_AD` `TRAFFIC` `CONV_CH` `DAILY_TGTS` `MONTHLY_TGTS` `REAL_MONTHLY` `CH_AD`(chid→[join,first,rev,sess]) `CH_AD_META` `BRAND_SALES`(날짜→브랜드→매출액(원), `6_raw_브랜드카테고리` 시트 원본 — 이름과 달리 카테고리 열은 없음, 브랜드 단위로만 운영).
- `C:\bang\gen_realtime_dashboard.py` — `read_*_sheet()` + `main()`. Excel COM으로 마스터 읽어 JSON 생성·push.
- `C:\bang\teams_notify.ps1` — Teams 알림(웹훅 URL, `#daily`/`#realtime` 앵커 분기).
- `C:\bang\pmkt_hourly_sync.ps1` — 시간별 동기화 오케스트레이션(BIZW/GA4 → Excel → JSON → push → Teams).
- `C:\bang\jiwoni-flow-spec.md` — 지원이 Power Automate Flow A/B 스펙(시스템 프롬프트 포함).
- `ad.html` / `ad-data.json` — 광고 성과 대시보드(아래 전용 섹션).
- `C:\bang\jiwon\scripts\build-ad-data.ps1` — 광고raw → ad-data.json 집계. `$COL` 인덱스 맵 주의.
- `C:\bang\jiwon\scripts\daily-ad-dashboard.ps1` — 09:30 오케스트레이션(빌드 → JSON 검증 → 실패 시 백업 복원 → push).
- `C:\bang\handover\인수인계.md` — 지원이 자동화 전반(Q&A봇·브리핑 3종·Airbridge 갭 보정). 광고 데이터가 어디서 오는지의 배경.

## 광고 성과 대시보드 (`ad.html`) — 같은 레포, 다른 파이프라인

KPI 대시보드와 **레포만 공유**할 뿐 데이터·생성 스크립트·담당이 전부 다르다. 위 R1~R8은 그대로 적용된다.

- **대시보드**: `ad.html` — 단일 파일 정적 SPA. 탭 4개(광고 성과 / 광고비용 편성 / DA상세 / SA상세). `https://wckmkt.github.io/dashboard/ad.html`
- **데이터**: `ad-data.json`(~12MB) — `C:\bang\jiwon\scripts\build-ad-data.ps1`이 대행사 광고raw 마스터파일을 Excel COM으로 읽어 생성. `daily-ad-dashboard.ps1`이 빌드→검증→push까지 수행하고, 작업스케줄러 **`지원이_광고대시보드`가 매일 09:30** 구동. 전체 실행 7~50분.
- **원천**: `(PMKT)2026년_광고raw_마스터파일_v2.0.xlsx` `로데이터` 시트(25+26 통합, 84컬럼, 약 50만행). OneDrive에서 `_adtmp`로 robocopy 후 읽는다. **암호화 컨테이너라 zipfile/openpyxl로는 못 연다 — Excel COM만 가능.**
- **스키마**: `meta`(generated/from/to/yoyFrom/yoyTo) + `dims` + `daily`(5개 차원별 일집계, YoY 밴드 포함) + `rows`(매체·캠페인·소재·기획전까지 granular, 올해 밴드만). 지표 11종 = `c ins su rv fo uv auv imp clk chRev asess`.
- **기간 밴드**: 최근 90일 + 전년 고정밴드(1/1~12/31). 매 실행마다 전체 재생성되므로 **지표를 추가해도 과거분이 자동으로 채워진다**(백필 불필요).

### 이 대시보드의 함정 (실제로 사고가 났던 지점)

1. **`$COL` 인덱스 맵은 컬럼 "순서"에 종속된다.** build-ad-data.ps1 상단의 `$COL` 해시는 마스터 시트의 컬럼 번호를 하드코딩한다. 대행사가 컬럼을 끼워넣으면 **에러 없이 엉뚱한 값**이 들어온다. 마스터 구조가 바뀌었다는 낌새가 있으면 `C:\bang\jiwon\scripts\_adtmp\header_dump.txt`(84컬럼 덤프)와 대조부터 한다. 없으면 같은 폴더의 `dump_headers.ps1`로 1행을 다시 덤프한다(작업용 폴더라 지워졌을 수 있음).
2. **채널ID 계열 지표(`uv` `auv` `asess` `chRev`)는 채널ID 단위로만 적재된다.** 소재명이 붙은 행에는 90일 내내 단 한 건도 없다(2026-08-13 확인). 그래서 소재 표에는 CP_UV·APP세션류를 넣지 않는다 — 넣으면 전량 0인 죽은 컬럼이 된다. 소재 단위로 볼 수 있는 건 MMP 계열(설치·가입·첫구매)뿐.
3. **매출·ROAS는 반드시 `rvOf()`/`roasOf()`를 거친다.** 상단 `MMP / 채널ID` 토글이 이걸로 동작한다. `o.rv`/`o.roas`를 직접 쓰면 토글이 조용히 무시된다(기획전 표·일별 추이에서 실제로 발생). **값을 고쳤으면 `revModeSeg` 핸들러의 재렌더 목록에도 그 함수를 추가**해야 한다 — 둘 중 하나만 고치면 증상이 그대로다.
4. **DA상세/SA상세 탭은 MMP 고정**이다(토글 UI가 tab1에만 있음). `state.revMode`는 전역이라 채널ID를 켜둔 채 넘어가면 기준이 조용히 바뀐다. 미해결 — 이 탭 수치를 논할 땐 MMP 기준임을 명시한다.
5. **기획전 미지정 행이 광고비의 35~41%**다. 프로모션별 실적 표는 이를 제외하므로 합계가 전체와 안 맞는다(카드 부제에 비중 자동 표기).
6. 소재행 기반 표(캠페인·소재·프로모션)는 `rows`를 쓰므로 **올해 90일만** 집계된다. YoY 비교가 필요하면 `daily`를 봐야 한다.
7. **마스터 갱신 직후 수동 실행하면 OneDrive 동기화 레이스에 걸린다.** 원본이 방금 저장됐으면 로컬엔 옛 버전만 내려와 있을 수 있고, 그대로 빌드하면 **에러 없이 전일 데이터가 통째로 빠진** JSON이 나오는 데다 검증 게이트(`daily>=100`)도 통과한다(2026-08-13 실사고). build-ad-data.ps1에 mtime 30초 안정화 대기를 넣어 막았고, `meta.srcModified`에 어느 마스터 버전으로 만들었는지 기록된다. **빌드 후에는 `meta.to`가 기대한 날짜인지 항상 확인할 것** — 로그의 `band1=` 줄과 orchestrator의 `검증:` 줄에 찍힌다.

## 커밋 / 배포 규율
- 데이터 커밋("data: … 갱신")은 파이프라인이 자동 생성한다.
- 코드 push가 거부되면 `git pull --rebase origin main`(무관한 ad-data 커밋과 경합) 후 재push.
- 커밋 메시지 끝: `Co-Authored-By: Claude <noreply@anthropic.com>`.
- `C:\bang` 파이프라인 레포는 명시적 요청 없이는 커밋하지 않는다.

## 반복 작업은 skill 사용
- 대시보드/파이프라인 수정 → **`dashboard-change`** skill(변경→프리뷰검증→커밋→배포검증 절차).
- 지원이에 새 질문 패턴/능력 추가 → **`jiwoni-capability`** skill(3단계 파이프라인·탭 스코프 규칙).
