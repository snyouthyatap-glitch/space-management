# 인수인계: 청년이봄야탑 공간관리 개선 작업 (2026-11-12)

이 문서는 claude.ai 대화에서 정리한 개선 작업을 Claude Code에서 이어서 하기 위한 문서다.
작업자는 사용자가 "시작"이라고 말하기 전에는 코드를 수정하지 않는다. 시작 전에는 목록 정리와 의견 제시만 한다.

## 0. 먼저 할 일

1. 저장소 `snyouthyatap-glitch/space-management`의 `main` 브랜치(기준 커밋 `48eb39f`)에서 새 브랜치를 만든다.
2. 같이 전달된 `space-management-wip.patch`를 적용한다. 항목 1과 2가 들어 있는 미완성 작업이다.
   ```bash
   git checkout -b ops-improvements 48eb39f
   git apply --check space-management-wip.patch && git apply space-management-wip.patch
   ```
3. 직원 프로그램을 빌드하고 셀프 테스트를 돌린다. 이전 환경에는 .NET이 없어서 빌드를 한 번도 못 했다.
   ```powershell
   cd staff-dashboard
   dotnet build
   dotnet run -- --self-test
   ```
4. 이 문서도 `docs/`에 같이 커밋한다. 그 전에 이 문서에 비밀값이 없는지 확인한다. 저장소는 공개 상태다.

## 1. 시스템 개요

- **이용자 웹** (`index.html`, `js/main.js`): GitHub Pages로 무료 서비스 중이다. 그래서 저장소를 비공개로 바꿀 수 없다.
- **관리자 웹** (`admin.html`, `js/admin.js`): 비밀번호로 로그인한다.
- **서버** (`APPS_SCRIPT_READY.gs`): 구글 Apps Script 웹앱이고, 데이터는 구글 시트에 있다. **GitHub에 올려도 자동 배포되지 않는다.** 사용자가 Apps Script 편집기에 코드를 붙여넣고 "배포 → 배포 관리 → 새 버전"으로 다시 배포해야 한다.
- **직원 알림 프로그램** (`staff-dashboard/`): C# WinForms(.NET, Windows) 트레이 프로그램이다. `install-dashboard.ps1`로 설치한다.
- **네이버 예약 동기화** (`automation/naver-booking-sync.js`): Playwright로 스마트플레이스 관리자 화면을 읽는다. 결과는 `outputs/naver-booking-sync.json`에 저장되고, 직원 프로그램이 이 파일을 읽는다.

### 현재 운영 흐름 (사용자 설명)

1. 이용자가 공간에서 QR을 찍어 웹에 접속한다. 공간 밖에서의 접속을 막으려고 주소 뒤에 `?entry=토큰`을 붙였다.
2. 전화번호 뒷 4자리로 로그인하면 자동으로 출석(방문일지)이 기록된다.
3. 이용하거나 예약한 설비의 일지를 제출한다. 대상은 커리어존, 커넥트룸, 프린터다.

## 2. 작업 목록

진행 상태는 다음 셋 중 하나다.
- ✅ 코드 작성 (패치에 포함)
- 🟡 일부 작성
- ⬜ 미착수

### 1. ✅ QR 출입 토큰 보안 수정

**문제**
- 토큰 `ybom-qr-2026`이 공개 저장소의 두 곳에 그대로 적혀 있었다. `APPS_SCRIPT_READY.gs`의 `QR_ENTRY_TOKEN`과 `js/main.js`의 `LOCAL_PREVIEW_ENTRY_TOKEN`이다.
- 그래서 외부에서 아래 일이 가능했다.
  - 뒷 4자리를 0000부터 9999까지 대입해 회원 정보(성별, 나이, 회원 ID)를 수집
  - 가짜 가입이나 가짜 일지 제출

**패치 내용**
- 토큰을 Script Properties의 `QR_ENTRY_TOKEN`에서 읽는다(`getQrEntryToken_`).
- 시트 메뉴에 `setupQrEntryToken`(「현장 QR 토큰 새로 발급」)을 추가했다. 무작위 토큰을 발급하고 QR로 만들 주소를 보여준다. 주소는 `QR_ENTRY_PAGE_URL` + `?entry=` + 토큰이다.
- `assertPhoneTailLookupAllowed_`: 뒷 4자리 조회를 5분에 60회까지로 제한했다. 전역 카운터이고 CacheService와 LockService를 쓴다.
- `js/main.js`에서 하드코딩된 토큰을 제거했다.

**남은 일**
- 사용자가 직접 배포해야 한다. 순서는 아래와 같고, 한 번에 이어서 해야 한다.
  1. Apps Script에 붙여넣고 새 버전으로 배포
  2. 메뉴에서 토큰 발급
  3. 새 QR 출력 후 교체
- 배포만 하고 토큰을 발급하지 않으면 QR 접속이 전부 막힌다.
- 옛 토큰은 git 기록에 남아 있으므로 반드시 재발급해야 한다.
- 남은 한계가 있다. QR 주소를 저장하거나 공유하면 밖에서도 접속할 수 있다. 접속 후 주소창의 토큰은 지워진다. 필요하면 매달 재발급하라고 안내한다.

### 2. ✅ 커리어존·커넥트룸 일지: 직원 대리 제출 간편화

**배경**
- 이용자가 일지를 안 내는 일이 많다. 그래서 직원이 PC에서 그 회원 계정으로 직접 로그인해 대신 제출하고 있었다.
- 직원 프로그램에도 "이용완료로 변경"(`adminCompleteReservationUsage`) 기능이 있다. 하지만 2인 이상 예약이면 `companions`를 빈 배열로 보내서 서버 검증에 실패한다. 이것이 원인으로 추정된다.

**패치 내용**

서버
- `getStaffMonthCheckIns_`를 추가했다. 그 달 방문일지를 `{date, name, memberId}` 형태로 중복 없이 모아 30초 동안 캐시한다.
- `handleAdminDashboardSnapshot_`가 이 결과를 `checkIns`로 반환한다.

직원 프로그램

| 파일 | 변경 내용 |
|---|---|
| `Models.cs` | `CalendarEntry.CheckedIn`(JsonIgnore), `StaffCheckIn`, `DashboardSnapshot.CheckIns` 추가 |
| `ReservationService.cs` | `ApplyCheckIns` 추가: 이름+날짜로 출석을 매칭하고, 회원 ID가 하나뿐이면 `MemberId`를 채운다. 시작 후 10분이 지났는데 일지가 없을 때, 출석했으면 `"일지 미제출"`, 아니면 `"노쇼"`로 판정한다. `MergeDuplicateReservationEntries`가 `MemberId`를 버리던 버그를 고쳤다. |
| `MainForm.cs` | 상단 「일지 미제출 N」 버튼(누르면 해당 날짜와 상태 필터로 이동), 예약 아래 「일지 대리 제출」 버튼, 상태 필터·색상 추가. 2인 이상 예약을 대리 제출할 때 `ShowCompanionDialog`로 동반자 성별(남성/여성)과 나이를 입력받는다. 웹 신청 예약은 서버에 저장된 동반자 정보를 쓰므로 입력을 생략한다. |
| `AppsScriptClient.cs` | `CompleteReservationUsageAsync`가 companions를 인자로 받는다. |
| `SelfTest.cs` | `TestMissingUsageLogStatus` 추가 |

**남은 일**
- 빌드와 셀프 테스트(컴파일을 아직 한 번도 확인하지 못했다)
- 실제 데이터로 동작 확인
- 네이버 예약자명이 지저분하면 이름 매칭이 실패한다. 그래서 항목 6과 같이 확인해야 한다.

### 3. ✅ 커리어존·커넥트룸 일지: 이용자 제출 유도

- 관련 코드: `js/main.js`의 `renderTodayReservations` 근처(약 919~1100행), 예약 카드, `RESERVATION_SUBMIT_UNLOCK_MINUTES = 20`
- 추정 원인:
  - 이용 시작 20분 전보다 일찍 도착해서 출석하면 제출 버튼이 잠겨 있다.
  - 자동 출석만 되면 할 일이 끝났다고 생각하고 화면을 닫는다.
- 제안:
  - 출석 직후 화면 위쪽에 "오늘 예약이 있어요 · 일지를 제출해야 이용이 기록돼요" 배너를 띄운다.
  - 제출 가능한 예약이 1건이면 자동으로 선택해서 바로 제출할 수 있게 한다.
  - 시간이 일러 잠겨 있으면 "HH:MM부터 제출 가능 · 그때 QR을 다시 찍어 주세요"라고 시각을 명시한다.
- 20분 규칙을 완화할지는 운영 정책이므로 사용자에게 먼저 묻는다.

### 4. ✅ 라운지 이용 안내 개선

- 증상: 라운지만 쓰는 이용자가 로그인 후 '라운지 이용' 버튼을 찾는다. 실제로는 출석만 하면 이용이 끝난다.
- 관련 코드: `index.html`의 `#facilitySection`(약 188~205행)에 있는 빠른 버튼(공간 예약, 프린터 이용)과 `.status-strip`의 "오늘 방문이 기록되었습니다."
- 제안(사용자에게 전달함):
  1. 로그인 직후 맨 위에 크게 "✓ 라운지 이용 체크 완료 · 따로 기록할 필요 없이 바로 이용하세요"를 띄운다.
  2. 빠른 버튼에 '라운지 이용'을 추가하고, 누르면 같은 안내가 나오게 한다.

### 5. ✅ 직원 프로그램에서 '내 일정'·'수동 일지' 제거

- 사용자가 거의 쓰지 않는다고 했다.
- `MainForm.cs`
  - 「수동 일지」 액션 버튼(약 228행, `ManualLogForm` 호출)
  - 날짜별 '내 일정' 패널: `_dayChecklist`, `_dayScheduleInput`, `_addScheduleButton`, `_deleteScheduleButton`, `_dayScheduleTitle`, `CalendarNote` 관련 UI. 약 600~650행.
- 함께 검토할 것:
  - `ManualLogForm.cs`, `AppsScriptClient`의 수동 일지 API, `SelfTest`의 `TestManualLogFlowAsync`
  - 달력 미리보기에 일정(notes)을 표시하는 부분(`CalendarEntryPresenter`)
  - `LocalState.CalendarNotes` 저장 위치
- 관리자 웹(`admin.html`)의 수동 일지는 이번 범위가 아니다. 지울지는 사용자에게 확인한다.
- 항목 2의 대리 제출은 수동 일지가 아니라 `adminCompleteReservationUsage`를 쓰므로 영향이 없다.

### 6. ✅ 네이버 예약 예약자명·연락처 정리

네이버 예약 목록의 형태는 세 가지다.

| 형태 | 화면 | 등록할 이름 | 등록할 연락처 |
|---|---|---|---|
| 일반 예약 | 예약자명 + 연락처 | 예약자명 | 연락처 |
| 대리예약 | `김예약 [대리예약]`, 아래 줄 `방문자: 이방문`, 연락처 두 줄 | **아래쪽 방문자 이름** | **아래쪽 연락처** |
| 관리자등록 | `최관리 [관리자등록]` | 앞의 예약자명 | 연락처 |

**문제**
- 괄호 같은 특수문자가 섞인다.
- '관리자등록'이나 '대리예약' 같은 이름이 아닌 단어가 예약자명에 들어간다.

**관련 코드**
- `automation/naver-booking-sync.js`(약 250~290행)
  - `bookerName` 추출이 상태어 뒤, 첫 전화번호 앞의 텍스트를 이름으로 쓴다. 그래서 배지 텍스트가 섞인다.
  - `phone`은 첫 번째 번호, `altPhone`은 두 번째 번호다. 대리예약이면 `altPhone`(방문자 연락처)을 써야 한다.
  - 상세 페이지 파싱(약 321행, `detailFieldValue(/^예약자$/)`)도 있다.
- `staff-dashboard/NaverBookingParser.cs`(89, 134, 150행): 예약자명 정리
- `APPS_SCRIPT_READY.gs`의 네이버 CSV 동기화: `mapNaverReservationRow_`(약 3758행). 같은 규칙을 적용할지 확인한다.

**구현 방향**
- 배지 단어(`대리예약`, `관리자등록` 등)와 괄호·특수문자를 제거한다.
- 대리예약이면 방문자 이름과 아래쪽 번호를 쓴다.
- 이름 정리 규칙은 한 함수로 모으고 테스트를 추가한다.

### 7. ✅ 뒷 4자리가 같은 기존 회원이 있을 때 신규 가입 경로

- 증상: 신규 이용자의 뒷 4자리가 기존 회원과 같으면 "○○님이 맞나요?" 확인창이 뜬다. 여기서 "제가 아니에요"를 누르면 처음 화면으로만 돌아가서 가입할 방법이 없다.
- 관련 코드
  - `index.html`의 `#memberConfirmSection`(약 79~89행)
  - `js/main.js`의 `memberRejectBtn` 핸들러(약 1941행, `resetToChoice()`)
  - 서버의 뒷 4자리 조회는 `APPS_SCRIPT_READY.gs` 약 1426행
- 요구: "제가 아니에요"를 누르면 **회원가입**이나 **메인 화면** 중 하나를 고르게 한다. 회원가입을 고르면 전체 번호로 가입하는 흐름으로 이어진다.
- 확인할 것: 서버 가입 처리가 뒷자리 중복이 아니라 전체 번호 중복만 막는지 확인한다.

### 8. ⏸ 보류 · 프린트 PC 잠금 프로그램 (출력 전 인증, 하루 10장 제한, 자동 일지)

- 배경:
  - 이용자가 QR을 찍지 않고 바로 프린터로 가서 프린터 일지가 빠진다.
  - 1인 하루 10장 제한은 웹 일지에만 적용돼 있다(서버 `getPrinterCountForDate_`, 약 2972행).
- 사용자가 이 방향에 동의했다("오 괜찮다").
- 복사·팩스 같은 프린터 자체 기능은 제한할 필요가 없다.

**동작**
1. 평소에는 프린터 대기열을 일시 정지해 둔다.
2. 출력 작업이 들어오면 PC 화면에 "전화번호 뒷 4자리" 입력 창을 띄운다.
3. Apps Script에서 회원 여부와 오늘 남은 장수를 확인한다.
4. 남은 장수 안이면 그 작업만 출력하고, 넘으면 안내한 뒤 작업을 취소한다.
5. 출력이 끝나면 다시 잠그고, 실제 장수를 프린터 일지에 자동으로 기록한다.

**환경 (사용자 확인 완료)**

| 항목 | 내용 |
|---|---|
| 프린터 | Brother MFC-L8900CDW |
| 연결 | 프린트 전용 PC와 USB(`USB001`) |
| OS | Windows 10 Pro 2009 |
| 인터넷 | script.google.com:443 접속 가능 |
| 대기열 | 같은 포트에 `Brother MFC-L8900CDW`와 `Brother MFC-L8900CDW series` 두 개가 있다. **둘 다 잠가야 한다.** `Brother PC-FAX v.3.2`도 같은 포트에 있다(출력 경로인지 확인 필요). |
| 가상 프린터 | PDF, XPS, OneNote, Hancom PDF는 종이가 나오지 않으므로 대상이 아니다. |
| 윈도우 계정 | 현재 이용자가 **관리자 계정**으로 쓰고 있다. |

**구현 메모**
- 이용자용 일반 사용자 계정을 새로 만들고, 관리자 계정은 직원 전용으로 둔다. 계정 생성 방법을 사용자에게 안내한다. 그렇지 않으면 이용자가 프로그램을 끄거나 대기열을 직접 풀 수 있다.
- 대기열 제어는 관리자 권한이 필요하다. 그래서 구조를 둘로 나눈다.
  - 윈도우 서비스: 대기열 감시, 일시 정지와 재개
  - 사용자 세션 UI: 번호 입력
- 장수 계산: 스풀러 작업의 TotalPages에 부수(copies)를 곱한다. 양면이나 N-up이면 실제 종이 장수와 다를 수 있어서 실기기 확인이 필요하다.
- 서버 API를 새로 만든다.
  - 남은 장수 조회
  - 출력 기록 저장(프린터 일지 시트, 출처 표시)
  - QR 토큰과는 별도로 프린트 PC 전용 키를 Script Properties에 둔다.
- 프린터가 와이파이에도 연결돼 있으면 PC를 거치지 않고 출력할 수 있다. 프린터에서 WLAN과 Wi-Fi Direct가 꺼져 있는지 사용자에게 확인을 요청해 둔 상태다(답변 전).
- 참고: Windows 10은 2025년 10월에 일반 보안 지원이 끝났다.

## 3. 배포·운영 주의

- Apps Script를 바꾸면 사용자가 직접 붙여넣고 새 버전으로 배포해야 한다. 웹앱 주소는 유지된다. 변경할 때마다 배포 절차를 짧게 안내한다.
- 프론트(`js/`, `*.html`) 변경은 `main`에 반영되면 GitHub Pages에 바로 나간다. 그래서 서버와 프론트의 호환 순서를 신경 쓴다.
- 직원 프로그램은 빌드한 뒤 직원 PC에 재설치해야 한다(`install-dashboard.ps1`, `설치하기.cmd`).
- 커밋과 푸시는 사용자가 요청할 때만 한다.
- 비밀값(토큰, 비밀번호, 프린트 PC 키)은 저장소에 절대 넣지 않는다.

## 4. 진행 결과 (2026-11-12, 브랜치 `ops-improvements`)

- 항목 1~7 코드 작성 완료. 항목 8은 보류.
- 확인한 것: 직원 프로그램 빌드(경고·오류 0), `dotnet run -- --self-test` 통과, `npm test`(예약자명 규칙, Apps Script 사본 일치 포함) 통과, 이용자 웹은 로컬 데모(`?demo=reservation`, 0000)로 화면 확인.
- 항목 3 운영 정책: "이용 20분 전부터"를 없애고 **예약 당일이면 언제든** 일지를 제출할 수 있게 했다(웹과 서버 모두).
- 항목 5: 직원 프로그램의 수동 일지 창과 '내 일정' 패널을 지웠다. 서버 `adminSubmitManualLog`와 관리자 웹의 수동 일지는 남겨 두었다. 예전에 저장한 일정은 상태 파일에 그대로 보존된다.
- 항목 6: 정리 규칙은 `automation/naver-booker.js`, `NaverBookingParser.NormalizeBookerName`, `cleanNaverBookerName_` 세 곳에 같은 내용으로 있다. 2026-11-10 실제 수집본의 대리예약 1건으로 확인했다. 관리자등록 실제 사례는 아직 확인하지 못했다.

### 배포 순서 (한 번에 이어서)

1. Apps Script에 `APPS_SCRIPT_READY.gs`를 붙여넣고 새 버전으로 배포한다.
2. 시트 메뉴 「현장 QR 토큰 새로 발급」을 실행하고, 새 QR을 출력해 교체한다.
3. 그다음 `main`에 병합해 이용자 웹(GitHub Pages)을 내보낸다. 서버보다 먼저 나가면 이른 시간 일지 제출이 옛 서버에서 거절된다.
4. 직원 프로그램을 빌드해 직원 PC에 다시 설치한다.
### 추가 작업 (2026-11-12 오후)

- 예전 작업 폴더(`Documents\청년이봄홈페이지\space-management`)의 미반영 작업 중 **네이버 로그인 관련만** 옮겼다: 로그인 버튼·상품 선택 수정, 설정 창 '네이버 로그인 설정'·'오류 로그 보기'. 오류 로그는 `%LOCALAPPDATA%\YouthSpringStaffDashboard\naver-sync.log`에 남는다.
- 네이버 예약 시트 맞추기(`adminEnsureNaverReservation`): 직원 프로그램이 5분마다 오늘 확정 예약을 시트에 없으면 추가하고, 오늘 취소된 예약은 시트에서 '취소'로 바꾼다. 일지가 제출됐거나 직원이 이미 처리한 예약은 바꾸지 않는다. 서버 배포가 필요하다.
- `SPACE_MANAGEMENT_REPO` 환경변수를 이 작업 폴더로 바꾸고, 수집 결과(`outputs/`)와 네이버 로그인 프로필(`automation/.naver-booking-state/`)을 옮겼다. `npm install`로 Playwright를 설치했다.
- 예전 폴더에 남겨 둔 것: '내 메모', 기타 화면 수정, task-management 앱, 포스터·설명서.
### 관리자 웹 삭제 (2026-11-12)

- 운영에서 쓰지 않아 `admin.html`, `js/admin.js`, `css/admin.css`를 지웠다. 직원 프로그램 설정 메뉴의 '관리자 페이지'도 없앴다.
- 서버에서 관리자 웹만 쓰던 기능을 지웠다: `adminSubmitManualLog`(수동 일지), `adminSessionStatus`, 그리고 그것만 쓰던 보조 함수(`getAdminManualSubject_`, `refreshArchivedPublicSheetsForAdminDate_`, `isValidLogDate_`).
- 관리자 비밀번호(`setupAdminAccessKey`)와 로그인(`adminChallenge`, `adminLogin`)은 직원 프로그램이 쓰므로 남겼다. 시트 메뉴 이름은 '관리자 비밀번호 설정'으로 바꿨다.