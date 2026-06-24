## 네이버 예약 CSV 동기화 안내

### 개요

이 기능은 네이버 스마트플레이스 예약 CSV를 구글드라이브 폴더에 올리면,
예약번호를 기준으로 기존 예약 시트와 비교해서 아래 작업을 수행하도록 설계되었습니다.

- 신규 예약 추가
- 기존 예약 수정
- 취소 예약 반영
- 구글 캘린더 일정 생성 / 수정

### 기준 키

비교 기준은 `예약번호`입니다.

즉:

- 같은 예약번호가 없으면 신규 예약
- 같은 예약번호가 있으면 수정 여부 확인
- 상태가 `취소`면 취소 상태로 반영

### Apps Script에서 추가된 함수

- `syncNaverReservationCsv()`
  - 수동 실행용
- `syncNaverReservationCsvInternal_(trashProcessedFiles)`
  - 내부 실행용

### 드라이브 폴더 이름

아래 이름의 구글드라이브 폴더가 필요합니다.

- `네이버예약CSV`

운영 방식:

1. 네이버 스마트플레이스에서 CSV 다운로드
2. 구글드라이브 `네이버예약CSV` 폴더에 업로드
3. Apps Script에서 `syncNaverReservationCsv()` 실행

필요하면 시간 기반 트리거로 자동 실행도 가능합니다.

### 생성 / 사용하는 시트

아래 시트가 자동 생성되거나 사용됩니다.

- `네이버예약_동기화`

### 시트 주요 컬럼

- `reservationId`
- `status`
- `bookerName`
- `maskedPhone`
- `usageDate`
- `startTime`
- `endTime`
- `placeName`
- `normalizedPlace`
- `headcount`
- `purpose`
- `calendarEventId`
- `lastSyncedAt`
- `sourceFileName`

### 캘린더 처리 방식

예약 1건당 캘린더 이벤트 1건을 생성하거나 수정합니다.

이벤트 제목 예시:

- `[예약] Joy 1 / 홍길동`
- `[취소] Joy 1 / 홍길동`

설명란에는 아래 정보가 들어갑니다.

- 예약번호
- 상태
- 예약자
- 전화번호(마스킹)
- 이용일시
- 인원
- 목적
- 취소사유

### CSV 구조 주의사항

네이버 CSV는 맨 위 2줄이 안내문이고,
실제 헤더는 3번째 줄에 있습니다.

스크립트에서는 이 구조를 고려해서 처리합니다.

### 현재 장소명 정규화 방식

아래 문자열을 제거해 비교에 사용합니다.

- `[동아리실] `
- `[Club room] `
- `3층 `
- `3rd floor `

예:

- `[동아리실] 3층 Joy 1` -> `Joy 1`
- `[Club room] 3rd floor Joy 1` -> `Joy 1`

### 취소 처리 방식

`상태`가 `취소`이면:

- 시트 상태를 취소로 반영
- 캘린더 제목 앞에 `[취소]` 표시
- 캘린더 색상은 회색으로 변경

### 권장 다음 단계

1. Apps Script에 이 코드 반영
2. 구글드라이브에 `네이버예약CSV` 폴더 생성
3. 테스트용 CSV 업로드
4. `syncNaverReservationCsv()` 수동 실행
5. 시트와 캘린더 반영 결과 확인
6. 문제 없으면 시간 기반 트리거 추가
