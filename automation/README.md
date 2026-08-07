# Booking Automation

네이버 예약 `예약자관리` 화면에서 아래 3개 빠른 필터를 순회해 데이터를 수집하는 스크립트입니다.

- `오늘이용`
- `확정대기`
- `오늘취소`

상품 필터는 먼저 아래 3개만 선택되도록 맞춥니다.

- `AI 커리어존 (청취Room1)`
- `AI 커리어존 (청취Room2)`
- `청년 커넥트룸`

## 실행

처음 한 번은 로그인 세션을 저장합니다.

```powershell
& 'C:\Users\SNYOUTH\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe' .\automation\naver-booking-sync.js --setup-login
```

로그인 후 콘솔에서 Enter를 누르면 프로필이 저장됩니다.

이후에는 일반 실행만 하면 됩니다.

```powershell
& 'C:\Users\SNYOUTH\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe' .\automation\naver-booking-sync.js
```

## 결과 파일

기본 저장 위치:

`outputs/naver-booking-sync.json`
