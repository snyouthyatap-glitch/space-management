const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const {
  cleanNaverBookerName,
  resolveNaverBooker,
  extractUseDateTime,
  isChangedReservation,
} = require("./naver-booker");

// APPS_SCRIPT_READY.gs에서 함수 여러 개를 꺼내 Node에서 실행한다(Apps Script 전용 객체는 가짜로 둔다).
function loadAppsScriptFunctions(names, returnName) {
  const source = fs.readFileSync(path.join(__dirname, "..", "APPS_SCRIPT_READY.gs"), "utf8");
  const bodies = names.map((name) => {
    const match = source.match(new RegExp(`function ${name}\\([^)]*\\) \\{[\\s\\S]*?\\n\\}`));
    assert.ok(match, `${name} 함수를 찾지 못했습니다.`);
    return match[0];
  });
  const stubs = "const TZ = 'Asia/Seoul'; const Utilities = { formatDate: () => '2026-11-12 12:00:00' };";
  return new Function(`${stubs}\n${bodies.join("\n")}\nreturn ${returnName};`)();
}

// staff-dashboard/SelfTest.cs의 TestNaverBookerNameCleanup과 같은 사례다.
const NAME_CASES = [
  ["한예시", "한예시"],
  ["김예약 [대리예약] 방문자: 이방문", "이방문"],
  ["김예약[대리예약]방문자:이방문", "이방문"],
  ["김예약대리예약방문자 이방문", "이방문"],
  ["김예약대리예약방문자: 이방문", "이방문"],
  ["이방문(010-1234-5678)", "이방문"],
  ["김예약대리예약", "김예약"],
  ["최관리 [관리자등록]", "최관리"],
  ["최관리관리자등록", "최관리"],
  ["예약자: 홍길동", "홍길동"],
  ["방문자：김봄", "김봄"],
  ["홍길동!!", "홍길동"],
  ["김 봄 010-1234-5678", "김 봄"],
  ["John Smith", "John Smith"],
  ["[관리자등록]", ""],
  ["", ""],
  // 2026-11 네이버 정책 변경: 이름 가운데·연락처 앞자리를 가린다(가짜 이름).
  ["박*예", "박*예"],
  ["박*예***-****-1234", "박*예"],
  ["김*****-****-1234", "김**"],
];

test("cleanNaverBookerName", () => {
  for (const [input, expected] of NAME_CASES) {
    assert.equal(cleanNaverBookerName(input), expected, input);
  }
});

test("APPS_SCRIPT_READY.gs의 cleanNaverBookerName_도 같은 결과를 낸다", () => {
  const source = fs.readFileSync(path.join(__dirname, "..", "APPS_SCRIPT_READY.gs"), "utf8");
  const match = source.match(/function cleanNaverBookerName_\(value\) \{[\s\S]*?\n\}/);
  assert.ok(match, "cleanNaverBookerName_ 함수를 찾지 못했습니다.");
  const gsClean = new Function(`${match[0]}; return cleanNaverBookerName_;`)();
  for (const [input, expected] of NAME_CASES) {
    assert.equal(gsClean(input), expected, input);
  }
});

test("APPS_SCRIPT_READY.gs의 formatNaverPhone_는 번호를 가리지 않고 저장한다", () => {
  const format = loadAppsScriptFunctions(["str_", "digits_", "formatNaverPhone_"], "formatNaverPhone_");
  assert.equal(format("010-1234-5678"), "010-1234-5678");
  assert.equal(format("01012345678"), "010-1234-5678");
  assert.equal(format("0101235678"), "010-123-5678");
  assert.equal(format("010-****-5678"), "010-****-5678");
  assert.equal(format(""), "");
});

test("이용일시: 24시간제와 '26.11.11.수' 형태, '변경' 표시를 읽는다", () => {
  // 실제 화면(2026-11-12 사용자 제공 캡처)과 같은 행 구조. 연락처는 가짜.
  const changedRow = "완료정변경010-0000-12341400000496변경26.11.11.수17:30~19:30AI 커리어존 (청취Room1)1-0원26.11.8.일00:30";
  assert.equal(extractUseDateTime(changedRow, "1400000496"), "26.11.11.수17:30~19:30");
  assert.equal(isChangedReservation(changedRow, "1400000496"), true);

  // 2026-11-10 실제 수집본 형태: 날짜 없이 24시간제 시간만
  const todayRow = "확정김예약대리예약방문자: 이방문010-1111-2222010-3333-4444140000014113:30~15:30AI 커리어존 (청취Room1)1-0원";
  assert.equal(extractUseDateTime(todayRow, "1400000141"), "13:30~15:30");
  assert.equal(isChangedReservation(todayRow, "1400000141"), false);

  // 예전 형태
  const oldRow = "확정한예시010-0000-3615131865724526. 8. 14.(금) 오후 6:00~9:00청년 커넥트룸";
  assert.equal(extractUseDateTime(oldRow, "1318657245"), "26. 8. 14.(금) 오후 6:00~9:00");
});

test("서버 diffNaverReservationDetails_: 바뀐 값만 고친다", () => {
  const diff = loadAppsScriptFunctions(
    [
      "str_", "digits_", "pad2_", "normalizeTime_", "normalizeSheetTime_", "normalizeSheetDate_",
      "safePositiveInteger_", "normalizeNameKey_", "cleanNaverBookerName_", "formatNaverPhone_",
      "nowText_", "diffNaverReservationDetails_",
    ],
    "diffNaverReservationDetails_"
  );
  const current = {
    status: "확정", bookerName: "정 변경", maskedPhone: "010-9999-1234", usageDate: "2026-11-11",
    startTime: "15:30", endTime: "17:30", placeName: "AI 커리어존 (청취Room1)", spaceType: "careerRoom1",
    headcount: "1", purpose: "면접 준비", rawUsageText: "26.11.11.수 15:30~17:30",
  };
  const sameButFormatted = {
    bookerName: "정변경", maskedPhone: "010-9999-1234", usageDate: "2026-11-11", startTime: "15:30",
    endTime: "17:30", placeName: "커리어존 Room1", spaceType: "careerRoom1", headcount: 1, purpose: "",
    rawUsageText: "다른 표기",
  };
  assert.deepEqual(diff(current, sameButFormatted, "확정"), {}, "표기만 다르면 바꾸지 않는다");

  const moved = { ...sameButFormatted, startTime: "17:30", endTime: "19:30", spaceType: "careerRoom2", placeName: "커리어존 Room2" };
  assert.deepEqual(
    Object.keys(diff(current, moved, "확정")).sort(),
    ["endTime", "placeName", "rawUsageText", "spaceType", "startTime"]
  );

  const confirmed = diff({ ...current, status: "신청" }, sameButFormatted, "확정");
  assert.equal(confirmed.status, "확정");

  // 예전에 가려 저장한 번호는 전체 번호로 고치고, 다른 사람 번호가 들어갔으면 바로잡는다.
  assert.equal(diff({ ...current, maskedPhone: "010-****-1234" }, sameButFormatted, "확정").maskedPhone, "010-9999-1234");
  assert.equal(diff(current, { ...sameButFormatted, maskedPhone: "010-5555-7070" }, "확정").maskedPhone, "010-5555-7070");
  assert.deepEqual(diff(current, { ...sameButFormatted, maskedPhone: "010-****-1234" }, "확정"), {});
});

test("서버: 네이버가 가린 이름(박*예)·연락처(끝 4자리)로도 회원 예약을 찾는다", () => {
  const api = loadAppsScriptFunctions(
    ["str_", "digits_", "normalizeNameKey_", "isMaskedName_", "maskedNameMatches_", "reservationBelongsToMember_",
      "formatNaverPhone_", "cleanNaverBookerName_"],
    "{ reservationBelongsToMember_, maskedNameMatches_, formatNaverPhone_, cleanNaverBookerName_ }"
  );
  const member = { name: "박지예", phoneLastDigits: "1234" };
  assert.equal(api.reservationBelongsToMember_({ bookerName: "박*예", maskedPhone: "***-****-1234" }, member), true);
  assert.equal(api.reservationBelongsToMember_({ bookerName: "박*예", maskedPhone: "***-****-9999" }, member), false);
  assert.equal(api.reservationBelongsToMember_({ bookerName: "박*예", maskedPhone: "" }, member), false, "가린 이름은 끝자리 필수");
  assert.equal(api.reservationBelongsToMember_({ bookerName: "박지예", maskedPhone: "" }, member), true);
  assert.equal(api.reservationBelongsToMember_({ bookerName: "김*예", maskedPhone: "***-****-1234" }, member), false);
  assert.equal(api.maskedNameMatches_("박*예", "박예"), false);
  assert.equal(api.formatNaverPhone_("***-****-1234"), "***-****-1234");
  assert.equal(api.cleanNaverBookerName_("박*예***-****-1234"), "박*예");

  const { findPhones, resolveNaverBooker } = require("./naver-booker");
  assert.deepEqual(findPhones("확정박*예***-****-12341415031923"), ["***-****-1234"]);
  assert.deepEqual(resolveNaverBooker("박*예", ["***-****-1234"]), { bookerName: "박*예", phone: "***-****-1234", proxyBooking: false });
});

test("서버 시트 고치기: 가린 이름이 기존 실명과 맞으면 실명을 둔다", () => {
  const diff = loadAppsScriptFunctions(
    [
      "str_", "digits_", "pad2_", "normalizeTime_", "normalizeSheetTime_", "normalizeSheetDate_",
      "safePositiveInteger_", "normalizeNameKey_", "isMaskedName_", "maskedNameMatches_", "cleanNaverBookerName_",
      "formatNaverPhone_", "nowText_", "diffNaverReservationDetails_",
    ],
    "diffNaverReservationDetails_"
  );
  const current = {
    status: "확정", bookerName: "박지예", maskedPhone: "010-9999-1234", usageDate: "2026-11-26",
    startTime: "16:00", endTime: "18:00", placeName: "커리어존 Room2", spaceType: "careerRoom2", headcount: "1", purpose: "공부",
  };
  const incoming = {
    bookerName: "박*예", maskedPhone: "***-****-1234", usageDate: "2026-11-26", startTime: "16:00", endTime: "18:00",
    placeName: "커리어존 Room2", spaceType: "careerRoom2", headcount: 1, purpose: "공부",
  };
  const changes = diff(current, incoming, "확정");
  assert.equal(changes.bookerName, undefined);
  assert.equal(changes.maskedPhone, undefined);
  // 예전에 '*'를 지워 잘못 저장한 이름(박예)과 빈 연락처는 가린 값으로 고친다.
  const fixed = diff({ ...current, bookerName: "박예", maskedPhone: "" }, incoming, "확정");
  assert.equal(fixed.bookerName, "박*예");
  assert.equal(fixed.maskedPhone, "***-****-1234");
});

test("서버 직원 운영 목록: 취소·반려·노쇼 웹 예약도 보낸다", () => {
  const visible = loadAppsScriptFunctions(
    ["str_", "visibleReservationRequestStatus_", "staffCalendarReservationRequestStatus_"],
    "staffCalendarReservationRequestStatus_"
  );
  for (const status of ["신청", "승인", "일지제출완료", "취소", "반려", "노쇼"]) assert.equal(visible(status), true, status);
  assert.equal(visible(""), false);
  const source = fs.readFileSync(path.join(__dirname, "..", "APPS_SCRIPT_READY.gs"), "utf8");
  assert.match(source, /cancelled: isCancelledNaverStatus_\(row\.status\) && !isNoShowNaverStatus_\(row\.status\)/);
});

test("서버: 승인 안 한 신청은 시작 시간이 되면 자동 취소", () => {
  const { createVirtualAppsScript } = require("../demo/virtual-apps-script");
  const app = createVirtualAppsScript();
  const { context, constants } = app;
  const { SHEET_NAMES, HEADERS } = constants;
  context.ensureDataSheetSchemas_();
  const today = context.today_();
  const at = (minutes) => context.Utilities.formatDate(new Date(Date.now() + minutes * 60000), "Asia/Seoul", "HH:mm");
  const startedAt = at(-5);
  const laterAt = at(60);
  if (startedAt > laterAt) return; // 자정 근처에는 건너뛴다.
  const add = (requestId, status, startTime) =>
    context.appendObjectRow_(SHEET_NAMES.RESERVATION_REQUESTS, HEADERS.reservationRequests, {
      requestId, status, name: "가짜", date: today, startTime, endTime: "23:59", spaceType: "careerRoom1",
    });
  add("r-started", "신청", startedAt);
  add("r-later", "신청", laterAt);
  add("r-approved", "승인", startedAt);
  assert.equal(context.expireUnapprovedReservationRequests_(), 1);
  const rows = context.getRowsAsObjects_(SHEET_NAMES.RESERVATION_REQUESTS, HEADERS.reservationRequests);
  const status = Object.fromEntries(rows.map((row) => [row.requestId, row.status]));
  assert.deepEqual(status, { "r-started": "취소", "r-later": "신청", "r-approved": "승인" });
  assert.equal(rows.find((row) => row.requestId === "r-started").decidedBy, "system");
});

test("서버 실적 기간: 월별·기간 지정·최대 366일", () => {
  const api = loadAppsScriptFunctions(
    [
      "str_", "pad2_", "isValidMonthKey_", "isPerformanceDateKey_", "performanceDayCount_",
      "performanceDatesInRange_", "resolvePerformanceRange_",
    ],
    "{ resolvePerformanceRange_, performanceDatesInRange_ }"
  );
  const source = fs.readFileSync(path.join(__dirname, "..", "APPS_SCRIPT_READY.gs"), "utf8");
  assert.match(source, /const PERFORMANCE_MAX_RANGE_DAYS = 366;/);
  globalThis.PERFORMANCE_MAX_RANGE_DAYS = 366;

  assert.deepEqual(api.resolvePerformanceRange_({ month: "2026-02" }), {
    month: "2026-02", startDate: "2026-02-01", endDate: "2026-02-28",
  });
  assert.deepEqual(api.resolvePerformanceRange_({ startDate: "2026-09-29", endDate: "2026-10-02" }), {
    month: "", startDate: "2026-09-29", endDate: "2026-10-02",
  });
  assert.deepEqual(api.performanceDatesInRange_("2026-12-30", "2027-01-02"), [
    "2026-12-30", "2026-12-31", "2027-01-01", "2027-01-02",
  ]);
  assert.throws(() => api.resolvePerformanceRange_({ startDate: "2026-10-02", endDate: "2026-09-29" }), /늦습니다/);
  assert.throws(() => api.resolvePerformanceRange_({ startDate: "2025-01-01", endDate: "2026-01-02" }), /최대 366일/);
  assert.throws(() => api.resolvePerformanceRange_({ startDate: "2026-02-30", endDate: "2026-03-01" }), /YYYY-MM-DD/);
  delete globalThis.PERFORMANCE_MAX_RANGE_DAYS;
});

test("상세 화면: 목록 행이 섞여 대리예약으로 보여도 이 예약의 전화번호를 쓴다", () => {
  const { parseDetailFields } = require("./naver-booking-sync");
  // 2026-11-14: 목록 컨테이너가 대리예약 행부터 시작해 일반 예약 3건이 모두 대리예약 연락처로 기록됐다.
  const fields = [
    { label: "전화번호", value: "예약번호" },
    { label: "예약자", value: "이예시" },
    { label: "전화번호", value: "010-2222-3333" },
  ];
  const parsed = parseDetailFields(fields, { proxyBooking: true, phone: "010-9999-0000" });
  assert.equal(parsed.phone, "010-2222-3333");
  assert.equal(parsed.proxyBooking, undefined);

  const proxy = parseDetailFields(
    [
      { label: "예약자", value: "김예약대리예약" },
      { label: "방문자", value: "이방문(010-3333-4444)" },
      { label: "전화번호", value: "010-1111-2222" },
    ],
    {}
  );
  assert.equal(proxy.phone, "010-3333-4444");
  assert.equal(proxy.proxyBooking, true);
});

test("상세 화면: 모든 칸을 잘리거나 붙지 않게 읽는다", () => {
  const { parseDetailFields } = require("./naver-booking-sync");
  // 2026-11-17 실제 상세 화면과 같은 칸 구성(이름·번호는 가짜). 화면 위쪽 메뉴 칸도 섞여 들어온다.
  const parsed = parseDetailFields([
    { label: "날짜 선택오늘취소적용", value: "로딩중" },
    { label: "35건", value: "내려받기" },
    { label: "예약자", value: "이예시" },
    { label: "전화번호", value: "010-2222-3333" },
    { label: "예약번호", value: "1400000001" },
    { label: "상품", value: "AI 커리어존 (청취Room1)" },
    { label: "이용일시", value: "2026. 11. 17.(화) 오후 4:00~오후 6:00(2시간)" },
    { label: "이용 인원(명)", value: "3" },
    { label: "이용 목적", value: "스터디 모임" },
    { label: "확정", value: "2026. 11. 17.(화) 오전 11:58파트너센터 (운영자:abcd****)" },
    { label: "신청", value: "2026. 11. 15.(일) 오후 4:53" },
  ]);
  assert.equal(parsed.bookerName, "이예시");
  assert.equal(parsed.phone, "010-2222-3333");
  assert.equal(parsed.reservationNo, "1400000001");
  assert.equal(parsed.productName, "AI 커리어존 (청취Room1)");
  assert.equal(parsed.useDateTime, "2026. 11. 17.(화) 오후 4:00~오후 6:00(2시간)");
  assert.equal(parsed.headcount, 3);
  assert.equal(parsed.purpose, "스터디 모임");
  assert.equal(parsed.confirmedAt, "2026. 11. 17.(화) 오전 11:58");
  assert.equal(parsed.requestedAt, "2026. 11. 15.(일) 오후 4:53");
});

test("한 달 목록: 이용일을 읽고 오늘이용·오늘취소를 골라낸다", () => {
  const { rowUseDate, splitUpcomingRows, monthListUrl, parseDetailFields } = require("./naver-booking-sync");
  assert.equal(rowUseDate({ useDateTime: "26.11.18.수 13:00~16:00" }), "2026-11-18");
  assert.equal(rowUseDate({ useDateTime: "2026. 11. 7.(토) 오전 10:00~오후 12:00(2시간)" }), "2026-11-07");
  assert.equal(rowUseDate({ useDateTime: "16:00~18:00" }), "");
  const rows = [
    { reservationNo: "1", status: "확정", useDateTime: "2026. 11. 17.(화) 오후 4:00~오후 6:00" },
    { reservationNo: "2", status: "취소", useDateTime: "26.11.17.화 10:00~12:00" },
    { reservationNo: "3", status: "확정", useDateTime: "26.11.18.수 10:00~12:00" },
    { reservationNo: "4", status: "신청", useDateTime: "26.11.17.화 13:00~14:00" },
  ];
  const split = splitUpcomingRows(rows, "2026-11-17");
  assert.deepEqual(split.today_use.map((row) => row.reservationNo), ["1"]);
  assert.deepEqual(split.today_cancel.map((row) => row.reservationNo), ["2"]);
  assert.match(monthListUrl("https://example.test/list", "2026-11-17", "2026-12-16"),
    /dateDropdownType=MONTH&startDateTime=2026-11-17&endDateTime=2026-12-16&dateFilter=USEDATE/);
  // 바뀐 예약은 상세 예약번호 옆에 '변경'이 붙어 온다(2026-11-17 실제 화면).
  assert.equal(parseDetailFields([{ label: "예약번호", value: "1400000001변경" }]).reservationNo, "1400000001");
});

test("대리예약은 방문자 이름과 아래쪽 연락처를 쓴다", () => {
  assert.deepEqual(
    resolveNaverBooker("김예약 [대리예약] 방문자: 이방문", ["010-1111-2222", "010-3333-4444"]),
    { bookerName: "이방문", phone: "010-3333-4444", proxyBooking: true }
  );
});

test("일반·관리자등록 예약은 첫 연락처를 쓴다", () => {
  assert.deepEqual(
    resolveNaverBooker("최관리 [관리자등록]", ["010-5555-6666"]),
    { bookerName: "최관리", phone: "010-5555-6666", proxyBooking: false }
  );
  assert.deepEqual(
    resolveNaverBooker("한예시", ["010-0000-3615"]),
    { bookerName: "한예시", phone: "010-0000-3615", proxyBooking: false }
  );
});
