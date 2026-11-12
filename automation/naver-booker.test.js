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

test("APPS_SCRIPT_READY.gs의 maskNaverPhone_는 가운데 자리를 가린다", () => {
  const source = fs.readFileSync(path.join(__dirname, "..", "APPS_SCRIPT_READY.gs"), "utf8");
  const pick = (name) => source.match(new RegExp(`function ${name}\\([^)]*\\) \\{[\\s\\S]*?\\n\\}`))[0];
  const mask = new Function(`${pick("str_")}; ${pick("digits_")}; ${pick("maskNaverPhone_")}; return maskNaverPhone_;`)();
  assert.equal(mask("010-1234-5678"), "010-****-5678");
  assert.equal(mask("01012345678"), "010-****-5678");
  assert.equal(mask("010-****-5678"), "010-****-5678");
  assert.equal(mask(""), "");
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
      "safePositiveInteger_", "normalizeNameKey_", "cleanNaverBookerName_", "maskNaverPhone_",
      "nowText_", "diffNaverReservationDetails_",
    ],
    "diffNaverReservationDetails_"
  );
  const current = {
    status: "확정", bookerName: "정 변경", maskedPhone: "010-****-1234", usageDate: "2026-11-11",
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
