// 가상 시트 데모 서버.
//   node demo/demo-server.js            → 가짜 데이터로 새로 시작
//   node demo/demo-server.js --keep     → 지난 데모 시트 이어서 사용
// 직원 프로그램:  "청년이봄 직원알림.exe" --virtual-sheet=http://127.0.0.1:8787/exec
// 이용자 QR 화면: http://demo.localhost:8787/?entry=demo-entry
// 가상 시트 보기: http://127.0.0.1:8787/sheet
// 모든 이름·연락처는 가짜다. 실제 구글 시트·네이버에는 아무것도 보내지 않는다.
const fs = require("node:fs");
const http = require("node:http");
const path = require("node:path");
const { createVirtualAppsScript } = require("./virtual-apps-script");

const PORT = Number(process.env.DEMO_PORT || 8787);
const ROOT = path.join(__dirname, "..");
const DATA_DIR = path.join(ROOT, "outputs", "demo");
const SHEET_FILE = path.join(DATA_DIR, "virtual-sheet.json");
const NAVER_FILE = path.join(DATA_DIR, "naver-booking-sync.json");
// 데모 전용 값(가상 서버에서만 통한다). 직원 프로그램 --virtual-sheet 모드가 같은 값을 쓴다.
const DEMO_ADMIN_PASSWORD = "virtual-demo-1234";
const DEMO_ENTRY_TOKEN = "demo-entry";

function todayText(offsetDays = 0) {
  const date = new Date(Date.now() + offsetDays * 86400000);
  return new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Seoul" }).format(date);
}

function koreanUseDateTime(date, start, end) {
  const [y, m, d] = date.split("-").map(Number);
  const day = "일월화수목금토"[new Date(Date.UTC(y, m - 1, d)).getUTCDay()];
  const hm = (value) => {
    const [h, min] = value.split(":").map(Number);
    return `${h < 12 ? "오전" : "오후"} ${h % 12 === 0 ? 12 : h % 12}:${String(min).padStart(2, "0")}`;
  };
  return `${y}. ${m}. ${d}.(${day}) ${hm(start)}~${hm(end)}`;
}

function seed(app) {
  const { context, constants } = app;
  const { SHEET_NAMES, HEADERS } = constants;
  context.ensureDataSheetSchemas_();
  app.properties[constants.ADMIN_ACCESS_HASH_PROPERTY] = context.sha256Hex_(DEMO_ADMIN_PASSWORD);
  app.properties[constants.QR_ENTRY_TOKEN_PROPERTY] = DEMO_ENTRY_TOKEN;

  const now = context.nowText_();
  const members = [
    ["m-kim", "김하늘", "여성", "1998-03-02", "01000001111"],
    ["m-lee", "이바다", "남성", "1995-07-15", "01000002222"],
    ["m-park", "박서준", "남성", "2000-11-30", "01000003333"],
    ["m-choi", "최노쇼", "여성", "1993-01-20", "01000004444"],
    ["m-jung", "정취소", "남성", "1997-05-05", "01000005555"],
    ["m-han", "한반려", "여성", "2001-09-09", "01000006666"],
  ];
  for (const [memberId, name, gender, birthDate, phone] of members) {
    context.appendObjectRow_(SHEET_NAMES.MEMBERS, HEADERS.members, {
      memberId, name, gender, birthDate, age: "", phone, phoneLastDigits: phone.slice(-4),
      isSeongnamResident: "관내", privacyConsent: true, optionalConsent: false, consentAt: now,
      status: "approved", role: "user", createdAt: now, updatedAt: now, lastUsedAt: "",
    });
  }

  const today = todayText();
  const naverRows = [
    ["1400000001", "확정", "김하늘", "010-0000-1111", "10:00", "12:00", "careerRoom1", "커리어존 Room1", 1, "면접 준비"],
    ["1400000002", "노쇼", "최노쇼", "010-0000-4444", "09:00", "10:00", "connectRoom", "커넥트룸", 2, "스터디"],
    ["1400000003", "취소", "정취소", "010-0000-5555", "13:00", "15:00", "careerRoom2", "커리어존 Room2", 1, "자기소개서"],
    ["1400000004", "확정", "이바다", "010-0000-2222", "16:00", "18:00", "connectRoom", "커넥트룸", 3, "회의"],
  ];
  for (const [reservationId, status, bookerName, maskedPhone, startTime, endTime, spaceType, placeName, headcount, purpose] of naverRows) {
    context.appendObjectRow_(SHEET_NAMES.NAVER_SYNC, HEADERS.naverSync, {
      reservationId, status, bookerName, maskedPhone, usageDate: today, startTime, endTime, placeName, spaceType,
      headcount, purpose, requestedAt: now, confirmedAt: now,
      cancelledAt: status === "확정" ? "" : now,
      cancelReason: status === "노쇼" ? "직원 프로그램에서 노쇼 처리" : status === "취소" ? "네이버 취소" : "",
      lastSyncedAt: now, sourceFileName: "demo", rawUsageText: `${startTime}~${endTime}`,
    });
  }

  const requests = [
    ["web-demo-1", "신청", "m-park", "박서준", "남성", "01000003333", "careerRoom2", "14:00", "15:00", "포트폴리오 정리"],
    ["web-demo-2", "반려", "m-han", "한반려", "여성", "01000006666", "careerRoom1", "11:00", "12:00", "개인 공부"],
    ["web-demo-3", "승인", "m-lee", "이바다", "남성", "01000002222", "careerRoom1", "19:00", "20:00", "온라인 면접"],
    ["web-demo-4", "취소", "m-jung", "정취소", "남성", "01000005555", "connectRoom", "12:00", "13:00", "모임"],
  ];
  // 내일 10시 신청을 15분 전에 받아 두고 아직 승인하지 않은 상태(직원 프로그램 '미승인 확인창' 확인용).
  const fifteenMinutesAgo = context.Utilities.formatDate(new Date(Date.now() - 15 * 60000), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  requests.push(["web-demo-5", "신청", "m-han", "한반려", "여성", "01000006666", "careerRoom1", "10:00", "12:00", "자격증 공부", todayText(1), fifteenMinutesAgo]);
  for (const [requestId, status, memberId, name, gender, phone, spaceType, startTime, endTime, purpose, date = today, requestedAt = now] of requests) {
    context.appendObjectRow_(SHEET_NAMES.RESERVATION_REQUESTS, HEADERS.reservationRequests, {
      requestId, status, memberId, name, memberNameKey: context.normalizeNameKey_(name), gender, age: 27, phone,
      phoneLastDigits: phone.slice(-4), spaceType, placeId: spaceType, placeName: context.getReservationSpaceLabel_(spaceType),
      date, startTime, endTime, purpose, headcount: 1, companionsJson: "[]", companionCount: 0,
      requestedAt, decidedAt: status === "신청" ? "" : now, decidedBy: status === "신청" ? "" : "demo",
      updatedAt: now,
    });
  }
}

// 직원 프로그램이 5분마다 읽는 네이버 결과 파일의 가짜본(스크래퍼 대신).
function writeNaverSnapshot() {
  const today = todayText();
  const later = todayText(3);
  const item = (reservationNo, status, bookerName, phone, date, start, end, productName, purpose, filterType) => ({
    reservationNo, status, bookerName, phone, proxyBooking: false, changed: false,
    useDateTime: koreanUseDateTime(date, start, end), productName, purpose, headcount: 1,
    rawText: `${status}${bookerName}${phone}${reservationNo}${start}~${end}${productName}`,
    filterType,
  });
  const upcoming = [
    item("1400000001", "확정", "김하늘", "010-0000-1111", today, "10:00", "12:00", "AI 커리어존 (청취Room1)", "면접 준비", "upcoming"),
    item("1400000002", "확정", "최노쇼", "010-0000-4444", today, "09:00", "10:00", "청년 커넥트룸", "스터디", "upcoming"),
    item("1400000003", "취소", "정취소", "010-0000-5555", today, "13:00", "15:00", "AI 커리어존 (청취Room2)", "자기소개서", "upcoming"),
    item("1400000004", "확정", "이바다", "010-0000-2222", today, "16:00", "18:00", "청년 커넥트룸", "회의", "upcoming"),
    // 앞으로의 예약: 시트에 미리 들어가야 웹 예약과 겹치지 않는다.
    item("1400000006", "확정", "송미래", "010-0000-8888", later, "10:00", "12:00", "AI 커리어존 (청취Room1)", "스터디", "upcoming"),
    item("1400000007", "취소", "강취소", "010-0000-9999", later, "13:00", "14:00", "청년 커넥트룸", "모임", "upcoming"),
  ];
  const snapshot = {
    syncedAt: new Date().toISOString(),
    summary: {},
    upcomingRange: { startDate: today, endDate: todayText(29) },
    results: {
      upcoming,
      today_use: [
        // 시트에서 노쇼로 바꾼 예약도 네이버에는 '확정'으로 남아 있다(노쇼가 덮이지 않는지 확인용).
        item("1400000001", "확정", "김하늘", "010-0000-1111", today, "10:00", "12:00", "AI 커리어존 (청취Room1)", "면접 준비", "today_use"),
        item("1400000002", "확정", "최노쇼", "010-0000-4444", today, "09:00", "10:00", "청년 커넥트룸", "스터디", "today_use"),
        item("1400000004", "확정", "이바다", "010-0000-2222", today, "16:00", "18:00", "청년 커넥트룸", "회의", "today_use"),
      ],
      pending_confirm: [
        item("1400000005", "신청", "윤새봄", "010-0000-7777", later, "14:00", "16:00", "AI 커리어존 (청취Room2)", "스터디", "pending_confirm"),
      ],
      today_cancel: [
        item("1400000003", "취소", "정취소", "010-0000-5555", today, "13:00", "15:00", "AI 커리어존 (청취Room2)", "자기소개서", "today_cancel"),
      ],
    },
  };
  fs.writeFileSync(NAVER_FILE, JSON.stringify(snapshot, null, 2), "utf8");
}

function loadApp() {
  const keep = process.argv.includes("--keep") && fs.existsSync(SHEET_FILE);
  if (keep) {
    const saved = JSON.parse(fs.readFileSync(SHEET_FILE, "utf8"));
    return createVirtualAppsScript(saved);
  }
  const app = createVirtualAppsScript();
  seed(app);
  return app;
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"]/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[ch]);
}

function renderSheets(app) {
  const wanted = () => true;
  const sections = app.spreadsheet.getSheets()
    .filter((sheet) => wanted(sheet.getName()))
    .map((sheet) => {
      const rows = sheet.rows.filter((row) => row.some((value) => value !== ""));
      const [header = [], ...body] = rows;
      return `<section><h2>${escapeHtml(sheet.getName())} <small>${body.length}줄</small></h2>
        <div class="wrap"><table><thead><tr>${header.map((h) => `<th>${escapeHtml(h)}</th>`).join("")}</tr></thead>
        <tbody>${body.map((row) => `<tr>${header.map((_, i) => `<td>${escapeHtml(row[i])}</td>`).join("")}</tr>`).join("")}</tbody></table></div></section>`;
    }).join("");
  return `<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
  <title>가상 시트</title><meta http-equiv="refresh" content="10">
  <style>body{font-family:"맑은 고딕",sans-serif;margin:16px;background:#f7f7f6;color:#222}h1{font-size:20px}
  h2{font-size:15px;margin:24px 0 6px}small{color:#888;font-weight:400}.wrap{overflow-x:auto;background:#fff;border:1px solid #ddd}
  table{border-collapse:collapse;font-size:12px;white-space:nowrap}th,td{border:1px solid #e5e5e5;padding:4px 8px;text-align:left}
  th{background:#4285f4;color:#fff;position:sticky;top:0}</style></head>
  <body><h1>가상 시트 (데모 · 10초마다 새로고침)</h1><p>오늘: ${todayText()} · 모든 이름·연락처는 가짜입니다.</p>${sections}</body></html>`;
}

function serveStatic(response, filePath, scriptUrl) {
  const resolved = path.resolve(ROOT, "." + filePath);
  if (!resolved.startsWith(ROOT) || !fs.existsSync(resolved) || fs.statSync(resolved).isDirectory()) {
    response.writeHead(404).end("not found");
    return;
  }
  const types = { ".html": "text/html; charset=utf-8", ".js": "application/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".png": "image/png", ".svg": "image/svg+xml", ".ico": "image/x-icon", ".json": "application/json" };
  let body = fs.readFileSync(resolved);
  if (filePath === "/js/config.js") {
    body = body.toString("utf8").replace(/GOOGLE_SCRIPT_URL = "[^"]*"/, `GOOGLE_SCRIPT_URL = "${scriptUrl}"`);
  }
  response.writeHead(200, { "Content-Type": types[path.extname(resolved)] || "application/octet-stream", "Cache-Control": "no-store" });
  response.end(body);
}

function main() {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  const app = loadApp();
  writeNaverSnapshot();
  const save = () => {
    fs.writeFileSync(SHEET_FILE, JSON.stringify({ sheets: app.spreadsheet.toJSON(), properties: app.properties }), "utf8");
    app.spreadsheet.changed = false;
  };
  save();

  const server = http.createServer((request, response) => {
    const url = new URL(request.url, `http://${request.headers.host || `127.0.0.1:${PORT}`}`);
    if (url.pathname === "/exec") {
      const parameter = Object.fromEntries(url.searchParams.entries());
      const output = app.context.doGet({ parameter });
      const action = parameter.action || "(상태 확인)";
      let ok = true;
      try { ok = JSON.parse(output.getContent().replace(/^[\w$]+\(|\);$/g, "")).ok !== false; } catch { /* 텍스트 응답 */ }
      console.log(`${new Date().toLocaleTimeString("ko-KR")}  ${ok ? "✓" : "✗"} ${action}${ok ? "" : " → " + output.getContent().slice(0, 160)}`);
      if (app.spreadsheet.changed) save();
      response.writeHead(200, { "Content-Type": `${output.mimeType}; charset=utf-8`, "Cache-Control": "no-store" });
      response.end(output.getContent());
      return;
    }
    if (url.pathname === "/sheet") {
      response.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
      response.end(renderSheets(app));
      return;
    }
    if (url.pathname === "/naver/reset") {
      writeNaverSnapshot();
      response.writeHead(200, { "Content-Type": "text/plain; charset=utf-8" }).end("가짜 네이버 결과를 다시 만들었습니다.");
      return;
    }
    const scriptUrl = `${url.protocol}//${url.host}/exec`;
    serveStatic(response, url.pathname === "/" ? "/index.html" : url.pathname, scriptUrl);
  });

  server.listen(PORT, "127.0.0.1", () => {
    console.log("가상 시트 데모 서버가 시작됐습니다. (실제 구글 시트·네이버에는 보내지 않습니다)");
    console.log(`  직원 프로그램 : "청년이봄 직원알림.exe" --virtual-sheet=http://127.0.0.1:${PORT}/exec`);
    console.log(`  이용자 QR 화면: http://demo.localhost:${PORT}/?entry=${DEMO_ENTRY_TOKEN}`);
    console.log(`  가상 시트 보기: http://127.0.0.1:${PORT}/sheet`);
    console.log(`  가짜 네이버 결과: ${NAVER_FILE}`);
  });
}

if (require.main === module) main();

module.exports = { DEMO_ADMIN_PASSWORD, DEMO_ENTRY_TOKEN, seed, loadApp };
