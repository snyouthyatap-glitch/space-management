#!/usr/bin/env node

const fs = require("fs");
const os = require("os");
const path = require("path");
const {
  cleanNaverBookerName,
  extractUseDateTime,
  findPhones,
  isChangedReservation,
  isNaverProxyBooking,
  resolveNaverBooker,
} = require("./naver-booker");

function loadPlaywright() {
  try {
    return require("playwright");
  } catch (error) {
    const candidates = [
      process.env.PLAYWRIGHT_NODE_MODULES,
      path.join(
        os.homedir(),
        ".cache",
        "codex-runtimes",
        "codex-primary-runtime",
        "dependencies",
        "node",
        "node_modules",
        "playwright"
      ),
    ].filter(Boolean);

    for (const candidate of candidates) {
      try {
        return require(candidate);
      } catch {
        // Try the next configured runtime path.
      }
    }

    throw error;
  }
}

const { chromium } = loadPlaywright();

const DEFAULT_BASE_URL =
  "https://partner.booking.naver.com/bizes/1002115/booking-list-view";
const DEFAULT_PRODUCTS = [
  "AI 커리어존 (청취Room1)",
  "AI 커리어존 (청취Room2)",
  "청년 커넥트룸",
];
// 이용일 기준 한 달(오늘부터 30일) 목록을 한 번 읽고, 오늘이용·오늘취소는 거기서 골라낸다.
// 확정대기는 날짜와 상관없이 모아 보여 주는 빠른 필터로 따로 읽는다.
const UPCOMING_DAYS = 30;
const QUICK_FILTERS = [{ label: "확정대기", key: "pending_confirm" }];
// 상세 화면은 새 예약·바뀐 예약만 연다(한 번 실행에 최대 개수). 나머지는 지난 결과를 다시 쓴다.
const DETAIL_FETCH_LIMIT = 12;

const REPO_ROOT = path.resolve(__dirname, "..");
const OUTPUT_DIR = path.join(REPO_ROOT, "outputs");
const STATE_DIR = path.join(__dirname, ".naver-booking-state");
const USER_DATA_DIR = process.env.NAVER_BOOKING_PROFILE_DIR || path.join(STATE_DIR, "chrome-profile");
const DEFAULT_OUTPUT_FILE = path.join(OUTPUT_DIR, "naver-booking-sync.json");
const DETAIL_CACHE_FILE = process.env.NAVER_DETAIL_CACHE_FILE || path.join(OUTPUT_DIR, "naver-detail-cache.json");

function seoulDate(offsetDays = 0) {
  return new Intl.DateTimeFormat("en-CA", { timeZone: "Asia/Seoul" }).format(
    new Date(Date.now() + offsetDays * 86400000)
  );
}

// '26.11.18.수 13:00~16:00', '2026. 11. 17.(화) 오후 4:00~…' → '2026-11-18'
function rowUseDate(row) {
  const match = String(row.useDateTime || "").match(/(\d{2,4})\.\s*(\d{1,2})\.\s*(\d{1,2})\./);
  if (!match) return "";
  const year = match[1].length === 2 ? 2000 + Number(match[1]) : Number(match[1]);
  return `${year}-${String(match[2]).padStart(2, "0")}-${String(match[3]).padStart(2, "0")}`;
}

function monthListUrl(baseUrl, startDate, endDate) {
  const url = new URL(baseUrl);
  url.searchParams.set("dateDropdownType", "MONTH");
  url.searchParams.set("startDateTime", startDate);
  url.searchParams.set("endDateTime", endDate);
  url.searchParams.set("dateFilter", "USEDATE");
  return url.toString();
}

// 목록 결과로 오늘이용·오늘취소를 만든다(예전 빠른 필터와 같은 모양).
function splitUpcomingRows(rows, today) {
  const status = (row) => normalizeReservationStatus(row.status);
  const todayRows = rows.filter((row) => rowUseDate(row) === today);
  return {
    today_use: todayRows
      .filter((row) => ["확정", "완료", "노쇼"].includes(status(row)))
      .map((row) => ({ ...row, filterType: "today_use", filterLabel: "오늘이용" })),
    today_cancel: todayRows
      .filter((row) => status(row) === "취소")
      .map((row) => ({ ...row, filterType: "today_cancel", filterLabel: "오늘취소" })),
  };
}

function loadDetailCache() {
  try {
    return JSON.parse(fs.readFileSync(DETAIL_CACHE_FILE, "utf8"));
  } catch {
    return {};
  }
}

function saveDetailCache(cache) {
  // 40일 넘게 다시 보이지 않은 예약은 지운다.
  const cutoff = Date.now() - 40 * 86400000;
  const kept = Object.fromEntries(
    Object.entries(cache).filter(([, item]) => Date.parse(item.seenAt || 0) >= cutoff)
  );
  fs.writeFileSync(DETAIL_CACHE_FILE, JSON.stringify(kept), "utf8");
}

function parseArgs(argv) {
  const out = {
    baseUrl: process.env.NAVER_BOOKING_BASE_URL || DEFAULT_BASE_URL,
    headless: false,
    setupLogin: false,
    output: process.env.NAVER_BOOKING_OUTPUT || DEFAULT_OUTPUT_FILE,
    timeoutMs: Number(process.env.NAVER_BOOKING_TIMEOUT_MS || 30000),
    products: DEFAULT_PRODUCTS.slice(),
    naverId: process.env.NAVER_BOOKING_ID || "",
    naverPassword: process.env.NAVER_BOOKING_PASSWORD || "",
    keepOpen: false,
    detail: process.env.NAVER_BOOKING_DETAIL !== "0",
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--headless") out.headless = true;
    else if (arg === "--no-detail") out.detail = false;
    else if (arg === "--setup-login") out.setupLogin = true;
    else if (arg === "--keep-open") out.keepOpen = true;
    else if (arg === "--base-url") out.baseUrl = argv[++i];
    else if (arg === "--output") out.output = path.resolve(argv[++i]);
    else if (arg === "--timeout-ms") out.timeoutMs = Number(argv[++i]);
    else if (arg === "--products") {
      out.products = argv[++i]
        .split(",")
        .map((v) => v.trim())
        .filter(Boolean);
    } else if (arg === "--help") {
      printHelp();
      process.exit(0);
    } else {
      throw new Error(`Unknown argument: ${arg}`);
    }
  }

  return out;
}

function printHelp() {
  console.log(`Usage:
  node automation/naver-booking-sync.js [options]

Options:
  --setup-login           Open the page and wait for manual login confirmation
  --headless              Run headless
  --no-detail             Skip opening each reservation detail page
  --keep-open             Keep the browser open after sync
  --base-url <url>        Booking list URL
  --output <path>         Output JSON path
  --timeout-ms <number>   Action timeout
  --products <a,b,c>      Product names to keep selected
  --help                  Show help
`);
}

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

async function waitForPageReady(page, timeoutMs) {
  await page.waitForLoadState("domcontentloaded", { timeout: timeoutMs });
  await page.waitForTimeout(1500);
}

async function setupLogin(page, timeoutMs) {
  console.log("로그인이 필요한 경우 브라우저에서 로그인한 뒤 Enter를 눌러 주세요.");
  await waitForPageReady(page, timeoutMs);
  if (process.stdin.isTTY) {
    await new Promise((resolve) => {
      process.stdin.resume();
      process.stdin.once("data", () => {
        process.stdin.pause();
        resolve();
      });
    });
    return;
  }

  // 직원 프로그램처럼 콘솔 입력이 없는 곳에서 실행되면, 로그인이 끝나 예약 화면이 열릴 때까지 기다린다.
  const deadline = Date.now() + Math.max(timeoutMs, 300000);
  while (Date.now() < deadline) {
    if (!(await isNaverLoginPage(page)) && page.url().includes("booking-list-view")) {
      return;
    }
    await page.waitForTimeout(1000);
  }

  throw new Error("NAVER_LOGIN_REQUIRED: 열린 브라우저에서 네이버 로그인을 마친 뒤 다시 시도해 주세요.");
}

async function isNaverLoginPage(page) {
  const url = page.url();
  if (/nid\.naver\.com\/nidlogin/i.test(url)) return true;
  return page.locator("#id, input[name='id'], input[name='pw'], #pw").first().isVisible().catch(() => false);
}

async function tryNaverCredentialLogin(page, config) {
  if (!(await isNaverLoginPage(page))) return;
  if (!config.naverId || !config.naverPassword) {
    throw new Error("NAVER_LOGIN_REQUIRED: saved Naver credentials are missing.");
  }

  const idInput = page.locator("#id, input[name='id']").first();
  const passwordInput = page.locator("#pw, input[name='pw']").first();
  await idInput.waitFor({ state: "visible", timeout: config.timeoutMs });
  await idInput.fill(config.naverId);
  await passwordInput.fill(config.naverPassword);
  // 네이버 로그인 화면은 가로형·세로형 로그인 버튼을 따로 그린다. 보이는 쪽을 누른다.
  await page.getByRole("button", { name: "로그인", exact: true }).click({
    timeout: config.timeoutMs,
  });
  await page.waitForURL((url) => !/nid\.naver\.com\/nidlogin/i.test(url.href), {
    timeout: Math.min(config.timeoutMs, 10000),
  }).catch(() => {});
  await waitForPageReady(page, config.timeoutMs);

  if (await isNaverLoginPage(page)) {
    const messages = await page.locator(".error_message, .error, [role='alert']").allTextContents();
    const reason = maskSecrets(messages.join(" ").replace(/\s+/g, " ").trim(), config).slice(0, 350);
    throw new Error(
      "NAVER_LOGIN_REQUIRED: " +
        (reason ||
          "네이버 로그인이 완료되지 않았습니다. 설정의 '네이버 로그인 설정'에서 직접 로그인하거나 추가 인증 여부를 확인해 주세요.")
    );
  }
}

// 오류 메시지에 아이디·비밀번호가 섞여 나오지 않게 가린다.
function maskSecrets(text, config) {
  let masked = String(text || "");
  if (config.naverId) masked = masked.split(config.naverId).join("[아이디]");
  if (config.naverPassword) masked = masked.split(config.naverPassword).join("[비밀번호]");
  return masked;
}

async function closeAnyDropdownOverlays(page) {
  await page.keyboard.press("Escape").catch(() => {});
  await page.waitForTimeout(300);
}

async function ensureProductSelection(page, products, timeoutMs) {
  await closeAnyDropdownOverlays(page);
  // 바뀐 예약자관리 화면은 상품 필터를 #dropdownHost로 그린다. 없으면 예전 방식으로 찾는다.
  const productButton = page.locator("#dropdownHost").first();
  if (await productButton.count()) {
    await productButton.click({ timeout: timeoutMs });
  } else {
    await page.getByRole("button", { name: /상품/ }).first().click({ timeout: timeoutMs });
  }
  await page.waitForTimeout(500);

  const overlay = page.locator("body");
  const allCheckbox = overlay.getByRole("checkbox", { name: "전체", exact: true }).first();
  if (await allCheckbox.isVisible().catch(() => false)) {
    const allRow = allCheckbox.locator("xpath=ancestor::*[self::label or self::li or self::div][1]");
    const checked = await allRow
      .evaluate((node) => {
        const candidate = node.querySelector("input[type='checkbox']");
        return candidate ? candidate.checked : /checked|selected|on/i.test(node.className);
      })
      .catch(() => false);
    if (checked) {
      await allCheckbox.click({ timeout: timeoutMs });
      await page.waitForTimeout(250);
    }
  }

  for (const product of products) {
    const productCheckbox = overlay.getByRole("checkbox", { name: product, exact: true }).first();
    if ((await productCheckbox.count()) && (await productCheckbox.isVisible().catch(() => false))) {
      const isChecked = await productCheckbox.evaluate((node) => node.checked).catch(() => false);
      if (!isChecked) await productCheckbox.click({ timeout: timeoutMs });
      await page.waitForTimeout(250);
      continue;
    }

    const productText = page.getByText(product, { exact: true }).first();
    await productText.waitFor({ state: "visible", timeout: timeoutMs });
    const productRow = productText.locator(
      "xpath=ancestor::*[self::label or self::li or self::div][1]"
    );
    const isChecked = await productRow
      .evaluate((node) => {
        const candidate = node.querySelector("input[type='checkbox']");
        return candidate ? candidate.checked : /checked|selected|on/i.test(node.className);
      })
      .catch(() => false);
    if (!isChecked) {
      await productText.click({ timeout: timeoutMs });
      await page.waitForTimeout(250);
    }
  }

  await closeAnyDropdownOverlays(page);
}

async function activateQuickFilter(page, label, timeoutMs) {
  const target = page.getByRole("button", { name: new RegExp(escapeRegExp(label)) }).first();
  await target.waitFor({ state: "visible", timeout: timeoutMs });
  await target.click({ timeout: timeoutMs });
  await page.waitForTimeout(1200);
}

async function collectRows(page) {
  return page.evaluate(() => {
    const clean = (value) => (value || "").replace(/\s+/g, " ").trim();
    const anchors = Array.from(document.querySelectorAll("a[href*='booking-list-view/bookings/']"));
    const map = new Map();

    for (const anchor of anchors) {
      const href = anchor.getAttribute("href") || "";
      const match = href.match(/bookings\/(\d+)/);
      if (!match) continue;
      const reservationNo = match[1];
      // 예약 한 건만 담은 가장 큰 요소를 행으로 쓴다.
      // (class에 'row'가 든 목록 전체를 행으로 잡으면 앞 예약의 연락처가 섞인다)
      const holdsOtherReservation = (node) =>
        Array.from(node.querySelectorAll("a[href*='booking-list-view/bookings/']"))
          .some((other) => !(other.getAttribute("href") || "").includes(`bookings/${reservationNo}`));
      // 행 글자는 상태(확정·신청 등)로 시작한다. 예약이 한 건뿐일 때 목록 머리글이나 화면 전체를 잡지 않게 한다.
      const startsWithStatus = (node) => /^(확정|신청|취소|완료|노쇼)/.test(clean(node.textContent));
      let row = anchor.closest("tr");
      if (!row || holdsOtherReservation(row)) {
        row = anchor;
        let node = anchor.parentElement;
        while (node && node !== document.body && !holdsOtherReservation(node)) {
          if (startsWithStatus(node)) row = node;
          node = node.parentElement;
        }
      }
      if (!row) continue;
      const text = clean(row.textContent);
      if (!text || text.length < 10) continue;
      if (!map.has(reservationNo) || map.get(reservationNo).text.length < text.length) {
        map.set(reservationNo, {
          reservationNo,
          href,
          text,
        });
      }
    }

    return Array.from(map.values()).map((item) => {
      const source = item.text;
      const extract = (pattern) => {
        const match = source.match(pattern);
        return match ? clean(match[1]) : null;
      };
      const status = extract(/^(확정|신청|취소|완료|노쇼)/);
      const phoneMatches = [...source.matchAll(/01[0-9]-\d{3,4}-\d{4}/g)].map((m) => m[0]);
      // 이름 정리는 브라우저 밖에서 naver-booker.js 규칙으로 한다.
      const bookerText = (() => {
        if (!status) return null;
        const firstPhone = phoneMatches[0] || "";
        return clean(firstPhone
          ? source.replace(status, "").split(firstPhone)[0]
          : source.replace(status, ""));
      })();
      const useDateTime =
        extract(/(\d{2}\.\s*\d{1,2}\.\s*\d{1,2}\.\([^)]+\)\s*(?:오전|오후)\s*\d{1,2}:\d{2}~(?:오전|오후)?\s*\d{1,2}:\d{2})/) ||
        extract(/(\d{2}\.\s*\d{1,2}\.\s*\d{1,2}\.\([^)]+\)\s*(?:오전|오후)\s*\d{1,2}:\d{2}~\d{1,2}:\d{2})/);

      const productCandidates = [
        "청년 커넥트룸",
        "AI 커리어존 (청취Room1)",
        "AI 커리어존 (청취Room2)",
      ];
      const productName = productCandidates.find((name) => source.includes(name)) || null;

      return {
        reservationNo: item.reservationNo,
        detailPath: item.href,
        rawText: source,
        status,
        bookerText,
        phones: phoneMatches,
        altPhone: phoneMatches[1] || null,
        useDateTime,
        productName,
      };
    });
  });
}

function applyBookerRule(row) {
  const { bookerText, phones, ...rest } = row;
  return {
    ...rest,
    ...resolveNaverBooker(bookerText, phones),
    // 브라우저 안 정규식은 '오후 6:00' 형태만 읽어서, 24시간제 행은 여기서 다시 읽는다.
    useDateTime: rest.useDateTime || extractUseDateTime(rest.rawText, rest.reservationNo),
    changed: isChangedReservation(rest.rawText, rest.reservationNo),
  };
}

function dedupeByReservation(rows) {
  const map = new Map();
  for (const row of rows) {
    const existing = map.get(row.reservationNo);
    if (!existing || (existing.rawText || "").length < (row.rawText || "").length) {
      map.set(row.reservationNo, row);
    }
  }
  return Array.from(map.values());
}

function detailUrl(baseUrl, detailPath) {
  if (!detailPath) return "";
  if (/^https?:\/\//i.test(detailPath)) return detailPath;
  return new URL(detailPath, baseUrl).toString();
}

function detailFieldValue(fields, labelPattern) {
  const field = (fields || []).find((item) => labelPattern.test(item.label || ""));
  return field ? field.value || "" : "";
}

// 같은 이름표가 여러 번 잡혀도(목록 머리글 등) 실제 번호가 든 값을 고른다.
function detailFieldPhone(fields, labelPattern) {
  for (const field of fields || []) {
    if (!labelPattern.test(field.label || "")) continue;
    const phone = findPhones(field.value)[0];
    if (phone) return phone;
  }
  return "";
}

function detailDateTime(value) {
  const match = String(value || "").match(
    /^\d{2,4}\.\s*\d{1,2}\.\s*\d{1,2}\.\s*(?:\([^)]*\))?\s*(?:오전|오후)?\s*\d{1,2}:\d{2}/
  );
  return match ? match[0].trim() : String(value || "").trim();
}

function parseHeadcount(value) {
  const match = String(value || "").match(/\d+/);
  return match ? Number(match[0]) : null;
}

function parseDetailFields(fields, row = {}) {
  const rawBooker = detailFieldValue(fields, /^예약자$/);
  // 실제 상세 화면: 예약자 = "김예약대리예약", 방문자 = "이방문(010-0000-0000)"
  const rawVisitor = detailFieldValue(fields, /^방문자(명)?$/);
  const visitorName = cleanNaverBookerName(rawVisitor);
  const visitorPhone =
    findPhones(rawVisitor)[0] ||
    detailFieldPhone(fields, /^방문자\s*(전화번호|연락처)$/) ||
    "";
  // 대리예약 여부는 이 예약의 상세 화면으로 판단한다. 목록 행 값은 상세에 예약자가 없을 때만 쓴다.
  const proxy = rawBooker
    ? isNaverProxyBooking(rawBooker) || Boolean(visitorName)
    : Boolean(row.proxyBooking) || Boolean(visitorName);
  const bookerName = visitorName || cleanNaverBookerName(rawBooker);
  // 대리예약 상세의 '전화번호'는 예약자 번호일 수 있으므로, 방문자 번호가 없으면 목록 값을 유지한다.
  const phone = visitorPhone || (proxy ? "" : detailFieldPhone(fields, /^전화번호$/));
  // 바뀐 예약은 상세 화면 예약번호 옆에 "변경"이 붙어 있어 숫자만 쓴다.
  const reservationNo = (detailFieldValue(fields, /^예약번호$/).match(/\d{6,}/) || [""])[0];
  const productName = detailFieldValue(fields, /^상품$/);
  const useDateTime = detailFieldValue(fields, /^이용일시$/);
  const purpose = detailFieldValue(fields, /^이용 목적$/);
  const headcount = parseHeadcount(detailFieldValue(fields, /^이용 인원/));
  // '확정' 값 뒤에는 처리한 곳(파트너센터·운영자)이 붙어 있어 날짜·시각만 남긴다.
  const requestedAt = detailDateTime(detailFieldValue(fields, /^신청$/));
  const confirmedAt = detailDateTime(detailFieldValue(fields, /^확정$/));

  return {
    ...(bookerName ? { bookerName } : {}),
    ...(phone ? { phone } : {}),
    ...(proxy ? { proxyBooking: true } : {}),
    ...(reservationNo ? { reservationNo } : {}),
    ...(productName ? { productName } : {}),
    ...(useDateTime ? { useDateTime } : {}),
    ...(purpose ? { purpose } : {}),
    ...(headcount ? { headcount } : {}),
    ...(requestedAt ? { requestedAt } : {}),
    ...(confirmedAt ? { confirmedAt } : {}),
  };
}

async function collectDetail(page) {
  return page.evaluate(() => {
    const clean = (value) => (value || "").replace(/\s+/g, " ").trim();
    const text = clean(document.body?.innerText || document.body?.textContent || "");
    const fields = [];
    const candidates = Array.from(
      document.querySelectorAll("dt, th, label, [class*='label'], [class*='title'], [class*='key']")
    );

    for (const labelNode of candidates) {
      const label = clean(labelNode.textContent);
      if (!label || label.length > 30) continue;
      const parent = labelNode.parentElement;
      const valueNode =
        labelNode.nextElementSibling ||
        parent?.querySelector("dd, td, [class*='value'], [class*='content'], [class*='desc']");
      const value = clean(valueNode?.textContent || "");
      if (!value || value === label || value.length > 200) continue;
      fields.push({ label, value });
    }

    return { detailText: text, detailFields: fields.slice(0, 80) };
  });
}

// 상세 화면 결과는 예약번호별로 저장해 두고, 목록 줄이 그대로면 다시 열지 않는다.
// cache: { [예약번호]: { signature, parsed, seenAt } }, budget: { remaining }
async function enrichRowsWithDetails(context, baseUrl, rows, timeoutMs, cache = {}, budget = { remaining: Infinity }) {
  const enriched = [];
  for (const row of rows) {
    const url = detailUrl(baseUrl, row.detailPath);
    if (!url) {
      enriched.push(row);
      continue;
    }

    const signature = String(row.rawText || "");
    const cached = cache[row.reservationNo];
    if (cached && cached.signature === signature) {
      cached.seenAt = new Date().toISOString();
      enriched.push({ ...row, ...cached.parsed, detailUrl: url, detailCollectedAt: cached.collectedAt, detailCached: true });
      continue;
    }
    if (budget.remaining <= 0) {
      // 이번에는 상세를 열지 못했다. 지난 상세 결과가 있으면 쓰고, 다음 실행에서 다시 연다.
      enriched.push({ ...row, ...(cached ? cached.parsed : {}), detailUrl: url, detailPending: true });
      continue;
    }
    budget.remaining -= 1;

    const detailPage = await context.newPage();
    detailPage.setDefaultTimeout(timeoutMs);
    try {
      await detailPage.goto(url, { waitUntil: "domcontentloaded", timeout: timeoutMs });
      await waitForPageReady(detailPage, timeoutMs);
      const detail = await collectDetail(detailPage);
      const parsedDetail = parseDetailFields(detail.detailFields, row);
      const collectedAt = new Date().toISOString();
      cache[row.reservationNo] = { signature, parsed: parsedDetail, collectedAt, seenAt: collectedAt };
      saveDetailCache(cache); // 중간에 시간이 끝나도 다음 실행에서 다시 열지 않게 바로 저장한다.
      enriched.push({
        ...row,
        ...parsedDetail,
        detailUrl: url,
        detailCollectedAt: collectedAt,
      });
    } catch (error) {
      enriched.push({
        ...row,
        detailUrl: url,
        detailError: error.message || String(error),
      });
    } finally {
      await detailPage.close().catch(() => {});
    }
  }
  return enriched;
}

function buildSummary(results) {
  return Object.fromEntries(
    Object.entries(results).map(([key, rows]) => [key, { count: rows.length }])
  );
}

function normalizeReservationStatus(value) {
  const status = String(value || "").replace(/\s+/g, "");
  if (status.includes("신청")) return "신청";
  if (status.includes("확정")) return "확정";
  if (status.includes("취소")) return "취소";
  if (status.includes("완료")) return "완료";
  if (status.includes("노쇼")) return "노쇼";
  return status;
}

function filterRowsByQuickFilter(key, rows) {
  if (key !== "pending_confirm") return rows;
  return rows.filter((row) => normalizeReservationStatus(row.status) === "신청");
}

async function sync(config) {
  ensureDir(OUTPUT_DIR);
  ensureDir(STATE_DIR);

  const context = await chromium.launchPersistentContext(USER_DATA_DIR, {
    // 설정의 '네이버 로그인 설정'이 여는 실제 크롬과 같은 브라우저를 써야 로그인 상태를 같이 쓴다.
    channel: "chrome",
    headless: config.headless,
    viewport: { width: 1440, height: 1200 },
    args: ["--lang=ko-KR"],
  });

  try {
    const page = context.pages()[0] || (await context.newPage());
    page.setDefaultTimeout(config.timeoutMs);
    await page.goto(config.baseUrl, { waitUntil: "domcontentloaded" });
    await waitForPageReady(page, config.timeoutMs);

    if (config.setupLogin) {
      await setupLogin(page, config.timeoutMs);
    } else {
      await tryNaverCredentialLogin(page, config);
      if (!page.url().includes("booking-list-view")) {
        await page.goto(config.baseUrl, { waitUntil: "domcontentloaded", timeout: config.timeoutMs });
        await waitForPageReady(page, config.timeoutMs);
        await tryNaverCredentialLogin(page, config);
      }
    }

    // 이용일 기준 오늘부터 30일 목록(모든 상태)을 연다.
    const today = seoulDate();
    await page.goto(monthListUrl(config.baseUrl, today, seoulDate(UPCOMING_DAYS - 1)), {
      waitUntil: "domcontentloaded",
      timeout: config.timeoutMs,
    });
    await waitForPageReady(page, config.timeoutMs);
    await ensureProductSelection(page, config.products, config.timeoutMs);
    await page.waitForTimeout(1200);

    const detailCache = loadDetailCache();
    const budget = { remaining: DETAIL_FETCH_LIMIT };
    const collectList = async (key, label) => {
      let rows = dedupeByReservation(await collectRows(page))
        .filter((row) => !row.productName || config.products.includes(row.productName))
        .map((row) => ({
          ...applyBookerRule(row),
          filterType: key,
          filterLabel: label,
          collectedAt: new Date().toISOString(),
        }));
      if (config.detail) {
        rows = await enrichRowsWithDetails(context, config.baseUrl, rows, config.timeoutMs, detailCache, budget);
      }
      return rows;
    };

    const results = {};
    results.upcoming = await collectList("upcoming", `이용일 ${UPCOMING_DAYS}일`);
    Object.assign(results, splitUpcomingRows(results.upcoming, today));
    for (const filter of QUICK_FILTERS) {
      await activateQuickFilter(page, filter.label, config.timeoutMs);
      results[filter.key] = filterRowsByQuickFilter(filter.key, await collectList(filter.key, filter.label));
    }
    saveDetailCache(detailCache);

    const payload = {
      upcomingRange: { startDate: today, endDate: seoulDate(UPCOMING_DAYS - 1) },
      syncedAt: new Date().toISOString(),
      baseUrl: config.baseUrl,
      products: config.products,
      summary: buildSummary(results),
      results,
    };

    ensureDir(path.dirname(config.output));
    fs.writeFileSync(config.output, JSON.stringify(payload, null, 2), "utf8");
    console.log(`Saved: ${config.output}`);
    console.log(JSON.stringify(payload.summary, null, 2));

    if (config.keepOpen) {
      console.log("브라우저를 열어 둔 상태입니다. 종료하려면 Ctrl+C를 누르세요.");
      await new Promise(() => {});
    }
  } finally {
    if (!config.keepOpen) {
      await context.close();
    }
  }
}

async function main() {
  const config = parseArgs(process.argv.slice(2));
  await sync(config);
}

if (require.main === module) {
  main().catch((error) => {
    console.error(error);
    process.exit(1);
  });
}

module.exports = { parseDetailFields, rowUseDate, splitUpcomingRows, monthListUrl };
