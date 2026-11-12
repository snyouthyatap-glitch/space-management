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
const QUICK_FILTERS = [
  { label: "오늘이용", key: "today_use" },
  { label: "확정대기", key: "pending_confirm" },
  { label: "오늘취소", key: "today_cancel" },
];

const REPO_ROOT = path.resolve(__dirname, "..");
const OUTPUT_DIR = path.join(REPO_ROOT, "outputs");
const STATE_DIR = path.join(__dirname, ".naver-booking-state");
const USER_DATA_DIR = path.join(STATE_DIR, "chrome-profile");
const DEFAULT_OUTPUT_FILE = path.join(OUTPUT_DIR, "naver-booking-sync.json");

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
      const row =
        anchor.closest("tr") ||
        anchor.closest("[class*='row']") ||
        anchor.closest("li") ||
        anchor.parentElement;
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
    findPhones(detailFieldValue(fields, /^방문자\s*(전화번호|연락처)$/))[0] ||
    "";
  const proxy = row.proxyBooking || isNaverProxyBooking(rawBooker) || Boolean(visitorName);
  const bookerName = visitorName || cleanNaverBookerName(rawBooker);
  // 대리예약 상세의 '전화번호'는 예약자 번호일 수 있으므로, 방문자 번호가 없으면 목록 값을 유지한다.
  const phone = visitorPhone || (proxy ? "" : detailFieldValue(fields, /^전화번호$/));
  const reservationNo = detailFieldValue(fields, /^예약번호$/);
  const productName = detailFieldValue(fields, /^상품$/);
  const useDateTime = detailFieldValue(fields, /^이용일시$/);
  const purpose = detailFieldValue(fields, /^이용 목적$/);
  const headcount = parseHeadcount(detailFieldValue(fields, /^이용 인원/));
  const requestedAt = detailFieldValue(fields, /^신청$/);
  const confirmedAt = detailFieldValue(fields, /^확정$/);

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

async function enrichRowsWithDetails(context, baseUrl, rows, timeoutMs) {
  const enriched = [];
  for (const row of rows) {
    const url = detailUrl(baseUrl, row.detailPath);
    if (!url) {
      enriched.push(row);
      continue;
    }

    const detailPage = await context.newPage();
    detailPage.setDefaultTimeout(timeoutMs);
    try {
      await detailPage.goto(url, { waitUntil: "domcontentloaded", timeout: timeoutMs });
      await waitForPageReady(detailPage, timeoutMs);
      const detail = await collectDetail(detailPage);
      const parsedDetail = parseDetailFields(detail.detailFields, row);
      enriched.push({
        ...row,
        ...parsedDetail,
        detailUrl: url,
        detailText: detail.detailText,
        detailFields: detail.detailFields,
        detailCollectedAt: new Date().toISOString(),
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

    await ensureProductSelection(page, config.products, config.timeoutMs);

    const results = {};
    for (const filter of QUICK_FILTERS) {
      await activateQuickFilter(page, filter.label, config.timeoutMs);
      let rows = dedupeByReservation(await collectRows(page)).map((row) => ({
        ...applyBookerRule(row),
        filterType: filter.key,
        filterLabel: filter.label,
        collectedAt: new Date().toISOString(),
      }));
      if (config.detail) {
        rows = await enrichRowsWithDetails(context, config.baseUrl, rows, config.timeoutMs);
      }
      rows = filterRowsByQuickFilter(filter.key, rows);
      results[filter.key] = rows;
    }

    const payload = {
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

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
