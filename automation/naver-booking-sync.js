#!/usr/bin/env node

const fs = require("fs");
const path = require("path");

const PLAYWRIGHT_NODE_MODULES =
  "C:\\Users\\SNYOUTH\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\node_modules";
process.env.NODE_PATH = process.env.NODE_PATH
  ? `${PLAYWRIGHT_NODE_MODULES}${path.delimiter}${process.env.NODE_PATH}`
  : PLAYWRIGHT_NODE_MODULES;
require("module").Module._initPaths();

const { chromium } = require("playwright");

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
    keepOpen: false,
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--headless") out.headless = true;
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
  await new Promise((resolve) => {
    process.stdin.resume();
    process.stdin.once("data", () => {
      process.stdin.pause();
      resolve();
    });
  });
}

async function closeAnyDropdownOverlays(page) {
  await page.keyboard.press("Escape").catch(() => {});
  await page.waitForTimeout(300);
}

async function ensureProductSelection(page, products, timeoutMs) {
  await closeAnyDropdownOverlays(page);
  await page.getByRole("button", { name: /상품/ }).click({ timeout: timeoutMs });
  await page.waitForTimeout(500);

  const overlay = page.locator("body");
  const allCheckbox = overlay.getByText("전체", { exact: true }).first();
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
        bookerName: status ? clean(source.replace(status, "").split(phoneMatches[0] || "")[0]) : null,
        phone: phoneMatches[0] || null,
        altPhone: phoneMatches[1] || null,
        useDateTime,
        productName,
      };
    });
  });
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

function buildSummary(results) {
  return Object.fromEntries(
    Object.entries(results).map(([key, rows]) => [key, { count: rows.length }])
  );
}

async function sync(config) {
  ensureDir(OUTPUT_DIR);
  ensureDir(STATE_DIR);

  const context = await chromium.launchPersistentContext(USER_DATA_DIR, {
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
    }

    await ensureProductSelection(page, config.products, config.timeoutMs);

    const results = {};
    for (const filter of QUICK_FILTERS) {
      await activateQuickFilter(page, filter.label, config.timeoutMs);
      const rows = dedupeByReservation(await collectRows(page)).map((row) => ({
        ...row,
        filterType: filter.key,
        filterLabel: filter.label,
        collectedAt: new Date().toISOString(),
      }));
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
