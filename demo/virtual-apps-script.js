// 가상 구글 시트 위에서 APPS_SCRIPT_READY.gs를 그대로 실행한다(데모·점검용).
// 실제 구글 시트나 네이버에는 아무것도 보내지 않는다.
const crypto = require("node:crypto");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const SCRIPT_PATH = path.join(__dirname, "..", "APPS_SCRIPT_READY.gs");

function chainable(target) {
  // 서식 변경(setBackground 등)은 데모에서 의미가 없으므로 받기만 하고 자기 자신을 돌려준다.
  return new Proxy(target, {
    get(object, key) {
      if (key in object) return object[key];
      if (typeof key === "string" && /^(set|auto|hide|show|protect|merge|insert|add|activate)/.test(key)) {
        return () => object.__proxy;
      }
      return undefined;
    },
  });
}

class VirtualRange {
  constructor(sheet, row, column, numRows, numColumns) {
    Object.assign(this, { sheet, row, column, numRows, numColumns });
    this.__proxy = chainable(this);
    return this.__proxy;
  }
  getRow() { return this.row; }
  getColumn() { return this.column; }
  getNumRows() { return this.numRows; }
  getNumColumns() { return this.numColumns; }
  getValues() {
    const rows = [];
    for (let r = 0; r < this.numRows; r += 1) {
      const row = [];
      for (let c = 0; c < this.numColumns; c += 1) row.push(this.sheet.cell(this.row + r, this.column + c));
      rows.push(row);
    }
    return rows;
  }
  getValue() { return this.sheet.cell(this.row, this.column); }
  getDisplayValues() { return this.getValues().map((row) => row.map((value) => String(value))); }
  getDisplayValue() { return String(this.getValue()); }
  setValues(values) {
    values.forEach((row, r) => row.forEach((value, c) => this.sheet.write(this.row + r, this.column + c, value)));
    return this.__proxy;
  }
  setValue(value) { this.sheet.write(this.row, this.column, value); return this.__proxy; }
  getFormula() { return this.sheet.formulas[`${this.row}:${this.column}`] || ""; }
  getFormulas() { return [[this.getFormula()]]; }
  setFormula(formula) { this.sheet.formulas[`${this.row}:${this.column}`] = formula; return this.__proxy; }
  clearContent() {
    for (let r = 0; r < this.numRows; r += 1)
      for (let c = 0; c < this.numColumns; c += 1) this.sheet.write(this.row + r, this.column + c, "");
    return this.__proxy;
  }
  clear() { return this.clearContent(); }
  sort() { return this.__proxy; }
  createTextFinder(text) {
    const range = this;
    const options = { entire: false, regex: false };
    const finder = {
      matchEntireCell(value) { options.entire = Boolean(value); return finder; },
      useRegularExpression(value) { options.regex = Boolean(value); return finder; },
      matchCase() { return finder; },
      findAll() {
        const pattern = options.regex
          ? new RegExp(options.entire ? `^(?:${text})$` : text)
          : null;
        const found = [];
        range.getValues().forEach((row, r) => row.forEach((value, c) => {
          const cell = String(value ?? "");
          const hit = pattern ? pattern.test(cell) : options.entire ? cell === text : cell.includes(text);
          if (hit) found.push(new VirtualRange(range.sheet, range.row + r, range.column + c, 1, 1));
        }));
        return found;
      },
      findNext() { return finder.findAll()[0] || null; },
    };
    return finder;
  }
}

class VirtualSheet {
  constructor(spreadsheet, name, rows = []) {
    this.spreadsheet = spreadsheet;
    this.name = name;
    this.rows = rows;
    this.formulas = {};
    this.__proxy = chainable(this);
    return this.__proxy;
  }
  getName() { return this.name; }
  setName(name) { this.name = name; return this.__proxy; }
  getSheetId() { return this.spreadsheet.sheets.indexOf(this.__proxy) + 1; }
  cell(row, column) {
    const value = (this.rows[row - 1] || [])[column - 1];
    return value === undefined || value === null ? "" : value;
  }
  write(row, column, value) {
    while (this.rows.length < row) this.rows.push([]);
    const target = this.rows[row - 1];
    while (target.length < column) target.push("");
    target[column - 1] = value instanceof Date ? value : value ?? "";
    this.spreadsheet.changed = true;
  }
  getLastRow() {
    for (let r = this.rows.length; r > 0; r -= 1) {
      if ((this.rows[r - 1] || []).some((value) => value !== "" && value !== null && value !== undefined)) return r;
    }
    return 0;
  }
  getLastColumn() {
    return this.rows.reduce((max, row) => {
      for (let c = row.length; c > max; c -= 1) if (row[c - 1] !== "" && row[c - 1] !== undefined) return c;
      return max;
    }, 0);
  }
  getMaxRows() { return Math.max(1000, this.rows.length); }
  getMaxColumns() { return Math.max(26, this.getLastColumn()); }
  getRange(row, column, numRows = 1, numColumns = 1) {
    if (typeof row === "string") {
      const match = row.match(/^([A-Z]+)(\d+)$/);
      if (!match) throw new Error(`가상 시트가 지원하지 않는 범위입니다: ${row}`);
      const col = match[1].split("").reduce((sum, ch) => sum * 26 + ch.charCodeAt(0) - 64, 0);
      return new VirtualRange(this.__proxy, Number(match[2]), col, 1, 1);
    }
    return new VirtualRange(this.__proxy, row, column, numRows, numColumns);
  }
  getDataRange() {
    return new VirtualRange(this.__proxy, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn()));
  }
  appendRow(values) {
    const row = this.getLastRow() + 1;
    values.forEach((value, index) => this.write(row, index + 1, value));
    return this.__proxy;
  }
  deleteRow(row) { this.rows.splice(row - 1, 1); this.spreadsheet.changed = true; return this.__proxy; }
  clear() { this.rows = []; this.formulas = {}; this.spreadsheet.changed = true; return this.__proxy; }
  clearContents() { return this.clear(); }
  copyTo(spreadsheet) {
    const copy = spreadsheet.insertSheet(`${this.name}의 사본`);
    copy.rows = JSON.parse(JSON.stringify(this.rows));
    return copy;
  }
}

class VirtualSpreadsheet {
  constructor(data = {}) {
    this.sheets = [];
    this.changed = false;
    for (const [name, rows] of Object.entries(data)) this.sheets.push(new VirtualSheet(this, name, rows));
  }
  getSheetByName(name) { return this.sheets.find((sheet) => sheet.getName() === name) || null; }
  getSheets() { return this.sheets.slice(); }
  insertSheet(name) {
    const sheet = new VirtualSheet(this, name || `시트${this.sheets.length + 1}`);
    this.sheets.push(sheet);
    this.changed = true;
    return sheet;
  }
  deleteSheet(sheet) { this.sheets = this.sheets.filter((item) => item !== sheet); this.changed = true; }
  getName() { return "가상 시트(데모)"; }
  getId() { return "virtual-demo-sheet"; }
  toJSON() {
    return Object.fromEntries(this.sheets.map((sheet) => [sheet.getName(), sheet.rows]));
  }
}

function formatDate(value, timeZone, pattern) {
  const date = value instanceof Date ? value : new Date(value);
  const parts = Object.fromEntries(
    new Intl.DateTimeFormat("en-CA", {
      timeZone: timeZone === "UTC" ? "UTC" : "Asia/Seoul",
      year: "numeric", month: "2-digit", day: "2-digit",
      hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23",
    }).formatToParts(date).map((part) => [part.type, part.value])
  );
  return pattern
    .replace("yyyy", parts.year).replace("MM", parts.month).replace("dd", parts.day)
    .replace("HH", parts.hour).replace("mm", parts.minute).replace("ss", parts.second);
}

function createServices(spreadsheet, properties) {
  const cache = new Map();
  const cacheApi = {
    get(key) {
      const item = cache.get(key);
      if (!item) return null;
      if (item.expires < Date.now()) { cache.delete(key); return null; }
      return item.value;
    },
    put(key, value, seconds = 600) { cache.set(key, { value: String(value), expires: Date.now() + seconds * 1000 }); },
    remove(key) { cache.delete(key); },
    removeAll(keys) { (keys || []).forEach((key) => cache.delete(key)); },
    getAll(keys) { return Object.fromEntries((keys || []).map((key) => [key, cacheApi.get(key)]).filter(([, v]) => v !== null)); },
    putAll(values, seconds) { Object.entries(values || {}).forEach(([k, v]) => cacheApi.put(k, v, seconds)); },
  };
  const propertyApi = {
    getProperty(key) { return Object.prototype.hasOwnProperty.call(properties, key) ? properties[key] : null; },
    setProperty(key, value) { properties[key] = String(value); spreadsheet.changed = true; return propertyApi; },
    deleteProperty(key) { delete properties[key]; return propertyApi; },
    getProperties() { return { ...properties }; },
    setProperties(values) { Object.assign(properties, values); return propertyApi; },
  };
  const ui = {
    alert(...args) { console.log("[시트 알림]", args.filter((arg) => typeof arg === "string").join(" ")); return "OK"; },
    prompt() { return { getSelectedButton: () => "CANCEL", getResponseText: () => "" }; },
    createMenu() { const menu = { addItem: () => menu, addSeparator: () => menu, addSubMenu: () => menu, addToUi: () => menu }; return menu; },
    Button: { OK: "OK", CANCEL: "CANCEL", YES: "YES", NO: "NO" },
    ButtonSet: { OK: "OK", OK_CANCEL: "OK_CANCEL", YES_NO: "YES_NO" },
  };
  return {
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet, getActive: () => spreadsheet, getUi: () => ui, flush() {} },
    CacheService: { getScriptCache: () => cacheApi },
    PropertiesService: { getScriptProperties: () => propertyApi },
    LockService: { getScriptLock: () => ({ waitLock() {}, tryLock: () => true, releaseLock() {}, hasLock: () => true }) },
    ContentService: {
      MimeType: { TEXT: "text/plain", JSON: "application/json", JAVASCRIPT: "application/javascript" },
      createTextOutput(text) {
        const output = {
          content: String(text), mimeType: "text/plain",
          setMimeType(type) { output.mimeType = type; return output; },
          getContent() { return output.content; },
        };
        return output;
      },
    },
    Utilities: {
      getUuid: () => crypto.randomUUID(),
      formatDate,
      sleep() {},
      DigestAlgorithm: { SHA_256: "sha256" },
      Charset: { UTF_8: "utf8" },
      computeDigest(algorithm, value) {
        return Array.from(crypto.createHash("sha256").update(String(value), "utf8").digest()).map((b) => (b > 127 ? b - 256 : b));
      },
      parseCsv(text) { return String(text).split(/\r?\n/).filter(Boolean).map((line) => line.split(",")); },
    },
    DriveApp: {
      getFolderById() { throw new Error("가상 시트 데모에서는 드라이브를 쓰지 않습니다."); },
      getFoldersByName() { return { hasNext: () => false }; },
    },
    ScriptApp: { getProjectTriggers: () => [], deleteTrigger() {}, newTrigger: () => ({ timeBased: () => ({ everyDays: () => ({ atHour: () => ({ create() {} }) }) }) }) },
    Session: { getScriptTimeZone: () => "Asia/Seoul" },
    Logger: { log: (...args) => console.log("[Logger]", ...args) },
  };
}

// 가상 시트에서 서버 코드를 실행한다. data: { 시트이름: [[...], ...] }
function createVirtualAppsScript({ sheets = {}, properties = {} } = {}) {
  const spreadsheet = new VirtualSpreadsheet(sheets);
  const props = { ...properties };
  const context = vm.createContext({ ...createServices(spreadsheet, props), console, Date, Math, JSON, Intl });
  const source = fs.readFileSync(SCRIPT_PATH, "utf8");
  vm.runInContext(`${source}\nthis.__constants = { SHEET_NAMES, HEADERS, ADMIN_ACCESS_HASH_PROPERTY, QR_ENTRY_TOKEN_PROPERTY };`, context, {
    filename: "APPS_SCRIPT_READY.gs",
  });
  return { context, spreadsheet, properties: props, constants: context.__constants };
}

module.exports = { createVirtualAppsScript };
