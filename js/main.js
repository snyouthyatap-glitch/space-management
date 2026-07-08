import { GOOGLE_SCRIPT_URL, QR_ENTRY_PARAM, QR_ENTRY_TOKEN } from "./config.js";
import { getTrustedNow, showToast, syncSystemTime, toggleLoading } from "./utils.js";

const STORAGE_KEY = "space-management-session";
const REMEMBERED_MEMBER_KEY = "space-management-remembered-member";
const REMEMBERED_LOUNGE_KEY = "space-management-remembered-lounge";
const LAST_VISIT_LOG_KEY = "space-management-last-visit-log";
const QR_ENTRY_SESSION_KEY = "space-management-qr-entry-ok";

const state = {
    currentMember: null,
    currentMode: null,
    proxyMember: null,
    adminSheetUrl: "",
    statsPeriod: "daily"
};

const sections = {
    gate: document.getElementById("entryGateSection"),
    choice: document.getElementById("choiceSection"),
    facilityEntry: document.getElementById("entrySection"),
    facilitySignup: document.getElementById("signupSection"),
    loungeEntry: document.getElementById("loungeSection"),
    loungeComplete: document.getElementById("loungeCompleteSection"),
    facility: document.getElementById("facilitySection"),
    admin: document.getElementById("adminSection")
};

function normalizeDigits(value) {
    return String(value || "").replace(/\D/g, "");
}

function isLocalPreview() {
    const { protocol, hostname } = window.location;
    return protocol === "file:" || hostname === "127.0.0.1" || hostname === "localhost";
}

function hasValidQrEntryToken() {
    if (sessionStorage.getItem(QR_ENTRY_SESSION_KEY) === "true") {
        return true;
    }
    if (isLocalPreview()) {
        return true;
    }
    const params = new URLSearchParams(window.location.search);
    const isValid = params.get(QR_ENTRY_PARAM) === QR_ENTRY_TOKEN;
    if (isValid) {
        sessionStorage.setItem(QR_ENTRY_SESSION_KEY, "true");
        const cleanUrl = `${window.location.origin}${window.location.pathname}${window.location.hash || ""}`;
        window.history.replaceState({}, document.title, cleanUrl);
    }
    return isValid;
}

function showEntryGate() {
    document.querySelector(".app-shell")?.classList.add("hidden");
    sections.gate?.classList.remove("hidden");
}

function todayString() {
    return getTrustedNow().toLocaleDateString("en-CA");
}

function formatDisplayDate(dateStr = todayString()) {
    return new Date(dateStr).toLocaleDateString("ko-KR", {
        year: "numeric",
        month: "long",
        day: "numeric"
    });
}

function birthDateToAge(birthDate) {
    if (!birthDate) {
        return 0;
    }
    const birth = new Date(birthDate);
    const today = getTrustedNow();
    let age = today.getFullYear() - birth.getFullYear();
    const monthDiff = today.getMonth() - birth.getMonth();
    if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birth.getDate())) {
        age -= 1;
    }
    return age;
}

function setSession(payload) {
    sessionStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
}

function getSession() {
    try {
        return JSON.parse(sessionStorage.getItem(STORAGE_KEY) || "null");
    } catch {
        return null;
    }
}

function clearSession() {
    sessionStorage.removeItem(STORAGE_KEY);
}

function getLastVisitLogRecord() {
    try {
        return JSON.parse(localStorage.getItem(LAST_VISIT_LOG_KEY) || "null");
    } catch {
        return null;
    }
}

function setLastVisitLogRecord(record) {
    localStorage.setItem(LAST_VISIT_LOG_KEY, JSON.stringify(record));
}

function clearExpiredVisitLogRecord() {
    const lastVisit = getLastVisitLogRecord();
    if (!lastVisit) {
        return;
    }
    if (lastVisit.date !== todayString()) {
        localStorage.removeItem(LAST_VISIT_LOG_KEY);
    }
}

function rememberFacilityMember(memberId, verifiedDate = todayString()) {
    localStorage.setItem(REMEMBERED_MEMBER_KEY, JSON.stringify({ memberId, verifiedDate }));
}

function getRememberedFacilityMember() {
    try {
        return JSON.parse(localStorage.getItem(REMEMBERED_MEMBER_KEY) || "null");
    } catch {
        return null;
    }
}

function clearRememberedFacilityMember() {
    localStorage.removeItem(REMEMBERED_MEMBER_KEY);
}

function rememberLoungeGuest(guest) {
    localStorage.setItem(REMEMBERED_LOUNGE_KEY, JSON.stringify(guest));
}

function getRememberedLoungeGuest() {
    try {
        return JSON.parse(localStorage.getItem(REMEMBERED_LOUNGE_KEY) || "null");
    } catch {
        return null;
    }
}

function clearRememberedLoungeGuest() {
    localStorage.removeItem(REMEMBERED_LOUNGE_KEY);
}

function showSection(name) {
    Object.entries(sections).forEach(([key, element]) => {
        if (element) {
            element.classList.toggle("hidden", key !== name);
        }
    });
}

function setAlert(elementId, message = "", type = "error") {
    const element = document.getElementById(elementId);
    if (!element) {
        return;
    }
    if (!message) {
        element.className = "alert hidden";
        element.textContent = "";
        return;
    }
    element.textContent = message;
    element.className = `alert alert-${type}`;
}

function sanitizeTime(value) {
    const digits = normalizeDigits(value);
    if (!digits) {
        return "";
    }

    if (digits.length <= 2) {
        return digits;
    }

    const clipped = digits.slice(0, 4);
    let rawHour = "";
    let rawMinute = "";

    if (clipped.length === 3) {
        rawHour = clipped.slice(0, 1);
        rawMinute = clipped.slice(1, 3);
    } else {
        rawHour = clipped.slice(0, 2);
        rawMinute = clipped.slice(2, 4);
    }

    const hh = Math.min(Number(rawHour || 0), 23).toString().padStart(2, "0");
    const mm = Math.min(Number(rawMinute || 0), 59).toString().padStart(2, "0");
    return `${hh}:${mm}`;
}

function formatTimeInputValue(value) {
    const digits = normalizeDigits(value).slice(0, 4);
    if (digits.length <= 3) {
        return digits;
    }
    return `${digits.slice(0, 2)}:${digits.slice(2, 4)}`;
}

function createFacilityMemberPayload(base) {
    const phone = normalizeDigits(base.phone);
    return {
        name: String(base.name || "").trim(),
        gender: String(base.gender || "").trim(),
        age: Number(base.age || 0),
        phone,
        phoneLastDigits: phone.slice(-4),
        isSeongnamResident: Boolean(base.isSeongnamResident),
        role: "user",
        status: "approved"
    };
}

function createLoungeGuestPayload(base) {
    const birthDate = String(base.birthDate || "").trim();
    return {
        gender: String(base.gender || "").trim(),
        birthDate,
        age: birthDateToAge(birthDate),
        isSeongnamResident: Boolean(base.isSeongnamResident),
        role: "lounge_guest"
    };
}

function collectCompanions(containerId) {
    const rows = document.querySelectorAll(`#${containerId} .companion-item`);
    return Array.from(rows)
        .map((row) => ({
            gender: row.querySelector(".comp-gender")?.value || "",
            age: Number(row.querySelector(".comp-age")?.value || 0)
        }))
        .filter((item) => item.gender && item.age);
}

function buildCompanionFields(containerId, count) {
    const container = document.getElementById(containerId);
    if (!container) {
        return;
    }

    container.innerHTML = "";
    for (let index = 1; index <= count; index += 1) {
        const row = document.createElement("div");
        row.className = "companion-item";
        row.innerHTML = `
            <span class="companion-label">추가 인원 ${index}</span>
            <select class="form-control comp-gender" required>
                <option value="">성별</option>
                <option value="남성">남성</option>
                <option value="여성">여성</option>
            </select>
            <input type="number" class="form-control comp-age" placeholder="나이" min="1" max="100" required>
        `;
        container.appendChild(row);
    }
}

function setupStepper(countInputId, containerId) {
    const input = document.getElementById(countInputId);
    if (!input) {
        return;
    }

    const refresh = () => {
        buildCompanionFields(containerId, Number(input.value || 0));
    };

    input.closest(".stepper-container")?.querySelectorAll(".stepper-btn").forEach((button) => {
        button.addEventListener("click", () => {
            const current = Number(input.value || 0);
            const delta = button.classList.contains("plus") ? 1 : -1;
            input.value = String(Math.max(0, Math.min(10, current + delta)));
            refresh();
        });
    });
}

function getFormCompanions(countInputId, containerId) {
    const count = Number(document.getElementById(countInputId).value || 0);
    const companions = collectCompanions(containerId);
    if (companions.length !== count) {
        throw new Error("추가 인원 정보를 모두 입력해 주세요.");
    }
    return companions;
}

function updateFacilityHeader(member) {
    const welcome = document.getElementById("facilityWelcome");
    const info = document.getElementById("facilityMemberInfo");
    const today = document.getElementById("todayLabel");

    if (welcome) {
        welcome.textContent = `${member.name}님 시설 이용 안내`;
    }
    if (info) {
        info.textContent = `${member.gender} · ${member.age}세 · 연락처 끝자리 ${member.phoneLastDigits || "-"}`;
    }
    if (today) {
        today.textContent = formatDisplayDate();
    }
}

function setCurrentMember(member, mode) {
    state.currentMember = member;
    state.currentMode = mode;
    if (mode === "facility") {
        updateFacilityHeader(member);
    }
}

function showLoungeComplete(guest, alreadySubmitted = false) {
    const title = document.getElementById("loungeCompleteTitle");
    const text = document.getElementById("loungeCompleteText");
    if (title) {
        title.textContent = alreadySubmitted
            ? "오늘 라운지 방문이 이미 확인되었습니다."
            : "라운지 방문이 기록되었습니다.";
    }
    if (text) {
        text.textContent = alreadySubmitted
            ? `${guest.gender} · ${guest.age}세 기준으로 오늘 방문 기록이 이미 있습니다.`
            : `${guest.gender} · ${guest.age}세 기준으로 오늘 방문 기록을 저장했습니다.`;
    }
    showSection("loungeComplete");
}

function updateProxyFieldVisibility() {
    const type = document.querySelector('input[name="proxyType"]:checked')?.value || "visit";
    document.getElementById("proxyPrinterFields")?.classList.toggle("hidden", type !== "printer");
    document.getElementById("proxyRoomFields")?.classList.toggle("hidden", !(type === "careerZone" || type === "connectRoom"));
}

function updateManualFieldVisibility() {
    const type = document.querySelector('input[name="manualType"]:checked')?.value || "visit";
    document.getElementById("manualPrinterFields")?.classList.toggle("hidden", type !== "printer");
    document.getElementById("manualRoomFields")?.classList.toggle("hidden", !(type === "careerZone" || type === "connectRoom"));
}

function setupTabs() {
    document.querySelectorAll(".tab").forEach((tab) => {
        tab.addEventListener("click", () => {
            const targetId = tab.dataset.target;
            const wrapper = tab.closest(".card");
            wrapper?.querySelectorAll(".tab").forEach((item) => item.classList.remove("active"));
            wrapper?.querySelectorAll(".tab-content").forEach((panel) => panel.classList.add("hidden"));
            tab.classList.add("active");
            document.getElementById(targetId)?.classList.remove("hidden");
        });
    });

    document.querySelectorAll(".admin-tab").forEach((tab) => {
        tab.addEventListener("click", () => {
            const targetId = tab.dataset.target;
            document.querySelectorAll(".admin-tab").forEach((item) => item.classList.remove("active"));
            document.querySelectorAll(".admin-tab-content").forEach((panel) => panel.classList.add("hidden"));
            tab.classList.add("active");
            document.getElementById(targetId)?.classList.remove("hidden");
        });
    });

    document.querySelectorAll(".period-btn").forEach((button) => {
        button.addEventListener("click", async () => {
            document.querySelectorAll(".period-btn").forEach((item) => item.classList.remove("active"));
            button.classList.add("active");
            await loadStats(button.dataset.period);
        });
    });
}

function setupTimeInputs() {
    document.querySelectorAll(".time-input").forEach((input) => {
        input.addEventListener("input", (event) => {
            const digits = normalizeDigits(event.target.value);
            if (digits.length <= 2) {
                event.target.value = digits;
                return;
            }
            event.target.value = formatTimeInputValue(digits);
        });

        input.addEventListener("blur", (event) => {
            event.target.value = sanitizeTime(event.target.value);
        });
    });
}

function setupNetworkWarning() {
    const warning = document.getElementById("onlineWarning");
    const update = () => warning?.classList.toggle("hidden", navigator.onLine);
    update();
    window.addEventListener("online", update);
    window.addEventListener("offline", update);
}

function setChoiceLoading(buttonId, isLoading) {
    const button = document.getElementById(buttonId);
    if (!button) {
        return;
    }
    button.disabled = isLoading;
    button.classList.toggle("loading", isLoading);
}

function setPageLoading(isLoading, message = "불러오는 중...") {
    let overlay = document.getElementById("pageLoadingOverlay");
    if (!overlay) {
        overlay = document.createElement("div");
        overlay.id = "pageLoadingOverlay";
        overlay.className = "page-loading-overlay hidden";
        overlay.innerHTML = `
            <div class="page-loading-card">
                <div class="page-loading-spinner"></div>
                <strong class="page-loading-title">불러오는 중...</strong>
                <span class="page-loading-text">잠시만 기다려주세요</span>
            </div>
        `;
        document.body.appendChild(overlay);
    }

    const title = overlay.querySelector(".page-loading-title");
    if (title) {
        title.textContent = message;
    }

    overlay.classList.toggle("hidden", !isLoading);
}

async function withPageLoading(buttonId, message, task) {
    setChoiceLoading(buttonId, true);
    setPageLoading(true, message);
    await new Promise((resolve) => setTimeout(resolve, 30));
    try {
        const result = await task();
        await new Promise((resolve) => setTimeout(resolve, 180));
        return result;
    } finally {
        setChoiceLoading(buttonId, false);
        setPageLoading(false);
    }
}

function fillSignupPhone(value) {
    const phoneInput = document.getElementById("signupPhone");
    if (phoneInput) {
        phoneInput.value = normalizeDigits(value);
    }
}

function resetToChoice(clearRememberedFacility = false) {
    clearSession();
    if (clearRememberedFacility) {
        clearRememberedFacilityMember();
    }
    state.currentMember = null;
    state.currentMode = null;
    state.proxyMember = null;
    document.getElementById("entryForm")?.reset();
    document.getElementById("signupForm")?.reset();
    document.getElementById("loungeForm")?.reset();
    document.getElementById("proxyEntryArea")?.classList.add("hidden");
    setAlert("entryAlert");
    setAlert("signupAlert");
    setAlert("loungeAlert");
    setAlert("facilityAlert");
    setAlert("adminAlert");
    showSection("choice");
}

function getCurrentMemberOrThrow() {
    if (!state.currentMember || state.currentMode !== "facility") {
        throw new Error("시설 이용 회원 정보를 찾을 수 없습니다. 처음 화면에서 다시 시작해 주세요.");
    }
    return state.currentMember;
}

function openProxyEntry(member) {
    state.proxyMember = member;
    const title = document.getElementById("proxyTargetName");
    if (title) {
        title.textContent = `${member.name} 회원 대리 제출`;
    }
    const dateInput = document.getElementById("proxyDateInput");
    if (dateInput) {
        dateInput.value = todayString();
    }
    document.getElementById("proxyEntryArea")?.classList.remove("hidden");
    const countInput = document.getElementById("proxyCompanionCount");
    if (countInput) {
        countInput.value = 0;
    }
    const companions = document.getElementById("proxyCompanions");
    if (companions) {
        companions.innerHTML = "";
    }
    updateProxyFieldVisibility();
}

function toApiMember(member) {
    return {
        id: member.id || "",
        name: member.name || "",
        gender: member.gender || "",
        age: Number(member.age || 0),
        phone: member.phone || "",
        phoneLastDigits: member.phoneLastDigits || "",
        role: member.role || "user"
    };
}

function callScript(action, payload = {}) {
    if (!GOOGLE_SCRIPT_URL) {
        return Promise.reject(new Error("Apps Script 주소가 설정되지 않았습니다."));
    }

    return new Promise((resolve, reject) => {
        const callbackName = `spaceManagementCallback_${Date.now()}_${Math.floor(Math.random() * 100000)}`;
        const script = document.createElement("script");
        const timeoutId = window.setTimeout(() => {
            cleanup();
            reject(new Error("요청 시간이 초과되었습니다."));
        }, 10000);

        function cleanup() {
            window.clearTimeout(timeoutId);
            delete window[callbackName];
            script.remove();
        }

        window[callbackName] = (response) => {
            cleanup();
            if (!response || response.ok === false) {
                reject(new Error(response?.error || "요청 처리 중 문제가 발생했습니다."));
                return;
            }
            resolve(response);
        };

        script.onerror = () => {
            cleanup();
            reject(new Error("Apps Script 연결에 실패했습니다."));
        };

        const url = new URL(GOOGLE_SCRIPT_URL);
        url.searchParams.set("action", action);
        url.searchParams.set("payload", JSON.stringify(payload));
        url.searchParams.set("callback", callbackName);
        script.src = url.toString();
        document.body.appendChild(script);
    });
}

async function loadAppSettings() {
    const response = await callScript("bootstrap");
    state.adminSheetUrl = response.adminSheetUrl || "";
    const sheetLink = document.getElementById("adminSheetLink");
    if (sheetLink && state.adminSheetUrl) {
        sheetLink.href = state.adminSheetUrl;
        sheetLink.classList.remove("hidden");
    }
}

async function ensureVisitLog(member, dateStr, source = "system", options = {}) {
    const lastVisit = getLastVisitLogRecord();
    if (lastVisit?.memberId === member.id && lastVisit?.date === dateStr) {
        return false;
    }

    const response = await callScript("submitVisitLog", {
        member: toApiMember(member),
        date: dateStr,
        source,
        createdBy: options.createdBy || "self",
        allowDuplicate: Boolean(options.allowDuplicate)
    });
    setLastVisitLogRecord({
        memberId: member.id,
        date: dateStr
    });
    return Boolean(response.created);
}

async function routeFacilityMember(member, source = "facility-qr") {
    setCurrentMember(member, "facility");
    await ensureVisitLog(member, todayString(), source);
    setSession({ mode: "facility", memberId: member.id });
    rememberFacilityMember(member.id, todayString());
    showSection("facility");
}

async function handleFacilityEntry(inputValue) {
    setAlert("entryAlert");
    const rawValue = String(inputValue || "").trim();

    if (!rawValue) {
        setAlert("entryAlert", "입력값을 확인해 주세요.");
        return;
    }

    const response = await callScript("resolveFacilityEntry", {
        inputValue: rawValue
    });

    if (response.mode === "admin") {
        if (response.adminSheetUrl) {
            state.adminSheetUrl = response.adminSheetUrl;
            const sheetLink = document.getElementById("adminSheetLink");
            if (sheetLink) {
                sheetLink.href = state.adminSheetUrl;
                sheetLink.classList.remove("hidden");
            }
        }
        await enterAdminMode();
        return;
    }

    if (response.mode === "ambiguous") {
        setAlert("entryAlert", response.message || "전체 전화번호를 다시 입력해 주세요.");
        return;
    }

    if (response.mode === "signup") {
        fillSignupPhone(response.suggestedPhone || rawValue);
        showSection("facilitySignup");
        return;
    }

    if (response.mode === "member" && response.member) {
        await routeFacilityMember(response.member, "facility-qr");
        return;
    }

    throw new Error("입력값을 처리하지 못했습니다.");
}

async function registerFacilityMember() {
    setAlert("signupAlert");
    const payload = createFacilityMemberPayload({
        name: document.getElementById("signupName")?.value,
        gender: document.getElementById("signupGender")?.value,
        age: document.getElementById("signupAge")?.value,
        phone: document.getElementById("signupPhone")?.value,
        isSeongnamResident: document.getElementById("signupSeongnamResident")?.checked
    });

    if (!payload.name || !payload.gender || !payload.age || !payload.phone) {
        setAlert("signupAlert", "입력 항목을 모두 확인해 주세요.");
        return;
    }

    const response = await callScript("registerFacilityMember", {
        member: payload
    });

    showToast(response.existing ? "기존 회원 정보를 불러왔습니다." : "회원 등록과 방문 확인이 완료되었습니다.", "success");
    await routeFacilityMember(response.member, response.existing ? "facility-signup-existing" : "facility-signup");
}

async function handleLoungeEntry() {
    const remembered = getRememberedLoungeGuest();
    if (!remembered || remembered.validDate !== todayString()) {
        clearRememberedLoungeGuest();
        document.getElementById("loungeForm")?.reset();
        showSection("loungeEntry");
        return;
    }

    const response = await callScript("getLoungeGuestById", {
        guestId: remembered.guestId
    });

    if (!response.guest || response.guest.validDate !== todayString()) {
        clearRememberedLoungeGuest();
        document.getElementById("loungeForm")?.reset();
        showSection("loungeEntry");
        return;
    }

    await ensureVisitLog(response.guest, todayString(), "lounge-remembered");
    showLoungeComplete(response.guest, true);
}

async function registerLoungeGuest() {
    setAlert("loungeAlert");
    const payload = createLoungeGuestPayload({
        gender: document.getElementById("loungeGender")?.value,
        birthDate: document.getElementById("loungeBirthDate")?.value,
        isSeongnamResident: document.getElementById("loungeSeongnamResident")?.checked
    });

    if (!payload.gender || !payload.birthDate || payload.age < 0) {
        setAlert("loungeAlert", "성별과 생년월일을 확인해 주세요.");
        return;
    }

    const existingRemembered = getRememberedLoungeGuest();
    if (existingRemembered?.validDate === todayString()) {
        const existingResponse = await callScript("getLoungeGuestById", {
            guestId: existingRemembered.guestId
        });
        if (existingResponse.guest?.validDate === todayString()) {
            await ensureVisitLog(existingResponse.guest, todayString(), "lounge-repeat");
            showLoungeComplete(existingResponse.guest, true);
            return;
        }
        clearRememberedLoungeGuest();
    }

    const response = await callScript("registerLoungeGuest", {
        guest: payload
    });
    const guest = response.guest;
    await ensureVisitLog(guest, todayString(), "lounge");
    rememberLoungeGuest({
        guestId: guest.id,
        validDate: todayString(),
        gender: guest.gender,
        birthDate: guest.birthDate,
        age: guest.age
    });
    showToast("라운지 방문이 기록되었습니다.", "success");
    showLoungeComplete(guest, false);
}

async function submitUsageRecord(type, data, member, options = {}) {
    await callScript("submitUsageRecord", {
        type,
        data: {
            ...data,
            date: data.date || todayString(),
            startTime: data.startTime ? sanitizeTime(data.startTime) : "",
            endTime: data.endTime ? sanitizeTime(data.endTime) : ""
        },
        member: toApiMember(member),
        createdBy: options.createdBy || "self"
    });
}

async function enterAdminMode() {
    showSection("admin");
    setSession({ mode: "admin" });
    state.proxyMember = null;
    await loadStats(state.statsPeriod);
}

async function searchMembers() {
    const term = String(document.getElementById("adminMemberSearchInput")?.value || "").trim();
    const resultsEl = document.getElementById("adminMemberSearchResults");
    if (!resultsEl) {
        return;
    }

    if (!term) {
        resultsEl.innerHTML = "";
        return;
    }

    const response = await callScript("searchMembers", { term });
    const matches = response.members || [];

    if (matches.length === 0) {
        resultsEl.innerHTML = '<div class="empty-mini">검색 결과가 없습니다.</div>';
        return;
    }

    resultsEl.innerHTML = matches.map((member) => `
        <button type="button" class="result-card" data-id="${member.id}">
            <strong>${member.name}</strong>
            <span>${member.gender} · ${member.age}세</span>
            <span>${member.phone || "-"}</span>
            <span class="result-status approved">등록 회원</span>
        </button>
    `).join("");

    resultsEl.querySelectorAll(".result-card").forEach((button) => {
        button.addEventListener("click", () => {
            const member = matches.find((item) => item.id === button.dataset.id);
            if (member) {
                openProxyEntry(member);
            }
        });
    });
}

async function handleProxySubmit(event) {
    event.preventDefault();
    if (!state.proxyMember) {
        return;
    }

    setAlert("adminAlert");

    const type = document.querySelector('input[name="proxyType"]:checked')?.value || "visit";
    const date = document.getElementById("proxyDateInput")?.value || todayString();

    if (type === "visit") {
        await ensureVisitLog(state.proxyMember, date, "admin-proxy", { createdBy: "admin-proxy" });
    } else if (type === "printer") {
        const count = Number(document.getElementById("proxyPrinterCount")?.value || 0);
        if (!count) {
            throw new Error("프린터 사용 매수를 입력해 주세요.");
        }
        await submitUsageRecord("printer", {
            date,
            count
        }, state.proxyMember, { createdBy: "admin-proxy" });
    } else {
        const startTime = document.getElementById("proxyStartTime")?.value;
        const endTime = document.getElementById("proxyEndTime")?.value;
        const purpose = document.getElementById("proxyPurpose")?.value;
        if (!startTime || !endTime || !String(purpose || "").trim()) {
            throw new Error("시작 시간, 종료 시간, 사용목적을 모두 입력해 주세요.");
        }
        const companions = getFormCompanions("proxyCompanionCount", "proxyCompanions");
        await submitUsageRecord(type, {
            date,
            startTime,
            endTime,
            purpose,
            companions
        }, state.proxyMember, { createdBy: "admin-proxy" });
    }

    showToast("대리 제출이 완료되었습니다.", "success");
    document.getElementById("proxyFacilityForm")?.reset();
    document.getElementById("proxyEntryArea")?.classList.add("hidden");
    state.proxyMember = null;
}

async function handleManualSubmit(event) {
    event.preventDefault();
    setAlert("adminAlert");
    const type = document.querySelector('input[name="manualType"]:checked')?.value || "visit";
    const manualMember = createFacilityMemberPayload({
        name: document.getElementById("manualName")?.value,
        gender: document.getElementById("manualGender")?.value,
        age: document.getElementById("manualAge")?.value,
        phone: document.getElementById("manualPhone")?.value
    });

    if (!manualMember.name || !manualMember.gender || !manualMember.age) {
        throw new Error("수동 제출 대상 정보를 확인해 주세요.");
    }

    const pseudoMember = {
        ...manualMember,
        id: "",
        role: "manual"
    };
    const date = document.getElementById("manualDate")?.value || todayString();

    if (type === "visit") {
        await ensureVisitLog(pseudoMember, date, "admin-manual", {
            createdBy: "admin-manual",
            allowDuplicate: true
        });
    } else if (type === "printer") {
        const count = Number(document.getElementById("manualPrinterCount")?.value || 0);
        if (!count) {
            throw new Error("프린터 사용 매수를 입력해 주세요.");
        }
        await submitUsageRecord("printer", {
            date,
            count
        }, pseudoMember, { createdBy: "admin-manual" });
    } else {
        const startTime = document.getElementById("manualStartTime")?.value;
        const endTime = document.getElementById("manualEndTime")?.value;
        const purpose = document.getElementById("manualPurpose")?.value;
        if (!startTime || !endTime || !String(purpose || "").trim()) {
            throw new Error("시작 시간, 종료 시간, 사용목적을 모두 입력해 주세요.");
        }
        const companions = getFormCompanions("manualCompanionCount", "manualCompanions");
        await submitUsageRecord(type, {
            date,
            startTime,
            endTime,
            purpose,
            companions
        }, pseudoMember, { createdBy: "admin-manual" });
    }

    showToast("수동 일지 제출이 완료되었습니다.", "success");
    document.getElementById("manualEntryForm")?.reset();
    document.getElementById("manualCompanionCount").value = 0;
    document.getElementById("manualCompanions").innerHTML = "";
    updateManualFieldVisibility();
}

async function loadStats(period) {
    state.statsPeriod = period;
    const response = await callScript("getStats", { period });
    document.getElementById("statTotalVisit").textContent = `${response.totals.visit}`;
    document.getElementById("statTotalPrint").textContent = `${response.totals.printer}`;
    document.getElementById("statTotalCareer").textContent = `${response.totals.careerZone}`;
    document.getElementById("statTotalConnect").textContent = `${response.totals.connectRoom}`;
    document.getElementById("statsRangeLabel").textContent = `${response.range.start} ~ ${response.range.end}`;
}

async function restoreSession() {
    const session = getSession();

    if (session?.mode === "admin") {
        await enterAdminMode();
        return;
    }

    const rememberedMember = session?.mode === "facility" && session.memberId
        ? { memberId: session.memberId, verifiedDate: todayString() }
        : getRememberedFacilityMember();

    if (!rememberedMember?.memberId) {
        return;
    }

    if (rememberedMember.verifiedDate !== todayString()) {
        clearSession();
        clearRememberedFacilityMember();
        return;
    }

    try {
        const response = await callScript("getFacilityMemberById", {
            memberId: rememberedMember.memberId
        });

        if (!response.member) {
            clearSession();
            clearRememberedFacilityMember();
            return;
        }

        await routeFacilityMember(response.member, "facility-remembered");
    } catch (error) {
        console.warn("Failed to restore facility session:", error);
        clearSession();
        clearRememberedFacilityMember();
    }
}

function setupListeners() {
    document.getElementById("loungeEntryBtn")?.addEventListener("click", async () => {
        try {
            await withPageLoading("loungeEntryBtn", "라운지 확인 중...", async () => {
                await handleLoungeEntry();
            });
        } catch (error) {
            console.error(error);
            showToast(error.message || "라운지 화면을 여는 중 문제가 발생했습니다.", "error");
        }
    });

    document.getElementById("facilityEntryBtn")?.addEventListener("click", async () => {
        await withPageLoading("facilityEntryBtn", "시설 화면 여는 중...", async () => {
            showSection("facilityEntry");
        });
    });
    document.getElementById("entryBackBtn")?.addEventListener("click", () => resetToChoice());
    document.getElementById("signupBackBtn")?.addEventListener("click", () => resetToChoice());
    document.getElementById("loungeBackBtn")?.addEventListener("click", () => resetToChoice());
    document.getElementById("loungeCompleteHomeBtn")?.addEventListener("click", () => resetToChoice());
    document.getElementById("loungeToFacilityBtn")?.addEventListener("click", () => showSection("facilityEntry"));
    document.getElementById("facilityExitBtn")?.addEventListener("click", () => resetToChoice());
    document.getElementById("facilityLogoutBtn")?.addEventListener("click", () => resetToChoice(true));
    document.getElementById("adminExitBtn")?.addEventListener("click", () => resetToChoice());

    document.getElementById("entryForm")?.addEventListener("submit", async (event) => {
        event.preventDefault();
        toggleLoading("entrySubmitBtn", true);
        try {
            await handleFacilityEntry(document.getElementById("entryCode")?.value);
        } catch (error) {
            console.error(error);
            setAlert("entryAlert", error.message || "조회 중 문제가 발생했습니다.");
        } finally {
            toggleLoading("entrySubmitBtn", false);
        }
    });

    document.getElementById("signupForm")?.addEventListener("submit", async (event) => {
        event.preventDefault();
        toggleLoading("signupSubmitBtn", true);
        try {
            await registerFacilityMember();
        } catch (error) {
            console.error(error);
            setAlert("signupAlert", error.message || "회원 등록 중 문제가 발생했습니다.");
        } finally {
            toggleLoading("signupSubmitBtn", false);
        }
    });

    document.getElementById("loungeForm")?.addEventListener("submit", async (event) => {
        event.preventDefault();
        toggleLoading("loungeSubmitBtn", true);
        try {
            await registerLoungeGuest();
        } catch (error) {
            console.error(error);
            setAlert("loungeAlert", error.message || "라운지 방문 처리 중 문제가 발생했습니다.");
        } finally {
            toggleLoading("loungeSubmitBtn", false);
        }
    });

    document.getElementById("refreshStatsBtn")?.addEventListener("click", () => loadStats(state.statsPeriod));
    document.getElementById("adminMemberSearchBtn")?.addEventListener("click", () => searchMembers());
    document.getElementById("closeProxyBtn")?.addEventListener("click", () => {
        document.getElementById("proxyEntryArea")?.classList.add("hidden");
        state.proxyMember = null;
    });

    document.querySelectorAll('input[name="proxyType"]').forEach((input) => {
        input.addEventListener("change", updateProxyFieldVisibility);
    });

    document.querySelectorAll('input[name="manualType"]').forEach((input) => {
        input.addEventListener("change", updateManualFieldVisibility);
    });

    document.getElementById("proxyFacilityForm")?.addEventListener("submit", async (event) => {
        event.preventDefault();
        toggleLoading("proxySubmitBtn", true);
        try {
            await handleProxySubmit(event);
        } catch (error) {
            console.error(error);
            setAlert("adminAlert", error.message || "대리 제출 중 문제가 발생했습니다.");
        } finally {
            toggleLoading("proxySubmitBtn", false);
        }
    });

    document.getElementById("manualEntryForm")?.addEventListener("submit", async (event) => {
        event.preventDefault();
        toggleLoading("manualSubmitBtn", true);
        try {
            await handleManualSubmit(event);
        } catch (error) {
            console.error(error);
            setAlert("adminAlert", error.message || "수동 제출 중 문제가 발생했습니다.");
        } finally {
            toggleLoading("manualSubmitBtn", false);
        }
    });

    document.getElementById("printerForm")?.addEventListener("submit", async (event) => {
        event.preventDefault();
        toggleLoading("printerSubmitBtn", true);
        try {
            const member = getCurrentMemberOrThrow();
            await submitUsageRecord("printer", {
                count: Number(document.getElementById("printerCount")?.value || 0)
            }, member);
            showToast("프린터 사용 기록을 제출했습니다.", "success");
            event.target.reset();
            document.getElementById("printerCount").value = 1;
        } catch (error) {
            console.error(error);
            setAlert("facilityAlert", error.message || "프린터 기록 제출 중 문제가 발생했습니다.");
        } finally {
            toggleLoading("printerSubmitBtn", false);
        }
    });

    document.getElementById("careerForm")?.addEventListener("submit", async (event) => {
        event.preventDefault();
        toggleLoading("careerSubmitBtn", true);
        try {
            const member = getCurrentMemberOrThrow();
            const companions = getFormCompanions("careerCompanionCount", "careerCompanions");
            await submitUsageRecord("careerZone", {
                startTime: document.getElementById("careerStartTime")?.value,
                endTime: document.getElementById("careerEndTime")?.value,
                purpose: document.getElementById("careerPurpose")?.value,
                companions
            }, member);
            showToast("커리어존 이용 기록을 제출했습니다.", "success");
            event.target.reset();
            document.getElementById("careerCompanionCount").value = 0;
            document.getElementById("careerCompanions").innerHTML = "";
        } catch (error) {
            console.error(error);
            setAlert("facilityAlert", error.message || "커리어존 기록 제출 중 문제가 발생했습니다.");
        } finally {
            toggleLoading("careerSubmitBtn", false);
        }
    });

    document.getElementById("connectForm")?.addEventListener("submit", async (event) => {
        event.preventDefault();
        toggleLoading("connectSubmitBtn", true);
        try {
            const member = getCurrentMemberOrThrow();
            const companions = getFormCompanions("connectCompanionCount", "connectCompanions");
            await submitUsageRecord("connectRoom", {
                startTime: document.getElementById("connectStartTime")?.value,
                endTime: document.getElementById("connectEndTime")?.value,
                purpose: document.getElementById("connectPurpose")?.value,
                companions
            }, member);
            showToast("커넥트존 이용 기록을 제출했습니다.", "success");
            event.target.reset();
            document.getElementById("connectCompanionCount").value = 0;
            document.getElementById("connectCompanions").innerHTML = "";
        } catch (error) {
            console.error(error);
            setAlert("facilityAlert", error.message || "커넥트존 기록 제출 중 문제가 발생했습니다.");
        } finally {
            toggleLoading("connectSubmitBtn", false);
        }
    });
}

async function init() {
    if (!hasValidQrEntryToken()) {
        showEntryGate();
        return;
    }

    await syncSystemTime();
    clearExpiredVisitLogRecord();
    await loadAppSettings();
    setupTabs();
    setupTimeInputs();
    setupNetworkWarning();
    setupStepper("careerCompanionCount", "careerCompanions");
    setupStepper("connectCompanionCount", "connectCompanions");
    setupStepper("proxyCompanionCount", "proxyCompanions");
    setupStepper("manualCompanionCount", "manualCompanions");
    setupListeners();
    updateProxyFieldVisibility();
    updateManualFieldVisibility();
    document.getElementById("proxyDateInput").value = todayString();
    document.getElementById("manualDate").value = todayString();
    await restoreSession();
}

init().catch((error) => {
    console.error("Initialization failed:", error);
    showToast("초기화 중 문제가 발생했습니다. 설정을 확인해 주세요.", "error");
});
