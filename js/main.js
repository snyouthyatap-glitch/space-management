import { GOOGLE_SCRIPT_URL, QR_ENTRY_PARAM } from "./config.js?v=20260724-3";
import { showToast, toggleLoading } from "./utils.js?v=20260724-3";

const QR_ENTRY_SESSION_KEY = "space-management-qr-entry-ok";
const PENDING_PRINTER_SUBMISSION_KEY = "space-management-pending-printer-submission";
const SCRIPT_REQUEST_TIMEOUT_MS = 25000;
let inMemoryEntryToken = "";

const state = {
    currentMember: null,
    currentMode: null,
    pendingMember: null,
    candidateMember: null,
    candidateSource: "facility-qr",
    verifiedSignupPhone: "",
    todayReservations: [],
    selectedReservationId: "",
    pendingPrinterSubmission: null
};

const sections = {
    gate: document.getElementById("entryGateSection"),
    facilityEntry: document.getElementById("entrySection"),
    memberConfirm: document.getElementById("memberConfirmSection"),
    facilitySignup: document.getElementById("signupSection"),
    birthUpdate: document.getElementById("birthUpdateSection"),
    facility: document.getElementById("facilitySection")
};

function normalizeDigits(value) {
    return String(value || "").replace(/\D/g, "");
}

function maskMemberName(name) {
    const value = String(name || "회원").trim();
    if (value.length <= 1) return value;
    if (value.length === 2) return `${value[0]}○`;
    return `${value[0]}${"○".repeat(value.length - 2)}${value[value.length - 1]}`;
}

function isLocalPreview() {
    const { protocol, hostname } = window.location;
    return protocol === "file:" || hostname === "127.0.0.1" || hostname === "localhost";
}

function hasValidQrEntryToken() {
    if (getEntryToken()) {
        return true;
    }
    const params = new URLSearchParams(window.location.search);
    const entryToken = String(params.get(QR_ENTRY_PARAM) || "").trim();
    if (entryToken) {
        setSessionValue(QR_ENTRY_SESSION_KEY, entryToken);
        const cleanUrl = `${window.location.origin}${window.location.pathname}${window.location.hash || ""}`;
        window.history.replaceState({}, document.title, cleanUrl);
    }
    return Boolean(entryToken) || isLocalPreview();
}

function getEntryToken() {
    return getSessionValue(QR_ENTRY_SESSION_KEY) || inMemoryEntryToken;
}

function getSessionValue(key) {
    try {
        return sessionStorage.getItem(key) || "";
    } catch {
        return "";
    }
}

function setSessionValue(key, value) {
    if (key === QR_ENTRY_SESSION_KEY) inMemoryEntryToken = value;
    try {
        sessionStorage.setItem(key, value);
    } catch {
        // Restricted storage must not block the current QR visit.
    }
}

function removeSessionValue(key) {
    try {
        sessionStorage.removeItem(key);
    } catch {
        // In-memory state still prevents duplicate requests during this page session.
    }
}

function createRequestId() {
    if (window.crypto?.randomUUID) {
        return window.crypto.randomUUID();
    }
    return `request-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function getPrinterSubmission(memberId, count) {
    const signature = `${memberId}:${todayString()}:${count}`;
    let pending = state.pendingPrinterSubmission;
    if (!pending) {
        try {
            pending = JSON.parse(getSessionValue(PENDING_PRINTER_SUBMISSION_KEY) || "null");
        } catch {
            pending = null;
        }
    }
    if (pending?.signature === signature && pending.requestId) {
        state.pendingPrinterSubmission = pending;
        return pending;
    }

    pending = { signature, requestId: createRequestId() };
    state.pendingPrinterSubmission = pending;
    setSessionValue(PENDING_PRINTER_SUBMISSION_KEY, JSON.stringify(pending));
    return pending;
}

function clearPrinterSubmission(requestId) {
    if (state.pendingPrinterSubmission?.requestId !== requestId) return;
    state.pendingPrinterSubmission = null;
    removeSessionValue(PENDING_PRINTER_SUBMISSION_KEY);
}

function showEntryGate() {
    document.querySelector(".app-shell")?.classList.add("hidden");
    sections.gate?.classList.remove("hidden");
}

function getSeoulDateParts() {
    const parts = new Intl.DateTimeFormat("en-US", {
        timeZone: "Asia/Seoul",
        year: "numeric",
        month: "2-digit",
        day: "2-digit"
    }).formatToParts(new Date());
    const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
    return {
        year: Number(values.year),
        month: Number(values.month),
        day: Number(values.day)
    };
}

function todayString() {
    const { year, month, day } = getSeoulDateParts();
    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function formatDisplayDate(dateStr = todayString()) {
    const [year, month, day] = dateStr.split("-").map(Number);
    return `${year}년 ${month}월 ${day}일`;
}

function birthDateToAge(birthDate) {
    if (!birthDate) {
        return 0;
    }
    const [birthYear, birthMonth, birthDay] = birthDate.split("-").map(Number);
    const today = getSeoulDateParts();
    let age = today.year - birthYear;
    if (today.month < birthMonth || (today.month === birthMonth && today.day < birthDay)) {
        age -= 1;
    }
    return age;
}

function parseBirthDateInput(value) {
    const digits = normalizeDigits(value);
    if (digits.length !== 8) {
        return "";
    }

    const year = Number(digits.slice(0, 4));
    const month = Number(digits.slice(4, 6));
    const day = Number(digits.slice(6, 8));

    if (year < 1900 || year > 2100 || month < 1 || month > 12 || day < 1 || day > 31) {
        return "";
    }

    const birthDate = `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    const date = new Date(Date.UTC(year, month - 1, day));
    if (date.getUTCFullYear() !== year || date.getUTCMonth() + 1 !== month || date.getUTCDate() !== day) {
        return "";
    }

    return birthDate;
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
    if (element) {
        element.className = "alert hidden";
        element.textContent = "";
    }
    if (message) {
        showToast(message, type);
    }
}

function sanitizeTime(value) {
    const digits = normalizeDigits(value);
    if (!digits) {
        return "";
    }

    if (digits.length < 3 || digits.length > 4) return "";
    const padded = digits.padStart(4, "0");
    const hour = Number(padded.slice(0, 2));
    const minute = Number(padded.slice(2, 4));
    if (hour > 23 || minute > 59) return "";
    const hh = String(hour).padStart(2, "0");
    const mm = String(minute).padStart(2, "0");
    return `${hh}:${mm}`;
}

function formatTimeInputValue(value) {
    const digits = normalizeDigits(value).slice(0, 4);
    if (digits.length <= 3) {
        return digits;
    }
    return `${digits.slice(0, 2)}:${digits.slice(2, 4)}`;
}

function isValidSignupPhone(phone) {
    return /^010\d{8}$/.test(phone);
}

function createFacilityMemberPayload(base) {
    const phone = normalizeDigits(base.phone);
    const birthDate = parseBirthDateInput(base.birthDate);
    return {
        name: String(base.name || "").trim(),
        gender: String(base.gender || "").trim(),
        birthDate,
        phone,
        isSeongnamResident: base.isSeongnamResident ? "관내" : "관외",
        privacyConsent: Boolean(base.privacyConsent)
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
    const displayAge = Number(member.age || 0);

    if (welcome) {
        welcome.textContent = `${member.name}님 시설 이용 안내`;
    }
    if (info) {
        info.textContent = `${member.gender} · ${displayAge}세 · 연락처 끝자리 ${member.phoneLastDigits || "-"}`;
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

function setupBirthDateInputs() {
    ["signupBirthDate", "birthUpdateInput"].forEach((inputId) => {
        const input = document.getElementById(inputId);
        if (!input) {
            return;
        }

        input.addEventListener("input", (event) => {
            event.target.value = normalizeDigits(event.target.value).slice(0, 8);
        });
    });
}

function getEntryPin() {
    return Array.from(document.querySelectorAll(".pin-input"))
        .map((input) => input.value)
        .join("");
}

function resetEntryStep() {
    document.getElementById("entryForm")?.classList.remove("hidden");
    document.getElementById("fullPhoneForm")?.classList.add("hidden");
    document.querySelectorAll(".pin-input").forEach((input) => {
        input.value = "";
    });
    const entryCode = document.getElementById("entryCode");
    const fullPhone = document.getElementById("fullPhoneInput");
    if (entryCode) entryCode.value = "";
    if (fullPhone) fullPhone.value = "";
    document.querySelector(".pin-input")?.focus();
}

function showFullPhoneStep(message) {
    document.getElementById("entryForm")?.classList.add("hidden");
    document.getElementById("fullPhoneForm")?.classList.remove("hidden");
    showToast(message, "info");
    document.getElementById("fullPhoneInput")?.focus();
}

function setupPinInputs() {
    const inputs = Array.from(document.querySelectorAll(".pin-input"));
    inputs.forEach((input, index) => {
        input.addEventListener("input", (event) => {
            const digits = normalizeDigits(event.target.value);
            if (digits.length > 1) {
                digits.slice(0, 4).split("").forEach((digit, offset) => {
                    if (inputs[index + offset]) inputs[index + offset].value = digit;
                });
            } else {
                event.target.value = digits.slice(-1);
            }
            document.getElementById("entryCode").value = getEntryPin();
            const nextIndex = Math.min(index + Math.max(digits.length, 1), inputs.length - 1);
            if (digits && index < inputs.length - 1) inputs[nextIndex].focus();
        });

        input.addEventListener("keydown", (event) => {
            if (event.key === "Backspace" && !input.value && index > 0) {
                inputs[index - 1].focus();
            }
            if (event.key === "ArrowLeft" && index > 0) inputs[index - 1].focus();
            if (event.key === "ArrowRight" && index < inputs.length - 1) inputs[index + 1].focus();
        });

        input.addEventListener("paste", (event) => {
            const digits = normalizeDigits(event.clipboardData?.getData("text")).slice(0, 4);
            if (digits.length <= 1) return;

            event.preventDefault();
            digits.split("").forEach((digit, offset) => {
                if (inputs[index + offset]) inputs[index + offset].value = digit;
            });
            document.getElementById("entryCode").value = getEntryPin();
            inputs[Math.min(index + digits.length - 1, inputs.length - 1)].focus();
        });

        input.addEventListener("focus", () => input.select());
    });

    if (!sections.facilityEntry?.classList.contains("hidden")) {
        inputs[0]?.focus();
    }
}

function setupNetworkWarning() {
    const warning = document.getElementById("onlineWarning");
    const update = () => warning?.classList.toggle("hidden", navigator.onLine);
    update();
    window.addEventListener("online", update);
    window.addEventListener("offline", update);
}

function showVerifiedSignupDetails(phone) {
    state.verifiedSignupPhone = phone;
    document.getElementById("signupPhoneCheckForm")?.classList.add("hidden");
    document.getElementById("signupForm")?.classList.remove("hidden");
    const description = document.getElementById("signupDescription");
    if (description) description.textContent = "등록된 회원 정보가 없어 신규 회원 등록을 진행합니다.";
    const verifiedText = document.getElementById("verifiedSignupPhoneText");
    if (verifiedText) verifiedText.textContent = `신규 등록 번호 ${phone.slice(0, 3)}-${phone.slice(3, 7)}-${phone.slice(7)}`;
    document.getElementById("signupName")?.focus();
}

function openSignupPhoneStep(value = "", isVerified = false) {
    const phoneInput = document.getElementById("signupPhone");
    const phone = normalizeDigits(value);
    if (phoneInput) {
        phoneInput.value = phone.length === 11 ? phone : "";
    }
    state.verifiedSignupPhone = "";
    const description = document.getElementById("signupDescription");
    if (description) description.textContent = "전체 전화번호로 기존 가입 여부를 먼저 확인합니다.";
    document.getElementById("signupPhoneCheckForm")?.classList.remove("hidden");
    document.getElementById("signupForm")?.classList.add("hidden");
    document.getElementById("signupForm")?.reset();
    showSection("facilitySignup");
    if (isVerified && phone.length === 11) {
        showVerifiedSignupDetails(phone);
    } else {
        phoneInput?.focus();
    }
}

function showMemberConfirmation(member, source = "facility-qr") {
    state.candidateMember = member;
    state.candidateSource = source;
    const name = document.getElementById("memberConfirmName");
    const info = document.getElementById("memberConfirmInfo");
    if (name) name.textContent = maskMemberName(member.name);
    if (info) info.textContent = `연락처 끝자리 ${member.phoneLastDigits || "-"}`;
    showSection("memberConfirm");
}

async function confirmCandidateMember() {
    if (!state.candidateMember) {
        throw new Error("회원 정보를 다시 확인해 주세요.");
    }
    const member = state.candidateMember;
    const source = state.candidateSource;
    await routeFacilityMember(member, source);
    state.candidateMember = null;
}

async function verifySignupPhone() {
    const phone = normalizeDigits(document.getElementById("signupPhone")?.value);
    if (!isValidSignupPhone(phone)) {
        setAlert("signupAlert", "휴대폰 번호 11자리를 정확히 입력해 주세요.");
        return;
    }

    const response = await callScript("resolveFacilityEntry", { inputValue: phone });
    if (response.mode === "member" && response.member) {
        showMemberConfirmation(response.member, "facility-phone-recheck");
        return;
    }

    if (response.mode !== "signup") {
        throw new Error("전화번호를 확인하지 못했습니다. 다시 입력해 주세요.");
    }

    showVerifiedSignupDetails(phone);
}

function resetToChoice() {
    state.currentMember = null;
    state.currentMode = null;
    state.pendingMember = null;
    state.candidateMember = null;
    state.candidateSource = "facility-qr";
    state.verifiedSignupPhone = "";
    state.todayReservations = [];
    state.selectedReservationId = "";
    document.getElementById("entryForm")?.reset();
    document.getElementById("signupForm")?.reset();
    document.getElementById("signupPhoneCheckForm")?.reset();
    document.getElementById("birthUpdateForm")?.reset();
    document.getElementById("printerForm")?.reset();
    document.getElementById("careerForm")?.reset();
    document.getElementById("connectForm")?.reset();
    renderTodayReservations([]);
    ["careerCompanionCount", "connectCompanionCount"].forEach((id) => {
        const input = document.getElementById(id);
        if (input) input.value = 0;
    });
    ["careerCompanions", "connectCompanions"].forEach((id) => {
        const container = document.getElementById(id);
        if (container) container.innerHTML = "";
    });
    document.querySelectorAll(".tab").forEach((tab) => {
        tab.classList.toggle("active", tab.dataset.target === "printerPanel");
    });
    document.querySelectorAll(".tab-content").forEach((panel) => {
        panel.classList.toggle("hidden", panel.id !== "printerPanel");
    });
    setAlert("entryAlert");
    setAlert("signupAlert");
    setAlert("birthUpdateAlert");
    setAlert("facilityAlert");
    resetEntryStep();
    showSection("facilityEntry");
}

function getCurrentMemberOrThrow() {
    if (!state.currentMember || state.currentMode !== "facility") {
        throw new Error("시설 이용 회원 정보를 찾을 수 없습니다. 처음 화면에서 다시 시작해 주세요.");
    }
    return state.currentMember;
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
        }, SCRIPT_REQUEST_TIMEOUT_MS);

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
        url.searchParams.set("payload", JSON.stringify({
            ...payload,
            entryToken: getEntryToken()
        }));
        url.searchParams.set("callback", callbackName);
        script.src = url.toString();
        document.body.appendChild(script);
    });
}

async function ensureVisitLog(member, source = "facility-qr") {
    const response = await callScript("submitVisitLog", {
        memberId: member.id,
        source
    });
    return response;
}

async function routeFacilityMember(member, source = "facility-qr") {
    const hasBirthDate = member.hasBirthDate ?? Boolean(member.birthDate);
    if (!hasBirthDate) {
        state.pendingMember = member;
        setAlert("birthUpdateAlert");
        document.getElementById("birthUpdateForm")?.reset();
        const memberName = document.getElementById("birthUpdateMemberName");
        if (memberName) memberName.textContent = maskMemberName(member.name);
        showSection("birthUpdate");
        document.getElementById("birthUpdateInput")?.focus();
        return;
    }

    const visitResponse = await ensureVisitLog(member, source);
    state.pendingMember = null;
    setCurrentMember(member, "facility");
    renderTodayReservations(visitResponse.reservations || []);
    showSection("facility");
}

function renderTodayReservations(reservations) {
    const section = document.getElementById("todayReservations");
    const list = document.getElementById("reservationList");
    const count = document.getElementById("reservationCount");
    const submitButton = document.getElementById("reservationSubmitBtn");
    if (!section || !list || !submitButton) return;

    state.todayReservations = Array.isArray(reservations) ? reservations : [];
    if (!state.todayReservations.some((reservation) => reservation.id === state.selectedReservationId && !reservation.submitted)) {
        state.selectedReservationId = "";
    }

    section.classList.toggle("hidden", state.todayReservations.length === 0);
    list.replaceChildren();
    if (count) count.textContent = `${state.todayReservations.length}건`;

    state.todayReservations.forEach((reservation) => {
        const item = document.createElement("button");
        item.type = "button";
        item.className = "reservation-item";
        item.dataset.reservationId = reservation.id;
        item.classList.toggle("selected", reservation.id === state.selectedReservationId);
        item.classList.toggle("submitted", Boolean(reservation.submitted));
        item.setAttribute("aria-pressed", String(reservation.id === state.selectedReservationId));
        if (reservation.submitted) item.setAttribute("aria-disabled", "true");

        const top = document.createElement("span");
        top.className = "reservation-item-top";
        const place = document.createElement("span");
        place.className = "reservation-place";
        place.textContent = reservation.placeName || "예약 공간";
        const time = document.createElement("span");
        time.className = "reservation-time";
        time.textContent = `${reservation.startTime || "-"} ~ ${reservation.endTime || "-"}`;
        top.append(place, time);

        const purpose = document.createElement("span");
        purpose.className = "reservation-purpose";
        purpose.textContent = reservation.purpose || "네이버 예약 이용";

        const meta = document.createElement("span");
        meta.className = "reservation-item-meta";
        const headcount = document.createElement("span");
        headcount.textContent = `예약 인원 ${reservation.headcount || 1}명`;
        meta.append(headcount);
        if (reservation.submitted) {
            const submitted = document.createElement("span");
            submitted.className = "reservation-state";
            submitted.textContent = "제출 완료";
            meta.append(submitted);
        }

        item.append(top, purpose, meta);
        item.addEventListener("click", () => {
            if (reservation.submitted) return;
            state.selectedReservationId = reservation.id;
            renderTodayReservations(state.todayReservations);
        });
        list.append(item);
    });

    renderReservationCompanions();
    updateReservationSubmitState();
}

function reservationCompanionsComplete() {
    const reservation = state.todayReservations.find((item) => item.id === state.selectedReservationId && !item.submitted);
    if (!reservation) return false;
    const expectedCount = Math.max(0, Number(reservation.headcount || 1) - 1);
    if (expectedCount === 0) return true;

    const rows = document.querySelectorAll("#reservationCompanions .companion-item");
    if (rows.length !== expectedCount) return false;
    return Array.from(rows).every((row) => {
        const gender = row.querySelector(".comp-gender")?.value || "";
        const age = Number(row.querySelector(".comp-age")?.value || 0);
        return Boolean(gender) && Number.isInteger(age) && age >= 1 && age <= 100;
    });
}

function updateReservationSubmitState() {
    const button = document.getElementById("reservationSubmitBtn");
    if (button) button.disabled = !reservationCompanionsComplete();
}

function renderReservationCompanions() {
    const section = document.getElementById("reservationCompanionSection");
    const container = document.getElementById("reservationCompanions");
    const countLabel = document.getElementById("reservationCompanionCount");
    if (!section || !container) return;

    const reservation = state.todayReservations.find((item) => item.id === state.selectedReservationId && !item.submitted);
    const companionCount = reservation ? Math.max(0, Number(reservation.headcount || 1) - 1) : 0;
    section.classList.toggle("hidden", companionCount === 0);
    if (countLabel) countLabel.textContent = companionCount > 0 ? `${companionCount}명` : "";

    const renderedReservationId = container.dataset.reservationId || "";
    const renderedCount = Number(container.dataset.companionCount || 0);
    if (!reservation || companionCount === 0) {
        container.replaceChildren();
        delete container.dataset.reservationId;
        delete container.dataset.companionCount;
        return;
    }
    if (renderedReservationId === reservation.id && renderedCount === companionCount) return;

    buildCompanionFields("reservationCompanions", companionCount);
    container.dataset.reservationId = reservation.id;
    container.dataset.companionCount = String(companionCount);
    container.querySelectorAll("select, input").forEach((input) => {
        input.addEventListener("input", updateReservationSubmitState);
        input.addEventListener("change", updateReservationSubmitState);
    });
}

async function submitSelectedReservation() {
    const member = getCurrentMemberOrThrow();
    const reservationId = state.selectedReservationId;
    if (!reservationId) throw new Error("제출할 예약 일정을 선택해 주세요.");
    const reservation = state.todayReservations.find((item) => item.id === reservationId && !item.submitted);
    if (!reservation) throw new Error("예약 정보를 다시 선택해 주세요.");

    const companionCount = Math.max(0, Number(reservation.headcount || 1) - 1);
    const companions = collectCompanions("reservationCompanions");
    if (companions.length !== companionCount) {
        throw new Error("동반 이용자의 성별과 나이를 모두 입력해 주세요.");
    }

    const response = await callScript("submitReservationUsage", {
        memberId: member.id,
        reservationId,
        companions
    });
    state.todayReservations = state.todayReservations.map((reservation) => (
        reservation.id === reservationId
            ? { ...reservation, submitted: true }
            : reservation
    ));
    state.selectedReservationId = "";
    renderTodayReservations(state.todayReservations);
    showToast(
        response.created ? "예약 내용으로 이용일지를 제출했습니다." : "이미 제출된 예약입니다.",
        "success"
    );
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

    if (response.mode === "ambiguous") {
        showFullPhoneStep(response.message || "전체 전화번호를 다시 입력해 주세요.");
        return;
    }

    if (response.mode === "signup") {
        const phone = normalizeDigits(rawValue);
        openSignupPhoneStep(response.suggestedPhone || rawValue, phone.length === 11);
        return;
    }

    if (response.mode === "member" && response.member) {
        showMemberConfirmation(response.member, "facility-qr");
        return;
    }

    throw new Error("입력값을 처리하지 못했습니다.");
}

async function updateMissingBirthDate() {
    setAlert("birthUpdateAlert");
    if (!state.pendingMember?.id) {
        throw new Error("회원 정보를 다시 확인해 주세요.");
    }

    if (state.pendingMember.hasBirthDate) {
        await routeFacilityMember(state.pendingMember, "facility-birthdate-update");
        return;
    }

    const birthDate = parseBirthDateInput(document.getElementById("birthUpdateInput")?.value);
    if (!birthDate || birthDateToAge(birthDate) < 0) {
        const message = "생년월일을 YYYYMMDD 형식으로 정확히 입력해 주세요.";
        setAlert("birthUpdateAlert", message);
        return;
    }

    const response = await callScript("updateMemberBirthDate", {
        memberId: state.pendingMember.id,
        birthDate,
        source: "facility-birthdate-update"
    });
    if (!response.member) {
        throw new Error("생년월일을 저장하지 못했습니다.");
    }

    state.pendingMember = response.member;
    if (response.visit) {
        state.pendingMember = null;
        setCurrentMember(response.member, "facility");
        renderTodayReservations(response.visit.reservations || []);
        showSection("facility");
    } else {
        await routeFacilityMember(response.member, "facility-birthdate-update");
    }
    showToast("생년월일이 저장되었습니다.", "success");
}

async function registerFacilityMember() {
    setAlert("signupAlert");
    const payload = createFacilityMemberPayload({
        name: document.getElementById("signupName")?.value,
        gender: document.getElementById("signupGender")?.value,
        birthDate: document.getElementById("signupBirthDate")?.value,
        phone: state.verifiedSignupPhone,
        isSeongnamResident: document.getElementById("signupSeongnamResident")?.checked,
        privacyConsent: document.getElementById("signupPrivacyConsent")?.checked
    });

    if (!payload.name || !payload.gender || !payload.birthDate || birthDateToAge(payload.birthDate) < 0 || !payload.phone) {
        setAlert("signupAlert", "입력 항목을 모두 확인해 주세요.");
        return;
    }

    if (!isValidSignupPhone(payload.phone)) {
        setAlert("signupAlert", "연락처는 01012341234 형식의 휴대폰 번호만 입력할 수 있습니다.");
        return;
    }

    if (!payload.privacyConsent) {
        setAlert("signupAlert", "개인정보 수집·이용(필수) 동의가 필요합니다.");
        return;
    }

    const response = await callScript("registerFacilityMember", {
        member: payload
    });

    if (response.existing) {
        showMemberConfirmation(response.member, "facility-signup-existing");
        return;
    }

    await routeFacilityMember(response.member, "facility-signup");
    showToast("회원 등록과 방문 확인이 완료되었습니다.", "success");
}

async function submitUsageRecord(type, data, member, requestId = "") {
    return callScript("submitUsageRecord", {
        type,
        data: {
            ...data,
            startTime: data.startTime ? sanitizeTime(data.startTime) : "",
            endTime: data.endTime ? sanitizeTime(data.endTime) : ""
        },
        memberId: member.id,
        requestId
    });
}

function setupListeners() {
    document.getElementById("signupBackBtn")?.addEventListener("click", () => resetToChoice());
    document.getElementById("facilityExitBtn")?.addEventListener("click", () => resetToChoice());

    document.getElementById("reservationSubmitBtn")?.addEventListener("click", async () => {
        toggleLoading("reservationSubmitBtn", true);
        try {
            await submitSelectedReservation();
        } catch (error) {
            console.error(error);
            showToast(error.message || "예약 일지 제출 중 문제가 발생했습니다.", "error");
        } finally {
            toggleLoading("reservationSubmitBtn", false);
            const button = document.getElementById("reservationSubmitBtn");
            if (button) button.disabled = !state.selectedReservationId;
        }
    });

    document.getElementById("memberConfirmBtn")?.addEventListener("click", async () => {
        toggleLoading("memberConfirmBtn", true);
        try {
            await confirmCandidateMember();
        } catch (error) {
            console.error(error);
            showToast(error.message || "회원 확인 중 문제가 발생했습니다.", "error");
        } finally {
            toggleLoading("memberConfirmBtn", false);
        }
    });

    document.getElementById("memberRejectBtn")?.addEventListener("click", () => {
        state.candidateMember = null;
        resetToChoice();
    });

    document.getElementById("entryForm")?.addEventListener("submit", async (event) => {
        event.preventDefault();
        const pin = getEntryPin();
        if (pin.length !== 4) {
            setAlert("entryAlert", "전화번호 뒷 4자리를 모두 입력해 주세요.");
            return;
        }
        toggleLoading("entrySubmitBtn", true);
        try {
            await handleFacilityEntry(pin);
        } catch (error) {
            console.error(error);
            setAlert("entryAlert", error.message || "조회 중 문제가 발생했습니다.");
        } finally {
            toggleLoading("entrySubmitBtn", false);
        }
    });

    document.getElementById("fullPhoneForm")?.addEventListener("submit", async (event) => {
        event.preventDefault();
        const phone = normalizeDigits(document.getElementById("fullPhoneInput")?.value);
        if (!isValidSignupPhone(phone)) {
            setAlert("entryAlert", "휴대폰 번호 11자리를 정확히 입력해 주세요.");
            return;
        }
        toggleLoading("fullPhoneSubmitBtn", true);
        try {
            await handleFacilityEntry(phone);
        } catch (error) {
            console.error(error);
            setAlert("entryAlert", error.message || "회원 확인 중 문제가 발생했습니다.");
        } finally {
            toggleLoading("fullPhoneSubmitBtn", false);
        }
    });

    document.getElementById("fullPhoneCancelBtn")?.addEventListener("click", () => {
        setAlert("entryAlert");
        resetEntryStep();
    });

    document.getElementById("fullPhoneInput")?.addEventListener("input", (event) => {
        event.target.value = normalizeDigits(event.target.value).slice(0, 11);
    });

    document.getElementById("birthUpdateForm")?.addEventListener("submit", async (event) => {
        event.preventDefault();
        toggleLoading("birthUpdateSubmitBtn", true);
        try {
            await updateMissingBirthDate();
        } catch (error) {
            console.error(error);
            const message = error.message || "생년월일 저장 중 문제가 발생했습니다.";
            setAlert("birthUpdateAlert", message);
        } finally {
            toggleLoading("birthUpdateSubmitBtn", false);
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

    document.getElementById("signupPhoneCheckForm")?.addEventListener("submit", async (event) => {
        event.preventDefault();
        toggleLoading("signupPhoneCheckBtn", true);
        try {
            await verifySignupPhone();
        } catch (error) {
            console.error(error);
            showToast(error.message || "전화번호 확인 중 문제가 발생했습니다.", "error");
        } finally {
            toggleLoading("signupPhoneCheckBtn", false);
        }
    });

    document.getElementById("signupPhoneChangeBtn")?.addEventListener("click", () => openSignupPhoneStep());

    document.getElementById("signupPhone")?.addEventListener("input", (event) => {
        const input = event.target;
        input.value = normalizeDigits(input.value).slice(0, 11);
    });

    document.getElementById("printerForm")?.addEventListener("submit", async (event) => {
        event.preventDefault();
        toggleLoading("printerSubmitBtn", true);
        try {
            const member = getCurrentMemberOrThrow();
            const count = Number(document.getElementById("printerCount")?.value || 0);
            const submission = getPrinterSubmission(member.id, count);
            const response = await submitUsageRecord("printer", { count }, member, submission.requestId);
            clearPrinterSubmission(submission.requestId);
            showToast(
                response.created === false
                    ? "이미 반영된 프린터 기록입니다."
                    : "프린터 사용 기록을 제출했습니다.",
                "success"
            );
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
            const purpose = String(document.getElementById("careerPurpose")?.value || "").trim();
            if (!purpose) {
                throw new Error("이용 목적을 입력해 주세요.");
            }
            const response = await submitUsageRecord("careerZone", {
                spaceType: document.getElementById("careerSpaceType")?.value,
                startTime: document.getElementById("careerStartTime")?.value,
                endTime: document.getElementById("careerEndTime")?.value,
                purpose,
                companions
            }, member);
            showToast(
                response?.linkedReservation
                    ? "네이버 예약 일정과 연결하여 이용 기록을 제출했습니다."
                    : "커리어존 이용 기록을 제출했습니다.",
                "success"
            );
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
            const response = await submitUsageRecord("connectRoom", {
                startTime: document.getElementById("connectStartTime")?.value,
                endTime: document.getElementById("connectEndTime")?.value,
                purpose: document.getElementById("connectPurpose")?.value,
                companions
            }, member);
            showToast(
                response?.linkedReservation
                    ? "네이버 예약 일정과 연결하여 이용 기록을 제출했습니다."
                    : "커넥트존 이용 기록을 제출했습니다.",
                "success"
            );
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

function init() {
    if (!hasValidQrEntryToken()) {
        showEntryGate();
        return;
    }

    showSection("facilityEntry");
    setupTabs();
    setupTimeInputs();
    setupBirthDateInputs();
    setupPinInputs();
    setupNetworkWarning();
    setupStepper("careerCompanionCount", "careerCompanions");
    setupStepper("connectCompanionCount", "connectCompanions");
    setupListeners();
}

try {
    init();
} catch (error) {
    console.error("Initialization failed:", error);
    showToast("초기화 중 문제가 발생했습니다. 설정을 확인해 주세요.", "error");
}
