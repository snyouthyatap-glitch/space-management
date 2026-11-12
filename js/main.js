import { GOOGLE_SCRIPT_URL, QR_ENTRY_PARAM } from "./config.js?v=20260724-3";
import { showToast, toggleLoading } from "./utils.js?v=20260724-3";

const QR_ENTRY_SESSION_KEY = "space-management-qr-entry-ok";
const PENDING_PRINTER_SUBMISSION_KEY = "space-management-pending-printer-submission";
const SCRIPT_REQUEST_TIMEOUT_MS = 25000;
const FACILITY_MIN_AGE = 19;
const FACILITY_MAX_AGE = 39;
const BOOKING_SLOT_START_HOUR = 9;
const BOOKING_SLOT_END_HOUR = 21;
const BOOKING_SLOT_MINUTES = 30;
const BOOKING_SPACE_CONFIG = {
    careerZone: {
        label: "커리어존",
        places: [
            { id: "careerRoom1", label: "Room1", placeName: "커리어존 Room1", maxHeadcount: 1 },
            { id: "careerRoom2", label: "Room2", placeName: "커리어존 Room2", maxHeadcount: 2 }
        ]
    },
    connectRoom: {
        label: "커넥트룸",
        places: [
            { id: "connectRoom", label: "커넥트룸", placeName: "커넥트룸", maxHeadcount: 8 }
        ]
    }
};
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
    autoSelectedReservationId: "",
    pendingPrinterSubmission: null,
    booking: {
        activeSpace: "",
        activePlaceId: "",
        visibleMonth: todayString().slice(0, 7),
        selectedDate: todayString(),
        selectedSlots: [],
        pendingStartSlot: "",
        selectionComplete: false,
        dragAnchorSlot: null,
        isDraggingSlots: false,
        availability: {
            key: "",
            reservations: [],
            reservedSlots: [],
            loading: false,
            error: ""
        }
    }
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

function getUrlParam(name) {
    return new URLSearchParams(window.location.search).get(name) || "";
}

function getDirectPage() {
    const page = String(window.location.hash || "").replace(/^#/, "").toLowerCase();
    return page === "signup" || page === "printer" ? page : "";
}

function isReservationDemoMode() {
    return isLocalPreview() && getUrlParam("demo") === "reservation";
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

function getSeoulNow() {
    const parts = new Intl.DateTimeFormat("en-US", {
        timeZone: "Asia/Seoul",
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        hourCycle: "h23"
    }).formatToParts(new Date());
    const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
    return {
        date: `${values.year}-${values.month}-${values.day}`,
        minutes: Number(values.hour) * 60 + Number(values.minute)
    };
}

function reservationSubmitStatus(reservation) {
    if (reservation?.submitted) {
        return { canSubmit: false, message: "일지 제출이 완료되었습니다.", state: "submitted" };
    }
    const reservationDate = reservation?.date || todayString();
    const now = getSeoulNow();
    if (reservationDate !== now.date) {
        return { canSubmit: false, message: "예약 당일에 일지를 제출할 수 있습니다.", state: "not-today" };
    }
    const start = timeToMinutes(reservation?.startTime);
    if (!Number.isFinite(start) || start <= 0) {
        return { canSubmit: false, message: "예약 시간을 확인한 뒤 일지를 제출할 수 있습니다.", state: "invalid-time" };
    }
    // 예약 당일이면 시간과 관계없이 제출할 수 있다(2026-11-12 운영 정책).
    return { canSubmit: true, message: "지금 일지를 제출할 수 있습니다.", state: "ready" };
}

function isWebReservation(reservation) {
    const source = String(reservation?.source || reservation?.reservationSource || "").toLowerCase();
    return source === "webreservation" || source === "web-reservation" || source === "web" || source === "reservationrequest";
}

function getReservationCompanions(reservation) {
    if (!Array.isArray(reservation?.companions)) return [];
    return reservation.companions
        .map((item) => ({
            gender: String(item?.gender || "").trim(),
            age: Number(item?.age || 0)
        }))
        .filter((item) => item.gender && Number.isFinite(item.age) && item.age > 0);
}

function normalizeTodayReservation(reservation) {
    if (!reservation || typeof reservation !== "object") return reservation;
    const source = reservation.source || reservation.reservationSource || "";
    const id = reservation.id || reservation.reservationId || reservation.webReservationId || reservation.naverReservationId || "";
    return {
        ...reservation,
        id,
        source,
        date: reservation.date || reservation.reservationDate || todayString(),
        startTime: reservation.startTime || reservation.start || "",
        endTime: reservation.endTime || reservation.end || "",
        placeName: reservation.placeName || reservation.spaceName || reservation.place || "예약 공간",
        purpose: reservation.purpose || reservation.memo || (isWebReservation({ source }) ? "웹 예약 신청" : "네이버 예약 이용"),
        headcount: Number(reservation.headcount || reservation.people || reservation.count || 1),
        submitted: Boolean(reservation.submitted || reservation.usageSubmitted)
    };
}

function parseDateString(dateStr) {
    const [year, month, day] = String(dateStr || "").split("-").map(Number);
    return new Date(year, month - 1, day || 1);
}

function formatDateString(date) {
    return [
        date.getFullYear(),
        String(date.getMonth() + 1).padStart(2, "0"),
        String(date.getDate()).padStart(2, "0")
    ].join("-");
}

function addDays(dateStr, days) {
    const date = parseDateString(dateStr);
    date.setDate(date.getDate() + days);
    return formatDateString(date);
}

function addMonths(monthStr, amount) {
    const [year, month] = String(monthStr || todayString().slice(0, 7)).split("-").map(Number);
    const date = new Date(year, month - 1 + amount, 1);
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function monthDisplayLabel(monthStr) {
    const [year, month] = String(monthStr).split("-").map(Number);
    return `${year}년 ${month}월`;
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

function isFacilityAgeEligible(age) {
    return Number.isInteger(age) && age >= FACILITY_MIN_AGE && age <= FACILITY_MAX_AGE;
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
    window.requestAnimationFrame(() => {
        window.scrollTo({ top: 0, behavior: "smooth" });
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
        welcome.textContent = `${member.name}님`;
    }
    if (info) {
        info.textContent = `${member.gender} · ${displayAge}세 · ${member.phoneLastDigits || "-"}`;
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
        const activateTab = () => {
            const targetId = tab.dataset.target;
            const wrapper = tab.closest(".card");
            wrapper?.querySelectorAll(".tab").forEach((item) => {
                item.classList.remove("active");
                item.setAttribute("aria-selected", "false");
            });
            wrapper?.querySelectorAll(".tab-content").forEach((panel) => panel.classList.add("hidden"));
            tab.classList.add("active");
            tab.setAttribute("aria-selected", "true");
            document.getElementById(targetId)?.classList.remove("hidden");
        };
        tab.setAttribute("aria-selected", tab.classList.contains("active") ? "true" : "false");
        tab.addEventListener("click", activateTab);
        tab.addEventListener("keydown", (event) => {
            if (event.key !== "Enter" && event.key !== " ") return;
            event.preventDefault();
            activateTab();
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

function showFacilityAgeRestriction(message) {
    state.currentMember = null;
    state.pendingMember = null;
    state.candidateMember = null;
    state.candidateSource = "facility-qr";
    showSection("facilityEntry");
    resetEntryStep();
    showToast(message || "시설 이용 가능 연령을 확인해 주세요.", "error");
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
    setMemberRejectChoiceVisible(false);
    showSection("memberConfirm");
}

function setMemberRejectChoiceVisible(visible) {
    document.getElementById("memberConfirmActions")?.classList.toggle("hidden", visible);
    document.getElementById("memberRejectChoice")?.classList.toggle("hidden", !visible);
}

function showLoungeNotice() {
    showMainFacilityView();
    const notice = document.getElementById("loungeNotice");
    if (!notice) return;
    notice.classList.remove("hidden", "lounge-notice-pulse");
    void notice.offsetWidth;
    notice.classList.add("lounge-notice-pulse");
    showToast("라운지는 지금처럼 출석만 하면 이용 기록이 끝나요. 따로 제출할 것이 없어요.", "success");
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
    if (response.mode === "restricted" || response.restricted) {
        showFacilityAgeRestriction(response.message);
        return;
    }
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
    state.autoSelectedReservationId = "";
    state.booking.activeSpace = "";
    state.booking.activePlaceId = "";
    state.booking.visibleMonth = todayString().slice(0, 7);
    state.booking.selectedDate = todayString();
    resetBookingSelection();
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
    openBookingMenu();
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
    return callScript("submitVisitLog", {
        memberId: member.id,
        source
    });
}

async function loadTodayReservations(member) {
    try {
        const response = await callScript("getTodayReservations", {
            memberId: member.id
        });
        if (state.currentMember?.id !== member.id || state.currentMode !== "facility") return;
        renderTodayReservations(response.reservations || response.todayReservations || []);
    } catch (error) {
        console.warn("Reservation loading failed:", error);
        if (state.currentMember?.id === member.id && state.currentMode === "facility") {
            showToast("예약 일정은 잠시 후 다시 확인해 주세요.", "info");
        }
    }
}

function enterFacility(member) {
    state.pendingMember = null;
    setCurrentMember(member, "facility");
    renderTodayReservations([]);
    showSection("facility");
    if (getDirectPage() === "printer") {
        window.requestAnimationFrame(showPrinterOnlyView);
    }
    if (isReservationDemoMode()) {
        renderTodayReservations(createDemoReservations());
        showToast("예약 미리보기 모드입니다. 실제 시트에는 저장되지 않습니다.", "info");
        return;
    }
    void loadTodayReservations(member);
}

function enterReservationDemo() {
    setCurrentMember(createDemoMember(), "facility");
    showSection("facility");
    renderTodayReservations(createDemoReservations());
    showToast("예약 미리보기 모드입니다. 실제 시트에는 저장되지 않습니다.", "info");
}

function createDemoMember() {
    return {
        id: "demo-member",
        name: "안치우",
        gender: "남성",
        age: 27,
        phoneLastDigits: "0000",
        hasBirthDate: true
    };
}

function createDemoReservations() {
    return [
        {
            id: "demo-reservation-career-1",
            demo: true,
            placeName: "커리어존 Room1",
            startTime: "14:00",
            endTime: "15:00",
            purpose: "진로 상담 준비",
            headcount: 2,
            submitted: false
        },
        {
            id: "demo-reservation-connect-1",
            demo: true,
            source: "webReservation",
            webReservationId: "demo-web-reservation-connect-1",
            placeName: "커넥트룸",
            startTime: "16:00",
            endTime: "18:00",
            purpose: "팀 프로젝트 회의",
            headcount: 4,
            companions: [
                { gender: "여성", age: 25 },
                { gender: "남성", age: 28 },
                { gender: "여성", age: 26 }
            ],
            submitted: false
        }
    ];
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

    await ensureVisitLog(member, source);
    enterFacility(member);
}

function renderTodayReservations(reservations) {
    const section = document.getElementById("todayReservations");
    const list = document.getElementById("reservationList");
    const count = document.getElementById("reservationCount");
    const submitButton = document.getElementById("reservationSubmitBtn");
    if (!section || !list || !submitButton) return;

    state.todayReservations = Array.isArray(reservations)
        ? reservations.map(normalizeTodayReservation).filter((reservation) => reservation?.id)
        : [];
    if (!state.todayReservations.some((reservation) => reservation.id === state.selectedReservationId && !reservation.submitted)) {
        state.selectedReservationId = "";
    }

    // 제출할 예약이 1건뿐이면 처음 한 번 자동으로 선택한다. 사용자가 선택을 풀면 다시 고르지 않는다.
    const pendingReservations = state.todayReservations.filter(
        (reservation) => !reservation.submitted && reservationSubmitStatus(reservation).canSubmit
    );
    if (
        !state.selectedReservationId &&
        pendingReservations.length === 1 &&
        state.autoSelectedReservationId !== pendingReservations[0].id
    ) {
        state.selectedReservationId = pendingReservations[0].id;
        state.autoSelectedReservationId = pendingReservations[0].id;
    }
    updateReservationNotice(pendingReservations.length);

    section.classList.toggle("hidden", state.todayReservations.length === 0);
    list.replaceChildren();
    if (count) count.textContent = `${state.todayReservations.length}건`;

    state.todayReservations.forEach((reservation) => {
        const submitStatus = reservationSubmitStatus(reservation);
        const item = document.createElement("button");
        item.type = "button";
        item.className = "reservation-item";
        item.dataset.reservationId = reservation.id;
        item.classList.toggle("selected", reservation.id === state.selectedReservationId);
        item.classList.toggle("submitted", Boolean(reservation.submitted));
        item.classList.toggle("not-ready", !submitStatus.canSubmit && !reservation.submitted);
        item.setAttribute("aria-pressed", String(reservation.id === state.selectedReservationId));
        if (reservation.submitted) item.setAttribute("aria-disabled", "true");

        const top = document.createElement("span");
        top.className = "reservation-item-top";
        const place = document.createElement("span");
        place.className = "reservation-place";
        place.textContent = reservation.placeName || "예약 공간";
        if (isWebReservation(reservation)) {
            const sourceBadge = document.createElement("span");
            sourceBadge.className = "reservation-source-badge";
            sourceBadge.textContent = "웹 예약";
            place.append(sourceBadge);
        }
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
        const stateLabel = document.createElement("span");
        stateLabel.className = `reservation-state ${submitStatus.state}`;
        stateLabel.textContent = submitStatus.message;
        meta.append(stateLabel);

        item.append(top, purpose, meta);
        item.addEventListener("click", () => {
            if (reservation.submitted) return;
            state.selectedReservationId = state.selectedReservationId === reservation.id ? "" : reservation.id;
            renderTodayReservations(state.todayReservations);
        });
        list.append(item);
    });

    renderReservationCompanions();
    updateReservationSubmitState();
}

// 제출할 예약이 있으면 '라운지는 기록할 필요 없음' 안내 대신 일지 제출 안내를 띄운다.
function updateReservationNotice(pendingCount) {
    document.getElementById("reservationNotice")?.classList.toggle("hidden", pendingCount === 0);
    document.getElementById("loungeNotice")?.classList.toggle("hidden", pendingCount > 0);
}

function scrollToReservationSubmit() {
    showMainFacilityView();
    window.setTimeout(() => {
        document.getElementById("todayReservations")?.scrollIntoView({ behavior: "smooth", block: "start" });
    }, 150);
}

function reservationCompanionsComplete() {
    const reservation = state.todayReservations.find((item) => item.id === state.selectedReservationId && !item.submitted);
    if (!reservation) return false;
    if (!reservationSubmitStatus(reservation).canSubmit) return false;
    const expectedCount = Math.max(0, Number(reservation.headcount || 1) - 1);
    if (expectedCount === 0) return true;
    if (isWebReservation(reservation)) {
        const savedCompanions = getReservationCompanions(reservation);
        if (savedCompanions.length === expectedCount) return true;
    }

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

function setupReservationSubmitClock() {
    window.setInterval(() => {
        if (state.currentMode !== "facility" || state.todayReservations.length === 0) return;
        renderTodayReservations(state.todayReservations);
    }, 60000);
}

function renderReservationCompanions() {
    const section = document.getElementById("reservationCompanionSection");
    const container = document.getElementById("reservationCompanions");
    const countLabel = document.getElementById("reservationCompanionCount");
    if (!section || !container) return;

    const reservation = state.todayReservations.find((item) => item.id === state.selectedReservationId && !item.submitted);
    const companionCount = reservation ? Math.max(0, Number(reservation.headcount || 1) - 1) : 0;
    const reuseWebCompanions = Boolean(reservation && isWebReservation(reservation) && companionCount > 0);
    section.classList.toggle("hidden", companionCount === 0 || reuseWebCompanions);
    if (countLabel) countLabel.textContent = companionCount > 0 ? `${companionCount}명` : "";

    const renderedReservationId = container.dataset.reservationId || "";
    const renderedCount = Number(container.dataset.companionCount || 0);
    if (!reservation || companionCount === 0 || reuseWebCompanions) {
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
    const submitStatus = reservationSubmitStatus(reservation);
    if (!submitStatus.canSubmit) throw new Error(submitStatus.message);

    const companionCount = Math.max(0, Number(reservation.headcount || 1) - 1);
    const savedCompanions = getReservationCompanions(reservation);
    const companions = isWebReservation(reservation) && savedCompanions.length === companionCount
        ? savedCompanions
        : collectCompanions("reservationCompanions");
    if (companions.length !== companionCount) {
        throw new Error("동반 이용자의 성별과 나이를 모두 입력해 주세요.");
    }

    if (reservation.demo) {
        state.todayReservations = state.todayReservations.map((reservation) => (
            reservation.id === reservationId
                ? { ...reservation, submitted: true }
                : reservation
        ));
        state.selectedReservationId = "";
        renderTodayReservations(state.todayReservations);
        showToast("미리보기 예약입니다. 실제 시트에는 저장하지 않았습니다.", "info");
        return;
    }

    const response = await callScript("submitReservationUsage", {
        memberId: member.id,
        reservationId,
        reservationSource: reservation.source || "",
        webReservationId: reservation.webReservationId || reservation.id,
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

function getBookingConfig(spaceType = state.booking.activeSpace) {
    return BOOKING_SPACE_CONFIG[spaceType] || BOOKING_SPACE_CONFIG.careerZone;
}

function getBookingPlace(spaceType = state.booking.activeSpace, placeId = state.booking.activePlaceId) {
    const config = getBookingConfig(spaceType);
    return config.places.find((place) => place.id === placeId) || config.places[0];
}

function resetBookingSelection() {
    state.booking.selectedSlots = [];
    state.booking.pendingStartSlot = "";
    state.booking.selectionComplete = false;
    state.booking.dragAnchorSlot = null;
    state.booking.isDraggingSlots = false;
}

function resetBookingForm() {
    document.getElementById("bookingRequestForm")?.reset();
    const headcountInput = document.getElementById("bookingHeadcount");
    if (headcountInput) headcountInput.value = 1;
    renderBookingCompanions();
}

function showBookingView(viewName) {
    ["spaceMenuView", "bookingCalendarView", "bookingTimeView"].forEach((id) => {
        document.getElementById(id)?.classList.toggle("hidden", id !== viewName);
    });
    const bookingActive = viewName !== "spaceMenuView";
    document.getElementById("facilitySummaryCard")?.classList.toggle("hidden", bookingActive);
    document.getElementById("facilityUsageCard")?.classList.toggle("hidden", bookingActive);
    document.getElementById("todayReservations")?.classList.toggle("booking-context-hidden", bookingActive);
    document.getElementById("spaceBookingCard")?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function showMainFacilityView() {
    document.getElementById("facilitySummaryCard")?.classList.remove("hidden");
    document.getElementById("facilityUsageCard")?.classList.remove("hidden");
    document.getElementById("spaceBookingCard")?.classList.remove("hidden");
    document.getElementById("todayReservations")?.classList.remove("booking-context-hidden");
    document.getElementById("printerMainBackBtn")?.classList.add("hidden");
    document.getElementById("bookingMenuMainBackBtn")?.classList.add("hidden");
    showBookingView("spaceMenuView");
    window.requestAnimationFrame(() => {
        window.scrollTo({ top: 0, behavior: "smooth" });
    });
}

function showSpaceBookingOnlyView() {
    showBookingView("spaceMenuView");
    document.getElementById("facilitySummaryCard")?.classList.add("hidden");
    document.getElementById("facilityUsageCard")?.classList.add("hidden");
    document.getElementById("todayReservations")?.classList.add("booking-context-hidden");
    const bookingCard = document.getElementById("spaceBookingCard");
    bookingCard?.classList.remove("hidden");
    document.getElementById("bookingMenuMainBackBtn")?.classList.remove("hidden");
    bookingCard?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function showPrinterOnlyView() {
    document.getElementById("facilitySummaryCard")?.classList.add("hidden");
    document.getElementById("spaceBookingCard")?.classList.add("hidden");
    document.getElementById("todayReservations")?.classList.add("booking-context-hidden");
    const usageCard = document.getElementById("facilityUsageCard");
    usageCard?.classList.remove("hidden");
    document.getElementById("printerMainBackBtn")?.classList.remove("hidden");
    document.getElementById("bookingMenuMainBackBtn")?.classList.add("hidden");
    usageCard?.scrollIntoView({ behavior: "smooth", block: "start" });
}

function openBookingCalendar(spaceType) {
    if (!BOOKING_SPACE_CONFIG[spaceType]) return;
    state.booking.activeSpace = spaceType;
    state.booking.activePlaceId = getBookingPlace(spaceType).id;
    state.booking.visibleMonth = state.booking.selectedDate.slice(0, 7);
    resetBookingSelection();
    renderBookingCalendar();
    showBookingView("bookingCalendarView");
}

function openTodayBooking(spaceType) {
    if (!BOOKING_SPACE_CONFIG[spaceType]) return;
    state.booking.activeSpace = spaceType;
    state.booking.activePlaceId = getBookingPlace(spaceType).id;
    state.booking.selectedDate = todayString();
    state.booking.visibleMonth = state.booking.selectedDate.slice(0, 7);
    resetBookingSelection();
    renderBookingTime();
    showBookingView("bookingTimeView");
    void refreshBookingAvailability();
}

function openBookingMenu() {
    state.booking.activeSpace = "";
    state.booking.activePlaceId = "";
    resetBookingSelection();
    document.getElementById("bookingMenuMainBackBtn")?.classList.add("hidden");
    document.getElementById("printerMainBackBtn")?.classList.add("hidden");
    showBookingView("spaceMenuView");
}

function changeBookingMonth(amount) {
    state.booking.visibleMonth = addMonths(state.booking.visibleMonth, amount);
    renderBookingCalendar();
}

function renderBookingCalendar() {
    const grid = document.getElementById("bookingCalendarGrid");
    const title = document.getElementById("bookingCalendarTitle");
    const monthLabel = document.getElementById("bookingMonthLabel");
    if (!grid) return;

    const config = getBookingConfig();
    if (title) title.textContent = `${config.label} 날짜 선택`;
    if (monthLabel) monthLabel.textContent = monthDisplayLabel(state.booking.visibleMonth);

    const [year, month] = state.booking.visibleMonth.split("-").map(Number);
    const firstDay = new Date(year, month - 1, 1);
    const startOffset = firstDay.getDay();
    const daysInMonth = new Date(year, month, 0).getDate();
    const today = todayString();

    grid.replaceChildren();
    for (let index = 0; index < startOffset; index += 1) {
        const empty = document.createElement("span");
        empty.className = "calendar-day-empty";
        grid.append(empty);
    }

    for (let day = 1; day <= daysInMonth; day += 1) {
        const dateStr = `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
        const button = document.createElement("button");
        button.type = "button";
        button.className = "calendar-day";
        button.textContent = String(day);
        button.classList.toggle("today", dateStr === today);
        button.classList.toggle("selected", dateStr === state.booking.selectedDate);
        button.disabled = dateStr < today;
        button.addEventListener("click", () => openBookingTime(dateStr));
        grid.append(button);
    }
}

function openBookingTime(dateStr) {
    state.booking.selectedDate = dateStr;
    resetBookingSelection();
    renderBookingTime();
    showBookingView("bookingTimeView");
    void refreshBookingAvailability();
}

function createTimeSlots() {
    const slots = [];
    for (let hour = BOOKING_SLOT_START_HOUR; hour < BOOKING_SLOT_END_HOUR; hour += 1) {
        slots.push(`${String(hour).padStart(2, "0")}:00`);
        slots.push(`${String(hour).padStart(2, "0")}:${String(BOOKING_SLOT_MINUTES).padStart(2, "0")}`);
    }
    return slots;
}

function timeToMinutes(time) {
    const [hour, minute] = String(time || "00:00").split(":").map(Number);
    return hour * 60 + minute;
}

function slotEndTime(slot) {
    const minutes = timeToMinutes(slot) + BOOKING_SLOT_MINUTES;
    return `${String(Math.floor(minutes / 60)).padStart(2, "0")}:${String(minutes % 60).padStart(2, "0")}`;
}

function bookingSlotRangeText() {
    if (!state.booking.selectionComplete || !state.booking.selectedSlots.length) return "";
    const slots = [...state.booking.selectedSlots].sort();
    return `${slots[0]} ~ ${slotEndTime(slots[slots.length - 1])}`;
}

function bookingAvailabilityKey(spaceType = state.booking.activeSpace, placeId = state.booking.activePlaceId, dateStr = state.booking.selectedDate) {
    return `${spaceType}:${placeId}:${dateStr}`;
}

function normalizeBookingReservation(reservation) {
    if (!reservation || typeof reservation !== "object") return null;
    return {
        ...reservation,
        id: reservation.id || reservation.reservationId || reservation.webReservationId || reservation.naverReservationId || createRequestId(),
        date: reservation.date || reservation.reservationDate || state.booking.selectedDate,
        placeId: reservation.placeId || reservation.roomId || reservation.spaceDetail || "",
        placeName: reservation.placeName || reservation.spaceName || reservation.place || "예약 공간",
        startTime: reservation.startTime || reservation.start || "",
        endTime: reservation.endTime || reservation.end || "",
        headcount: Number(reservation.headcount || reservation.people || reservation.count || 1)
    };
}

function normalizeReservedSlots(slots) {
    if (!Array.isArray(slots)) return [];
    return slots
        .map((slot) => typeof slot === "string" ? slot : slot?.time || slot?.slot || "")
        .filter(Boolean);
}

function normalizeAvailabilityResponse(response) {
    const reservations = Array.isArray(response?.reservations)
        ? response.reservations.map(normalizeBookingReservation).filter(Boolean)
        : [];
    const reservedSlots = [
        ...normalizeReservedSlots(response?.reservedSlots),
        ...normalizeReservedSlots(response?.unavailableSlots),
        ...normalizeReservedSlots(response?.bookedSlots)
    ];
    return { reservations, reservedSlots: [...new Set(reservedSlots)] };
}

function matchesBookingPlace(reservation, spaceType, placeId = state.booking.activePlaceId) {
    const place = getBookingPlace(spaceType, placeId);
    const reservationPlaceId = String(reservation.placeId || reservation.roomId || reservation.spaceDetail || "");
    const placeName = String(reservation.placeName || reservation.spaceType || "");
    if (reservationPlaceId) return reservationPlaceId === place.id;
    return placeName.includes(place.placeName);
}

function createDemoBookingReservations() {
    const today = todayString();
    return [
        ...createDemoReservations().map((item) => ({ ...item, date: today })),
        {
            id: "demo-booking-career-2",
            demo: true,
            date: addDays(today, 1),
            placeName: "커리어존 Room2",
            startTime: "10:30",
            endTime: "12:00",
            purpose: "취업 준비",
            headcount: 1
        },
        {
            id: "demo-booking-connect-2",
            demo: true,
            date: addDays(today, 2),
            placeName: "커넥트룸",
            startTime: "13:00",
            endTime: "15:00",
            purpose: "스터디 모임",
            headcount: 5
        }
    ];
}

function getBookingReservations(spaceType, dateStr) {
    if (isReservationDemoMode() || isLocalPreview()) {
        return createDemoBookingReservations().filter((reservation) => (
            (reservation.date || todayString()) === dateStr && matchesBookingPlace(reservation, spaceType)
        ));
    }
    const key = bookingAvailabilityKey(spaceType, state.booking.activePlaceId, dateStr);
    if (state.booking.availability.key !== key) return [];
    return state.booking.availability.reservations;
}

function getBookingReservedSlots(spaceType, dateStr) {
    if (isReservationDemoMode() || isLocalPreview()) return [];
    const key = bookingAvailabilityKey(spaceType, state.booking.activePlaceId, dateStr);
    if (state.booking.availability.key !== key) return [];
    return state.booking.availability.reservedSlots;
}

function isSlotBooked(slot, reservations) {
    const start = timeToMinutes(slot);
    const end = timeToMinutes(slotEndTime(slot));
    return reservations.some((reservation) => (
        start < timeToMinutes(reservation.endTime) && end > timeToMinutes(reservation.startTime)
    ));
}

function isBookingSlotUnavailable(slot, reservations, reservedSlots) {
    return reservedSlots.includes(slot) || isSlotBooked(slot, reservations);
}

async function refreshBookingAvailability() {
    if (isReservationDemoMode() || isLocalPreview()) {
        state.booking.availability = {
            key: bookingAvailabilityKey(),
            reservations: getBookingReservations(state.booking.activeSpace, state.booking.selectedDate),
            reservedSlots: [],
            loading: false,
            error: ""
        };
        renderBookingTime();
        return;
    }

    const member = getCurrentMemberOrThrow();
    const place = getBookingPlace();
    const key = bookingAvailabilityKey();
    state.booking.availability = {
        key,
        reservations: [],
        reservedSlots: [],
        loading: true,
        error: ""
    };
    renderBookingTime();

    try {
        const response = await callScript("getSpaceAvailability", {
            memberId: member.id,
            spaceType: state.booking.activeSpace,
            placeId: place.id,
            placeName: place.placeName,
            date: state.booking.selectedDate
        });
        if (state.booking.availability.key !== key) return;
        const normalized = normalizeAvailabilityResponse(response);
        state.booking.availability = {
            key,
            reservations: normalized.reservations,
            reservedSlots: normalized.reservedSlots,
            loading: false,
            error: ""
        };
    } catch (error) {
        console.warn("Space availability loading failed:", error);
        if (state.booking.availability.key !== key) return;
        state.booking.availability = {
            key,
            reservations: [],
            reservedSlots: [],
            loading: false,
            error: error.message || "예약 현황을 불러오지 못했습니다."
        };
    }
    resetBookingSelection();
    renderBookingTime();
}

function renderBookingStatus(reservations) {
    const list = document.getElementById("bookingStatusList");
    if (!list) return;
    list.replaceChildren();
    if (!reservations.length) {
        const empty = document.createElement("div");
        empty.className = "booking-status-empty";
        empty.textContent = "현재 등록된 예약이 없습니다.";
        list.append(empty);
        return;
    }

    reservations.forEach((reservation) => {
        const item = document.createElement("div");
        item.className = "booking-status-item";
        item.innerHTML = `
            <strong>${reservation.startTime || "-"} ~ ${reservation.endTime || "-"}</strong>
            <span>${reservation.placeName || "예약 공간"} · ${reservation.headcount || 1}명</span>
        `;
        list.append(item);
    });
}

function renderBookingPlaceSelector() {
    const selector = document.getElementById("bookingPlaceSelector");
    if (!selector) return;

    const config = getBookingConfig();
    selector.classList.toggle("hidden", config.places.length <= 1);
    selector.replaceChildren();

    if (config.places.length <= 1) return;

    config.places.forEach((place) => {
        const button = document.createElement("button");
        button.type = "button";
        button.className = "booking-place-option";
        button.classList.toggle("active", place.id === state.booking.activePlaceId);
        button.textContent = `${place.label} · 최대 ${place.maxHeadcount}명`;
        button.addEventListener("click", () => {
            state.booking.activePlaceId = place.id;
            resetBookingSelection();
            resetBookingForm();
            renderBookingTime();
            void refreshBookingAvailability();
        });
        selector.append(button);
    });
}

function getBookingHeadcount() {
    const place = getBookingPlace();
    if (place.maxHeadcount <= 1) return 1;
    const input = document.getElementById("bookingHeadcount");
    const value = Number(input?.value || 1);
    if (!Number.isInteger(value) || value < 1) return 1;
    return Math.min(value, place.maxHeadcount);
}

function renderBookingCompanions() {
    const place = getBookingPlace();
    const headcountGroup = document.getElementById("bookingHeadcountGroup");
    const headcountInput = document.getElementById("bookingHeadcount");
    const section = document.getElementById("bookingCompanionSection");
    const countLabel = document.getElementById("bookingCompanionCount");
    const companions = document.getElementById("bookingCompanions");
    if (!headcountGroup || !headcountInput || !section || !companions) return;

    const isSingleFixed = place.maxHeadcount <= 1;
    headcountGroup.classList.toggle("hidden", isSingleFixed);
    headcountInput.max = String(place.maxHeadcount);
    if (isSingleFixed) {
        headcountInput.value = 1;
        section.classList.add("hidden");
        companions.replaceChildren();
        return;
    }

    const headcount = getBookingHeadcount();
    headcountInput.value = String(headcount);
    const companionCount = Math.max(0, headcount - 1);
    section.classList.toggle("hidden", companionCount === 0);
    if (countLabel) countLabel.textContent = companionCount > 0 ? `${companionCount}명` : "";
    if (companionCount === 0) {
        companions.replaceChildren();
        companions.dataset.companionCount = "0";
        updateBookingSubmitState();
        return;
    }
    if (companions.dataset.companionCount === String(companionCount)) {
        updateBookingSubmitState();
        return;
    }
    buildCompanionFields("bookingCompanions", companionCount);
    companions.dataset.companionCount = String(companionCount);
    companions.querySelectorAll("select, input").forEach((input) => {
        input.addEventListener("input", updateBookingSubmitState);
        input.addEventListener("change", updateBookingSubmitState);
    });
    updateBookingSubmitState();
}

function bookingCompanionsComplete() {
    const companionCount = Math.max(0, getBookingHeadcount() - 1);
    if (companionCount === 0) return true;
    const companions = collectCompanions("bookingCompanions");
    return companions.length === companionCount;
}

function updateBookingSubmitState() {
    const submitButton = document.getElementById("bookingRequestSubmitBtn");
    const availabilityPending = state.booking.availability.loading || Boolean(state.booking.availability.error);
    if (submitButton) submitButton.disabled = availabilityPending || !bookingSlotRangeText() || !bookingCompanionsComplete();
}

function bookingSelectionMessage(rangeText) {
    if (state.booking.availability.loading) return "예약 현황을 불러오는 중입니다.";
    if (state.booking.availability.error) return state.booking.availability.error;
    if (rangeText) return `선택 시간 ${rangeText}`;
    if (state.booking.pendingStartSlot) return `시작 시간 ${state.booking.pendingStartSlot} 선택됨. 종료 시간을 눌러 주세요.`;
    return "시작 시간을 선택해 주세요.";
}

function selectedSlotRange(startSlot, endSlot, reservations, reservedSlots = []) {
    const allSlots = createTimeSlots();
    const startIndex = allSlots.indexOf(startSlot);
    const endIndex = allSlots.indexOf(endSlot);
    if (startIndex < 0 || endIndex < 0) return [];

    const [start, end] = [startIndex, endIndex].sort((a, b) => a - b);
    const range = start === end ? [allSlots[start]] : allSlots.slice(start, end);
    if (range.some((slot) => isBookingSlotUnavailable(slot, reservations, reservedSlots))) {
        return [];
    }
    return range;
}

function handleTimeSlotClick(slot, booked) {
    if (booked || state.booking.availability.loading || state.booking.availability.error) return;
    const reservations = getBookingReservations(state.booking.activeSpace, state.booking.selectedDate);
    const reservedSlots = getBookingReservedSlots(state.booking.activeSpace, state.booking.selectedDate);

    if (!state.booking.pendingStartSlot || state.booking.selectionComplete) {
        state.booking.pendingStartSlot = slot;
        state.booking.selectedSlots = [slot];
        state.booking.selectionComplete = false;
        renderBookingTime();
        return;
    }

    const range = selectedSlotRange(state.booking.pendingStartSlot, slot, reservations, reservedSlots);
    if (!range.length) {
        showToast("예약이 있는 시간을 포함해서 선택할 수 없습니다.", "error");
        state.booking.pendingStartSlot = slot;
        state.booking.selectedSlots = [slot];
        state.booking.selectionComplete = false;
        renderBookingTime();
        return;
    }

    state.booking.selectedSlots = range;
    state.booking.selectionComplete = true;
    renderBookingTime();
}

function renderBookingTime() {
    const grid = document.getElementById("bookingTimeGrid");
    const title = document.getElementById("bookingTimeTitle");
    const dateLabel = document.getElementById("bookingTimeDate");
    const summary = document.getElementById("bookingSelectionSummary");
    if (!grid) return;

    const config = getBookingConfig();
    const place = getBookingPlace();
    const reservations = getBookingReservations(state.booking.activeSpace, state.booking.selectedDate);
    const reservedSlots = getBookingReservedSlots(state.booking.activeSpace, state.booking.selectedDate);
    if (title) {
        title.textContent = place.id === "connectRoom" ? `${config.label} 예약` : `${config.label} ${place.label} 예약`;
    }
    if (dateLabel) dateLabel.textContent = formatDisplayDate(state.booking.selectedDate);
    const headcountInput = document.getElementById("bookingHeadcount");
    if (headcountInput) {
        headcountInput.max = String(place.maxHeadcount);
        if (Number(headcountInput.value || 1) > place.maxHeadcount) {
            headcountInput.value = String(place.maxHeadcount);
        }
    }
    renderBookingPlaceSelector();
    renderBookingCompanions();
    renderBookingStatus(reservations);

    grid.replaceChildren();
    createTimeSlots().forEach((slot) => {
        const booked = isBookingSlotUnavailable(slot, reservations, reservedSlots);
        const selected = state.booking.selectedSlots.includes(slot);
        const button = document.createElement("button");
        button.type = "button";
        button.className = "time-slot";
        button.textContent = slot;
        button.dataset.slot = slot;
        button.disabled = booked;
        button.classList.toggle("selected", selected);
        button.classList.toggle("pending", state.booking.pendingStartSlot === slot && !state.booking.selectionComplete);
        button.classList.toggle("booked", booked);
        button.addEventListener("click", () => handleTimeSlotClick(slot, booked));
        grid.append(button);
    });

    const rangeText = bookingSlotRangeText();
    if (summary) {
        summary.classList.toggle("error", Boolean(state.booking.availability.error));
        summary.textContent = bookingSelectionMessage(rangeText);
    }
    updateBookingSubmitState();
}

async function submitBookingRequest() {
    const member = getCurrentMemberOrThrow();
    const config = getBookingConfig();
    const place = getBookingPlace();
    const rangeText = bookingSlotRangeText();
    const purpose = String(document.getElementById("bookingPurpose")?.value || "").trim();
    const headcount = getBookingHeadcount();
    const companionCount = Math.max(0, headcount - 1);
    const companions = collectCompanions("bookingCompanions");
    if (!rangeText) throw new Error("예약할 시간을 선택해 주세요.");
    if (!purpose) throw new Error("이용 목적을 입력해 주세요.");
    if (!Number.isInteger(headcount) || headcount < 1 || headcount > place.maxHeadcount) {
        throw new Error(`${place.placeName}은 1~${place.maxHeadcount}명까지 신청할 수 있습니다.`);
    }
    if (companions.length !== companionCount) {
        throw new Error("동반 이용자의 성별과 나이를 모두 입력해 주세요.");
    }
    const slots = [...state.booking.selectedSlots].sort();
    const currentReservations = getBookingReservations(state.booking.activeSpace, state.booking.selectedDate);
    const currentReservedSlots = getBookingReservedSlots(state.booking.activeSpace, state.booking.selectedDate);
    if (slots.some((slot) => isBookingSlotUnavailable(slot, currentReservations, currentReservedSlots))) {
        throw new Error("선택한 시간에 이미 예약이 있습니다. 시간을 다시 선택해 주세요.");
    }

    const payload = {
        memberId: member.id,
        memberName: member.name || "",
        phoneLastDigits: member.phoneLastDigits || "",
        source: "webReservation",
        spaceType: state.booking.activeSpace,
        placeId: place.id,
        placeName: place.placeName,
        spaceLabel: config.label,
        date: state.booking.selectedDate,
        startTime: slots[0],
        endTime: slotEndTime(slots[slots.length - 1]),
        purpose,
        headcount,
        companions,
        requestId: createRequestId()
    };

    if (isReservationDemoMode() || isLocalPreview()) {
        showToast("데모 예약 신청입니다. 실제 시트에는 저장하지 않았습니다.", "info");
        resetBookingForm();
        state.booking.selectedSlots = [];
        state.booking.pendingStartSlot = "";
        state.booking.selectionComplete = false;
        renderBookingTime();
        return payload;
    }

    const response = await callScript("submitSpaceReservationRequest", payload);
    showToast("예약 신청이 접수되었습니다.", "success");
    resetBookingForm();
    resetBookingSelection();
    if (Array.isArray(response.todayReservations) || Array.isArray(response.reservationsForToday)) {
        renderTodayReservations(response.todayReservations || response.reservationsForToday);
    } else {
        void loadTodayReservations(member);
    }
    await refreshBookingAvailability();
    return response;
}

async function handleFacilityEntry(inputValue) {
    setAlert("entryAlert");
    const rawValue = String(inputValue || "").trim();

    if (!rawValue) {
        setAlert("entryAlert", "입력값을 확인해 주세요.");
        return;
    }

    if (isReservationDemoMode()) {
        if (normalizeDigits(rawValue) !== "0000") {
            setAlert("entryAlert", "데모 모드에서는 0000을 입력해 주세요.");
            return;
        }
        enterReservationDemo();
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

    if (response.mode === "restricted" || response.restricted) {
        showFacilityAgeRestriction(response.message);
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
    if (response.mode === "restricted" || response.restricted) {
        showFacilityAgeRestriction(response.message);
        return;
    }
    if (!response.member) {
        throw new Error("생년월일을 저장하지 못했습니다.");
    }

    state.pendingMember = response.member;
    if (response.visit) {
        enterFacility(response.member);
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

    const registrationAge = birthDateToAge(payload.birthDate);
    if (!payload.name || !payload.gender || !payload.birthDate || registrationAge < 0 || !payload.phone) {
        setAlert("signupAlert", "입력 항목을 모두 확인해 주세요.");
        return;
    }

    if (!isFacilityAgeEligible(registrationAge)) {
        setAlert(
            "signupAlert",
            `청년이봄 시설은 만 19세 이상 만 39세 이하만 가입할 수 있습니다. 현재 만 ${registrationAge}세로 회원가입 및 이용이 불가능합니다.`
        );
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

    if (response.mode === "restricted" || response.restricted) {
        showFacilityAgeRestriction(response.message);
        return;
    }

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
            updateReservationSubmitState();
        }
    });

    document.getElementById("quickLoungeBtn")?.addEventListener("click", showLoungeNotice);
    document.getElementById("reservationNoticeBtn")?.addEventListener("click", scrollToReservationSubmit);
    document.getElementById("quickBookingBtn")?.addEventListener("click", showSpaceBookingOnlyView);
    document.getElementById("quickPrinterBtn")?.addEventListener("click", showPrinterOnlyView);
    document.getElementById("printerMainBackBtn")?.addEventListener("click", showMainFacilityView);
    document.getElementById("bookingMenuMainBackBtn")?.addEventListener("click", showMainFacilityView);

    document.querySelectorAll("[data-booking-space]").forEach((button) => {
        button.addEventListener("click", () => openTodayBooking(button.dataset.bookingSpace));
    });
    document.getElementById("bookingCalendarBackBtn")?.addEventListener("click", openBookingMenu);
    document.getElementById("bookingSpaceBackBtn")?.addEventListener("click", openBookingMenu);
    document.getElementById("bookingTimeBackBtn")?.addEventListener("click", () => {
        resetBookingSelection();
        renderBookingCalendar();
        showBookingView("bookingCalendarView");
    });
    document.getElementById("bookingPrevMonthBtn")?.addEventListener("click", () => changeBookingMonth(-1));
    document.getElementById("bookingNextMonthBtn")?.addEventListener("click", () => changeBookingMonth(1));
    document.getElementById("bookingHeadcount")?.addEventListener("input", () => {
        const input = document.getElementById("bookingHeadcount");
        const place = getBookingPlace();
        if (!input) return;
        input.max = String(place.maxHeadcount);
        if (Number(input.value || 1) > place.maxHeadcount) input.value = String(place.maxHeadcount);
        renderBookingCompanions();
    });
    document.getElementById("bookingRequestForm")?.addEventListener("submit", async (event) => {
        event.preventDefault();
        toggleLoading("bookingRequestSubmitBtn", true);
        try {
            await submitBookingRequest();
        } catch (error) {
            console.error(error);
            setAlert("facilityAlert", error.message || "예약 신청 중 문제가 발생했습니다.");
        } finally {
            toggleLoading("bookingRequestSubmitBtn", false);
            renderBookingTime();
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
        setMemberRejectChoiceVisible(true);
    });
    document.getElementById("memberRejectSignupBtn")?.addEventListener("click", () => {
        state.candidateMember = null;
        // 뒷 4자리가 같은 기존 회원이 있으므로 전체 번호로 가입 여부부터 확인한다.
        openSignupPhoneStep("", false);
    });
    document.getElementById("memberRejectHomeBtn")?.addEventListener("click", () => {
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
                    : "커넥트룸 이용 기록을 제출했습니다.",
                "success"
            );
            event.target.reset();
            document.getElementById("connectCompanionCount").value = 0;
            document.getElementById("connectCompanions").innerHTML = "";
        } catch (error) {
            console.error(error);
            setAlert("facilityAlert", error.message || "커넥트룸 기록 제출 중 문제가 발생했습니다.");
        } finally {
            toggleLoading("connectSubmitBtn", false);
        }
    });
}

function init() {
    const directPage = getDirectPage();
    if (!hasValidQrEntryToken() && directPage !== "signup") {
        showEntryGate();
        return;
    }

    if (isReservationDemoMode()) {
        setupTabs();
        setupTimeInputs();
        setupBirthDateInputs();
        setupPinInputs();
        setupNetworkWarning();
        setupStepper("careerCompanionCount", "careerCompanions");
        setupStepper("connectCompanionCount", "connectCompanions");
        setupListeners();
        setupReservationSubmitClock();
        showSection("facilityEntry");
        showToast("데모 모드입니다. 전화번호 뒷자리 0000을 입력해 주세요.", "info");
        return;
    }

    showSection("facilityEntry");
    void callScript("warmup").catch(() => {
        // Warmup is best-effort and must never block the entry form.
    });
    setupTabs();
    setupTimeInputs();
    setupBirthDateInputs();
    setupPinInputs();
    setupNetworkWarning();
    setupStepper("careerCompanionCount", "careerCompanions");
    setupStepper("connectCompanionCount", "connectCompanions");
    setupListeners();
    setupReservationSubmitClock();

    if (directPage === "signup") {
        openSignupPhoneStep();
    }
}

try {
    init();
} catch (error) {
    console.error("Initialization failed:", error);
    showToast("초기화 중 문제가 발생했습니다. 설정을 확인해 주세요.", "error");
}
