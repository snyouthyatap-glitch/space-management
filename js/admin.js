import { GOOGLE_SCRIPT_URL } from "./config.js?v=20260724-3";

const SESSION_KEY = "space-management-admin-session";
const PRINTER_REQUEST_KEY = "space-management-admin-printer-request";
const REQUEST_TIMEOUT_MS = 25000;

const state = {
    sessionToken: "",
    subjectType: "member",
    member: null,
    guestToken: "",
    pendingPrinterRequest: null
};

const logTypeLabels = {
    visit: "방문일지",
    printer: "프린터",
    careerRoom1: "커리어존 Room1",
    careerRoom2: "커리어존 Room2",
    connectRoom: "커넥트룸"
};

function normalizeDigits(value) {
    return String(value || "").replace(/\D/g, "");
}

function getSeoulDateParts() {
    const parts = new Intl.DateTimeFormat("en-CA", {
        timeZone: "Asia/Seoul",
        year: "numeric",
        month: "2-digit",
        day: "2-digit"
    }).formatToParts(new Date());
    const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
    return {
        year: values.year,
        month: values.month,
        day: values.day
    };
}

function todayString() {
    const { year, month, day } = getSeoulDateParts();
    return `${year}-${month}-${day}`;
}

function formatDate(date) {
    const [year, month, day] = String(date).split("-");
    return `${year}.${month}.${day}`;
}

function formatPhone(phone) {
    const digits = normalizeDigits(phone);
    return digits.length === 11
        ? `${digits.slice(0, 3)}-${digits.slice(3, 7)}-${digits.slice(7)}`
        : digits;
}

function setLoading(buttonId, loading, loadingText) {
    const button = document.getElementById(buttonId);
    if (!button) return;
    const text = button.querySelector(".button-text");
    const spinner = button.querySelector(".spinner");
    if (loading) {
        button.dataset.originalText = text?.textContent || "";
        if (text && loadingText) text.textContent = loadingText;
        button.disabled = true;
        button.classList.add("loading");
        if (spinner) spinner.style.display = "inline-block";
    } else {
        if (text && button.dataset.originalText) text.textContent = button.dataset.originalText;
        button.disabled = false;
        button.classList.remove("loading");
        if (spinner) spinner.style.display = "none";
    }
}

function showToast(message, type = "success") {
    const container = document.getElementById("toastContainer");
    const toast = document.createElement("div");
    toast.className = `toast ${type}`;
    toast.setAttribute("role", type === "error" ? "alert" : "status");
    toast.textContent = message;
    container.appendChild(toast);
    window.setTimeout(() => toast.remove(), 4500);
}

function setMessage(id, message = "") {
    const element = document.getElementById(id);
    if (element) element.textContent = message;
}

function getSessionValue(key) {
    try {
        return sessionStorage.getItem(key) || "";
    } catch {
        return "";
    }
}

function setSessionValue(key, value) {
    try {
        sessionStorage.setItem(key, value);
    } catch {
        // The current in-memory session remains usable.
    }
}

function removeSessionValue(key) {
    try {
        sessionStorage.removeItem(key);
    } catch {
        // No action required.
    }
}

function callScript(action, payload = {}) {
    if (!GOOGLE_SCRIPT_URL) {
        return Promise.reject(new Error("Apps Script 주소가 설정되지 않았습니다."));
    }

    return new Promise((resolve, reject) => {
        const callbackName = `spaceAdminCallback_${Date.now()}_${Math.floor(Math.random() * 100000)}`;
        const script = document.createElement("script");
        const timeoutId = window.setTimeout(() => {
            cleanup();
            reject(new Error("요청 시간이 초과되었습니다."));
        }, REQUEST_TIMEOUT_MS);

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
            adminSession: state.sessionToken
        }));
        url.searchParams.set("callback", callbackName);
        script.src = url.toString();
        document.body.appendChild(script);
    });
}

async function sha256Hex(value) {
    const data = new TextEncoder().encode(String(value));
    const digest = await crypto.subtle.digest("SHA-256", data);
    return Array.from(new Uint8Array(digest))
        .map((byte) => byte.toString(16).padStart(2, "0"))
        .join("");
}

function showWorkspace() {
    document.getElementById("loginSection").classList.add("hidden");
    document.getElementById("workspaceSection").classList.remove("hidden");
    document.getElementById("logoutBtn").classList.remove("hidden");
    document.getElementById(state.subjectType === "guest" ? "guestGender" : "memberPhone").focus();
}

function logout(message = "") {
    state.sessionToken = "";
    state.subjectType = "member";
    state.member = null;
    state.guestToken = "";
    removeSessionValue(SESSION_KEY);
    document.querySelector('input[name="subjectMode"][value="member"]').checked = true;
    document.getElementById("memberInputPanel").classList.remove("hidden");
    document.getElementById("guestInputPanel").classList.add("hidden");
    resetMember();
    document.getElementById("workspaceSection").classList.add("hidden");
    document.getElementById("logoutBtn").classList.add("hidden");
    document.getElementById("loginSection").classList.remove("hidden");
    document.getElementById("adminPassword").value = "";
    setMessage("loginMessage", message);
    document.getElementById("adminPassword").focus();
}

async function login(password) {
    const challengeResponse = await callScript("adminChallenge");
    const passwordHash = await sha256Hex(password);
    const proof = await sha256Hex(`${challengeResponse.challenge}:${passwordHash}`);
    const response = await callScript("adminLogin", {
        challenge: challengeResponse.challenge,
        proof
    });
    state.sessionToken = response.sessionToken;
    setSessionValue(SESSION_KEY, state.sessionToken);
    setMessage("loginMessage");
    showWorkspace();
}

function setManualLogEnabled(enabled) {
    document.getElementById("manualLogFields").disabled = !enabled;
    document.getElementById("manualLogSection").classList.toggle("disabled-section", !enabled);
}

function resetMember() {
    state.member = null;
    document.getElementById("memberSearchForm").classList.remove("hidden");
    document.getElementById("memberResult").classList.add("hidden");
    document.getElementById("changeMemberBtn").classList.add("hidden");
    setManualLogEnabled(state.subjectType === "guest");
    document.getElementById("memberPhone").value = "";
    setMessage("memberMessage");
    setMessage("manualLogMessage");
}

function setSubjectMode(subjectType) {
    state.subjectType = subjectType === "guest" ? "guest" : "member";
    state.guestToken = state.subjectType === "guest" ? createRequestId() : "";
    state.pendingPrinterRequest = null;
    removeSessionValue(PRINTER_REQUEST_KEY);
    resetMember();

    const isGuest = state.subjectType === "guest";
    document.getElementById("memberInputPanel").classList.toggle("hidden", isGuest);
    document.getElementById("guestInputPanel").classList.toggle("hidden", !isGuest);
    document.getElementById("guestGender").value = "";
    document.getElementById("guestAge").value = "";
    document.getElementById(isGuest ? "guestGender" : "memberPhone").focus();
}

function renderMember(member) {
    state.subjectType = "member";
    state.member = member;
    document.getElementById("memberName").textContent = member.name;
    document.getElementById("memberMeta").textContent =
        `${member.gender || "-"} · ${member.birthDate ? `${member.age}세` : "생년월일 미등록"}`;
    document.getElementById("memberPhoneText").textContent = formatPhone(member.phone);
    document.getElementById("memberSearchForm").classList.add("hidden");
    document.getElementById("memberResult").classList.remove("hidden");
    document.getElementById("changeMemberBtn").classList.remove("hidden");
    setManualLogEnabled(true);
    setMessage("memberMessage");
    document.getElementById("logType").focus();
}

function getGuestSubject() {
    const gender = document.getElementById("guestGender").value;
    const age = Number(document.getElementById("guestAge").value);
    if (!["남성", "여성"].includes(gender)) {
        throw new Error("비회원 성별을 선택해 주세요.");
    }
    if (!Number.isInteger(age) || age < 1 || age > 100) {
        throw new Error("비회원 만 나이를 1세부터 100세 사이로 입력해 주세요.");
    }
    return {
        name: "비회원",
        gender,
        age
    };
}

function isRoomType(logType) {
    return ["careerRoom1", "careerRoom2", "connectRoom"].includes(logType);
}

function updateConditionalFields() {
    const logType = document.getElementById("logType").value;
    const printerFields = document.getElementById("printerFields");
    const roomFields = document.getElementById("roomFields");
    printerFields.classList.toggle("hidden", logType !== "printer");
    roomFields.classList.toggle("hidden", !isRoomType(logType));

    document.getElementById("printerCount").required = logType === "printer";
    ["startTime", "endTime", "purpose"].forEach((id) => {
        document.getElementById(id).required = isRoomType(logType);
    });
}

function renderCompanions() {
    const count = Math.max(0, Math.min(10, Number(document.getElementById("companionCount").value || 0)));
    document.getElementById("companionCount").value = String(count);
    const container = document.getElementById("companionFields");
    container.replaceChildren();
    for (let index = 1; index <= count; index += 1) {
        const row = document.createElement("div");
        row.className = "companion-row";

        const label = document.createElement("strong");
        label.textContent = `추가 인원 ${index}`;

        const gender = document.createElement("select");
        gender.className = "companion-gender";
        gender.required = true;
        gender.setAttribute("aria-label", `추가 인원 ${index} 성별`);
        gender.innerHTML = '<option value="">성별</option><option value="남성">남성</option><option value="여성">여성</option>';

        const age = document.createElement("input");
        age.className = "companion-age";
        age.type = "number";
        age.min = "1";
        age.max = "100";
        age.required = true;
        age.placeholder = "나이";
        age.setAttribute("aria-label", `추가 인원 ${index} 나이`);

        row.append(label, gender, age);
        container.appendChild(row);
    }
}

function collectCompanions() {
    return Array.from(document.querySelectorAll(".companion-row")).map((row) => {
        const gender = row.querySelector(".companion-gender").value;
        const age = Number(row.querySelector(".companion-age").value);
        if (!gender || !Number.isInteger(age) || age < 1 || age > 100) {
            throw new Error("추가 인원의 성별과 나이를 모두 입력해 주세요.");
        }
        return { gender, age };
    });
}

function createRequestId() {
    if (crypto.randomUUID) return crypto.randomUUID();
    return `admin-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function getPrinterRequest(memberId, date, count) {
    const signature = `${memberId}:${date}:${count}`;
    let pending = state.pendingPrinterRequest;
    if (!pending) {
        try {
            pending = JSON.parse(getSessionValue(PRINTER_REQUEST_KEY) || "null");
        } catch {
            pending = null;
        }
    }
    if (pending?.signature === signature && pending.requestId) {
        state.pendingPrinterRequest = pending;
        return pending.requestId;
    }
    pending = { signature, requestId: createRequestId() };
    state.pendingPrinterRequest = pending;
    setSessionValue(PRINTER_REQUEST_KEY, JSON.stringify(pending));
    return pending.requestId;
}

function clearPrinterRequest(requestId) {
    if (state.pendingPrinterRequest?.requestId !== requestId) return;
    state.pendingPrinterRequest = null;
    removeSessionValue(PRINTER_REQUEST_KEY);
}

function addRecentEntry(member, date, logType, response) {
    document.getElementById("emptyRecentRow")?.remove();
    const row = document.createElement("tr");
    const resultText = response.logCreated === false ? "기존 기록 확인" : "저장 완료";
    [
        new Intl.DateTimeFormat("ko-KR", {
            timeZone: "Asia/Seoul",
            hour: "2-digit",
            minute: "2-digit"
        }).format(new Date()),
        member.name,
        formatDate(date),
        logTypeLabels[logType],
        resultText
    ].forEach((value, index) => {
        const cell = document.createElement("td");
        cell.textContent = value;
        if (index === 4) cell.className = "status-success";
        row.appendChild(cell);
    });
    document.getElementById("recentBody").prepend(row);
}

function resetLogDetails() {
    document.getElementById("printerCount").value = "1";
    document.getElementById("startTime").value = "";
    document.getElementById("endTime").value = "";
    document.getElementById("purpose").value = "";
    document.getElementById("companionCount").value = "0";
    renderCompanions();
}

async function submitManualLog() {
    const isGuest = state.subjectType === "guest";
    if (!isGuest && !state.member) throw new Error("회원을 먼저 조회해 주세요.");
    const subject = isGuest ? getGuestSubject() : state.member;
    const logType = document.getElementById("logType").value;
    const date = document.getElementById("usageDate").value;
    const data = {};
    let requestId = createRequestId();

    if (logType === "printer") {
        data.count = Number(document.getElementById("printerCount").value || 0);
        requestId = getPrinterRequest(
            isGuest ? state.guestToken : state.member.id,
            date,
            data.count
        );
    } else if (isRoomType(logType)) {
        data.startTime = document.getElementById("startTime").value;
        data.endTime = document.getElementById("endTime").value;
        data.purpose = document.getElementById("purpose").value.trim();
        data.companions = collectCompanions();
    }

    const response = await callScript("adminSubmitManualLog", {
        subjectType: state.subjectType,
        memberId: isGuest ? "" : state.member.id,
        guest: isGuest ? {
            token: state.guestToken,
            gender: subject.gender,
            age: subject.age
        } : null,
        logType,
        date,
        data,
        requestId
    });
    if (logType === "printer") clearPrinterRequest(requestId);
    addRecentEntry(subject, date, logType, response);
    resetLogDetails();
    if (isGuest) setSubjectMode("guest");

    if (response.linkedReservation) {
        showToast("네이버 예약과 연결하여 일지를 저장했습니다.");
    } else if (response.logCreated === false) {
        showToast("이미 등록된 기록을 확인했습니다.");
    } else {
        showToast("수동 일지를 저장했습니다.");
    }
}

function handleAuthError(error) {
    const message = error?.message || "요청 처리 중 문제가 발생했습니다.";
    if (message.includes("관리자 인증이 만료")) {
        logout(message);
        return true;
    }
    return false;
}

function setupListeners() {
    document.getElementById("loginForm").addEventListener("submit", async (event) => {
        event.preventDefault();
        setMessage("loginMessage");
        setLoading("loginBtn", true, "확인 중...");
        try {
            await login(document.getElementById("adminPassword").value);
        } catch (error) {
            setMessage("loginMessage", error.message || "관리자 인증에 실패했습니다.");
        } finally {
            setLoading("loginBtn", false);
        }
    });

    document.getElementById("logoutBtn").addEventListener("click", () => logout());

    document.querySelectorAll('input[name="subjectMode"]').forEach((input) => {
        input.addEventListener("change", (event) => {
            if (event.target.checked) setSubjectMode(event.target.value);
        });
    });

    document.getElementById("memberSearchForm").addEventListener("submit", async (event) => {
        event.preventDefault();
        setMessage("memberMessage");
        const phone = normalizeDigits(document.getElementById("memberPhone").value);
        if (!/^010\d{8}$/.test(phone)) {
            setMessage("memberMessage", "휴대폰 번호 11자리를 정확히 입력해 주세요.");
            return;
        }
        setLoading("memberSearchBtn", true, "조회 중...");
        try {
            const response = await callScript("adminLookupMember", { phone });
            renderMember(response.member);
        } catch (error) {
            if (!handleAuthError(error)) {
                setMessage("memberMessage", error.message || "회원 조회에 실패했습니다.");
            }
        } finally {
            setLoading("memberSearchBtn", false);
        }
    });

    document.getElementById("memberPhone").addEventListener("input", (event) => {
        event.target.value = normalizeDigits(event.target.value).slice(0, 11);
    });

    document.getElementById("changeMemberBtn").addEventListener("click", () => {
        resetMember();
        document.getElementById("memberPhone").focus();
    });

    document.getElementById("logType").addEventListener("change", updateConditionalFields);
    document.getElementById("companionCount").addEventListener("input", renderCompanions);

    document.getElementById("manualLogForm").addEventListener("submit", async (event) => {
        event.preventDefault();
        setMessage("manualLogMessage");
        setLoading("manualSubmitBtn", true, "저장 중...");
        try {
            await submitManualLog();
        } catch (error) {
            if (!handleAuthError(error)) {
                setMessage("manualLogMessage", error.message || "수동 일지 저장에 실패했습니다.");
                showToast(error.message || "수동 일지 저장에 실패했습니다.", "error");
            }
        } finally {
            setLoading("manualSubmitBtn", false);
        }
    });
}

async function init() {
    const today = todayString();
    document.getElementById("todayText").textContent = formatDate(today);
    document.getElementById("usageDate").value = today;
    document.getElementById("usageDate").max = today;
    setupListeners();
    updateConditionalFields();
    renderCompanions();
    setSubjectMode("member");

    state.sessionToken = getSessionValue(SESSION_KEY);
    if (!state.sessionToken) {
        document.getElementById("adminPassword").focus();
        return;
    }
    try {
        await callScript("adminSessionStatus");
        showWorkspace();
    } catch {
        logout("관리자 인증이 만료되었습니다. 다시 로그인해 주세요.");
    }
}

init();
