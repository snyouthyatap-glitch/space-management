import { 
    doc, getDoc, setDoc, addDoc, collection, serverTimestamp, updateDoc,
    query, where, orderBy, getDocs, limit, Timestamp
} from "https://www.gstatic.com/firebasejs/10.8.0/firebase-firestore.js";
import { getTrustedNow, toggleLoading, showAlert, calcAge, categorizeMember } from "./utils.js";
import { GOOGLE_SCRIPT_URL } from "./config.js";

// Global helper for user-related state
let isProcessingVisit = false;

export async function submitRecord(db, type, data, btnId, currentUserData, isRetry = false, targetUser = null) {
    const todayStr = getTrustedNow().toLocaleDateString('en-CA');
    const isVerified = (localStorage.getItem('lastVerifiedDate') === todayStr);
    const isAdmin = (currentUserData && currentUserData.role === 'admin');

    if (!isVerified && !isAdmin && !targetUser) {
        alert('현장 방문 인증이 필요합니다. 센터 내 비치된 QR 코드를 찍어주세요.');
        return false;
    }

    if (type === 'careerZone' || type === 'connectRoom') {
        if (data.startTime && data.startTime.length === 4 && !data.startTime.includes(':')) {
            data.startTime = data.startTime.substring(0, 2) + ':' + data.startTime.substring(2, 4);
        }
        if (data.endTime && data.endTime.length === 4 && !data.endTime.includes(':')) {
            data.endTime = data.endTime.substring(0, 2) + ':' + data.endTime.substring(2, 4);
        }
    }

    const u = targetUser || currentUserData;
    if (!u) {
        if (btnId) toggleLoading(btnId, false);
        return false;
    }

    const u_id = u.uid || u.id; // UID 필드명 호환성 처리

    if ((type === 'careerZone' || type === 'connectRoom' || type === 'printer') && u) {
        data.gender = u.gender || '미상';
        data.age = u.age || calcAge(u.birthDate) || '미상';
    }

    if (btnId) toggleLoading(btnId, true);

    const logEntry = {
        type, ...data,
        date: todayStr, // 관리자 통계 집계용 날짜 필드 추가
        userName: u.name, userId: u_id,
        email: u.email || '', sheetSent: false,
        timestamp: serverTimestamp()
    };

    let docRefs = [];
    const firebasePromise = (async () => {
        try {
            const collectionMap = { 'visit': 'logs_visit', 'printer': 'logs_printer', 'careerZone': 'logs_career', 'connectRoom': 'logs_connect' };
            const targetCol = collectionMap[type] || 'all_logs';
            const p1 = addDoc(collection(db, "records", u_id, type), logEntry);
            const p2 = addDoc(collection(db, targetCol), logEntry);
            const results = await Promise.allSettled([p1, p2]);
            results.forEach(res => { if (res.status === 'fulfilled') docRefs.push(res.value); });
            return docRefs.length > 0;
        } catch (e) { return false; }
    })();

    const sheetPromise = syncToSheet(type, data, u);
    const [firebaseSuccess, sheetSuccess] = await Promise.all([firebasePromise, sheetPromise]);

    if (sheetSuccess && firebaseSuccess) {
        for (const ref of docRefs) {
            updateDoc(ref, { sheetSent: true }).catch(() => { });
        }
    }

    if (!sheetSuccess && !isRetry) {
        const pending = JSON.parse(localStorage.getItem('pendingRecords') || '[]');
        pending.push({ type, data, timestamp: getTrustedNow().toISOString(), targetUser });
        localStorage.setItem('pendingRecords', JSON.stringify(pending));
        if (typeof window.updateSyncUI === 'function') window.updateSyncUI();
    }

    if (btnId) toggleLoading(btnId, false);
    return sheetSuccess;
}

export async function ensureVisitLog(db, targetUid, userData, currentUserData) {
    const todayStr = getTrustedNow().toLocaleDateString('en-CA');
    const visitDocRef = doc(db, "users", targetUid, "visitLog", todayStr);
    try {
        const visitDoc = await getDoc(visitDocRef);
        if (!visitDoc.exists()) {
            await setDoc(visitDocRef, { submitted: true, timestamp: serverTimestamp() });
            return await submitRecord(db, 'visit', {
                date: todayStr, name: userData.name,
                gender: userData.gender, age: userData.age || calcAge(userData.birthDate)
            }, null, currentUserData, false, { uid: targetUid, ...userData });
        }
        return true;
    } catch (e) { return false; }
}

export async function checkAndSubmitDailyVisit(db, user, currentUserData) {
    if (!user || !currentUserData || isProcessingVisit) return false;
    if (currentUserData.role === 'admin' || !currentUserData.role) return false;

    const todayStr = getTrustedNow().toLocaleDateString('en-CA');
    const isVerified = (localStorage.getItem('lastVerifiedDate') === todayStr);
    if (!isVerified) return false;

    isProcessingVisit = true;
    try {
        const success = await ensureVisitLog(db, user.uid, currentUserData, currentUserData);
        if (success) {
            document.getElementById('visitAlert')?.classList.remove('hidden');
        }
        return success;
    } finally {
        isProcessingVisit = false;
    }
}

export function setupUserListeners(db, auth, currentUserData) {
    // Tab Switching
    document.querySelectorAll('.tab').forEach(tab => {
        tab.addEventListener('click', (e) => {
            const container = e.target.closest('.card');
            container.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
            container.querySelectorAll('.tab-content').forEach(c => c.classList.add('hidden'));
            e.target.classList.add('active');
            document.getElementById(e.target.dataset.target).classList.remove('hidden');
        });
    });

    // Time Input Auto-Format (1400 -> 14:00) & Strict Validation (HH: 00-23, MM: 00-59)
    document.querySelectorAll('.time-input').forEach(input => {
        input.addEventListener('input', (e) => {
            let val = e.target.value.replace(/[^0-9]/g, '');
            if (val.length > 4) val = val.substring(0, 4);

            let hh = val.substring(0, 2);
            let mm = val.substring(2, 4);

            // HH 검증 (00-23)
            if (hh.length === 2) {
                if (parseInt(hh, 10) > 23) hh = '23';
            }
            // MM 검증 (00-59)
            if (mm.length === 2) {
                if (parseInt(mm, 10) > 59) mm = '59';
            }

            let formatted = hh;
            if (val.length > 2) {
                formatted += ':' + mm;
            }
            
            e.target.value = formatted;
        });

        // Blur 시 비정상적인 포맷 방지 (예: 1: -> 01:00)
        input.addEventListener('blur', (e) => {
            let val = e.target.value.replace(/[^0-9]/g, '');
            if (!val) return;
            
            if (val.length < 4) {
                val = val.padStart(4, '0');
                let hh = val.substring(0, 2);
                let mm = val.substring(2, 4);
                if (parseInt(hh) > 23) hh = '23';
                if (parseInt(mm) > 59) mm = '59';
                e.target.value = `${hh}:${mm}`;
            }
        });
    });

    // Background Auto Sync
    async function processPendingRecords() {
        if (!navigator.onLine) return;
        const pending = JSON.parse(localStorage.getItem('pendingRecords') || '[]');
        if (pending.length === 0) return;

        console.log(`[AutoSync] Processing ${pending.length} pending records...`);
        const newPending = [];
        
        for (const item of pending) {
            try {
                // Background sync uses a flag to prevent redundant UI alerts during auto-process
                const success = await submitRecord(db, item.type, item.data, null, currentUserData, true, item.targetUser);
                if (!success) newPending.push(item);
            } catch (e) {
                newPending.push(item);
            }
        }
        
        localStorage.setItem('pendingRecords', JSON.stringify(newPending));
        if (newPending.length === 0) {
            console.log("[AutoSync] All records synchronized.");
        }
    }

    window.addEventListener('online', processPendingRecords);
    // 앱 초기화 시 온라인 상태라면 즉시 동기화 시도
    if (navigator.onLine) processPendingRecords();

    // --- 추가된 로직: 하이브리드 스텝퍼 및 동적 필드 제어 ---
    const setupStepper = (countId, listId) => {
        const countInput = document.getElementById(countId);
        const listContainer = document.getElementById(listId);
        if (!countInput || !listContainer) return;

        const updateFields = () => {
            const count = parseInt(countInput.value, 10);
            const currentFields = listContainer.querySelectorAll('.companion-item').length;
            
            if (count > currentFields) {
                // 필드 추가
                for (let i = currentFields + 1; i <= count; i++) {
                    const item = document.createElement('div');
                    item.className = 'companion-item';
                    item.style.background = 'rgba(255, 255, 255, 0.3)';
                    item.style.padding = '10px';
                    item.style.borderRadius = '10px';
                    item.innerHTML = `
                        <div style="display: flex; gap: 8px; align-items: center;">
                            <span style="font-size: 0.75rem; font-weight: 700; color: #92400e; min-width: 50px;">동행인 ${i}</span>
                            <select class="form-control comp-gender" style="flex: 1; padding: 5px;" required>
                                <option value="">성별</option>
                                <option value="남성">남성</option>
                                <option value="여성">여성</option>
                            </select>
                            <input type="number" class="form-control comp-age" placeholder="나이" style="flex: 1; padding: 5px;" min="1" max="100" required>
                            <span style="font-size: 0.75rem; color: #92400e;">세</span>
                        </div>
                    `;
                    listContainer.appendChild(item);
                }
            } else if (count < currentFields) {
                // 필드 삭제
                for (let i = currentFields; i > count; i--) {
                    if (listContainer.lastElementChild) {
                        listContainer.lastElementChild.remove();
                    }
                }
            }
        };

        // +/- 버튼 이벤트 연결
        const container = countInput.closest('.stepper-container');
        if (container) {
            container.querySelectorAll('.stepper-btn').forEach(btn => {
                btn.addEventListener('click', () => {
                    let current = parseInt(countInput.value, 10);
                    if (btn.classList.contains('plus')) {
                        if (current < 10) countInput.value = current + 1;
                    } else if (btn.classList.contains('minus')) {
                        if (current > 0) countInput.value = current - 1;
                    }
                    updateFields();
                });
            });
        }
    };

    setupStepper('cCompCount', 'cCompList');
    setupStepper('rCompCount', 'rCompList');

    // --- 히스토리 네비게이션 및 데이터 로딩 ---
    const userArea = document.getElementById('userArea');
    const historySection = document.getElementById('historySection');

    const toggleHistory = (show, mode = 'history') => {
        if (show) {
            userArea?.classList.add('hidden');
            historySection?.classList.remove('hidden');
            
            const calendarView = document.getElementById('calendarView');
            const timelineView = document.getElementById('timelineView');
            const titleEl = document.getElementById('historySectionTitle');

            if (mode === 'calendar') {
                calendarView?.classList.remove('hidden');
                timelineView?.classList.add('hidden');
                if (titleEl) titleEl.innerHTML = '<i class="fa-solid fa-calendar-check"></i> 나의 방문 달력';
            } else {
                calendarView?.classList.add('hidden');
                timelineView?.classList.remove('hidden');
                if (titleEl) titleEl.innerHTML = '<i class="fa-solid fa-receipt"></i> 나의 공간 이용 내역';
            }

            // 데이터 로드 시작
            const uid = auth.currentUser?.uid;
            if (uid) loadUserHistory(db, uid);
            else if (currentUserData?.uid) loadUserHistory(db, currentUserData.uid);
            else loadUserHistory(db, 'preview-user-id');
        } else {
            userArea?.classList.remove('hidden');
            historySection?.classList.add('hidden');
        }
    };

    async function loadUserHistory(db, userId) {
        const timelineEl = document.getElementById('usageTimeline');
        if (!timelineEl) return;
        timelineEl.innerHTML = '<p class="text-center py-4">기록을 불러오는 중입니다...</p>';

        const types = ['visit', 'printer', 'careerZone', 'connectRoom'];
        const allLogs = [];

        try {
            await Promise.all(types.map(async (type) => {
                try {
                    const q = query(collection(db, "records", userId, type), limit(30)); 
                    const snap = await getDocs(q);
                    snap.forEach(d => {
                        const data = d.data();
                        allLogs.push({ id: d.id, logType: type, ...data });
                    });
                } catch (err) {
                    console.warn(`Query failed for type ${type}:`, err);
                }
            }));

            // 시간순 정렬 (timestamp가 없는 경우 현재 시간으로 간주)
            allLogs.sort((a, b) => {
                const tA = a.timestamp?.seconds || (new Date().getTime() / 1000);
                const tB = b.timestamp?.seconds || (new Date().getTime() / 1000);
                return tB - tA;
            });

            renderCalendar(allLogs.filter(l => l.logType === 'visit'));
            renderUsageTimeline(allLogs.filter(l => l.logType !== 'visit'));
        } catch (e) {
            console.error(e);
            timelineEl.innerHTML = '<p class="text-center py-4 text-danger">기록 로딩 실패</p>';
        }
    }

    function renderCalendar(visits) {
        const grid = document.getElementById('calendarGrid');
        const label = document.getElementById('currentMonthLabel');
        if (!grid || !label) return;

        const now = getTrustedNow();
        const year = now.getFullYear();
        const month = now.getMonth();
        label.textContent = `${year}년 ${month + 1}월`;

        const firstDay = new Date(year, month, 1).getDay();
        const lastDate = new Date(year, month + 1, 0).getDate();

        // 방문한 날짜 셋 (YYYY-MM-DD 로컬 기준 정규화)
        const visitDates = new Set();
        visits.forEach(v => {
            if (v.timestamp && v.timestamp.seconds) {
                const d = new Date(v.timestamp.seconds * 1000);
                const iso = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
                visitDates.add(iso);
            } else if (v.date) {
                visitDates.add(v.date);
            }
        });

        grid.innerHTML = '';
        // 요일 헤더
        const days = ['일','월','화','수','목','금','토'];
        days.forEach((day, idx) => {
            const el = document.createElement('div');
            el.style.fontWeight = '800'; el.style.fontSize = '0.75rem';
            el.style.color = (idx === 0) ? '#ff4040' : (idx === 6 ? '#2563eb' : '#92400e');
            el.style.opacity = '0.6';
            el.textContent = day;
            grid.appendChild(el);
        });

        // 전월 빈 칸
        for (let i = 0; i < firstDay; i++) grid.appendChild(document.createElement('div'));

        // 해당 월의 날짜들
        for (let i = 1; i <= lastDate; i++) {
            const dayEl = document.createElement('div');
            dayEl.className = 'calendar-day';
            const dateStr = `${year}-${String(month+1).padStart(2,'0')}-${String(i).padStart(2,'0')}`;
            
            if (i === now.getDate()) dayEl.classList.add('active');
            if (visitDates.has(dateStr)) dayEl.classList.add('has-visit');
            
            dayEl.textContent = i;
            grid.appendChild(dayEl);
        }
    }

    function renderUsageTimeline(logs) {
        const timelineEl = document.getElementById('usageTimeline');
        if (logs.length === 0) {
            timelineEl.innerHTML = '<p class="text-center py-4">최근 이용 내역이 없습니다.</p>';
            return;
        }

        const typeLabels = { 'printer': '🖨️ 프린터 이용', 'careerZone': '🎧 커리어존 이용', 'connectRoom': '🤝 커넥트룸 이용' };

        timelineEl.innerHTML = logs.map(log => {
            const dateVal = log.timestamp ? new Date(log.timestamp.seconds * 1000) : new Date();
            const dateStr = dateVal.toLocaleDateString('ko-KR', { month: 'long', day: 'numeric', hour: '2-digit', minute: '2-digit' });
            
            let details = '';
            if (log.logType === 'printer') details = `${log.count}매 출력`;
            else if (log.logType === 'careerZone') details = `${log.place} (${log.startTime}~${log.endTime})`;
            else if (log.logType === 'connectRoom') details = `${log.startTime}~${log.endTime}`;

            return `
                <div class="timeline-item">
                    <span class="timeline-date">${dateStr}</span>
                    <span class="timeline-content">${typeLabels[log.logType] || log.logType}</span>
                    <span class="timeline-details">${details}</span>
                </div>
            `;
        }).join('');
    }

    document.getElementById('showHistoryBtn')?.addEventListener('click', () => toggleHistory(true, 'history'));
    document.getElementById('showCalendarBtn')?.addEventListener('click', () => toggleHistory(true, 'calendar'));
    document.getElementById('closeHistoryBtn')?.addEventListener('click', () => toggleHistory(false));

    // Forms ... (Rest of listeners)
    document.getElementById('printerForm')?.addEventListener('submit', async (e) => {
        e.preventDefault();
        const count = parseInt(document.getElementById('printerCount').value, 10);
        const uid = auth.currentUser?.uid || currentUserData?.uid;
        
        if (!uid) {
            alert('사용자 정보를 찾을 수 없습니다.');
            return;
        }

        // 오늘 출력량 조회 로직
        try {
            const now = new Date();
            const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
            const startTimestamp = Timestamp.fromDate(startOfToday);

            const q = query(
                collection(db, "records", uid, "printer"),
                where("timestamp", ">=", startTimestamp)
            );
            
            const snap = await getDocs(q);
            let todayTotal = 0;
            snap.forEach(doc => {
                todayTotal += (doc.data().count || 0);
            });

            if (todayTotal >= 10) {
                alert(`오늘의 출력 한도(10장)를 모두 소진하셨습니다. (현재: ${todayTotal}장)`);
                return;
            }

            if (todayTotal + count > 10) {
                alert(`한도를 초과합니다. 오늘은 현재 ${10 - todayTotal}장까지만 추가 출력이 가능합니다.`);
                return;
            }

            const btn = e.target.querySelector('button[type="submit"]');
            btn.id = btn.id || 'btn_submit_print';
            submitRecord(db, 'printer', { count }, btn.id, currentUserData).then(success => {
                if (success) e.target.reset();
            });
        } catch (err) {
            console.error("Printer validation error:", err);
            alert('출력량을 확인하는 중 오류가 발생했습니다. 잠시 후 다시 시도해주세요.');
        }
    });

    // Career Zone Form
    document.getElementById('careerForm')?.addEventListener('submit', (e) => {
        e.preventDefault();
        const compCount = parseInt(document.getElementById('cCompCount').value, 10);
        if (compCount > 0 && !document.getElementById('careerPrivacy').checked) {
            alert('동행인의 정보 수집 동의여부를 확인해주세요.');
            return;
        }
        
        const companionRows = Array.from(e.target.querySelectorAll('.companion-row'));
        const companions = companionRows.map(row => ({
            gender: row.querySelector('.comp-gender').value,
            age: row.querySelector('.comp-age').value
        }));

        submitRecord(db, 'careerZone', {
            place: document.getElementById('cPlace').value,
            startTime: document.getElementById('cStartTime').value,
            endTime: document.getElementById('cEndTime').value,
            purpose: document.getElementById('cPurpose').value,
            companionCount: compCount,
            companions: companions
        }, btn.id, currentUserData).then(success => {
            if (success) {
                e.target.reset();
                document.getElementById('cCompList').innerHTML = '';
            }
        });
    });

    // Connect Room Form
    document.getElementById('connectForm')?.addEventListener('submit', (e) => {
        e.preventDefault();
        const compCount = parseInt(document.getElementById('rCompCount').value, 10);
        if (compCount > 0 && !document.getElementById('connectPrivacy').checked) {
            alert('동행인의 정보 수집 동의여부를 확인해주세요.');
            return;
        }
        
        const companionRows = Array.from(e.target.querySelectorAll('.companion-row'));
        const companions = companionRows.map(row => ({
            gender: row.querySelector('.comp-gender').value,
            age: row.querySelector('.comp-age').value
        }));

        submitRecord(db, 'connectRoom', {
            startTime: document.getElementById('rStartTime').value,
            endTime: document.getElementById('rEndTime').value,
            purpose: document.getElementById('rPurpose').value,
            companionCount: compCount,
            companions: companions
        }, btn.id, currentUserData).then(success => {
            if (success) {
                e.target.reset();
                document.getElementById('rCompList').innerHTML = '';
            }
        });
    });
}
/**
 * Google Sheets 실시간 연동 함수 (apps_script.js 규격 준수)
 */
async function syncToSheet(type, data, user) {
    if (!GOOGLE_SCRIPT_URL) return false;

    try {
        const todayStr = getTrustedNow().toLocaleDateString('en-CA');
        let sheetType = type;
        if (type === 'careerZone') sheetType = 'career';
        else if (type === 'connectRoom') sheetType = 'connect';

        const payload = {
            sheetType: sheetType,
            date: data.date || todayStr, // 날짜 강제 주입
            name: user.name,
            gender: user.gender || '미상',
            age: user.age || 20
        };

        if (sheetType === 'printer') {
            payload.count = data.printerCount || data.count || 1;
        } else if (sheetType === 'career' || sheetType === 'connect') {
            payload.startTime = data.startTime;
            payload.endTime = data.endTime;
            payload.purpose = data.purpose || "공간 이용";
            
            // 본인 + 동행인 데이터 통합 집계
            const users = [{ gender: payload.gender, age: payload.age }];
            if (Array.isArray(data.companions)) users.push(...data.companions);
            
            payload.companionsDetail = data.companionCount ? `동반자 ${data.companionCount}명 (${users.map(u => `${u.gender}/${u.age}`).join(', ')})` : "";
            
            // 분류기 가동 및 집계
            if (sheetType === 'career') {
                payload.place = (type === 'careerZone') ? 'AI 커리어존' : '청년 커넥트룸';
                payload.maleCount = 0; payload.femaleCount = 0;
            } else if (sheetType === 'connect') {
                payload.m20 = 0; payload.f20 = 0; payload.m30 = 0; payload.f30 = 0; payload.m19 = 0; payload.f19 = 0;
            }

            users.forEach(u => {
                const cat = categorizeMember(u.gender, u.age);
                if (sheetType === 'career') {
                    if (cat.startsWith('m')) payload.maleCount++; else payload.femaleCount++;
                } else if (sheetType === 'connect') {
                    if (payload.hasOwnProperty(cat)) payload[cat]++;
                }
            });
        }

        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 10000);
        await fetch(GOOGLE_SCRIPT_URL, {
            method: 'POST', mode: 'no-cors',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload),
            signal: controller.signal
        });
        clearTimeout(timeoutId);
        return true;
    } catch (e) {
        console.error("Sheet Sync Error:", e);
        return false;
    }
}
