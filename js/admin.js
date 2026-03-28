import { 
    collection, query, where, getDocs, updateDoc, doc, getDoc, limit, onSnapshot, serverTimestamp 
} from "https://www.gstatic.com/firebasejs/10.8.0/firebase-firestore.js";
import { getTrustedNow } from "./utils.js";
import { submitRecord, ensureVisitLog } from "./user.js";

export async function loadAdminStats(db, currentUserData, period = 'daily') {
    if (!currentUserData || currentUserData.role !== 'admin') return;

    const refreshBtn = document.getElementById('refreshStatsBtn');
    let originalBtnContent = '';
    if (refreshBtn) {
        originalBtnContent = refreshBtn.innerHTML;
        refreshBtn.disabled = true;
        refreshBtn.innerHTML = '<i class="fa-solid fa-sync fa-spin"></i>';
    }

    const now = getTrustedNow();
    const stats = { visit: 0, printer: 0, career: 0, connect: 0 };
    let startDate = new Date(now);
    startDate.setHours(0, 0, 0, 0);

    if (period === 'weekly') {
        const day = startDate.getDay();
        const diff = startDate.getDate() - day + (day === 0 ? -6 : 1);
        startDate.setDate(diff);
    } else if (period === 'monthly') {
        startDate.setDate(1);
    }

    const startDateStr = startDate.toLocaleDateString('en-CA');
    const todayStr = now.toLocaleDateString('en-CA');

    try {
        const logCollections = ['logs_visit', 'logs_printer', 'logs_career', 'logs_connect'];
        const queryPromises = logCollections.map(colName => {
            let q = (period === 'daily') ? 
                query(collection(db, colName), where("date", "==", todayStr)) : 
                query(collection(db, colName), where("date", ">=", startDateStr));
            return getDocs(q);
        });

        const snapshots = await Promise.all(queryPromises);
        snapshots.forEach((querySnapshot, idx) => {
            const colName = logCollections[idx];
            querySnapshot.forEach((doc) => {
                const data = doc.data();
                if (colName === 'logs_visit') stats.visit++;
                else if (colName === 'logs_printer') stats.printer++;
                else if (colName === 'logs_career') stats.career++;
                else if (colName === 'logs_connect') stats.connect++;
            });
        });

        document.getElementById('statTotalVisit').textContent = `${stats.visit}명`;
        document.getElementById('statTotalPrint').textContent = `${stats.printer}건`;

        // UI에 존재하지 않는 항목은 조건부 업데이트 (에러 방지)
        const careerEl = document.getElementById('statTotalCareer');
        if (careerEl) careerEl.textContent = `${stats.career}건`;
        
        const connectEl = document.getElementById('statTotalConnect');
        if (connectEl) connectEl.textContent = `${stats.connect}건`;
    } finally {
        refreshBtn.disabled = false;
        refreshBtn.innerHTML = originalBtnContent;
    }
}

export async function loadPendingUsers(db) {
    const listEl = document.getElementById('pendingUserList');
    const searchInput = document.getElementById('adminUserSearch');
    const searchTerm = searchInput ? searchInput.value.trim().toLowerCase() : '';

    try {
        let q = searchTerm ? 
            query(collection(db, "users"), limit(100)) : 
            query(collection(db, "users"), where("status", "==", "pending"), limit(50));

        const querySnapshot = await getDocs(q);
        const users = [];
        querySnapshot.forEach(docSnap => users.push({ id: docSnap.id, ...docSnap.data() }));
        renderUserList(users, listEl, false, db);
    } catch (e) {
        console.error("loadPendingUsers Error:", e);
        listEl.innerHTML = `<tr><td colspan="4" class="text-center">로딩 오류 (권한 부족)<br><span style="font-size:0.7rem; color:red;">운영 중인 관리자 계정으로 로그인해야 접근 가능합니다.</span></td></tr>`;
    }
}

// --- 회원 검색 및 대리 작성 추가 ---
let selectedProxyUser = null;

export async function searchUsers(db, searchTerm) {
    const resultsEl = document.getElementById('adminMemberSearchResults');
    try {
        const q = query(collection(db, "users"), limit(100)); // 검색 범위 확대
        const querySnapshot = await getDocs(q);
        const results = [];
        querySnapshot.forEach(docSnap => {
            const u = docSnap.data();
            if (u.name.includes(searchTerm)) results.push({ id: docSnap.id, ...u });
        });
        renderSearchResults(results, resultsEl);
    } catch (e) {
        resultsEl.innerHTML = '<p class="text-center p-3">검색 중 오류가 발생했습니다.</p>';
    }
}

function renderSearchResults(users, resultsEl) {
    if (users.length === 0) {
        resultsEl.innerHTML = '<div class="text-center p-5 opacity-50"><i class="fa-solid fa-user-slash mb-2" style="font-size:2rem;"></i><p>검색 결과가 없습니다.</p></div>';
        return;
    }
    
    let html = '<div class="results-grid" style="display:grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap:12px;">';
    users.forEach(u => {
        html += `
            <div class="search-result-card fade-in" data-uid="${u.id}" data-name="${u.name}" 
                 style="background:white; padding:16px; border-radius:18px; border:2px solid #f3f4f6; cursor:pointer; transition:all 0.2s;">
                <div style="display:flex; align-items:center; gap:12px;">
                    <div style="width:36px; height:36px; background:#fef3c7; border-radius:10px; display:flex; align-items:center; justify-content:center; color:#92400e;">
                        <i class="fa-solid fa-user"></i>
                    </div>
                    <div>
                        <div style="font-weight:800; color:#451a03;">${u.name}</div>
                        <div style="font-size:0.75rem; color:#92400e; opacity:0.6;">${u.phone || '연락처 미등록'}</div>
                    </div>
                </div>
            </div>
        `;
    });
    html += '</div>';
    resultsEl.innerHTML = html;

    resultsEl.querySelectorAll('.search-result-card').forEach(card => {
        card.addEventListener('click', (e) => {
            const targetCard = e.currentTarget;
            const uid = targetCard.dataset.uid;
            const name = targetCard.dataset.name;
            const userObj = users.find(u => u.id === uid);
            
            selectedProxyUser = userObj;
            document.getElementById('proxyTargetName').textContent = name;
            document.getElementById('proxyEntryArea').classList.remove('hidden');
            
            // 오늘 날짜 기본값 세팅 (YYYY-MM-DD 형식)
            const proxyDateInput = document.getElementById('proxyDateInput');
            if (proxyDateInput) proxyDateInput.value = getTrustedNow().toLocaleDateString('en-CA');
            
            // 시각적 피드백 (선택 효과)
            resultsEl.querySelectorAll('.search-result-card').forEach(c => {
                c.style.borderColor = '#f3f4f6';
                c.style.background = 'white';
            });
            targetCard.style.borderColor = 'var(--primary-main)';
            targetCard.style.background = '#fffdf5';
            
            // 폼으로 스크롤
            document.getElementById('proxyEntryArea').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
        });
    });
}

function renderUserList(users, listEl, isMock, db) {
    let html = '';
    users.forEach((u) => {
        const isPending = u.status === 'pending';
        html += `
            <tr id="row-${u.id}">
                <td><strong>${u.name}</strong><br><span style="color:var(--gray-text); font-size:0.7rem;">${u.userId || ''}</span></td>
                <td>${u.gender || '-'} | ${u.age || '-'}세<br><span style="color:var(--gray-text); font-size:0.7rem;">${u.phone || '-'}</span></td>
                <td>${u.createdAt ? new Date(u.createdAt.seconds * 1000).toLocaleDateString() : '-'}</td>
                <td>
                    <div class="d-flex gap-1" style="display:flex; gap:4px;">
                        ${isPending ? `
                            <button class="btn btn-primary btn-sm approve-btn" data-uid="${u.id}" style="padding:4px 8px; font-size:0.75rem;">승인</button>
                            <button class="btn btn-danger btn-sm reject-btn" data-uid="${u.id}" style="padding:4px 8px; font-size:0.75rem;">거절</button>
                        ` : `<span class="badge ${u.status === 'approved' ? 'badge-success' : 'badge-danger'}">${u.status === 'approved' ? '승인됨' : '거절됨'}</span>`}
                    </div>
                </td>
            </tr>
        `;
    });
    listEl.innerHTML = html || '<tr><td colspan="4" class="text-center">조건에 맞는 사용자가 없습니다.</td></tr>';

    // 승인 이벤트 리스너 등록
    listEl.querySelectorAll('.approve-btn').forEach(btn => {
        btn.addEventListener('click', async (e) => {
            const uid = e.target.dataset.uid;
            if (confirm('이 사용자를 승인하시겠습니까?')) {
                if (isMock) {
                    alert('테스트 모드: 승인 완료 알림이 시뮬레이션됩니다.');
                    document.getElementById(`row-${uid}`).remove();
                } else if (db) {
                    const userRef = doc(db, "users", uid);
                    const userDoc = await getDoc(userRef);
                    if (userDoc.exists()) {
                        await updateDoc(userRef, { status: 'approved' });
                        if (typeof ensureVisitLog === 'function') await ensureVisitLog(db, uid, userDoc.data(), userDoc.data());
                        loadPendingUsers(db);
                    }
                }
            }
        });
    });

    // 거절 이벤트 리스너 등록
    listEl.querySelectorAll('.reject-btn').forEach(btn => {
        btn.addEventListener('click', async (e) => {
            const uid = e.target.dataset.uid;
            if (confirm('가입 신청을 거절하시겠습니까?')) {
                if (isMock) {
                    alert('테스트 모드: 거절 완료 알림이 시뮬레이션됩니다.');
                    document.getElementById(`row-${uid}`).remove();
                } else if (db) {
                    const userRef = doc(db, "users", uid);
                    await updateDoc(userRef, { status: 'rejected' });
                    alert('가입 신청이 거절되었습니다.');
                    loadPendingUsers(db);
                }
            }
        });
    });
}

export function setupAdminListeners(db) {
    // Period selection
    document.querySelectorAll('.period-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const period = e.currentTarget.dataset.period;
            document.querySelectorAll('.period-btn').forEach(b => b.classList.remove('active'));
            e.currentTarget.classList.add('active');
            
            // 전역 변수가 아닌 파라미터로 명시적 전달
            loadAdminStats(db, { role: 'admin' }, period);
        });
    });

    // Admin Tabs
    document.querySelectorAll('.admin-tab').forEach(tab => {
        tab.addEventListener('click', (e) => {
            document.querySelectorAll('.admin-tab').forEach(t => t.classList.remove('active'));
            document.querySelectorAll('.admin-tab-content').forEach(c => c.classList.add('hidden'));
            e.target.classList.add('active');
            document.getElementById(e.target.dataset.target).classList.remove('hidden');
        });
    });

    // Admin Mode Toggles
    document.getElementById('adminModeBtn')?.addEventListener('click', () => {
        document.getElementById('userArea').classList.add('hidden');
        document.getElementById('facilityLogSection').classList.add('hidden');
        document.getElementById('adminArea').classList.remove('hidden');
        loadAdminStats(db, null, 'daily');
        loadPendingUsers(db);
    });

    document.getElementById('exitAdminModeBtn')?.addEventListener('click', () => {
        document.getElementById('adminArea').classList.add('hidden');
        document.getElementById('userArea').classList.remove('hidden');
        document.getElementById('facilityLogSection').classList.remove('hidden');
    });

    // Refresh button
    document.getElementById('refreshStatsBtn')?.addEventListener('click', () => {
        loadAdminStats(db, { role: 'admin' }, 'daily');
    });

    // Member Search listeners
    const searchInput = document.getElementById('adminMemberSearchInput');
    const searchBtn = document.getElementById('adminMemberSearchBtn');
    
    if (searchBtn && searchInput) {
        searchBtn.addEventListener('click', () => {
            const term = searchInput.value.trim();
            if (term) searchUsers(db, term);
        });
        searchInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') searchBtn.click();
        });
    }

    // Proxy form submission
    const proxyForm = document.getElementById('proxyFacilityForm');
    if (proxyForm) {
        // 서비스 타입 변경 시 필드 토글
        proxyForm.querySelectorAll('input[name="proxyType"]').forEach(radio => {
            radio.addEventListener('change', (e) => {
                const type = e.target.value;
                document.querySelectorAll('.proxy-fields-panel').forEach(p => p.classList.add('hidden'));
                
                if (type === 'visit') {
                    document.getElementById('proxy-visit-fields').classList.remove('hidden');
                } else if (type === 'printer') {
                    document.getElementById('proxy-printer-fields').classList.remove('hidden');
                } else {
                    document.getElementById('proxy-room-fields').classList.remove('hidden');
                    // 룸 이용 초기화 시 동행인 리스트 비우기
                    const companionsList = document.getElementById('proxyCompanionsList');
                    if (companionsList) companionsList.innerHTML = '';
                }
            });
        });

        // [New] 동행인 추가 버튼 연동
        const addCompanionBtn = document.getElementById('addProxyCompanionBtn');
        const companionsList = document.getElementById('proxyCompanionsList');
        if (addCompanionBtn && companionsList) {
            addCompanionBtn.addEventListener('click', () => {
                const row = document.createElement('div');
                row.className = 'proxy-companion-row';
                row.innerHTML = `
                    <select class="proxy-field-input cp-gender" required>
                        <option value="남성">남성</option>
                        <option value="여성">여성</option>
                    </select>
                    <input type="number" class="proxy-field-input cp-age" placeholder="나이" min="1" max="100" required>
                    <button type="button" class="btn-remove-companion" title="삭제">
                        <i class="fa-solid fa-xmark"></i>
                    </button>
                `;
                
                // 삭제 이벤트
                row.querySelector('.btn-remove-companion').addEventListener('click', () => row.remove());
                companionsList.appendChild(row);
            });
        }

        proxyForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            if (!selectedProxyUser) return;

            const formData = new FormData(proxyForm);
            const type = formData.get('proxyType');
            const selectedDate = formData.get('proxyDate') || getTrustedNow().toLocaleDateString('en-CA');
            let data = {
                date: selectedDate
            };

            // 동행인 상세 데이터 수집
            const companions = [];
            document.querySelectorAll('.proxy-companion-row').forEach(row => {
                const gender = row.querySelector('.cp-gender').value;
                const age = row.querySelector('.cp-age').value;
                if (gender && age) {
                    companions.push({ gender, age });
                }
            });

            // 타입별 데이터 구성
            if (type === 'printer') {
                data.printerCount = parseInt(formData.get('printerCount'));
            } else if (type === 'careerZone' || type === 'connectRoom') {
                data.startTime = formData.get('startTime');
                data.endTime = formData.get('endTime');
                data.purpose = formData.get('purpose');
                data.companions = companions; // 상세 정보 배열 전달
                data.companionCount = companions.length; // 총 인원
            } else if (type === 'visit') {
                data.note = "관리자 대리 방문 기록";
            }

            const submitBtnId = 'proxySubmitBtn';
            const success = await submitRecord(db, type, data, submitBtnId, { role: 'admin' }, false, selectedProxyUser);
            
            if (success) {
                const typeName = { visit: '방문', printer: '프린터', careerZone: '커리어존', connectRoom: '커넥트룸' }[type];
                alert(`${selectedProxyUser.name} 회원님의 [${typeName}] 일지가 제출되었습니다.`);
                proxyForm.reset();
                document.getElementById('proxyEntryArea').classList.add('hidden');
                selectedProxyUser = null;
            } else {
                alert('대리 제출에 실패했습니다. 다시 시도해 주세요.');
            }
        });
    }
}
