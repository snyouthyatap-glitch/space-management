import { initializeApp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-app.js";
import { getAuth, onAuthStateChanged } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-auth.js";
import { getFirestore, doc, getDoc } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-firestore.js";
import { firebaseConfig } from "./config.js";
import { syncSystemTime, calcAge, getTrustedNow } from "./utils.js";
import { setupAuthListeners } from "./auth.js";
import { setupUserListeners, checkAndSubmitDailyVisit, updateUserData } from "./user.js";
import { loadAdminStats, loadPendingUsers, setupAdminListeners, updateAdminProfile } from "./admin.js";

// Initialize Firebase
const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

console.log("App Initializing...");

// Global State
let currentUserData = null;

// UI Elements (Main)
const authSection = document.getElementById('authSection');
const mainSection = document.getElementById('mainSection');
const userNameDisplay = document.getElementById('userNameDisplay');

// Role-based UI Update
async function updateUIForRole(profile) {
    const { role, status, name, uid } = profile;
    if (userNameDisplay) userNameDisplay.textContent = name;
    authSection?.classList.add('hidden');
    mainSection?.classList.remove('hidden');


    // Reset visibility
    ['adminBadge', 'adminModeBtn', 'pendingArea', 'userArea', 'adminArea', 'facilityLogSection', 'visitAlert', 'visitRequiredMsg'].forEach(id => {
        document.getElementById(id)?.classList.add('hidden');
    });

    if (role === 'admin') {
        // [Admin 전용 모드] 관리자는 관리자 영역과 사용자 기능(일지 작성)모두 노출
        document.getElementById('adminArea')?.classList.remove('hidden');
        document.getElementById('adminBadge')?.classList.remove('hidden');
        
        // 관리자용 UI에서 사용자 전환 버튼 제거 (몰입감 향상)
        document.getElementById('adminModeBtn')?.classList.add('hidden');
        const exitBtn = document.getElementById('exitAdminModeBtn');
        if (exitBtn) exitBtn.classList.add('hidden');

        loadAdminStats(db, currentUserData, 'daily');
        loadPendingUsers(db);

        // 관리자도 자가 일지 작성이 가능하도록 섹션 노출 (인증 우회 적용됨)
        document.getElementById('facilityLogSection')?.classList.remove('hidden');
        document.getElementById('visitAlert')?.classList.remove('hidden');
        document.getElementById('visitRequiredMsg')?.classList.add('hidden');
    } else if (status === 'approved') {
        // [User 모드] 관리자가 아닐 때만 사용자 기능 노출
        document.getElementById('userArea')?.classList.remove('hidden');
        
        // 방문 일지 제출 여부 확인 후 인터페이스 노출 결정
        const success = await checkAndSubmitDailyVisit(db, auth.currentUser, currentUserData);
        if (success) {
            document.getElementById('facilityLogSection')?.classList.remove('hidden');
            document.getElementById('visitRequiredMsg')?.classList.add('hidden');
            document.getElementById('visitAlert')?.classList.remove('hidden');
        } else {
            document.getElementById('visitRequiredMsg')?.classList.remove('hidden');
        }
    } else {
        document.getElementById('pendingArea')?.classList.remove('hidden');
    }

    // 오늘 날짜 표시 (필요한 경우 텍스트로만 표시)
    const dateStr = getTrustedNow().toLocaleDateString('ko-KR', { year: 'numeric', month: 'long', day: 'numeric' });
    document.querySelectorAll('.today-date').forEach(el => {
        if (el.tagName !== 'INPUT') el.textContent = dateStr;
    });
}

// Auth State Subscriber
onAuthStateChanged(auth, async (user) => {
    if (user) {
        localStorage.setItem('lastLoginDate', getTrustedNow().toLocaleDateString('en-CA'));
        try {
            const userDoc = await getDoc(doc(db, "users", user.uid));
            if (userDoc.exists()) {
                const profile = userDoc.data();
                currentUserData = {
                    uid: user.uid, email: user.email, ...profile,
                    age: profile.birthDate ? calcAge(profile.birthDate) : (profile.age || '')
                };
                updateUserData(currentUserData); // 1. 유저 데이터 전역 주입 선행
                updateAdminProfile(currentUserData); // 1.5 관리자 데이터 전역 주입
                updateUIForRole(currentUserData); // 2. 이후 UI 및 시설 일지 체크 진행
            }
        } catch (e) { console.error("Profile load fail:", e); }
    } else {
        currentUserData = null;
        updateUserData(null); // 로그아웃 시 데이터 초기화
        updateAdminProfile(null); // 로그아웃 시 데이터 초기화
        mainSection?.classList.add('hidden');
        authSection?.classList.remove('hidden');
    }
});

// Bootstrap
async function init() {
    // QR 인증 파라미터 확인 (?auth=yatap)
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('auth') === 'yatap') {
        const todayStr = getTrustedNow().toLocaleDateString('en-CA');
        localStorage.setItem('lastVerifiedDate', todayStr);
        console.log("QR 인증 성공: " + todayStr);
        
        // URL에서 파라미터 제거하여 주소창을 깔끔하게 유지
        const newUrl = window.location.pathname;
        window.history.replaceState({}, document.title, newUrl);
    }

    // 리스너를 먼저 등록하여 오프라인 상태에서도 UI 전환이 즉시 가능하도록 함
    setupAuthListeners(auth, db, currentUserData, { updateUIForRole });
    setupUserListeners(db, auth);
    setupAdminListeners(db);
    
    // 시간 동기화는 비동기로 진행 (초기화 블록 방지)
    syncSystemTime().then(() => {
        console.log("백그라운드 시간 동기화 완료");
    });

    // 관리자 전용 종료 버튼 연동
    const adminLogoutBtn = document.getElementById('logoutActionBtnAdmin');
    if (adminLogoutBtn) {
        adminLogoutBtn.addEventListener('click', () => {
            auth.signOut().then(() => {
                location.reload(); // 세션 종료 후 첫 화면으로
            });
        });
    }
}

init();

