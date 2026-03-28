import { 
    createUserWithEmailAndPassword, signInWithEmailAndPassword, 
    signOut, updateProfile, setPersistence, 
    browserLocalPersistence, browserSessionPersistence 
} from "https://www.gstatic.com/firebasejs/10.8.0/firebase-auth.js";
import { doc, setDoc, serverTimestamp } from "https://www.gstatic.com/firebasejs/10.8.0/firebase-firestore.js";
import { showAlert, toggleLoading, calcAge } from "./utils.js";

export function setupAuthListeners(auth, db, currentUserData, callbacks) {
    const { updateUIForRole } = callbacks;

    // Login Form
    const loginForm = document.getElementById('loginForm');
    if (loginForm) {
        loginForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const userId = document.getElementById('loginId').value.trim();
            const email = userId + '@yatap.local';
            const pw = document.getElementById('loginPw').value;
            const keepLoggedIn = document.getElementById('keepLoggedIn').checked;

            toggleLoading('loginBtn', true);
            try {
                const persistenceType = keepLoggedIn ? browserLocalPersistence : browserSessionPersistence;
                await setPersistence(auth, persistenceType);
                await signInWithEmailAndPassword(auth, email, pw);
            } catch (error) {
                showAlert('loginAlert', "로그인 실패. 아이디 또는 비밀번호를 확인하세요.", 'error');
            } finally {
                toggleLoading('loginBtn', false);
            }
        });
    }

    // Signup Form
    const signupForm = document.getElementById('signupForm');
    if (signupForm) {
        signupForm.addEventListener('submit', async (e) => {
            e.preventDefault();
            const userId = document.getElementById('signupId').value.trim();
            const email = userId + '@yatap.local';
            const pw = document.getElementById('signupPw').value;
            const pwConfirm = document.getElementById('signupPwConfirm').value;
            
            if (pw !== pwConfirm) {
                showAlert('signupAlert', '비밀번호가 서로 일치하지 않습니다.', 'error');
                return;
            }

            const name = document.getElementById('signupName').value;

            const birthRaw = document.getElementById('signupBirth').value.replace(/[^0-9]/g, '');
            if (birthRaw.length !== 8) {
                showAlert('signupAlert', '생년월일 8자리를 모두 입력해주세요.', 'error');
                return;
            }

            const yyyy = birthRaw.substring(0, 4);
            const mm = birthRaw.substring(4, 6);
            const dd = birthRaw.substring(6, 8);
            const mmNum = parseInt(mm, 10);
            const ddNum = parseInt(dd, 10);
            if (mmNum < 1 || mmNum > 12 || ddNum < 1 || ddNum > 31) {
                showAlert('signupAlert', '유효하지 않은 월 또는 일입니다.', 'error');
                return;
            }

            const birthDate = `${yyyy}-${mm}-${dd}`;
            const gender = document.getElementById('signupGender').value;
            const phone = document.getElementById('signupPhone').value;

            const ageString = calcAge(birthDate);
            if (!ageString) {
                showAlert('signupAlert', '유효하지 않은 생년월일입니다.', 'error');
                return;
            }
            const age = parseInt(ageString, 10);

            if (age < 20 || age > 39) {
                showAlert('signupAlert', '죄송합니다. 본 프로그램은 만 20세~39세 청년만 가입 가능합니다.', 'error');
                return;
            }

            if (!document.getElementById('signupPrivacy').checked) {
                showAlert('signupAlert', '개인정보 수집 및 이용에 동의해주세요.', 'error');
                return;
            }

            toggleLoading('signupBtn', true);
            try {
                const userCredential = await createUserWithEmailAndPassword(auth, email, pw);
                const user = userCredential.user;
                await updateProfile(user, { displayName: name });
                await setDoc(doc(db, "users", user.uid), {
                    name, birthDate, age, gender, phone, userId,
                    status: 'pending',
                    createdAt: serverTimestamp()
                });

                alert('회원가입 신청이 완료되었습니다!\n관리자 승인 후(최대 24시간 이내) 이용 가능합니다.');
                signupForm.reset();
                document.getElementById('signupView').classList.add('hidden');
                showAlert('loginAlert', '가입 신청 완료! 관리자 승인을 기다려주세요.', 'success');
            } catch (error) {
                let msg = error.message;
                if (error.code === 'auth/email-already-in-use') msg = "이미 사용 중인 아이디입니다.";
                else if (error.code === 'auth/weak-password') msg = "비밀번호는 최소 6자리 이상이어야 합니다.";
                showAlert('signupAlert', msg, 'error');
            } finally {
                toggleLoading('signupBtn', false);
            }
        });
    }

    // View Toggles
    const signupBtn = document.getElementById('showSignupBtn');
    console.log("Signup button element found:", !!signupBtn);
    signupBtn?.addEventListener('click', (e) => {
        console.log("Signup button clicked!");
        e.preventDefault();
        document.getElementById('loginView').classList.add('hidden');
        document.getElementById('signupView').classList.remove('hidden');
    });

    document.getElementById('showLoginBtn')?.addEventListener('click', () => {
        document.getElementById('signupView').classList.add('hidden');
        document.getElementById('loginView').classList.remove('hidden');
    });

    // Logout
    document.getElementById('logoutActionBtn')?.addEventListener('click', () => signOut(auth));
    document.getElementById('pendingLogoutBtn')?.addEventListener('click', () => signOut(auth));

    // Password Visibility Toggle
    document.querySelectorAll('.pw-toggle').forEach(btn => {
        btn.addEventListener('click', () => {
            const targetId = btn.getAttribute('data-target');
            const targetInput = document.getElementById(targetId);
            if (!targetInput) return;
            
            if (targetInput.type === 'password') {
                targetInput.type = 'text';
                btn.classList.replace('fa-eye', 'fa-eye-slash');
            } else {
                targetInput.type = 'password';
                btn.classList.replace('fa-eye-slash', 'fa-eye');
            }
        });
    });

    // Auto-Hyphen for Phone
    document.getElementById('signupPhone')?.addEventListener('input', (e) => {
        let val = e.target.value.replace(/[^0-9]/g, '');
        if (val.length > 3 && val.length <= 7) {
            val = val.substring(0, 3) + '-' + val.substring(3);
        } else if (val.length > 7) {
            val = val.substring(0, 3) + '-' + val.substring(3, 7) + '-' + val.substring(7, 11);
        }
        e.target.value = val;
    });

    // Auto-Spacing for Birth Date (YYYY / MM / DD)
    document.getElementById('signupBirth')?.addEventListener('input', (e) => {
        let val = e.target.value.replace(/[^0-9]/g, '');
        if (val.length > 4 && val.length <= 6) {
            val = val.substring(0, 4) + ' / ' + val.substring(4);
        } else if (val.length > 6) {
            val = val.substring(0, 4) + ' / ' + val.substring(4, 6) + ' / ' + val.substring(6, 8);
        }
        e.target.value = val;
    });
}
