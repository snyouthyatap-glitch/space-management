let systemTimeOffset = 0;

export async function syncSystemTime() {
    const timeApis = [
        'https://worldtimeapi.org/api/timezone/Asia/Seoul',
        'https://timeapi.io/api/Time/current/zone?timeZone=Asia/Seoul'
    ];

    for (const api of timeApis) {
        try {
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), 3000);
            const response = await fetch(api, { signal: controller.signal });
            clearTimeout(timeoutId);
            
            if (response.ok) {
                const data = await response.json();
                const serverTimeStr = data.datetime || data.dateTime || data.currentDateTime;
                const serverTime = new Date(serverTimeStr).getTime();
                const localTime = Date.now();
                systemTimeOffset = serverTime - localTime;
                console.log("서버 시간 동기화 완료 (" + api + ")");
                return;
            }
        } catch (e) {
            console.warn(api + " 동기화 실패 또는 타임아웃:", e);
        }
    }
    console.warn("모든 시간 서버 동기화 실패, 로컬 시간 사용");
}

export function getTrustedNow() {
    return new Date(Date.now() + systemTimeOffset);
}

export function showAlert(elementId, message, type = 'error') {
    const el = document.getElementById(elementId);
    if (!el) return;
    el.textContent = message;
    el.className = `alert alert-${type}`;
    el.classList.remove('hidden');
    setTimeout(() => el.classList.add('hidden'), 5000);
}

export function toggleLoading(btnId, isLoading) {
    const btn = document.getElementById(btnId);
    if (!btn) return;
    const btnText = btn.querySelector('.btn-text');
    const spinner = btn.querySelector('.spinner');

    if (isLoading) {
        btn.disabled = true;
        btn.classList.add('loading');
        if (btnText) {
            btn.setAttribute('data-original-text', btnText.textContent);
            btnText.textContent = '제출 중...';
        }
        if (spinner) spinner.style.display = 'inline-block';
    } else {
        btn.disabled = false;
        btn.classList.remove('loading');
        if (btnText) {
            const originalText = btn.getAttribute('data-original-text');
            if (originalText) btnText.textContent = originalText;
        }
        if (spinner) spinner.style.display = 'none';
    }
}

export function calcAge(birthStr) {
    if (!birthStr) return '';
    const parts = birthStr.split('-');
    if (parts.length === 3) {
        const m = parseInt(parts[1], 10);
        const d = parseInt(parts[2], 10);
        if (m < 1 || m > 12 || d < 1 || d > 31) return null;
    }

    const birth = new Date(birthStr);
    if (isNaN(birth.getTime())) return null;

    const today = getTrustedNow();
    let age = today.getFullYear() - birth.getFullYear();
    const m = today.getMonth() - birth.getMonth();
    if (m < 0 || (m === 0 && today.getDate() < birth.getDate())) {
        age--;
    }
    return String(age);
}
/**
 * 성별과 나이를 기반으로 시트 통계용 카테고리(m20, f19 등)를 반환합니다.
 */
export function categorizeMember(gender, age) {
    const ageNum = parseInt(age) || 20;
    const isMale = (gender === '남성');
    const prefix = isMale ? 'm' : 'f';

    if (ageNum <= 19) return prefix + '19';
    if (ageNum >= 20 && ageNum <= 29) return prefix + '20';
    if (ageNum >= 30 && ageNum <= 39) return prefix + '30';
    return prefix + '30'; // 40대 이상도 30대 칸에 합산 (시트 규격 준수)
}

/**
 * 토스트 알림 표시 함수
 * @param {string} message - 표시할 메시지
 * @param {string} type - 'success' | 'error'
 */
export function showToast(message, type = 'success') {
    const container = document.getElementById('toastContainer');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = `toast-item ${type}`;
    
    // 아이콘 설정
    const icon = type === 'success' 
        ? '<i class="fa-solid fa-circle-check" style="color: #10b981;"></i>' 
        : '<i class="fa-solid fa-circle-xmark" style="color: #ef4444;"></i>';

    toast.innerHTML = `${icon} <span>${message}</span>`;
    container.appendChild(toast);

    // 알림음 재생 (선택 사항)
    const sound = document.getElementById('notificationSound');
    if (sound) {
        sound.currentTime = 0;
        sound.play().catch(() => {}); // 유저 상호작용 없으면 차단될 수 있음
    }

    // 4초 후 자동 제거
    setTimeout(() => {
        toast.remove();
    }, 4000);
}
