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

/**
 * 토스트 알림 표시 함수
 * @param {string} message - 표시할 메시지
 * @param {string} type - 'success' | 'error' | 'info'
 */
export function showToast(message, type = 'success') {
    const container = document.getElementById('toastContainer');
    if (!container) return;

    const normalizedType = ['success', 'error', 'info'].includes(type) ? type : 'info';
    const toast = document.createElement('div');
    toast.className = `toast-item ${normalizedType}`;
    toast.setAttribute('role', normalizedType === 'error' ? 'alert' : 'status');

    const icon = document.createElement('span');
    icon.setAttribute('aria-hidden', 'true');
    icon.className = normalizedType === 'success'
        ? 'icon icon-circle-check'
        : normalizedType === 'error'
            ? 'icon icon-circle-xmark'
            : 'icon icon-circle-info';

    const text = document.createElement('span');
    text.textContent = message;
    toast.append(icon, text);
    container.appendChild(toast);

    setTimeout(() => {
        toast.remove();
    }, 4000);
}
