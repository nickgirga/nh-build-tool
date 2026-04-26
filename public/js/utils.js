// Copies the given text to the clipboard and briefly highlights the button by ID.
export async function copyToClipboard(btnId, text) {
    await navigator.clipboard.writeText(text);
    const btn = document.getElementById(btnId);
    if (btn) {
        btn.classList.add('copied');
        setTimeout(() => btn.classList.remove('copied'), 1500);
    }
}

// Shows the toast notification for 2 seconds.
export function showToast() {
    const toast = document.getElementById('toast');
    toast.classList.add('show');
    setTimeout(() => toast.classList.remove('show'), 2000);
}

// Formats a string of 10 digits into (XXX) XXX-XXXX.
export function formatPhoneNumber(text) {
    const digits = text.replace(/\D/g, '');
    if (digits.length === 10) {
        return `(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`;
    }
    return text;
}
