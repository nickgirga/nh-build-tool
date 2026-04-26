import { USERS_TABLE_KEY } from './constants.js';
import { workbookData } from './workbook.js';
import { formatPhoneNumber } from './utils.js';

// DOM references for the Users Table editable cells.
export const usersCells = {
    ticket: document.getElementById('usersTicket'),
    username: document.getElementById('usersUsername'),
    email: document.getElementById('usersEmail'),
    teams: document.getElementById('usersTeams'),
    cmicAccount: document.getElementById('usersCmicAccount'),
    password: document.getElementById('usersPassword')
};

// Makes the Users Table cells editable and attaches save/tab/phone-formatting handlers.
export function initUsersEditableCells() {
    Object.values(usersCells).forEach(cell => {
        cell.contentEditable = 'true';
        cell.addEventListener('input', () => {
            if (cell === usersCells.teams) {
                const digits = cell.textContent.replace(/\D/g, '');
                if (digits.length === 10) {
                    const formatted = formatPhoneNumber(cell.textContent);
                    if (cell.textContent !== formatted) {
                        const selection = window.getSelection();
                        const range = selection.rangeCount > 0 ? selection.getRangeAt(0) : null;
                        const offset = range ? range.startOffset : cell.textContent.length;
                        cell.textContent = formatted;
                        if (range) {
                            const newOffset = Math.min(offset, formatted.length);
                            range.setStart(cell.firstChild || cell, newOffset);
                            range.collapse(true);
                            selection.removeAllRanges();
                            selection.addRange(range);
                        }
                    }
                }
            }
            saveUsersTable();
            updateCopyEmailCmdBtnState();
        });
        cell.addEventListener('keydown', (e) => {
            if (e.key === 'Tab') {
                e.preventDefault();
                const cells = Object.values(usersCells);
                const currentIndex = cells.indexOf(cell);
                const nextIndex = e.shiftKey ? currentIndex - 1 : currentIndex + 1;
                if (nextIndex >= 0 && nextIndex < cells.length) {
                    cells[nextIndex].focus();
                } else if (nextIndex >= cells.length) {
                    document.getElementById('copyUsersBtn').focus();
                }
            }
        });
    });
}

// Makes the Users Table label cells clickable to copy their adjacent value.
export function initUsersLabelCells() {
    const labels = document.querySelectorAll('.users-table .label-cell');
    labels.forEach(label => {
        label.style.cursor = 'pointer';
        label.addEventListener('click', async () => {
            const valueCell = label.nextElementSibling;
            if (valueCell && valueCell.textContent) {
                await navigator.clipboard.writeText(valueCell.textContent);
                const toast = document.getElementById('toast');
                toast.classList.add('show');
                setTimeout(() => toast.classList.remove('show'), 2000);
            }
        });
    });
}

// Enables/disables the Copy Email Command button based on whether username and domain are filled.
export function updateCopyEmailCmdBtnState() {
    const btn = document.getElementById('copyEmailCmdBtn');
    const hasUsername = usersCells.username.textContent.trim().length > 0;
    const emailDomainSelect = document.getElementById('emailDomainSelect');
    let hasDomain = false;
    if (emailDomainSelect.value === '__other__') {
        hasDomain = document.getElementById('emailDomainOther').value.trim().length > 0;
    } else {
        hasDomain = emailDomainSelect.value.length > 0;
    }

    if (hasUsername && hasDomain) {
        btn.disabled = false;
        btn.removeAttribute('title');
    } else if (!hasUsername && !hasDomain) {
        btn.disabled = true;
        btn.title = 'Fill the Username Field and Choose an Email Domain';
    } else if (!hasDomain) {
        btn.disabled = true;
        btn.title = 'Choose an Email Domain';
    } else {
        btn.disabled = true;
        btn.title = 'Fill the Username Field';
    }
}

// Shows the custom email domain input row.
export function showEmailDomainOther() {
    document.getElementById('emailDomainOtherRow').classList.add('show');
}

// Hides the custom email domain input row.
export function hideEmailDomainOther() {
    document.getElementById('emailDomainOtherRow').classList.remove('show');
}

// Persists the Users Table content to localStorage.
export function saveUsersTable() {
    const data = {
        ticket: usersCells.ticket.textContent,
        username: usersCells.username.textContent,
        email: usersCells.email.textContent,
        teams: usersCells.teams.textContent,
        cmicAccount: usersCells.cmicAccount.textContent,
        password: usersCells.password.textContent
    };
    localStorage.setItem(USERS_TABLE_KEY, JSON.stringify(data));
}

// Restores the Users Table content from localStorage.
export function loadUsersTable() {
    const stored = localStorage.getItem(USERS_TABLE_KEY);
    if (stored) {
        try {
            const data = JSON.parse(stored);
            usersCells.ticket.textContent = data.ticket || '';
            usersCells.username.textContent = data.username || '';
            usersCells.email.textContent = data.email || '';
            usersCells.teams.textContent = data.teams || '';
            usersCells.cmicAccount.textContent = data.cmicAccount || '';
            usersCells.password.textContent = data.password || '';
        } catch (e) {
            console.error('Failed to load users table:', e);
        }
    }
    updateCopyEmailCmdBtnState();
}

// Clears all Users Table cells and removes the saved entry from localStorage.
export function clearUsersTable() {
    Object.values(usersCells).forEach(cell => {
        cell.textContent = '';
    });
    localStorage.removeItem(USERS_TABLE_KEY);
    updateCopyEmailCmdBtnState();
}

// If Mirror is unchecked, copies the current CMiC settings into the Users Table CMiC field.
export function updateUsersTableCmic() {
    if (document.getElementById('mirrorCheckbox').checked) return;

    const cmicSettings = document.getElementById('cmicUserSettings').textContent;
    const cmicAccess = document.getElementById('cmicUserAccess').textContent;
    if (cmicSettings && cmicAccess && !cmicSettings.includes('Select') && !cmicAccess.includes('Select')) {
        usersCells.cmicAccount.textContent = `${cmicSettings} - ${cmicAccess}`;
    }
}

// Returns the Users Table as an HTML string (full version with all fields).
export function getUsersTableHtml() {
    const rows = [
        ['Ticket:', usersCells.ticket.textContent],
        ['Username:', usersCells.username.textContent],
        ['Email:', usersCells.email.textContent],
        ['Teams:', usersCells.teams.textContent],
        ['CMiC Account:', usersCells.cmicAccount.textContent],
        ['Password:', usersCells.password.textContent]
    ];
    
    let html = '<table style="border-collapse: collapse;">';
    rows.forEach(([label, value]) => {
        html += `<tr><td style="border: 1px solid #000; padding: 4px 8px; width: 120px; white-space: nowrap;">${label}</td><td style="border: 1px solid #000; padding: 4px 8px;">${value}</td></tr>`;
    });
    html += '</table>';
    return html;
}

// Returns the Users Table as an HTML string (response version, limited fields).
export function getUsersTableHtmlResponse() {
    const rows = [
        ['Username:', usersCells.username.textContent],
        ['Email:', usersCells.email.textContent],
        ['Teams:', usersCells.teams.textContent]
    ];
    
    let html = '<table style="border-collapse: collapse;">';
    rows.forEach(([label, value]) => {
        html += `<tr><td style="border: 1px solid #000; padding: 4px 8px; width: 120px; white-space: nowrap;">${label}</td><td style="border: 1px solid #000; padding: 4px 8px;">${value}</td></tr>`;
    });
    html += '</table>';
    return html;
}

// Returns the Users Table as a plain text string (full version).
export function getUsersTablePlainText() {
    const rows = [
        ['Ticket:', usersCells.ticket.textContent],
        ['Username:', usersCells.username.textContent],
        ['Email:', usersCells.email.textContent],
        ['Teams:', usersCells.teams.textContent],
        ['CMiC Account:', usersCells.cmicAccount.textContent],
        ['Password:', usersCells.password.textContent]
    ];
    
    return rows.map(([label, value]) => `${label}\t${value}`).join('\n');
}

// Returns the Users Table as a plain text string (response version).
export function getUsersTablePlainTextResponse() {
    const rows = [
        ['Username:', usersCells.username.textContent],
        ['Email:', usersCells.email.textContent],
        ['Teams:', usersCells.teams.textContent]
    ];
    
    return rows.map(([label, value]) => `${label}\t${value}`).join('\n');
}

// Writes both HTML and plain text representations to the clipboard for rich pasting.
export async function copyUsersTableAsHtml(html, plainText) {
    if (ClipboardItem && navigator.clipboard) {
        try {
            const blobHtml = new Blob([html], { type: 'text/html' });
            const blobText = new Blob([plainText], { type: 'text/plain' });
            
            const item = new ClipboardItem({
                'text/html': blobHtml,
                'text/plain': blobText
            });
            
            await navigator.clipboard.write([item]);
        } catch (e) {
            console.warn('HTML clipboard not supported, falling back to plain text');
            await navigator.clipboard.writeText(plainText);
        }
    } else {
        await navigator.clipboard.writeText(plainText);
    }
}

// Copy handler for the full Users Table.
document.getElementById('copyUsersBtn').addEventListener('click', async () => {
    await copyUsersTableAsHtml(getUsersTableHtml(), getUsersTablePlainText());
    const copyUsersBtn = document.getElementById('copyUsersBtn');
    copyUsersBtn.classList.add('copied');
    setTimeout(() => copyUsersBtn.classList.remove('copied'), 1500);
});

// Copy handler for the response-version Users Table.
document.getElementById('copyUsersResponseBtn').addEventListener('click', async () => {
    await copyUsersTableAsHtml(getUsersTableHtmlResponse(), getUsersTablePlainTextResponse());
    const copyUsersResponseBtn = document.getElementById('copyUsersResponseBtn');
    copyUsersResponseBtn.classList.add('copied');
    setTimeout(() => copyUsersResponseBtn.classList.remove('copied'), 1500);
});

// Clear handler for the Users Table.
document.getElementById('clearUsersBtn').addEventListener('click', () => {
    clearUsersTable();
    const clearUsersBtn = document.getElementById('clearUsersBtn');
    clearUsersBtn.classList.add('copied');
    setTimeout(() => clearUsersBtn.classList.remove('copied'), 1500);
});

// Generates and copies the email command for the current username + domain.
document.getElementById('copyEmailCmdBtn').addEventListener('click', () => {
    const username = usersCells.username.textContent.trim();
    if (!username) return;

    const emailDomainSelect = document.getElementById('emailDomainSelect');
    let emailDomain = '';
    if (emailDomainSelect.value === '__other__') {
        emailDomain = document.getElementById('emailDomainOther').value.trim();
    } else {
        emailDomain = emailDomainSelect.value;
    }

    const emailAddress = username + '@' + emailDomain;
    usersCells.email.textContent = emailAddress;
    saveUsersTable();

    const command = workbookData.emailTemplate1 + username + '@' + emailDomain + workbookData.emailTemplate2 + username + workbookData.emailTemplate3;
    navigator.clipboard.writeText(command);
    
    const btn = document.getElementById('copyEmailCmdBtn');
    btn.classList.add('copied');
    setTimeout(() => btn.classList.remove('copied'), 1500);
});
