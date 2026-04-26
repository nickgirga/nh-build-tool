import { showToast } from './utils.js';

// DOM references for the Ticket Log Table Row cells.
export const cells = {
    ticket: document.getElementById('ticket'),
    dateReceived: document.getElementById('dateReceived'),
    dateWorked: document.getElementById('dateWorked'),
    name: document.getElementById('name'),
    user: document.getElementById('user'),
    email: document.getElementById('email'),
    jobTitle: document.getElementById('jobTitle'),
    empType: document.getElementById('empType'),
    startDate: document.getElementById('startDate'),
    division: document.getElementById('division'),
    location: document.getElementById('location'),
    phone: document.getElementById('phone'),
    employeeId: document.getElementById('employeeId')
};

const copyBtn = document.getElementById('copyBtn');

// Populates the Ticket Log Table Row with parsed ticket data and attaches click-to-copy handlers.
export function updateTicketLogCells(data) {
    cells.ticket.textContent = data.ticket || '';
    cells.dateReceived.textContent = data.dateReceived || '';
    cells.dateWorked.textContent = data.dateWorked || '';
    cells.name.textContent = data.name || '';
    cells.user.textContent = data.user || '';
    cells.email.textContent = data.email || '';
    cells.jobTitle.textContent = data.jobTitle || '';
    const tableRowEmpType = data.empType && /craft/i.test(data.empType) ? 'Craft' : 'Admin';
    cells.empType.textContent = tableRowEmpType;
    cells.startDate.textContent = data.startDate || '';
    cells.division.textContent = data.division || '';
    cells.location.textContent = data.location || '';
    cells.phone.textContent = data.phone || '';
    cells.employeeId.textContent = data.employeeId || '';

    Object.values(cells).forEach(cell => {
        cell.onclick = async () => {
            if (cell.textContent) {
                await navigator.clipboard.writeText(cell.textContent);
                showToast();
            }
        };
    });
}

// Returns the Ticket Log row values as a tab-delimited string for Excel pasting.
export function getRowString() {
    return [
        cells.ticket.textContent,
        cells.dateReceived.textContent,
        cells.dateWorked.textContent,
        cells.name.textContent,
        cells.user.textContent,
        cells.email.textContent,
        cells.jobTitle.textContent,
        cells.empType.textContent,
        cells.startDate.textContent,
        cells.division.textContent,
        cells.location.textContent,
        cells.phone.textContent,
        cells.employeeId.textContent
    ].join('\t');
}

// Copy button handler for the Ticket Log row.
copyBtn.addEventListener('click', async () => {
    const rowString = getRowString();
    await navigator.clipboard.writeText(rowString);
    copyBtn.classList.add('copied');
    setTimeout(() => copyBtn.classList.remove('copied'), 1500);
});
