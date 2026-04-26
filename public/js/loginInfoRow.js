import { showToast } from './utils.js';

// DOM references for the Login Info Table Row cells.
export const loginInfoCells = {
    employeeId: document.getElementById('loginEmployeeId'),
    name: document.getElementById('loginName'),
    division: document.getElementById('loginDivision'),
    startDate: document.getElementById('loginStartDate'),
    workLocation: document.getElementById('loginWorkLocation'),
    ticket: document.getElementById('loginTicket'),
    dateReceived: document.getElementById('loginDateReceived'),
    userName: document.getElementById('loginUserName'),
    email: document.getElementById('loginEmail'),
    teamsNumber: document.getElementById('loginTeamsNumber'),
    password: document.getElementById('loginPassword'),
    solutionTeam: document.getElementById('loginSolutionTeam'),
    agentName: document.getElementById('loginAgentName')
};

const copyLoginInfoBtn = document.getElementById('copyLoginInfoBtn');

// Populates the Login Info Table Row with parsed ticket data and attaches click-to-copy handlers.
export function updateLoginInfoCells(data, agentName) {
    loginInfoCells.employeeId.textContent = data.employeeId || '';
    loginInfoCells.name.textContent = data.name || '';
    loginInfoCells.division.textContent = data.division || '';
    loginInfoCells.startDate.textContent = data.startDate || '';
    loginInfoCells.workLocation.textContent = '';
    loginInfoCells.ticket.textContent = data.ticket || '';
    loginInfoCells.dateReceived.textContent = data.dateReceived || '';
    loginInfoCells.userName.textContent = '';
    loginInfoCells.email.textContent = '';
    loginInfoCells.teamsNumber.textContent = '';
    loginInfoCells.password.textContent = '';
    loginInfoCells.solutionTeam.textContent = data.solutionTeam || '';
    loginInfoCells.agentName.textContent = agentName || '';

    Object.values(loginInfoCells).forEach(cell => {
        cell.onclick = async () => {
            if (cell.textContent) {
                await navigator.clipboard.writeText(cell.textContent);
                showToast();
            }
        };
    });
}

// Returns the Login Info row values as a tab-delimited string for Excel pasting.
export function getLoginInfoRowString() {
    return [
        loginInfoCells.employeeId.textContent,
        loginInfoCells.name.textContent,
        loginInfoCells.division.textContent,
        loginInfoCells.startDate.textContent,
        loginInfoCells.workLocation.textContent,
        loginInfoCells.ticket.textContent,
        loginInfoCells.dateReceived.textContent,
        loginInfoCells.userName.textContent,
        loginInfoCells.email.textContent,
        loginInfoCells.teamsNumber.textContent,
        loginInfoCells.password.textContent,
        loginInfoCells.solutionTeam.textContent,
        loginInfoCells.agentName.textContent
    ].join('\t');
}

// Copy button handler for the Login Info row.
copyLoginInfoBtn.addEventListener('click', async () => {
    const rowString = getLoginInfoRowString();
    await navigator.clipboard.writeText(rowString);
    copyLoginInfoBtn.classList.add('copied');
    setTimeout(() => copyLoginInfoBtn.classList.remove('copied'), 1500);
});
