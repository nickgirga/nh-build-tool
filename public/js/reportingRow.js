import { showToast } from './utils.js';

// DOM references for the Reporting Table Row cells.
export const reportingCells = {
    ticket: document.getElementById('reportTicket'),
    name: document.getElementById('reportName'),
    division: document.getElementById('reportDivision'),
    dateReceived: document.getElementById('reportDateReceived'),
    startDate: document.getElementById('reportStartDate'),
    busDays: document.getElementById('reportBusDays'),
    solutionTeam: document.getElementById('reportSolutionTeam'),
    agentName: document.getElementById('reportAgentName')
};

const copyReportingBtn = document.getElementById('copyReportingBtn');

// Populates the Reporting Table Row with parsed ticket data and attaches click-to-copy handlers.
export function updateReportingCells(data, agentName) {
    reportingCells.ticket.textContent = data.ticket || '';
    reportingCells.name.textContent = data.name || '';
    reportingCells.division.textContent = data.division || '';
    reportingCells.dateReceived.textContent = data.dateReceived || '';
    reportingCells.startDate.textContent = data.startDate || '';
    reportingCells.busDays.textContent = '';
    reportingCells.busDays.setAttribute('data-formula', '=NETWORKDAYS.INTL(INDEX(E:E,ROW()),INDEX(F:F,ROW()))');
    reportingCells.solutionTeam.textContent = data.solutionTeam || '';
    reportingCells.agentName.textContent = agentName || '';

    Object.values(reportingCells).forEach(cell => {
        cell.onclick = async () => {
            if (cell.textContent) {
                await navigator.clipboard.writeText(cell.textContent);
                showToast();
            }
        };
    });
}

// Returns the Reporting row values as a tab-delimited string (includes the bus-days Excel formula).
export function getReportingRowString() {
    const busDaysValue = reportingCells.busDays.getAttribute('data-formula') || reportingCells.busDays.textContent;
    return [
        reportingCells.ticket.textContent,
        reportingCells.name.textContent,
        reportingCells.division.textContent,
        reportingCells.dateReceived.textContent,
        reportingCells.startDate.textContent,
        busDaysValue,
        reportingCells.solutionTeam.textContent,
        reportingCells.agentName.textContent
    ].join('\t');
}

// Copy button handler for the Reporting row.
copyReportingBtn.addEventListener('click', async () => {
    const rowString = getReportingRowString();
    await navigator.clipboard.writeText(rowString);
    copyReportingBtn.classList.add('copied');
    setTimeout(() => copyReportingBtn.classList.remove('copied'), 1500);
});
