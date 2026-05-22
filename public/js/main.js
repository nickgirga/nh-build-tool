import {
    STORAGE_KEY, AGENT_NAME_KEY, MIRROR_PATTERN
} from './constants.js';
import {
    workbookData, parseWorkbook, loadFromStorage, populateClipboardUtilities,
    populateDropdowns, updateFileStatus, resetWorkbookData
} from './workbook.js';
import {
    updateLookups, findBestJobTitleMatch, findBestLocationMatch,
    selectDropdownValue
} from './lookups.js';
import { generateUsersPassword } from './password.js';
import { parseTicketData } from './ticketParser.js';
import { updateTicketLogCells } from './ticketLogRow.js';
import { updateLoginInfoCells } from './loginInfoRow.js';
import { updateReportingCells } from './reportingRow.js';
import {
    adSection, adDivisionEl, cmicSection, cmicCodeEl, supervisorSection,
    supervisorDisplayEl, setCurrentSupervisor, updateAdSections
} from './adSections.js';
import {
    usersCells, initUsersEditableCells, initUsersLabelCells,
    loadUsersTable, clearUsersTable, updateUsersTableCmic,
    updateCopyEmailCmdBtnState, showEmailDomainOther, hideEmailDomainOther
} from './usersTable.js';
import { copyToClipboard } from './utils.js';
import {
    showInfoBanner, clearBanners, buildPreParseNotes, buildPostParseNotes,
    buildErrors, buildWarning, buildSlaWarning,
    renderNotes, renderErrors, renderWarning, renderSlaWarning
} from './banners.js';

// DOM references for main controls
const excelFileInput = document.getElementById('excelFile');
const clearAllStorageBtn = document.getElementById('clearAllStorageBtn');
const agentNameInput = document.getElementById('agentName');
const ticketDataInput = document.getElementById('ticketData');
const generateBtn = document.getElementById('generateBtn');
const clearBtn = document.getElementById('clearBtn');
const outputSection = document.getElementById('outputSection');

// --- Excel File Handling ---

excelFileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = (event) => {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = window.XLSX.read(data, { type: 'array' });
            parseWorkbook(workbook);
            populateDropdowns();
            populateClipboardUtilities();
            updateFileStatus(true);
            updateLookups();
            updateChecklistFabHref();
        } catch (error) {
            console.error('Error parsing Excel file:', error);
            const fileStatus = document.getElementById('fileStatus');
            fileStatus.textContent = 'Error parsing file. Please ensure it is a valid .xlsx file.';
            fileStatus.classList.remove('loaded');
            fileStatus.classList.add('show');
        }
    };
    reader.readAsArrayBuffer(file);
});

// --- Lookup Dropdown Event Listeners ---

document.getElementById('jobTypeSelect').addEventListener('change', () => {
    document.getElementById('lookupWarning').classList.remove('show');
    updateLookups();
});

document.getElementById('jobTitleSelect').addEventListener('change', () => {
    document.getElementById('lookupWarning').classList.remove('show');
    updateLookups();
    updateUsersTableCmic();
});

document.getElementById('locationSelect').addEventListener('change', () => {
    document.getElementById('lookupWarning').classList.remove('show');
    updateLookups();
    updateUsersTableCmic();
});

// --- Lookup Result Copy Buttons ---

document.getElementById('copyAdRolesPerTypeBtn').addEventListener('click', () => {
    copyToClipboard('copyAdRolesPerTypeBtn', document.getElementById('adRolesPerType').textContent);
});

document.getElementById('copyAdRolesPerJobTitleBtn').addEventListener('click', () => {
    copyToClipboard('copyAdRolesPerJobTitleBtn', document.getElementById('adRolesPerJobTitle').textContent);
});

document.getElementById('copyAdRolesPerLocationBtn').addEventListener('click', () => {
    copyToClipboard('copyAdRolesPerLocationBtn', document.getElementById('adRolesPerLocation').textContent);
});

document.getElementById('copyCmicUserSettingsBtn').addEventListener('click', () => {
    copyToClipboard('copyCmicUserSettingsBtn', document.getElementById('cmicUserSettings').textContent);
});

document.getElementById('copyCmicUserAccessBtn').addEventListener('click', () => {
    copyToClipboard('copyCmicUserAccessBtn', document.getElementById('cmicUserAccess').textContent);
});

document.getElementById('copyO365GroupsBtn').addEventListener('click', () => {
    copyToClipboard('copyO365GroupsBtn', document.getElementById('o365Groups').textContent);
});

document.getElementById('copyAreaCodesBtn').addEventListener('click', () => {
    copyToClipboard('copyAreaCodesBtn', document.getElementById('areaCodes').textContent);
});

document.getElementById('copyEmailCommandSetupScriptBtn').addEventListener('click', () => {
    copyToClipboard('copyEmailCommandSetupScriptBtn', document.getElementById('emailCommandSetupScript').textContent);
});

document.getElementById('copyLogonScriptBtn').addEventListener('click', () => {
    copyToClipboard('copyLogonScriptBtn', document.getElementById('logonScript').textContent);
});

document.getElementById('copyCmicPasswordBtn').addEventListener('click', () => {
    copyToClipboard('copyCmicPasswordBtn', document.getElementById('cmicPassword').textContent);
});

document.getElementById('copyAllAdRolesBtn').addEventListener('click', () => {
    const adRolesPerType = document.getElementById('adRolesPerType').textContent;
    const adRolesPerJobTitle = document.getElementById('adRolesPerJobTitle').textContent;
    const adRolesPerLocation = document.getElementById('adRolesPerLocation').textContent;

    const allRoles = [adRolesPerType, adRolesPerJobTitle, adRolesPerLocation]
        .filter(r => r && !r.includes('Select') && r.includes(';'))
        .join(' ');

    copyToClipboard('copyAllAdRolesBtn', allRoles);
});

// --- Main Generate Button ---

generateBtn.addEventListener('click', async () => {
    if (!ticketDataInput.value.includes('HR: New Hire')) {
        alert("The Pasted Ticket Data doesn't appear to be valid. Ensure you have selected all text on the ticket before copying and pasting into the text field.");
        return;
    }

    showInfoBanner();

    const agentNameBanner = document.getElementById('agentNameBanner');
    if (!agentNameInput.value.trim()) {
        agentNameBanner.classList.add('show');
        agentNameInput.classList.add('agent-name-error');
    } else {
        agentNameBanner.classList.remove('show');
        agentNameInput.classList.remove('agent-name-error');
    }

    let notes = buildPreParseNotes(ticketDataInput.value);
    renderNotes(notes);

    const data = parseTicketData(ticketDataInput.value);

    if (!data.dateReceived) {
        alert("The 'Date Received' could not be found. Ensure you are at the top of the ticket when copying, as FreshService will hide this info as you scroll down.");
    }

    renderErrors(buildErrors(data, ticketDataInput.value));
    renderWarning(buildWarning(data));
    renderSlaWarning(buildSlaWarning(data));

    notes = notes.concat(buildPostParseNotes(data, ticketDataInput.value));
    renderNotes(notes);

    updateTicketLogCells(data);
    updateLoginInfoCells(data, agentNameInput.value);
    updateReportingCells(data, agentNameInput.value);

    if (workbookData.adPerType.length > 0) {
        const lookupSection = document.getElementById('lookupSection');
        const resultsSection = document.getElementById('resultsSection');

        document.getElementById('jobTypeSelect').selectedIndex = 0;
        document.getElementById('jobTitleSelect').selectedIndex = 0;
        document.getElementById('locationSelect').selectedIndex = 0;
        document.getElementById('emailDomainSelect').selectedIndex = 0;
        hideEmailDomainOther();
        document.getElementById('emailDomainOther').value = '';

        if (data.empType) {
            selectDropdownValue('jobTypeSelect', data.empTypeForLookup);
        }
        if (data.jobTitle) {
            selectDropdownValue('jobTitleSelect', findBestJobTitleMatch(data.jobTitle, data.cmicCode));
        }
        if (data.location) {
            selectDropdownValue('locationSelect', findBestLocationMatch(data.location, data.empTypeForLookup, ticketDataInput.value.toLowerCase(), data.cmicCode));
        }
        if (data.emailDomain) {
            const emailDomainClean = data.emailDomain.toLowerCase();
            const matchedDomain = workbookData.emailDomains.find(d => d.toLowerCase() === emailDomainClean);
            if (matchedDomain) {
                document.getElementById('emailDomainSelect').value = matchedDomain;
            } else {
                document.getElementById('emailDomainSelect').value = '__other__';
                showEmailDomainOther();
                document.getElementById('emailDomainOther').value = data.emailDomain;
            }
        }
        updateCopyEmailCmdBtnState();

        updateLookups();
        lookupSection.style.display = 'flex';
        resultsSection.style.display = 'flex';
        document.getElementById('lookupWarning').classList.add('show');

        const isMirror = MIRROR_PATTERN.test(ticketDataInput.value);
        document.getElementById('mirrorCheckbox').checked = isMirror;
        if (!isMirror && !usersCells.cmicAccount.textContent.trim()) {
            updateUsersTableCmic();
        }
    }

    clearBtn.classList.add('show');
    resetAnimations();
    outputSection.style.display = 'flex';
    document.getElementById('loginInfoSection').style.display = 'flex';
    document.getElementById('reportingSection').style.display = 'flex';

    if (!usersCells.ticket.textContent.trim()) {
        if (data.ticket && data.solutionTeam) {
            const lastName = data.solutionTeam.split(' ').pop();
            usersCells.ticket.textContent = `${data.ticket} - ${lastName}`;
        } else if (data.ticket) {
            usersCells.ticket.textContent = data.ticket;
        }
    }

    updateAdSections(data);

    await generateUsersPassword();
});

// --- Clear Button ---

clearBtn.addEventListener('click', () => {
    ticketDataInput.value = '';
    updateTicketLogCells({});
    updateLoginInfoCells({}, '');
    updateReportingCells({}, '');
    resetAnimations();
    supervisorDisplayEl.textContent = '';
    adDivisionEl.textContent = '';
    cmicCodeEl.textContent = '';
    setCurrentSupervisor('');
    clearBanners();
    clearBtn.classList.remove('show');
    document.getElementById('mirrorCheckbox').checked = false;
    clearUsersTable();
});

// --- Animation Reset ---

function resetAnimations() {
    const sections = [outputSection, adSection, cmicSection, supervisorSection];
    const loginInfoSection = document.getElementById('loginInfoSection');
    const reportingSection = document.getElementById('reportingSection');
    
    sections.forEach(section => {
        section.style.display = 'none';
        section.style.animation = 'none';
        section.style.animationDelay = '';
    });
    
    if (loginInfoSection) {
        loginInfoSection.style.display = 'none';
    }
    
    if (reportingSection) {
        reportingSection.style.display = 'none';
    }
    
    setTimeout(() => {
        sections.forEach(section => {
            section.style.animation = '';
        });
    }, 10);
}

// --- Email Domain & Mirror Checkboxes ---

document.getElementById('emailDomainSelect').addEventListener('change', () => {
    const emailDomainSelect = document.getElementById('emailDomainSelect');
    const emailDomainOther = document.getElementById('emailDomainOther');
    if (emailDomainSelect.value === '__other__') {
        showEmailDomainOther();
    } else {
        hideEmailDomainOther();
        emailDomainOther.value = '';
    }
    updateCopyEmailCmdBtnState();
});

document.getElementById('emailDomainOther').addEventListener('input', () => {
    updateCopyEmailCmdBtnState();
});

document.getElementById('mirrorCheckbox').addEventListener('change', () => {
    if (!document.getElementById('mirrorCheckbox').checked) {
        updateUsersTableCmic();
    }
});

// --- Checklist FAB ---

const checklistFab = document.getElementById('checklistFab');

function updateChecklistFabHref() {
    const hash = workbookData.checklistHash;
    if (hash) {
        checklistFab.href = `https://markdown-online.gitlab.io/?preview=1#fileContents=${hash}`;
    } else {
        checklistFab.href = '#';
    }
}

checklistFab.addEventListener('click', (e) => {
    if (!workbookData.checklistHash) {
        e.preventDefault();
        alert('No checklist hash found. Ensure the workbook has a value in Metadata cell B3.');
    }
});

// --- Agent Name ---

const savedAgentName = localStorage.getItem(AGENT_NAME_KEY);
if (savedAgentName) {
    agentNameInput.value = savedAgentName;
}

agentNameInput.addEventListener('input', () => {
    localStorage.setItem(AGENT_NAME_KEY, agentNameInput.value);
    if (agentNameInput.value.trim()) {
        agentNameInput.classList.remove('agent-name-error');
    }
});

// --- Clear All Storage ---

clearAllStorageBtn.addEventListener('click', () => {
    console.log('Clear button clicked');
    if (confirm('This will clear all stored Excel data, agent name, and the users table. Are you sure?')) {
        console.log('Confirmed, clearing...');
        localStorage.clear();
        resetWorkbookData();
        document.getElementById('jobTypeSelect').innerHTML = '';
        document.getElementById('jobTitleSelect').innerHTML = '';
        document.getElementById('locationSelect').innerHTML = '';
        document.getElementById('emailDomainSelect').innerHTML = '';
        hideEmailDomainOther();
        document.getElementById('emailDomainOther').value = '';
        document.getElementById('adRolesPerType').textContent = '';
        document.getElementById('adRolesPerJobTitle').textContent = '';
        document.getElementById('adRolesPerLocation').textContent = '';
        document.getElementById('cmicUserSettings').textContent = '';
        document.getElementById('cmicUserAccess').textContent = '';
        document.getElementById('o365Groups').textContent = '';
        document.getElementById('areaCodes').textContent = '';
        document.getElementById('emailCommandSetupScript').textContent = '';
        document.getElementById('logonScript').textContent = '';
        document.getElementById('cmicPassword').textContent = '';
        document.getElementById('clipboardSection').style.display = 'none';
        excelFileInput.value = '';
        agentNameInput.value = '';
        clearUsersTable();
        updateFileStatus(false);
        updateChecklistFabHref();
    }
});

// --- Initialization ---

if (loadFromStorage()) {
    populateDropdowns();
    populateClipboardUtilities();
    updateFileStatus(true);
    updateLookups();
    updateChecklistFabHref();
} else {
    updateFileStatus(false);
}

loadUsersTable();
initUsersEditableCells();
initUsersLabelCells();
