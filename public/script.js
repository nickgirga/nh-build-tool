const XLS = window.XLSX || window.xlsx;

const STORAGE_KEY = 'nhbuildtool_workbook';
const AGENT_NAME_KEY = 'nhbuildtool_agentName';
const USERS_TABLE_KEY = 'nhbuildtool_usersTable';
const MIRROR_PATTERN = /\bmirror\b/i;

const excelFileInput = document.getElementById('excelFile');
// const clearExcelBtn = document.getElementById('clearExcelBtn'); // removed - use "Clear All Local Storage" instead
const clearAllStorageBtn = document.getElementById('clearAllStorageBtn');
const fileStatus = document.getElementById('fileStatus');
const agentNameInput = document.getElementById('agentName');

let workbookData = {
    adPerType: [],
    jobTitles: [],
    locationMap: [],
    adPerLocation: [],
    jobSoftwares: [],
    cmic: [],
    cmicLocationMap: []
};

function parseSheetToArrays(worksheet, headerRow = 1) {
    const data = [];
    const range = XLS.utils.decode_range(worksheet['!ref']);
    
    for (let row = headerRow; row <= range.e.r; row++) {
        const rowData = [];
        for (let col = range.e.c; col >= 0; col--) {
            const cellRef = XLS.utils.encode_cell({ r: row, c: col });
            const cell = worksheet[cellRef];
            rowData.push(cell ? String(cell.v).trim() : '');
        }
        if (rowData.some(v => v)) {
            data.push(rowData.reverse());
        }
    }
    return data;
}

function parseWorkbook(wb) {
    const sheets = {};
    
    const expectedSheets = [
        'AD Per Type', 'Job Titles', 'Location Map', 'AD Per Location',
        'Job Softwares', 'CMiC', 'CMiC Location Map'
    ];
    
    for (const sheetName of expectedSheets) {
        if (wb.Sheets[sheetName]) {
            sheets[sheetName] = parseSheetToArrays(wb.Sheets[sheetName]);
        }
    }
    
    const getColumnA = (sheetData) => {
        return [...new Set(sheetData.map(row => row[0]).filter(v => v))];
    };
    
    workbookData = {
        adPerType: sheets['AD Per Type'] || [],
        jobTitles: sheets['Job Titles'] || [],
        locationMap: sheets['Location Map'] || [],
        adPerLocation: sheets['AD Per Location'] || [],
        jobSoftwares: sheets['Job Softwares'] || [],
        cmic: sheets['CMiC'] || [],
        cmicLocationMap: sheets['CMiC Location Map'] || []
    };
    
    const dataToStore = {
        timestamp: Date.now(),
        ...workbookData
    };
    
    localStorage.setItem(STORAGE_KEY, JSON.stringify(dataToStore));
    
    return workbookData;
}

function loadFromStorage() {
    const stored = localStorage.getItem(STORAGE_KEY);
    if (stored) {
        try {
            const data = JSON.parse(stored);
            workbookData = {
                adPerType: data.adPerType || [],
                jobTitles: data.jobTitles || [],
                locationMap: data.locationMap || [],
                adPerLocation: data.adPerLocation || [],
                jobSoftwares: data.jobSoftwares || [],
                cmic: data.cmic || [],
                cmicLocationMap: data.cmicLocationMap || []
            };
            return true;
        } catch (e) {
            console.error('Failed to load from storage:', e);
            return false;
        }
    }
    return false;
}

function populateDropdowns() {
    const jobTypeSelect = document.getElementById('jobTypeSelect');
    const jobTitleSelect = document.getElementById('jobTitleSelect');
    const locationSelect = document.getElementById('locationSelect');
    
    jobTypeSelect.innerHTML = '';
    jobTitleSelect.innerHTML = '';
    locationSelect.innerHTML = '';
    
    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = 'Select...';
    defaultOption.selected = true;
    jobTypeSelect.appendChild(defaultOption.cloneNode(true));
    jobTitleSelect.appendChild(defaultOption.cloneNode(true));
    locationSelect.appendChild(defaultOption.cloneNode(true));
    
    workbookData.adPerType.forEach(type => {
        const option = document.createElement('option');
        option.value = type[0];
        option.textContent = type[0];
        jobTypeSelect.appendChild(option);
    });
    
    workbookData.jobTitles.forEach(title => {
        const option = document.createElement('option');
        option.value = title[0];
        option.textContent = title[0];
        jobTitleSelect.appendChild(option);
    });
    
    workbookData.adPerLocation.forEach(loc => {
        const option = document.createElement('option');
        option.value = loc[0];
        option.textContent = loc[0];
        locationSelect.appendChild(option);
    });
}

function vLookup(lookupValue, sheetData, keyColIndex, returnColIndex) {
    for (const row of sheetData) {
        if (row[keyColIndex] && row[keyColIndex].toLowerCase() === lookupValue.toLowerCase()) {
            return row[returnColIndex] || '';
        }
    }
    return '';
}

function xLookup(conditions, sheetData, returnColIndex) {
    for (const row of sheetData) {
        let match = true;
        for (const { colIndex, value } of conditions) {
            if (colIndex >= row.length || row[colIndex].toLowerCase() !== value.toLowerCase()) {
                match = false;
                break;
            }
        }
        if (match && returnColIndex < row.length) {
            return row[returnColIndex] || '';
        }
    }
    return '';
}

function findAreaForLocation(location) {
    return vLookup(location, workbookData.locationMap, 0, 1);
}

function findCmicGroupForLocation(location) {
    return vLookup(location, workbookData.cmicLocationMap, 1, 0);
}

function findBestJobTitleMatch(parsedJobTitle) {
    if (!parsedJobTitle || !workbookData.jobTitles.length) return parsedJobTitle;
    
    const lowerTitle = parsedJobTitle.toLowerCase();
    const parsedWords = lowerTitle.split(/\s+/);
    
    let bestMatch = null;
    let bestMatchWordCount = 0;
    
    for (const row of workbookData.jobTitles) {
        const excelTitle = row[0];
        if (!excelTitle) continue;
        
        const excelWords = excelTitle.toLowerCase().split(/\s+/);
        let matchingWords = 0;
        
        for (const excelWord of excelWords) {
            if (parsedWords.includes(excelWord)) {
                matchingWords++;
            }
        }
        
        if (matchingWords === excelWords.length && excelWords.length > bestMatchWordCount) {
            bestMatch = excelTitle;
            bestMatchWordCount = excelWords.length;
        }
    }
    
    if (bestMatch) return bestMatch;
    
    for (const row of workbookData.jobTitles) {
        const excelTitle = row[0];
        if (excelTitle && lowerTitle.includes(excelTitle.toLowerCase())) {
            return excelTitle;
        }
    }
    
    return parsedJobTitle;
}

function findBestLocationMatch(parsedLocation, empTypeForLookup, textLower) {
    if (empTypeForLookup === 'Craft') {
        return 'Craft';
    }
    
    const uniqueLocation = new RegExp('s.......n e....y');
    const uniqueLocation2 = new RegExp('s.k c..........n');
    const uniqueLocation3 = new RegExp('s.k b......s');
    
    if (textLower && uniqueLocation.test(textLower)) {
        for (const row of workbookData.adPerLocation) {
            const excelLocation = row[0];
            if (excelLocation && uniqueLocation.test(excelLocation.toLowerCase())) {
                return excelLocation;
            }
        }
    }
    
    if (textLower && (uniqueLocation2.test(textLower) || uniqueLocation3.test(textLower))) {
        for (const row of workbookData.adPerLocation) {
            const excelLocation = row[0];
            if (excelLocation && uniqueLocation2.test(excelLocation.toLowerCase())) {
                return excelLocation;
            }
        }
    }
    
    if (!parsedLocation || !workbookData.adPerLocation.length) return parsedLocation;
    
    const lowerLocation = parsedLocation.toLowerCase();
    
    if (lowerLocation.includes('santa ana')) {
        return 'Orange County';
    }
    
    if (lowerLocation.includes('atlanta') || lowerLocation.includes('raleigh')) {
        return 'Cambridge';
    }
    
    if (lowerLocation.includes('honolulu')) {
        return 'Hawaii';
    }
    
    for (const row of workbookData.adPerLocation) {
        const excelLocation = row[0];
        if (excelLocation && lowerLocation.includes(excelLocation.toLowerCase())) {
            return excelLocation;
        }
    }
    
    for (const row of workbookData.adPerLocation) {
        const excelLocation = row[0];
        if (excelLocation && excelLocation.toLowerCase().includes(lowerLocation)) {
            return excelLocation;
        }
    }
    
    return parsedLocation;
}

function updateLookups() {
    const selectedJobType = document.getElementById('jobTypeSelect').value;
    const selectedJobTitle = document.getElementById('jobTitleSelect').value;
    const selectedLocation = document.getElementById('locationSelect').value;
    
    const adRolesPerType = selectedJobType ? vLookup(selectedJobType, workbookData.adPerType, 0, 1) : '';
    document.getElementById('adRolesPerType').textContent = adRolesPerType || (selectedJobType ? 'No matching role/type' : 'Select Job Type above');
    
    const area = selectedLocation ? findAreaForLocation(selectedLocation) : '';
    const adRolesPerJobTitle = (area && selectedJobTitle) ? xLookup(
        [
            { colIndex: 0, value: area },
            { colIndex: 1, value: selectedJobTitle }
        ],
        workbookData.jobSoftwares,
        2
    ) : '';
    document.getElementById('adRolesPerJobTitle').textContent = adRolesPerJobTitle || (area && selectedJobTitle ? 'No matching role/location' : 'Select Job Title & Location above');
    
    const adRolesPerLocation = selectedLocation ? vLookup(selectedLocation, workbookData.adPerLocation, 0, 1) : '';
    document.getElementById('adRolesPerLocation').textContent = adRolesPerLocation || (selectedLocation ? 'No matching role/location' : 'Select Location above');
    
    const cmicGroup = selectedLocation ? findCmicGroupForLocation(selectedLocation) : '';
    const cmicUserSettings = (cmicGroup && selectedJobTitle) ? xLookup(
        [
            { colIndex: 0, value: cmicGroup },
            { colIndex: 1, value: selectedJobTitle }
        ],
        workbookData.cmic,
        2
    ) : '';
    document.getElementById('cmicUserSettings').textContent = cmicUserSettings || (cmicGroup && selectedJobTitle ? 'No matching role/location' : 'Select Job Title & Location above');
    
    const cmicUserAccess = (cmicGroup && selectedJobTitle) ? xLookup(
        [
            { colIndex: 0, value: cmicGroup },
            { colIndex: 1, value: selectedJobTitle }
        ],
        workbookData.cmic,
        3
    ) : '';
    document.getElementById('cmicUserAccess').textContent = cmicUserAccess || (cmicGroup && selectedJobTitle ? 'No matching role/location' : 'Select Job Title & Location above');
    
    const o365Groups = selectedLocation ? vLookup(selectedLocation, workbookData.adPerLocation, 0, 2) : '';
    document.getElementById('o365Groups').textContent = o365Groups || (selectedLocation ? 'No matching role/location' : 'Select Location above');
    
    const areaCodes = selectedLocation ? vLookup(selectedLocation, workbookData.adPerLocation, 0, 3) : '';
    document.getElementById('areaCodes').textContent = areaCodes || (selectedLocation ? 'No matching role/location' : 'Select Location above');
}

function findValueInSheet(sheetData, searchValue) {
    const lowerSearch = searchValue.toLowerCase();
    for (const row of sheetData) {
        if (row[0] && row[0].toLowerCase() === lowerSearch) {
            return row[0];
        }
    }
    return searchValue;
}

function selectDropdownValue(dropdownId, value) {
    const select = document.getElementById(dropdownId);
    const options = select.options;
    for (let i = 0; i < options.length; i++) {
        if (options[i].value.toLowerCase() === value.toLowerCase()) {
            select.selectedIndex = i;
            return true;
        }
    }
    return false;
}

excelFileInput.addEventListener('change', async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = (event) => {
        try {
            const data = new Uint8Array(event.target.result);
            const workbook = XLS.read(data, { type: 'array' });
            parseWorkbook(workbook);
            populateDropdowns();
            updateFileStatus(true);
        } catch (error) {
            console.error('Error parsing Excel file:', error);
            fileStatus.textContent = 'Error parsing file. Please ensure it is a valid .xlsx file.';
            fileStatus.classList.remove('loaded');
        }
    };
    reader.readAsArrayBuffer(file);
});

// clearExcelBtn removed - use "Clear All Local Storage" button instead

function updateFileStatus(loaded) {
    const lookupSection = document.getElementById('lookupSection');
    const resultsSection = document.getElementById('resultsSection');
    
    if (loaded && workbookData.adPerType.length > 0) {
        const stored = localStorage.getItem(STORAGE_KEY);
        let dateStr = '';
        if (stored) {
            try {
                const data = JSON.parse(stored);
                if (data.timestamp) {
                    const d = new Date(data.timestamp);
                    const m = d.getMonth() + 1;
                    const day = d.getDate();
                    const y = d.getFullYear().toString().slice(-2);
                    dateStr = ` [Last saved ${m}/${day}/${y}]`;
                }
            } catch (e) {}
        }
        fileStatus.textContent = 'Excel data loaded from local storage' + dateStr;
        fileStatus.classList.add('loaded');
        lookupSection.style.display = 'flex';
        resultsSection.style.display = 'flex';
        updateLookups();
    } else {
        fileStatus.textContent = '';
        fileStatus.classList.remove('loaded');
        lookupSection.style.display = 'none';
        resultsSection.style.display = 'none';
    }
}

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

function updateUsersTableCmic() {
    const isMirror = MIRROR_PATTERN.test(ticketDataInput.value);
    if (isMirror) return;
    
    const cmicSettings = document.getElementById('cmicUserSettings').textContent;
    const cmicAccess = document.getElementById('cmicUserAccess').textContent;
    if (cmicSettings && cmicAccess && !cmicSettings.includes('Select') && !cmicAccess.includes('Select')) {
        usersCells.cmicAccount.textContent = `${cmicSettings} - ${cmicAccess}`;
    }
}

document.getElementById('copyAdRolesPerTypeBtn').addEventListener('click', async () => {
    await navigator.clipboard.writeText(document.getElementById('adRolesPerType').textContent);
    showToast();
});

document.getElementById('copyAdRolesPerJobTitleBtn').addEventListener('click', async () => {
    await navigator.clipboard.writeText(document.getElementById('adRolesPerJobTitle').textContent);
    showToast();
});

document.getElementById('copyAdRolesPerLocationBtn').addEventListener('click', async () => {
    await navigator.clipboard.writeText(document.getElementById('adRolesPerLocation').textContent);
    showToast();
});

document.getElementById('copyCmicUserSettingsBtn').addEventListener('click', async () => {
    await navigator.clipboard.writeText(document.getElementById('cmicUserSettings').textContent);
    showToast();
});

document.getElementById('copyCmicUserAccessBtn').addEventListener('click', async () => {
    await navigator.clipboard.writeText(document.getElementById('cmicUserAccess').textContent);
    showToast();
});

document.getElementById('copyO365GroupsBtn').addEventListener('click', async () => {
    await navigator.clipboard.writeText(document.getElementById('o365Groups').textContent);
    showToast();
});

document.getElementById('copyAreaCodesBtn').addEventListener('click', async () => {
    await navigator.clipboard.writeText(document.getElementById('areaCodes').textContent);
    showToast();
});

document.getElementById('copyAllAdRolesBtn').addEventListener('click', async () => {
    const adRolesPerType = document.getElementById('adRolesPerType').textContent;
    const adRolesPerJobTitle = document.getElementById('adRolesPerJobTitle').textContent;
    const adRolesPerLocation = document.getElementById('adRolesPerLocation').textContent;
    
    const allRoles = [adRolesPerType, adRolesPerJobTitle, adRolesPerLocation]
        .filter(r => r && !r.includes('Select') && r.includes(';'))
        .join(' ');
    
    await navigator.clipboard.writeText(allRoles);
    showToast();
});

const ticketDataInput = document.getElementById('ticketData');
const generateBtn = document.getElementById('generateBtn');
const clearBtn = document.getElementById('clearBtn');
const outputSection = document.getElementById('outputSection');
const copyBtn = document.getElementById('copyBtn');
const errorLog = document.getElementById('errorLog');
const warningLog = document.getElementById('warningLog');
const slaWarningLog = document.getElementById('slaWarningLog');
const infoLog = document.getElementById('infoLog');
const notesLog = document.getElementById('notesLogId');
const toast = document.getElementById('toast');

const cells = {
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

const loginInfoCells = {
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

const reportingCells = {
    ticket: document.getElementById('reportTicket'),
    name: document.getElementById('reportName'),
    division: document.getElementById('reportDivision'),
    dateReceived: document.getElementById('reportDateReceived'),
    startDate: document.getElementById('reportStartDate'),
    busDays: document.getElementById('reportBusDays'),
    solutionTeam: document.getElementById('reportSolutionTeam'),
    agentName: document.getElementById('reportAgentName')
};

const usersCells = {
    ticket: document.getElementById('usersTicket'),
    username: document.getElementById('usersUsername'),
    email: document.getElementById('usersEmail'),
    teams: document.getElementById('usersTeams'),
    cmicAccount: document.getElementById('usersCmicAccount'),
    password: document.getElementById('usersPassword')
};

function formatPhoneNumber(text) {
    const digits = text.replace(/\D/g, '');
    if (digits.length === 10) {
        return `(${digits.slice(0, 3)}) ${digits.slice(3, 6)}-${digits.slice(6)}`;
    }
    return text;
}

function initUsersEditableCells() {
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

function initUsersLabelCells() {
    const labels = document.querySelectorAll('.users-table .label-cell');
    labels.forEach(label => {
        label.style.cursor = 'pointer';
        label.addEventListener('click', async () => {
            const valueCell = label.nextElementSibling;
            if (valueCell && valueCell.textContent) {
                await navigator.clipboard.writeText(valueCell.textContent);
                showToast();
            }
        });
    });
}

function saveUsersTable() {
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

function loadUsersTable() {
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
}

function clearUsersTable() {
    Object.values(usersCells).forEach(cell => {
        cell.textContent = '';
    });
    localStorage.removeItem(USERS_TABLE_KEY);
}

const WORD_LIST_URL = 'https://cdn.jsdelivr.net/gh/bitcoin/bips@master/bip-0039/english.txt';

let _wordListPromise = null;

function loadWordList() {
    if (_wordListPromise) return _wordListPromise;

    _wordListPromise = fetch(WORD_LIST_URL)
        .then(r => r.text())
        .then(text => {
            return text.split('\n')
                .map(w => w.trim())
                .filter(w => w.length >= 4 && w.length <= 8);
        })
        .catch(err => {
            console.error('Failed to load word list:', err);
            _wordListPromise = null;
            throw err;
        });

    return _wordListPromise;
}

// Pre-fetch on page load
loadWordList();

async function generateUsersPassword() {
    const words = await loadWordList();
    const arr = new Uint32Array(3);
    crypto.getRandomValues(arr);
    const chosen = [
        words[arr[0] % words.length],
        words[arr[1] % words.length],
        words[arr[2] % words.length]
    ];
    const password = chosen.map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join('') + Math.floor(Math.random() * 10);

    usersCells.password.textContent = password;
    saveUsersTable();

    const btn = document.getElementById('generatePasswordBtn');
    btn.classList.add('copied');
    setTimeout(() => btn.classList.remove('copied'), 1500);
}

function getUsersTableHtml() {
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

function getUsersTableHtmlResponse() {
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

function getUsersTablePlainText() {
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

function getUsersTablePlainTextResponse() {
    const rows = [
        ['Username:', usersCells.username.textContent],
        ['Email:', usersCells.email.textContent],
        ['Teams:', usersCells.teams.textContent]
    ];
    
    return rows.map(([label, value]) => `${label}\t${value}`).join('\n');
}

async function copyUsersTableAsHtml(html, plainText) {
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

const months = { Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6, Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12 };

function parseDateReceived(str) {
    if (!str.startsWith('reported ')) return '';
    const start = str.indexOf('(');
    const end = str.indexOf(')');
    if (start === -1 || end === -1 || end <= start) return '';
    const inner = str.substring(start + 1, end);
    const currentYear = new Date().getFullYear();
    
    const matchWithYear = inner.match(/(\w+)\s+(\d+)\s+(\d{4})/);
    if (matchWithYear) {
        const month = months[matchWithYear[1]];
        if (month) {
            return `${month}/${matchWithYear[2]}/${matchWithYear[3]}`;
        }
    }
    
    const matchNoYear = inner.match(/(\w+)\s+(\d+)/);
    if (matchNoYear) {
        const month = months[matchNoYear[1]];
        if (month) {
            return `${month}/${matchNoYear[2]}/${currentYear}`;
        }
    }
    
    return '';
}

function parseStartDate(str) {
    const idx = str.indexOf('Start date:');
    if (idx !== -1) {
        const inner = str.substring(idx + 11).trim();
        const parts = inner.split(', ');
        if (parts.length >= 2) {
            const monthDay = parts[1].split(' ');
            if (monthDay.length >= 2) {
                const month = months[monthDay[0]];
                if (month) {
                    const year = parts[2] ? parts[2].trim().split(' ')[0] : '';
                    return `${month}/${monthDay[1]}/${year}`;
                }
            }
        }
    }
    return '';
}

function getBusinessDays(startDateStr, endDateStr) {
    const startDate = new Date(startDateStr);
    const endDate = new Date(endDateStr);
    
    const start = new Date(startDate);
    start.setHours(0, 0, 0, 0);
    const end = new Date(endDate);
    end.setHours(0, 0, 0, 0);
    
    if (isNaN(start.getTime()) || isNaN(end.getTime())) {
        return null;
    }
    
    let count = 0;
    const current = new Date(start);
    
    while (current < end) {
        const dayOfWeek = current.getDay();
        if (dayOfWeek !== 0 && dayOfWeek !== 6) {
            count++;
        }
        current.setDate(current.getDate() + 1);
    }
    
    while (current > end) {
        const dayOfWeek = current.getDay();
        if (dayOfWeek !== 0 && dayOfWeek !== 6) {
            count--;
        }
        current.setDate(current.getDate() - 1);
    }
    
    return count;
}

function parseTicketData(text) {
    const data = {
        ticket: '',
        dateReceived: '',
        dateWorked: '',
        name: '',
        user: '',
        email: '',
        jobTitle: '',
        empType: '',
        empTypeForLookup: '',
        startDate: '',
        division: '',
        location: '',
        phone: '',
        employeeId: '',
        cmicCode: '',
        supervisor: '',
        startDateFallback: false,
        startDateSource: '',
        solutionTeam: ''
    };

    const lines = text.split('\n').map(l => l.trim()).filter(l => l);

    for (const line of lines) {
        if (line.includes('#SR-')) {
            const idx = line.indexOf('#SR-');
            const num = line.substring(idx + 4).trim();
            if (num) {
                data.ticket = num;
                break;
            }
        }
    }

    for (const line of lines) {
        if (line.startsWith('reported ')) {
            const result = parseDateReceived(line);
            if (result) {
                data.dateReceived = result;
                break;
            }
        }
    }

    const today = new Date();
    data.dateWorked = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;

    for (const line of lines) {
        if (line.includes('New Hire') && line.includes(', ')) {
            const match = line.match(/,\s+(\S+)\s+-/);
            if (match) {
                const firstName = match[1];
                const beforeComma = line.substring(0, line.indexOf(','));
                const lastDash = beforeComma.lastIndexOf(' - ');
                if (lastDash !== -1) {
                    const lastName = beforeComma.substring(lastDash + 3).trim();
                    if (firstName && lastName) {
                        data.name = `${firstName} ${lastName}`;
                        break;
                    }
                }
            }
        }
    }

    if (!data.name) {
        let firstName = '', lastName = '';
        for (let i = 0; i < lines.length; i++) {
            if (lines[i] === 'Legal First Name' && lines[i + 1]) {
                firstName = lines[i + 1].trim();
            }
            if (lines[i] === 'Legal Last Name' && lines[i + 1]) {
                lastName = lines[i + 1].trim();
            }
        }
        if (firstName && lastName) {
            data.name = `${firstName} ${lastName}`;
        }
    }

    for (let i = 0; i < lines.length; i++) {
        if (lines[i] === 'Job Title' && lines[i + 1]) {
            data.jobTitle = lines[i + 1].trim();
            break;
        }
    }

for (let i = 0; i < lines.length; i++) {
        if (lines[i] === 'Employee Type' && lines[i + 1]) {
            data.empType = lines[i + 1].trim();
            break;
        }
    }
    
    const textLower = text.toLowerCase();
    
    data.empTypeForLookup = data.empType;
    
    if (/consultant/i.test(textLower)) {
        data.empTypeForLookup = 'Consultants';
    } else if (/temp/i.test(data.empType)) {
        data.empTypeForLookup = 'Temp';
    }
    
    if (/intern/i.test(data.jobTitle)) {
        data.empTypeForLookup = 'Intern';
    }
    
    if (!data.empTypeForLookup || data.empTypeForLookup === 'Admin') {
        data.empTypeForLookup = 'Non-Temp';
    }
    
    let hrNewHireDate = null;
    for (const line of lines) {
        if (line.includes('HR: New Hire')) {
            const nonSpace = line.split(' ')[0];
            if (nonSpace) {
                const parts = nonSpace.split('/');
                if (parts.length >= 2) {
                    let month = parseInt(parts[0]);
                    let day = parseInt(parts[1]);
                    let year = parts[2] ? parseInt(parts[2]) : new Date().getFullYear();
                    if (year < 100) {
                        year = year + 2000;
                    }
                    if (month >= 1 && month <= 12 && day >= 1 && day <= 31 && year >= 2000) {
                        data.startDate = `${month}/${day}/${year}`;
                        hrNewHireDate = data.startDate;
                    }
                }
                break;
            }
        }
    }

    if (!data.startDate) {
        for (const line of lines) {
            if (line.includes('Start date:')) {
                const result = parseStartDate(line);
                if (result) {
                    data.startDate = result;
                    data.startDateSource = 'the title';
                    break;
                }
            }
        }
    }

    if (!data.startDate) {
        for (let i = 0; i < lines.length; i++) {
            if (lines[i] === 'Start Date' && lines[i + 1]) {
                const rawDate = lines[i + 1].trim();
                const parsed = new Date(rawDate);
                if (!isNaN(parsed.getTime())) {
                    data.startDate = `${parsed.getMonth() + 1}/${parsed.getDate()}/${parsed.getFullYear()}`;
                } else {
                    data.startDate = rawDate;
                }
                data.startDateSource = 'the ticket body';
                data.startDateFallback = true;
                break;
            }
        }
    }

    if (!hrNewHireDate && data.startDate) {
        data.startDateFallback = true;
    }

    for (let i = 0; i < lines.length; i++) {
        if (lines[i] === 'Division' && lines[i + 1]) {
            data.division = lines[i + 1].trim();
            break;
        }
    }

    for (let i = 0; i < lines.length; i++) {
        if (lines[i].includes('CMiC Company Code') && lines[i + 1]) {
            data.cmicCode = lines[i + 1].trim();
            break;
        }
    }

for (let i = 0; i < lines.length; i++) {
        if (lines[i] === "Employee's Supervisor") {
            data.supervisor = lines[i + 1] ? lines[i + 1].trim() : '';
            break;
        }
    }
    
    for (let i = 0; i < lines.length; i++) {
        if (lines[i].includes('Contact information')) {
            data.solutionTeam = lines[i + 2] ? lines[i + 2].trim() : '';
            break;
        }
    }
    
    for (let i = 0; i < lines.length; i++) {
        if (lines[i] === 'Location' && lines[i + 1]) {
            data.location = lines[i + 1].trim();
            break;
        }
    }

    for (let i = 0; i < lines.length; i++) {
        if (lines[i].includes('Employee ID') && lines[i + 1]) {
            const match = lines[i + 1].match(/\d+/);
            if (match) {
                data.employeeId = match[0];
                break;
            }
        }
    }

    if (!data.employeeId && data.empType && /temp/i.test(data.empType)) {
        data.employeeId = 'N/A';
    }

    return data;
}

function showToast() {
    toast.classList.add('show');
    setTimeout(() => {
        toast.classList.remove('show');
    }, 2000);
}

function updateOutput(data) {
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
    loginInfoCells.agentName.textContent = agentNameInput.value || '';
    
    reportingCells.ticket.textContent = data.ticket || '';
    reportingCells.name.textContent = data.name || '';
    reportingCells.division.textContent = data.division || '';
    reportingCells.dateReceived.textContent = data.dateReceived || '';
    reportingCells.startDate.textContent = data.startDate || '';
    reportingCells.busDays.textContent = '';
    reportingCells.busDays.setAttribute('data-formula', '=NETWORKDAYS.INTL(INDEX(E:E,ROW()),INDEX(F:F,ROW()))');
    reportingCells.solutionTeam.textContent = data.solutionTeam || '';
    reportingCells.agentName.textContent = agentNameInput.value || '';

    Object.values(cells).forEach(cell => {
        cell.onclick = async () => {
            if (cell.textContent) {
                await navigator.clipboard.writeText(cell.textContent);
                showToast();
            }
        };
    });
    
    Object.values(loginInfoCells).forEach(cell => {
        cell.onclick = async () => {
            if (cell.textContent) {
                await navigator.clipboard.writeText(cell.textContent);
                showToast();
            }
        };
    });
}

function getRowString() {
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

function getLoginInfoRowString() {
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

function getReportingRowString() {
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

const adSection = document.getElementById('adSection');
const adDivisionEl = document.getElementById('adDivision');
const copyAdBtn = document.getElementById('copyAdBtn');

const cmicSection = document.getElementById('cmicSection');
const cmicCodeEl = document.getElementById('cmicCode');
const copyCmicBtn = document.getElementById('copyCmicBtn');

const supervisorSection = document.getElementById('supervisorSection');
const supervisorDisplayEl = document.getElementById('supervisorDisplay');
const copySupervisorNameBtn = document.getElementById('copySupervisorNameBtn');
const copySupervisorIdBtn = document.getElementById('copySupervisorIdBtn');

let currentSupervisor = '';

function formatAdDivision(divisionStr) {
    if (!divisionStr) return '';
    const match = divisionStr.match(/^(\d+)\s+(.+)$/);
    if (match) {
        return `${match[2]} (${match[1]})`;
    }
    return divisionStr;
}

function formatCmicCode(codeStr) {
    if (!codeStr) return '';
    const match = codeStr.match(/^(\d+)\s*-\s*(.+)$/);
    if (match) {
        return `${match[2]} (${match[1]})`;
    }
    return codeStr;
}

generateBtn.addEventListener('click', async () => {
    infoLog.textContent = 'While this tool aims to reduce the tedium of moving data around and to make unique requests stand out, you should not rely on it for all of your information. You must manually review the ticket when you work the account. Always double check and validate the data before using it!';
    infoLog.classList.add('show');
    
    const rehirePattern = /\bre\s*hire\b/i;
    const nonStandardPattern = /Non-Standard Computer/i;
    const notes = [];
    if (rehirePattern.test(ticketDataInput.value)) {
        notes.push('This ticket appears to be for a rehire.');
    }
    if (MIRROR_PATTERN.test(ticketDataInput.value)) {
        notes.push('This ticket appears to request a mirror.');
    }
    if (nonStandardPattern.test(ticketDataInput.value)) {
        notes.push('This ticket appears to request a non-standard computer.');
    }
    if (/DO NOT add New Hire to.*Mailing Distribution List/i.test(ticketDataInput.value)) {
        notes.push('This ticket appears to request no mailing list.');
    }
    
    const lines = ticketDataInput.value.split('\n');
    for (let i = 0; i < lines.length; i++) {
        if (lines[i].includes('CMiC Company Code') && lines[i + 1]) {
            const nextLine = lines[i + 1];
            if (!/S.{7}n/.test(nextLine)) {
                notes.push('This ticket appears to be for another company.');
            }
            break;
        }
    }
    
    let legalFirstName = '';
    let legalLastName = '';
    let preferredName = '';
    for (let i = 0; i < lines.length; i++) {
        if (lines[i].trim() === 'Legal First Name') {
            for (let j = i + 1; j < lines.length && j <= i + 5; j++) {
                if (lines[j].trim()) {
                    legalFirstName = lines[j].trim();
                    break;
                }
            }
        }
        if (lines[i].trim() === 'Legal Last Name') {
            for (let j = i + 1; j < lines.length && j <= i + 5; j++) {
                if (lines[j].trim()) {
                    legalLastName = lines[j].trim();
                    break;
                }
            }
        }
        if (lines[i].trim() === 'Preferred Name') {
            for (let j = i + 1; j < lines.length && j <= i + 5; j++) {
                if (lines[j].trim()) {
                    preferredName = lines[j].trim();
                    break;
                }
            }
        }
    }
    const fullLegalName = `${legalFirstName} ${legalLastName}`.toLowerCase();
    if (preferredName && legalFirstName && preferredName.toLowerCase() !== legalFirstName.toLowerCase() && preferredName.toLowerCase() !== fullLegalName) {
        notes.push('This ticket appears to request a rename.');
    }
    
    const needsRenameForChars = preferredName && legalFirstName && legalLastName &&
        (preferredName.toLowerCase() === legalFirstName.toLowerCase() || preferredName.toLowerCase() === fullLegalName) &&
        `${legalFirstName} ${legalLastName}`.length > 19;
    if (needsRenameForChars) {
        notes.push('This ticket might need a rename due to character limits.');
    }
    
    if (notes.length > 0) {
        notesLog.innerHTML = `<div class="notes-title">Notes</div><ul>${notes.map(n => `<li>${n}</li>`).join('')}</ul>`;
        notesLog.classList.add('show');
    } else {
        notesLog.classList.remove('show');
    }
    
    const data = parseTicketData(ticketDataInput.value);
    
    if (!data.dateReceived) {
        alert("The 'Date Received' could not be found. Ensure you are at the top of the ticket when copying, as FreshService will hide this info as you scroll down.");
    }
    
    const errors = [];
    if (!data.dateReceived) errors.push("Date Received");
    if (!data.ticket) errors.push("Ticket #");
    if (!data.name) errors.push("Name");
    if (!data.jobTitle) errors.push("Job Title");
    if (!data.empType) errors.push("Employee Type");
    if (!data.startDate) errors.push("Start Date");
    if (!data.division) errors.push("Division");
    if (!data.cmicCode) errors.push("CMiC Code");
    if (!data.supervisor) errors.push("Supervisor");
    if (!data.location) errors.push("Location");
    if (!data.employeeId) errors.push("Employee ID");
    if (!data.solutionTeam) errors.push("Solution Team Member");
    
    if (workbookData.adPerType.length > 0) {
        const empTypeMatched = workbookData.adPerType.some(row => row[0] && row[0].toLowerCase() === (data.empTypeForLookup || '').toLowerCase());
        if (data.empTypeForLookup && !empTypeMatched) {
            errors.push("Employee Type (For Lookup)");
        }
        
        const jobTitleMatched = workbookData.jobTitles.some(row => row[0] && row[0].toLowerCase() === (findBestJobTitleMatch(data.jobTitle) || '').toLowerCase());
        if (data.jobTitle && !jobTitleMatched) {
            errors.push("Job Title (For Lookup)");
        }
        
        const locationMatched = workbookData.adPerLocation.some(row => row[0] && row[0].toLowerCase() === (findBestLocationMatch(data.location, data.empTypeForLookup, ticketDataInput.value.toLowerCase()) || '').toLowerCase());
        if (data.location && !locationMatched) {
            errors.push("Location (For Lookup)");
        }
    }
    
    if (errors.length > 0) {
        errorLog.innerHTML = `<div class="error-title">Failed to parse:</div><ul>${errors.map(e => `<li>${e}</li>`).join('')}</ul>`;
        errorLog.classList.add('has-errors');
    } else {
        errorLog.classList.remove('has-errors');
    }

    if (data.startDateFallback) {
        const sourceNote = data.startDateSource ? ` (from ${data.startDateSource})` : '';
        warningLog.innerHTML = `<div class="warning-title">Parsed using fallback method:</div><ul><li>Start Date${sourceNote}</li></ul>`;
        warningLog.classList.add('has-warnings');
    } else {
        warningLog.classList.remove('has-warnings');
    }

    const businessDays = data.dateReceived && data.startDate 
        ? getBusinessDays(data.dateReceived, data.startDate) 
        : null;
    
    if (businessDays !== null && businessDays < 10) {
        slaWarningLog.innerHTML = `<div class="sla-warning-title">SLA Violation:</div><ul><li>Business days between Date Received and Start Date: ${businessDays} days (less than 10 days)</li></ul>`;
        slaWarningLog.classList.add('has-sla-warning');
    } else {
        slaWarningLog.classList.remove('has-sla-warning');
    }

    const isIntern = /intern/i.test(data.jobTitle);
    
    const textLower = ticketDataInput.value.toLowerCase();
    const hasKeyboard = textLower.includes('keyboard');
    const hasMouse = textLower.includes('mouse');
    const hasMonitor = textLower.includes('monitor');
    
    let keyboardRequested = false;
    let mouseRequested = false;
    let monitorRequested = false;
    
    const ticketLines = ticketDataInput.value.split('\n');
    for (let i = 0; i < ticketLines.length; i++) {
        const line = ticketLines[i].trim().toLowerCase();
        const nextLine = ticketLines[i + 1] ? ticketLines[i + 1].trim().toLowerCase() : '';
        
        if (hasKeyboard && line === 'keyboard') {
            if (nextLine !== '' && nextLine !== 'none' && !nextLine.includes('n/a')) {
                keyboardRequested = true;
            }
        }
        if (hasMouse && line === 'mouse') {
            if (nextLine !== '' && nextLine !== 'none' && !nextLine.includes('n/a')) {
                mouseRequested = true;
            }
        }
        if (hasMonitor && line === 'monitor') {
            if (nextLine !== '' && nextLine !== 'none' && !nextLine.includes('n/a')) {
                monitorRequested = true;
            }
        }
    }
    
    if (!isIntern && !keyboardRequested && !mouseRequested && !monitorRequested) {
        notes.push('This ticket appears to not request any peripherals.');
    }
    
    if (isIntern && (keyboardRequested || mouseRequested || monitorRequested)) {
        notes.push('This ticket appears to have requested peripherals, despite being for an intern.');
    }
    
    if (notes.length > 0) {
        notesLog.innerHTML = `<div class="notes-title">Notes</div><ul>${notes.map(n => `<li>${n}</li>`).join('')}</ul>`;
        notesLog.classList.add('show');
    } else {
        notesLog.classList.remove('show');
    }
    
    updateOutput(data);
    
    if (workbookData.adPerType.length > 0) {
        const lookupSection = document.getElementById('lookupSection');
        const resultsSection = document.getElementById('resultsSection');
        
        document.getElementById('jobTypeSelect').selectedIndex = 0;
        document.getElementById('jobTitleSelect').selectedIndex = 0;
        document.getElementById('locationSelect').selectedIndex = 0;
        
        if (data.empType) {
            selectDropdownValue('jobTypeSelect', data.empTypeForLookup);
        }
        if (data.jobTitle) {
            selectDropdownValue('jobTitleSelect', findBestJobTitleMatch(data.jobTitle));
        }
        if (data.location) {
            selectDropdownValue('locationSelect', findBestLocationMatch(data.location, data.empTypeForLookup, ticketDataInput.value.toLowerCase()));
        }
        
        updateLookups();
        lookupSection.style.display = 'flex';
        resultsSection.style.display = 'flex';
        document.getElementById('lookupWarning').classList.add('show');
        
        const isMirror = MIRROR_PATTERN.test(ticketDataInput.value);
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
    
    const adFormatted = formatAdDivision(data.division);
    adDivisionEl.textContent = adFormatted;
    setTimeout(() => {
        adSection.style.display = 'flex';
    }, 100);
    adSection.style.animationDelay = '0.1s';
    
    const cmicFormatted = formatCmicCode(data.cmicCode);
    cmicCodeEl.textContent = cmicFormatted;
    setTimeout(() => {
        cmicSection.style.display = 'flex';
    }, 200);
    cmicSection.style.animationDelay = '0.2s';
    
    supervisorDisplayEl.textContent = data.supervisor || '';
    currentSupervisor = data.supervisor || '';
    
    const hasSupervisorId = /\(\d+\)/.test(currentSupervisor);
    copySupervisorIdBtn.style.display = hasSupervisorId ? 'flex' : 'none';
    
    setTimeout(() => {
        supervisorSection.style.display = 'flex';
    }, 300);
    supervisorSection.style.animationDelay = '0.3s';

    await generateUsersPassword();
});

function resetAnimations() {
    const sections = [outputSection, adSection, cmicSection, supervisorSection];
    const loginInfoSection = document.getElementById('loginInfoSection');
    
    sections.forEach(section => {
        section.style.display = 'none';
        section.style.animation = 'none';
        section.style.animationDelay = '';
    });
    
    if (loginInfoSection) {
        loginInfoSection.style.display = 'none';
    }
    
    const reportingSection = document.getElementById('reportingSection');
    if (reportingSection) {
        reportingSection.style.display = 'none';
    }
    
    setTimeout(() => {
        sections.forEach(section => {
            section.style.animation = '';
        });
    }, 10);
}

copyBtn.addEventListener('click', async () => {
    const rowString = getRowString();
    await navigator.clipboard.writeText(rowString);
    copyBtn.classList.add('copied');
    setTimeout(() => {
        copyBtn.classList.remove('copied');
    }, 1500);
});

document.getElementById('copyLoginInfoBtn').addEventListener('click', async () => {
    const rowString = getLoginInfoRowString();
    await navigator.clipboard.writeText(rowString);
    document.getElementById('copyLoginInfoBtn').classList.add('copied');
    setTimeout(() => {
        document.getElementById('copyLoginInfoBtn').classList.remove('copied');
    }, 1500);
});

document.getElementById('copyReportingBtn').addEventListener('click', async () => {
    const rowString = getReportingRowString();
    await navigator.clipboard.writeText(rowString);
    document.getElementById('copyReportingBtn').classList.add('copied');
    setTimeout(() => {
        document.getElementById('copyReportingBtn').classList.remove('copied');
    }, 1500);
});

copyAdBtn.addEventListener('click', async () => {
    await navigator.clipboard.writeText(adDivisionEl.textContent);
    copyAdBtn.classList.add('copied');
    setTimeout(() => {
        copyAdBtn.classList.remove('copied');
    }, 1500);
});

copyCmicBtn.addEventListener('click', async () => {
    await navigator.clipboard.writeText(cmicCodeEl.textContent);
    copyCmicBtn.classList.add('copied');
    setTimeout(() => {
        copyCmicBtn.classList.remove('copied');
    }, 1500);
});

copySupervisorNameBtn.addEventListener('click', async () => {
    const name = currentSupervisor.replace(/\s*\(\d+\)\s*$/, '').trim();
    await navigator.clipboard.writeText(name);
    copySupervisorNameBtn.classList.add('copied');
    setTimeout(() => {
        copySupervisorNameBtn.classList.remove('copied');
    }, 1500);
});

copySupervisorIdBtn.addEventListener('click', async () => {
    const match = currentSupervisor.match(/\((\d+)\)/);
    const id = match ? match[1] : '';
    await navigator.clipboard.writeText(id);
    copySupervisorIdBtn.classList.add('copied');
    setTimeout(() => {
        copySupervisorIdBtn.classList.remove('copied');
    }, 1500);
});

document.getElementById('copyUsersBtn').addEventListener('click', async () => {
    await copyUsersTableAsHtml(getUsersTableHtml(), getUsersTablePlainText());
    const copyUsersBtn = document.getElementById('copyUsersBtn');
    copyUsersBtn.classList.add('copied');
    setTimeout(() => {
        copyUsersBtn.classList.remove('copied');
    }, 1500);
});

document.getElementById('copyUsersResponseBtn').addEventListener('click', async () => {
    await copyUsersTableAsHtml(getUsersTableHtmlResponse(), getUsersTablePlainTextResponse());
    const copyUsersResponseBtn = document.getElementById('copyUsersResponseBtn');
    copyUsersResponseBtn.classList.add('copied');
    setTimeout(() => {
        copyUsersResponseBtn.classList.remove('copied');
    }, 1500);
});

document.getElementById('clearUsersBtn').addEventListener('click', () => {
    clearUsersTable();
    const clearUsersBtn = document.getElementById('clearUsersBtn');
    clearUsersBtn.classList.add('copied');
    setTimeout(() => {
        clearUsersBtn.classList.remove('copied');
    }, 1500);
});

document.getElementById('generatePasswordBtn').addEventListener('click', () => generateUsersPassword());

clearBtn.addEventListener('click', () => {
    ticketDataInput.value = '';
    updateOutput({});
    resetAnimations();
    supervisorDisplayEl.textContent = '';
    adDivisionEl.textContent = '';
    cmicCodeEl.textContent = '';
    currentSupervisor = '';
    errorLog.classList.remove('has-errors');
    warningLog.classList.remove('has-warnings');
    slaWarningLog.classList.remove('has-sla-warning');
    infoLog.classList.remove('show');
    document.getElementById('lookupWarning').classList.remove('show');
    notesLog.classList.remove('show');
    clearBtn.classList.remove('show');
    clearUsersTable();
});

if (loadFromStorage()) {
    populateDropdowns();
    updateFileStatus(true);
}

loadUsersTable();
initUsersEditableCells();
initUsersLabelCells();

const savedAgentName = localStorage.getItem(AGENT_NAME_KEY);
if (savedAgentName) {
    agentNameInput.value = savedAgentName;
}

agentNameInput.addEventListener('input', () => {
    localStorage.setItem(AGENT_NAME_KEY, agentNameInput.value);
});

clearAllStorageBtn.addEventListener('click', () => {
    console.log('Clear button clicked');
    if (confirm('This will clear all stored Excel data, agent name, and the users table. Are you sure?')) {
        console.log('Confirmed, clearing...');
        localStorage.clear();
        workbookData = {
            adPerType: [],
            jobTitles: [],
            locationMap: [],
            adPerLocation: [],
            jobSoftwares: [],
            cmic: [],
            cmicLocationMap: []
        };
        document.getElementById('jobTypeSelect').innerHTML = '';
        document.getElementById('jobTitleSelect').innerHTML = '';
        document.getElementById('locationSelect').innerHTML = '';
        document.getElementById('adRolesPerType').textContent = '';
        document.getElementById('adRolesPerJobTitle').textContent = '';
        document.getElementById('adRolesPerLocation').textContent = '';
        document.getElementById('cmicUserSettings').textContent = '';
        document.getElementById('cmicUserAccess').textContent = '';
        document.getElementById('o365Groups').textContent = '';
        document.getElementById('areaCodes').textContent = '';
        excelFileInput.value = '';
        agentNameInput.value = '';
        clearUsersTable();
        updateFileStatus(false);
    }
});