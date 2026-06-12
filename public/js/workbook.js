import { XLS, STORAGE_KEY, EXPECTED_MAJOR_VERSION, EXPECTED_MINOR_VERSION } from './constants.js';

// In-memory cache of parsed Excel workbook data.
export let workbookData = {
    adPerType: [],
    jobTitles: [],
    locationMap: [],
    adPerLocation: [],
    jobSoftwares: [],
    cmic: [],
    cmicLocationMap: [],
    emailDomains: [],
    emailCommandSetupScript: [],
    logonScript: '',
    cmicPassword: '',
    emailTemplate1: '',
    emailTemplate2: '',
    emailTemplate3: '',
    adobeCcAdGroup: '',
    version: '',
    checklistHash: '',
    pimActivationHash: '',
    trackerUrl: ''
};

// Resets the in-memory workbook cache to empty defaults.
export function resetWorkbookData() {
    workbookData = {
        adPerType: [],
        jobTitles: [],
        locationMap: [],
        adPerLocation: [],
        jobSoftwares: [],
        cmic: [],
        cmicLocationMap: [],
        emailDomains: [],
        emailCommandSetupScript: [],
        logonScript: '',
        cmicPassword: '',
        emailTemplate1: '',
        emailTemplate2: '',
        emailTemplate3: '',
        adobeCcAdGroup: '',
        version: '',
        checklistHash: '',
        pimActivationHash: '',
        trackerUrl: ''
    };
}

// Checks the workbook version string and shows an alert if major/minor exceed expectations.
export function checkWorkbookVersion(versionString) {
    if (!versionString || typeof versionString !== 'string') return;
    const parts = versionString.split('.');
    const major = parseInt(parts[0], 10);
    const minor = parseInt(parts[1], 10);
    if (!isNaN(major) && major > EXPECTED_MAJOR_VERSION) {
        alert('This workbook is very likely not to work with the tool, as the major version is greater than expected.');
        return;
    }
    if (!isNaN(minor) && minor > EXPECTED_MINOR_VERSION) {
        alert('This workbook may not work as expected with the tool, as the minor version is greater than expected.');
    }
}

// Converts an Excel worksheet into an array of row arrays (skips the header row).
export function parseSheetToArrays(worksheet, headerRow = 1) {
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

// Parses the uploaded Excel workbook and caches the extracted lookup data.
export function parseWorkbook(wb) {
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
    
    const emailDomains = [];
    const emailCommandSetupScript = [];
    let logonScript = '';
    let cmicPassword = '';
    let emailTemplate1 = '';
    let emailTemplate2 = '';
    let emailTemplate3 = '';
    let adobeCcAdGroup = '';
    let version = '';
    let checklistHash = '';
    let pimActivationHash = '';
    let trackerUrl = '';

    if (wb.Sheets['Metadata']) {
        const metadataSheet = wb.Sheets['Metadata'];
        const versionRef = XLS.utils.encode_cell({ r: 1, c: 1 });
        const versionCell = metadataSheet[versionRef];
        if (versionCell && versionCell.v) {
            version = String(versionCell.v).trim();
        }
        checkWorkbookVersion(version);

        const hashRef = XLS.utils.encode_cell({ r: 2, c: 1 });
        const hashCell = metadataSheet[hashRef];
        if (hashCell && hashCell.v) {
            checklistHash = String(hashCell.v).trim();
        }

        const pimHashRef = XLS.utils.encode_cell({ r: 3, c: 1 });
        const pimHashCell = metadataSheet[pimHashRef];
        if (pimHashCell && pimHashCell.v) {
            pimActivationHash = String(pimHashCell.v).trim();
        }

        const trackerRef = XLS.utils.encode_cell({ r: 4, c: 1 });
        const trackerCell = metadataSheet[trackerRef];
        if (trackerCell && trackerCell.v) {
            trackerUrl = String(trackerCell.v).trim();
        }
    }

    if (wb.Sheets['Master Search']) {
        const masterSheet = wb.Sheets['Master Search'];
        for (let row = 7; row <= 12; row++) {
            const cellRef = XLS.utils.encode_cell({ r: row, c: 9 });
            const cell = masterSheet[cellRef];
            if (cell && cell.v) {
                emailDomains.push(String(cell.v).trim());
            }
        }
        for (let row = 6; row <= 8; row++) {
            const cellRef = XLS.utils.encode_cell({ r: row, c: 0 });
            const cell = masterSheet[cellRef];
            if (cell && cell.v) {
                emailCommandSetupScript.push(String(cell.v).trim());
            }
        }
        const logonRef = XLS.utils.encode_cell({ r: 4, c: 0 });
        const logonCell = masterSheet[logonRef];
        if (logonCell && logonCell.v) {
            logonScript = String(logonCell.v).trim();
        }
        const cmicPwRef = XLS.utils.encode_cell({ r: 4, c: 4 });
        const cmicPwCell = masterSheet[cmicPwRef];
        if (cmicPwCell && cmicPwCell.v) {
            cmicPassword = String(cmicPwCell.v).trim();
        }
        const t1Ref = XLS.utils.encode_cell({ r: 14, c: 9 });
        const t1Cell = masterSheet[t1Ref];
        if (t1Cell && t1Cell.v) emailTemplate1 = String(t1Cell.v).trim();
        const t2Ref = XLS.utils.encode_cell({ r: 14, c: 10 });
        const t2Cell = masterSheet[t2Ref];
        if (t2Cell && t2Cell.v) emailTemplate2 = String(t2Cell.v).trim();
        const t3Ref = XLS.utils.encode_cell({ r: 14, c: 11 });
        const t3Cell = masterSheet[t3Ref];
        if (t3Cell && t3Cell.v) emailTemplate3 = String(t3Cell.v).trim();

        const adobeRef = XLS.utils.encode_cell({ r: 4, c: 1 });
        const adobeCell = masterSheet[adobeRef];
        if (adobeCell && adobeCell.f) {
            const formula = adobeCell.f;
            const matches = formula.match(/"([^"]*)"/g);
            if (matches) {
                for (const match of matches) {
                    const value = match.replace(/"/g, '');
                    if (value && value !== 'Marketing') {
                        adobeCcAdGroup = value;
                        break;
                    }
                }
            }
        }
    }

    workbookData = {
        adPerType: sheets['AD Per Type'] || [],
        jobTitles: sheets['Job Titles'] || [],
        locationMap: sheets['Location Map'] || [],
        adPerLocation: sheets['AD Per Location'] || [],
        jobSoftwares: sheets['Job Softwares'] || [],
        cmic: sheets['CMiC'] || [],
        cmicLocationMap: sheets['CMiC Location Map'] || [],
        emailDomains: emailDomains,
        emailCommandSetupScript: emailCommandSetupScript,
        logonScript: logonScript,
        cmicPassword: cmicPassword,
        emailTemplate1: emailTemplate1,
        emailTemplate2: emailTemplate2,
        emailTemplate3: emailTemplate3,
        adobeCcAdGroup: adobeCcAdGroup,
        version: version,
        checklistHash: checklistHash,
        pimActivationHash: pimActivationHash,
        trackerUrl: trackerUrl
    };
    
    const dataToStore = {
        timestamp: Date.now(),
        ...workbookData
    };
    
    localStorage.setItem(STORAGE_KEY, JSON.stringify(dataToStore));
    
    return workbookData;
}

// Restores workbook data from localStorage if available.
export function loadFromStorage() {
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
                cmicLocationMap: data.cmicLocationMap || [],
                emailDomains: data.emailDomains || [],
                emailCommandSetupScript: data.emailCommandSetupScript || [],
                logonScript: data.logonScript || '',
                cmicPassword: data.cmicPassword || '',
                emailTemplate1: data.emailTemplate1 || '',
                emailTemplate2: data.emailTemplate2 || '',
                emailTemplate3: data.emailTemplate3 || '',
                adobeCcAdGroup: data.adobeCcAdGroup || '',
                version: data.version || '',
                checklistHash: data.checklistHash || '',
                pimActivationHash: data.pimActivationHash || '',
                trackerUrl: data.trackerUrl || ''
            };
            checkWorkbookVersion(workbookData.version);
            return true;
        } catch (e) {
            console.error('Failed to load from storage:', e);
            return false;
        }
    }
    return false;
}

// Displays the Email Command Setup Script, Logon Script, and CMiC Password in the Clipboard Utilities section.
export function populateClipboardUtilities() {
    const emailCommandEl = document.getElementById('emailCommandSetupScript');
    const logonScriptEl = document.getElementById('logonScript');
    const cmicPasswordEl = document.getElementById('cmicPassword');

    emailCommandEl.textContent = workbookData.emailCommandSetupScript.join('\n');
    logonScriptEl.textContent = workbookData.logonScript;
    cmicPasswordEl.textContent = workbookData.cmicPassword;

    const hasData = workbookData.emailCommandSetupScript.length > 0 || workbookData.logonScript || workbookData.cmicPassword;
    document.getElementById('clipboardSection').style.display = hasData ? 'flex' : 'none';
}

// Fills the Job Information dropdowns with values from the parsed workbook.
export function populateDropdowns() {
    const jobTypeSelect = document.getElementById('jobTypeSelect');
    const jobTitleSelect = document.getElementById('jobTitleSelect');
    const locationSelect = document.getElementById('locationSelect');
    const emailDomainSelect = document.getElementById('emailDomainSelect');

    jobTypeSelect.innerHTML = '';
    jobTitleSelect.innerHTML = '';
    locationSelect.innerHTML = '';
    emailDomainSelect.innerHTML = '';

    const defaultOption = document.createElement('option');
    defaultOption.value = '';
    defaultOption.textContent = 'Select...';
    defaultOption.selected = true;
    jobTypeSelect.appendChild(defaultOption.cloneNode(true));
    jobTitleSelect.appendChild(defaultOption.cloneNode(true));
    locationSelect.appendChild(defaultOption.cloneNode(true));
    emailDomainSelect.appendChild(defaultOption.cloneNode(true));

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

    workbookData.emailDomains.forEach(domain => {
        const option = document.createElement('option');
        option.value = domain;
        option.textContent = domain;
        emailDomainSelect.appendChild(option);
    });

    const otherOption = document.createElement('option');
    otherOption.value = '__other__';
    otherOption.textContent = 'Other';
    emailDomainSelect.appendChild(otherOption);
}

// Updates the file status text and shows/hides lookup sections based on whether data is loaded.
export function updateFileStatus(loaded) {
    const fileStatus = document.getElementById('fileStatus');
    const lookupSection = document.getElementById('lookupSection');
    const resultsSection = document.getElementById('resultsSection');
    const powerBiSection = document.getElementById('powerBiSection');
    const workbookWarning = document.getElementById('workbookWarning');

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
        fileStatus.classList.add('show');
        lookupSection.style.display = 'flex';
        resultsSection.style.display = 'flex';
        if (powerBiSection) powerBiSection.style.display = 'flex';
        if (workbookWarning) workbookWarning.classList.remove('show');
    } else {
        fileStatus.textContent = '';
        fileStatus.classList.remove('loaded');
        fileStatus.classList.remove('show');
        lookupSection.style.display = 'none';
        resultsSection.style.display = 'none';
        if (powerBiSection) powerBiSection.style.display = 'none';
        if (workbookWarning) workbookWarning.classList.add('show');
    }
}
