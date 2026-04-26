import { workbookData } from './workbook.js';
import { findBestJobTitleMatch, findBestLocationMatch } from './lookups.js';
import { getBusinessDays } from './ticketParser.js';
import { MIRROR_PATTERN } from './constants.js';

// Shows the permanent info banner reminding users to manually validate ticket data.
export function showInfoBanner() {
    const infoLog = document.getElementById('infoLog');
    infoLog.textContent = 'While this tool aims to reduce the tedium of moving data around and to make unique requests stand out, you should not rely on it for all of your information. You must manually review the ticket when you work the account. Always double check and validate the data before using it!';
    infoLog.classList.add('show');
}

// Hides and resets all banner sections (errors, warnings, SLA, info, notes, lookup warning).
export function clearBanners() {
    document.getElementById('errorLog').classList.remove('has-errors');
    document.getElementById('warningLog').classList.remove('has-warnings');
    document.getElementById('slaWarningLog').classList.remove('has-sla-warning');
    document.getElementById('infoLog').classList.remove('show');
    document.getElementById('agentNameBanner').classList.remove('show');
    document.getElementById('notesLogId').classList.remove('show');
    document.getElementById('lookupWarning').classList.remove('show');
}

// Builds the list of informational notes from raw ticket text (before parsing).
export function buildPreParseNotes(ticketText) {
    const notes = [];
    const lines = ticketText.split('\n');
    const textLower = ticketText.toLowerCase();

    if (/\bre\s*hire\b/i.test(ticketText)) {
        notes.push('This ticket appears to be for a rehire.');
    }
    if (MIRROR_PATTERN.test(ticketText)) {
        notes.push('This ticket appears to request a mirror.');
    }
    if (/Non-Standard Computer/i.test(ticketText)) {
        notes.push('This ticket appears to request a non-standard computer.');
    }
    if (/DO NOT add New Hire to.*Mailing Distribution List/i.test(ticketText)) {
        notes.push('This ticket appears to request no mailing list.');
    }

    for (let i = 0; i < lines.length; i++) {
        if (lines[i].includes('CMiC Company Code') && lines[i + 1]) {
            if (!/S.{7}n/.test(lines[i + 1])) {
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

    let parsedEmailDomain = '';
    for (let i = 0; i < lines.length; i++) {
        if (lines[i].includes('Desired Email Domain Name') && lines[i + 1]) {
            parsedEmailDomain = lines[i + 1].trim().replace(/^@/, '');
            break;
        }
    }
    const normalDomainPattern = new RegExp('s.......n\\.com');
    const uniqueDomainPattern = new RegExp('s.........c\\.com');
    const hasUniqueEmailDomain = (parsedEmailDomain && !normalDomainPattern.test(parsedEmailDomain.toLowerCase())) || uniqueDomainPattern.test(textLower);
    if (hasUniqueEmailDomain) {
        notes.push('This ticket appears to request a unique email domain.');
    }

    return notes;
}

// Builds notes that depend on parsed ticket data (e.g. peripheral checks for interns).
export function buildPostParseNotes(data, ticketText) {
    const notes = [];
    const textLower = ticketText.toLowerCase();
    const isIntern = /intern/i.test(data.jobTitle);
    const hasKeyboard = textLower.includes('keyboard');
    const hasMouse = textLower.includes('mouse');
    const hasMonitor = textLower.includes('monitor');

    let keyboardRequested = false;
    let mouseRequested = false;
    let monitorRequested = false;

    const ticketLines = ticketText.split('\n');
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

    return notes;
}

// Builds the list of parse/validation errors based on extracted ticket data.
export function buildErrors(data, ticketText) {
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
    if (!data.emailDomain) errors.push("Email Domain");
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

        const locationMatched = workbookData.adPerLocation.some(row => row[0] && row[0].toLowerCase() === (findBestLocationMatch(data.location, data.empTypeForLookup, ticketText.toLowerCase()) || '').toLowerCase());
        if (data.location && !locationMatched) {
            errors.push("Location (For Lookup)");
        }

        const emailDomainMatched = workbookData.emailDomains.some(domain => domain.toLowerCase() === (data.emailDomain || '').toLowerCase());
        if (data.emailDomain && !emailDomainMatched) {
            errors.push("Email Domain (For Lookup)");
        }
    }

    return errors;
}

// Checks whether the start date had to be extracted via a fallback method.
export function buildWarning(data) {
    if (data.startDateFallback) {
        const sourceNote = data.startDateSource ? ` (from ${data.startDateSource})` : '';
        return { message: `Start Date${sourceNote}` };
    }
    return null;
}

// Checks for an SLA violation (fewer than 10 business days between received and start).
export function buildSlaWarning(data) {
    const businessDays = data.dateReceived && data.startDate
        ? getBusinessDays(data.dateReceived, data.startDate)
        : null;
    if (businessDays !== null && businessDays < 10) {
        return { days: businessDays };
    }
    return null;
}

// Renders the Notes banner. Pass an empty array to hide it.
export function renderNotes(notes) {
    const notesLog = document.getElementById('notesLogId');
    if (notes.length > 0) {
        notesLog.innerHTML = `<div class="notes-title">Notes</div><ul>${notes.map(n => `<li>${n}</li>`).join('')}</ul>`;
        notesLog.classList.add('show');
    } else {
        notesLog.classList.remove('show');
    }
}

// Renders the Errors banner. Pass an empty array to hide it.
export function renderErrors(errors) {
    const errorLog = document.getElementById('errorLog');
    if (errors.length > 0) {
        errorLog.innerHTML = `<div class="error-title">Failed to parse:</div><ul>${errors.map(e => `<li>${e}</li>`).join('')}</ul>`;
        errorLog.classList.add('has-errors');
    } else {
        errorLog.classList.remove('has-errors');
    }
}

// Renders the fallback-method Warning banner.
export function renderWarning(warning) {
    const warningLog = document.getElementById('warningLog');
    if (warning) {
        warningLog.innerHTML = `<div class="warning-title">Parsed using fallback method:</div><ul><li>${warning.message}</li></ul>`;
        warningLog.classList.add('has-warnings');
    } else {
        warningLog.classList.remove('has-warnings');
    }
}

// Renders the SLA Violation banner.
export function renderSlaWarning(sla) {
    const slaWarningLog = document.getElementById('slaWarningLog');
    if (sla) {
        slaWarningLog.innerHTML = `<div class="sla-warning-title">SLA Violation:</div><ul><li>Business days between Date Received and Start Date: ${sla.days} days (less than 10 days)</li></ul>`;
        slaWarningLog.classList.add('has-sla-warning');
    } else {
        slaWarningLog.classList.remove('has-sla-warning');
    }
}
