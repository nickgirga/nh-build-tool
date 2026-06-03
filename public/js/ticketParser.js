import { MONTHS } from './constants.js';
import { checkForUniqueEmail } from './lookups.js';

// Extracts the Date Received from the FreshService ticket header line.
export function parseDateReceived(str) {
    if (!str.startsWith('reported ')) return '';
    const start = str.indexOf('(');
    const end = str.indexOf(')');
    if (start === -1 || end === -1 || end <= start) return '';
    const inner = str.substring(start + 1, end);
    const currentYear = new Date().getFullYear();
    
    const matchWithYear = inner.match(/(\w+)\s+(\d+)\s+(\d{4})/);
    if (matchWithYear) {
        const month = MONTHS[matchWithYear[1]];
        if (month) {
            return `${month}/${matchWithYear[2]}/${matchWithYear[3]}`;
        }
    }
    
    const matchNoYear = inner.match(/(\w+)\s+(\d+)/);
    if (matchNoYear) {
        const month = MONTHS[matchNoYear[1]];
        if (month) {
            return `${month}/${matchNoYear[2]}/${currentYear}`;
        }
    }
    
    return '';
}

// Extracts the Start Date from the ticket subject/body line containing "Start date:".
export function parseStartDate(str) {
    const idx = str.indexOf('Start date:');
    if (idx !== -1) {
        const inner = str.substring(idx + 11).trim();
        const parts = inner.split(', ');
        if (parts.length >= 2) {
            const monthDay = parts[1].split(' ');
            if (monthDay.length >= 2) {
                const month = MONTHS[monthDay[0]];
                if (month) {
                    const year = parts[2] ? parts[2].trim().split(' ')[0] : '';
                    return `${month}/${monthDay[1]}/${year}`;
                }
            }
        }
    }
    return '';
}

// Calculates business days between two MM/DD/YYYY date strings (ignores weekends).
export function getBusinessDays(startDateStr, endDateStr) {
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

// Parses pasted FreshService ticket text into structured fields for the build tool.
export function parseTicketData(text) {
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
        solutionTeam: '',
        emailDomain: ''
    };

    const lines = text.split('\n').map(l => l.trim()).filter(l => l);

    // Extract ticket number
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

    // Extract date received
    for (const line of lines) {
        if (line.startsWith('reported ')) {
            const result = parseDateReceived(line);
            if (result) {
                data.dateReceived = result;
                break;
            }
        }
    }

    // Set date worked to today
    const today = new Date();
    data.dateWorked = `${today.getMonth() + 1}/${today.getDate()}/${today.getFullYear()}`;

    // Extract name from "New Hire" subject line
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

    // Fallback name extraction from Legal First/Last Name fields
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

    // Extract job title
    for (let i = 0; i < lines.length; i++) {
        if (lines[i] === 'Job Title' && lines[i + 1]) {
            data.jobTitle = lines[i + 1].trim();
            break;
        }
    }

    // Extract employee type
    for (let i = 0; i < lines.length; i++) {
        if (lines[i] === 'Employee Type' && lines[i + 1]) {
            data.empType = lines[i + 1].trim();
            break;
        }
    }
    
    const textLower = text.toLowerCase();
    
    // Normalize employee type for lookup purposes
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
    
    // Extract start date from HR: New Hire line
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

    // Fallback start date from ticket title
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

    // Second fallback start date from Start Date field
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

    // Extract division
    for (let i = 0; i < lines.length; i++) {
        if (lines[i] === 'Division' && lines[i + 1]) {
            data.division = lines[i + 1].trim();
            break;
        }
    }

    // Extract CMiC company code
    for (let i = 0; i < lines.length; i++) {
        if (lines[i].includes('CMiC Company Code') && lines[i + 1]) {
            data.cmicCode = lines[i + 1].trim();
            break;
        }
    }

    // If CMiC Company Code is "Other", fallback to "CMiC Company Code not listed above"
    if (/other/i.test(data.cmicCode)) {
        for (let i = 0; i < lines.length; i++) {
            if (lines[i].includes('CMiC Company Code not listed above') && lines[i + 1]) {
                data.cmicCode = lines[i + 1].trim();
                break;
            }
        }
    }

    // Extract supervisor
    for (let i = 0; i < lines.length; i++) {
        if (lines[i] === "Employee's Supervisor") {
            data.supervisor = lines[i + 1] ? lines[i + 1].trim() : '';
            break;
        }
    }
    
    // Extract solution team member from Contact information / Requester information section
    for (let i = 0; i < lines.length; i++) {
        if (lines[i].includes('Contact information') || lines[i].includes('Requester information')) {
            data.solutionTeam = lines[i + 2] ? lines[i + 2].trim() : '';
            break;
        }
    }
    
    // Extract location
    for (let i = 0; i < lines.length; i++) {
        if (lines[i] === 'Location' && lines[i + 1]) {
            data.location = lines[i + 1].trim();
            break;
        }
    }

    // Extract employee ID
    for (let i = 0; i < lines.length; i++) {
        if (lines[i].includes('Employee ID') && lines[i + 1]) {
            const match = lines[i + 1].match(/\d+/);
            if (match) {
                data.employeeId = match[0];
                break;
            }
        }
    }

    // Temp employees don't get an employee ID
    if (!data.employeeId && data.empType && /temp/i.test(data.empType)) {
        data.employeeId = 'N/A';
    }

    // Extract desired email domain
    for (let i = 0; i < lines.length; i++) {
        if (lines[i].includes('Desired Email Domain Name') && lines[i + 1]) {
            data.emailDomain = lines[i + 1].trim().replace(/^@/, '');
            break;
        }
    }

    data.emailDomain = checkForUniqueEmail(data.emailDomain, textLower);

    return data;
}
