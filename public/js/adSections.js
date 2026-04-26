// DOM references for the AD Department, CMiC Company, and Manager sections.
export const adSection = document.getElementById('adSection');
export const adDivisionEl = document.getElementById('adDivision');
export const copyAdBtn = document.getElementById('copyAdBtn');

export const cmicSection = document.getElementById('cmicSection');
export const cmicCodeEl = document.getElementById('cmicCode');
export const copyCmicBtn = document.getElementById('copyCmicBtn');

export const supervisorSection = document.getElementById('supervisorSection');
export const supervisorDisplayEl = document.getElementById('supervisorDisplay');
export const copySupervisorNameBtn = document.getElementById('copySupervisorNameBtn');
export const copySupervisorIdBtn = document.getElementById('copySupervisorIdBtn');

// Stores the currently displayed supervisor string.
export let currentSupervisor = '';

// Updates the supervisor state from outside this module (used by clear handlers).
export function setCurrentSupervisor(value) {
    currentSupervisor = value;
}

// Reformats a division string from "123 Name" to "Name (123)" for Active Directory.
export function formatAdDivision(divisionStr) {
    if (!divisionStr) return '';
    const match = divisionStr.match(/^(\d+)\s+(.+)$/);
    if (match) {
        return `${match[2]} (${match[1]})`;
    }
    return divisionStr;
}

// Reformats a CMiC code from "123 - Name" to "Name (123)" for Active Directory.
export function formatCmicCode(codeStr) {
    if (!codeStr) return '';
    const match = codeStr.match(/^(\d+)\s*-\s*(.+)$/);
    if (match) {
        return `${match[2]} (${match[1]})`;
    }
    return codeStr;
}

// Updates the AD Department, CMiC Company, and Manager sections with parsed data and staged reveal animations.
export function updateAdSections(data) {
    const adFormatted = formatAdDivision(data.division);
    adDivisionEl.textContent = adFormatted;
    if (adFormatted) {
        setTimeout(() => {
            adSection.style.display = 'flex';
        }, 100);
        adSection.style.animationDelay = '0.1s';
    }
    
    const cmicFormatted = formatCmicCode(data.cmicCode);
    cmicCodeEl.textContent = cmicFormatted;
    if (cmicFormatted) {
        setTimeout(() => {
            cmicSection.style.display = 'flex';
        }, 200);
        cmicSection.style.animationDelay = '0.2s';
    }
    
    supervisorDisplayEl.textContent = data.supervisor || '';
    currentSupervisor = data.supervisor || '';
    
    const hasSupervisorId = /\(\d+\)/.test(currentSupervisor);
    copySupervisorIdBtn.style.display = hasSupervisorId ? 'flex' : 'none';
    
    if (currentSupervisor) {
        setTimeout(() => {
            supervisorSection.style.display = 'flex';
        }, 300);
        supervisorSection.style.animationDelay = '0.3s';
    }
}

// Copy handler for AD Department.
copyAdBtn.addEventListener('click', async () => {
    await navigator.clipboard.writeText(adDivisionEl.textContent);
    copyAdBtn.classList.add('copied');
    setTimeout(() => copyAdBtn.classList.remove('copied'), 1500);
});

// Copy handler for CMiC Company.
copyCmicBtn.addEventListener('click', async () => {
    await navigator.clipboard.writeText(cmicCodeEl.textContent);
    copyCmicBtn.classList.add('copied');
    setTimeout(() => copyCmicBtn.classList.remove('copied'), 1500);
});

// Copy handler for Supervisor Name (strips the ID suffix).
copySupervisorNameBtn.addEventListener('click', async () => {
    const name = currentSupervisor.replace(/\s*\(\d+\)\s*$/, '').trim();
    await navigator.clipboard.writeText(name);
    copySupervisorNameBtn.classList.add('copied');
    setTimeout(() => copySupervisorNameBtn.classList.remove('copied'), 1500);
});

// Copy handler for Supervisor ID (extracts the number in parentheses).
copySupervisorIdBtn.addEventListener('click', async () => {
    const match = currentSupervisor.match(/\((\d+)\)/);
    const id = match ? match[1] : '';
    await navigator.clipboard.writeText(id);
    copySupervisorIdBtn.classList.add('copied');
    setTimeout(() => copySupervisorIdBtn.classList.remove('copied'), 1500);
});
