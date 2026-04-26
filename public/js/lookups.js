import { workbookData } from './workbook.js';

// Simple vertical lookup: finds the first row where keyColIndex matches lookupValue (case-insensitive) and returns returnColIndex.
export function vLookup(lookupValue, sheetData, keyColIndex, returnColIndex) {
    for (const row of sheetData) {
        if (row[keyColIndex] && row[keyColIndex].toLowerCase() === lookupValue.toLowerCase()) {
            return row[returnColIndex] || '';
        }
    }
    return '';
}

// Multi-condition lookup: returns the value at returnColIndex when all conditions match.
export function xLookup(conditions, sheetData, returnColIndex) {
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

// Returns the area name mapped to a location from the Location Map sheet.
export function findAreaForLocation(location) {
    return vLookup(location, workbookData.locationMap, 0, 1);
}

// Returns the CMiC group mapped to a location from the CMiC Location Map sheet.
export function findCmicGroupForLocation(location) {
    return vLookup(location, workbookData.cmicLocationMap, 1, 0);
}

// Attempts to match a parsed job title against the workbook's canonical job titles using word overlap.
export function findBestJobTitleMatch(parsedJobTitle) {
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

// Checks if the ticket text contains a specific unique email domain and returns it if found.
export function checkForUniqueEmail(parsedDomain, textLower) {
    const uniqueDomain = new RegExp('s.........c\\.com');

    if (textLower && uniqueDomain.test(textLower)) {
        for (const domain of workbookData.emailDomains) {
            if (domain && uniqueDomain.test(domain.toLowerCase())) {
                return domain;
            }
        }
    }

    return parsedDomain;
}

// Attempts to match a parsed location against the workbook's canonical locations with special-case overrides.
export function findBestLocationMatch(parsedLocation, empTypeForLookup, textLower) {
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

// Updates all lookup result fields in the UI based on the currently selected dropdown values.
export function updateLookups() {
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

    const adobeBanner = document.getElementById('adobeBanner');
    if (selectedJobTitle === 'Marketing' && workbookData.adobeCcAdGroup) {
        document.getElementById('adobeCcAdGroup').textContent = workbookData.adobeCcAdGroup;
        adobeBanner.classList.add('show');
    } else {
        adobeBanner.classList.remove('show');
    }
}

// Finds an exact case-insensitive match for a value in the first column of a sheet.
export function findValueInSheet(sheetData, searchValue) {
    const lowerSearch = searchValue.toLowerCase();
    for (const row of sheetData) {
        if (row[0] && row[0].toLowerCase() === lowerSearch) {
            return row[0];
        }
    }
    return searchValue;
}

// Selects a dropdown option by case-insensitive value comparison.
export function selectDropdownValue(dropdownId, value) {
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
