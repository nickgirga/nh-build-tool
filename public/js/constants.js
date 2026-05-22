// Global reference to the SheetJS library loaded from CDN in index.html
export const XLS = window.XLSX || window.xlsx;

// localStorage keys for persisting app data
export const STORAGE_KEY = 'nhbuildtool_workbook';
export const AGENT_NAME_KEY = 'nhbuildtool_agentName';
export const USERS_TABLE_KEY = 'nhbuildtool_usersTable';

// Regex to detect mirror requests in ticket text
export const MIRROR_PATTERN = /\bmirror\b/i;

// BIP-39 English word list URL for passphrase generation
export const WORD_LIST_URL = 'https://cdn.jsdelivr.net/gh/bitcoin/bips@master/bip-0039/english.txt';

// Mapping of abbreviated month names to numbers (1-12)
export const MONTHS = { Jan: 1, Feb: 2, Mar: 3, Apr: 4, May: 5, Jun: 6, Jul: 7, Aug: 8, Sep: 9, Oct: 10, Nov: 11, Dec: 12 };

// Expected workbook version numbers (from 'Metadata' sheet cell B2)
export const EXPECTED_MAJOR_VERSION = 2;
export const EXPECTED_MINOR_VERSION = 2;
