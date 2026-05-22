import { WORD_LIST_URL } from './constants.js';
import { usersCells, saveUsersTable } from './usersTable.js';

let _wordListPromise = null;

// Fetches and filters the BIP-39 word list to words between 4-8 characters.
export function loadWordList() {
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

// Generates a random passphrase from three capitalized words + a digit and updates the Users Table.
export async function generateUsersPassword() {
    const words = await loadWordList();
    let password = '';
    while (password.length < 15) {
        const arr = new Uint32Array(3);
        crypto.getRandomValues(arr);
        const chosen = [
            words[arr[0] % words.length],
            words[arr[1] % words.length],
            words[arr[2] % words.length]
        ];
        password = chosen.map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join('') + Math.floor(Math.random() * 10);
    }

    usersCells.password.textContent = password;
    saveUsersTable();

    const btn = document.getElementById('generatePasswordBtn');
    btn.classList.add('copied');
    setTimeout(() => btn.classList.remove('copied'), 1500);
}

// Pre-fetch the word list on page load.
loadWordList();

// Generate Password button handler.
document.getElementById('generatePasswordBtn').addEventListener('click', () => generateUsersPassword());
