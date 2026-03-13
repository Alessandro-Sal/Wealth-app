/**
 * Silently fetches notes from the cloud and restores the local storage
 * Execute this when the app initializes
 */
function syncNotesOnStartup() {
    if (typeof google !== 'undefined' && google.script) {
        google.script.run
            .withSuccessHandler(function(cloudNotesDict) {
                // Overwrite local storage with the secure cloud database
                localStorage.setItem('ticker_notes', JSON.stringify(cloudNotesDict));
            })
            .withFailureHandler(function(e) {
                console.warn("Failed to sync notes on startup:", e);
            })
            .getNotesFromCloud();
    }
}

// Call it directly so it runs when the JS executes
syncNotesOnStartup();
/**
 * Opens the notes sheet for a specific ticker
 */
function openTickerNotes(ticker, event) {
    if (event) {
        event.preventDefault();
        event.stopPropagation();
    }
    
    const sheet = document.getElementById('ticker-notes-sheet');
    const backdrop = document.getElementById('common-backdrop');
    const title = document.getElementById('notes-sheet-title');
    const input = document.getElementById('ticker-notes-input');
    const hiddenTicker = document.getElementById('current-notes-ticker');
    
    if (!sheet || !title || !input || !hiddenTicker) return;
    
    // Set the target ticker
    hiddenTicker.value = ticker;
    title.innerText = `${ticker} NOTES`;
    
    // Load existing note from LocalStorage
    let notesData = {};
    try {
        notesData = JSON.parse(localStorage.getItem('ticker_notes') || '{}');
    } catch(e) { console.error("Error reading notes"); }
    
    input.value = notesData[ticker] || '';
    
    // Show the modal
    if (backdrop) backdrop.classList.add('active');
    sheet.classList.add('active');
}

/**
 * Closes the notes sheet cleanly
 */
function closeTickerNotesSheet() {
    const sheet = document.getElementById('ticker-notes-sheet');
    const backdrop = document.getElementById('common-backdrop');
    if (sheet) sheet.classList.remove('active');
    
    const activeSheets = document.querySelectorAll('.sheet-modal.active');
    if (backdrop && activeSheets.length === 0) {
        backdrop.classList.remove('active');
    }
}

/**
 * Saves the note locally for speed and syncs it to Google Sheets for safety
 */
function saveTickerNote(event) {
    const ticker = document.getElementById('current-notes-ticker').value;
    const noteText = document.getElementById('ticker-notes-input').value.trim();
    
    if (!ticker) return;
    
    // 1. Local Storage save (instant speed)
    let notesData = {};
    try {
        notesData = JSON.parse(localStorage.getItem('ticker_notes') || '{}');
    } catch(e) {}
    
    if (noteText === '') {
        delete notesData[ticker]; 
    } else {
        notesData[ticker] = noteText;
    }
    localStorage.setItem('ticker_notes', JSON.stringify(notesData));
    
    // 2. UI Button feedback (Syncing state)
    const btn = event.currentTarget;
    const oldText = btn.innerText;
    const oldBg = btn.style.background;
    
    btn.innerText = '☁️ SYNCING...';
    btn.style.pointerEvents = 'none';
    
    // 3. Google Sheets Backup (Bulletproof storage)
    if (typeof google !== 'undefined' && google.script) {
        google.script.run
            .withSuccessHandler(function() {
                // Cloud save successful
                btn.innerText = '✅ SECURELY SAVED';
                btn.style.background = 'var(--success, #34c759)';
                setTimeout(() => {
                    btn.innerText = oldText;
                    btn.style.background = oldBg;
                    btn.style.pointerEvents = 'auto';
                    closeTickerNotesSheet();
                }, 800);
            })
            .withFailureHandler(function(err) {
                // Network error handling
                btn.innerText = '⚠️ SAVED LOCALLY (SYNC FAILED)';
                btn.style.background = 'var(--warning, #ff9500)';
                setTimeout(() => {
                    btn.innerText = oldText;
                    btn.style.background = oldBg;
                    btn.style.pointerEvents = 'auto';
                }, 2000);
            })
            .saveNoteToCloud(ticker, noteText);
    } else {
        // Fallback if running outside Apps Script environment
        setTimeout(() => closeTickerNotesSheet(), 500);
    }
}

/**
 * Utility: Check if a ticker has a saved note (useful for UI indicators)
 */
function hasTickerNote(ticker) {
    try {
        const notesData = JSON.parse(localStorage.getItem('ticker_notes') || '{}');
        return !!notesData[ticker];
    } catch(e) {
        return false;
    }
}

/**
 * Saves or updates a note for a specific ticker in the Google Sheet.
 * If the note is empty, it deletes the row.
 */
function saveNoteToCloud(ticker, noteText) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("App_Notes");
    
    // Create the sheet automatically if it doesn't exist yet
    if (!sheet) {
        sheet = ss.insertSheet("App_Notes");
        sheet.appendRow(["Ticker", "Note", "Last Update"]);
        sheet.getRange("A1:C1").setFontWeight("bold");
    }
    
    const data = sheet.getDataRange().getValues();
    let found = false;
    
    // Scan existing rows to update or delete
    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === ticker) {
            if (noteText === "") {
                sheet.deleteRow(i + 1); 
            } else {
                sheet.getRange(i + 1, 2).setValue(noteText);
                sheet.getRange(i + 1, 3).setValue(new Date());
            }
            found = true;
            break;
        }
    }
    
    // If not found and the note is not empty, add a new row at the bottom
    if (!found && noteText !== "") {
        sheet.appendRow([ticker, noteText, new Date()]);
    }
    
    return true;
}

/**
 * Fetches all notes from the cloud to restore them on app load.
 * Returns an object like: { "AAPL": "Buy under 150", "TSLA": "Target 300" }
 */
function getNotesFromCloud() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("App_Notes");
    if (!sheet) return {};
    
    const data = sheet.getDataRange().getValues();
    let notesDict = {};
    
    for (let i = 1; i < data.length; i++) {
        const ticker = data[i][0];
        const note = data[i][1];
        if (ticker) notesDict[ticker] = note;
    }
    
    return notesDict;
}