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

// ============================================================================
// --- TICKER NOTES MODULE (FRONTEND) ---
// ============================================================================

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
    
    hiddenTicker.value = ticker;
    title.innerText = `${ticker} NOTES`;
    
    let notesData = {};
    try {
        notesData = JSON.parse(localStorage.getItem('ticker_notes') || '{}');
    } catch(e) { console.error("Error reading notes"); }
    
    input.value = notesData[ticker] || '';
    
    if (backdrop) backdrop.classList.add('active');
    sheet.classList.add('active');
}

function closeTickerNotesSheet() {
    const sheet = document.getElementById('ticker-notes-sheet');
    const backdrop = document.getElementById('common-backdrop');
    if (sheet) sheet.classList.remove('active');
    
    const activeSheets = document.querySelectorAll('.sheet-modal.active');
    if (backdrop && activeSheets.length === 0) {
        backdrop.classList.remove('active');
    }
}

function saveTickerNote(event) {
    const ticker = document.getElementById('current-notes-ticker').value;
    const noteText = document.getElementById('ticker-notes-input').value.trim();
    
    if (!ticker) return;
    
    let notesData = {};
    try { notesData = JSON.parse(localStorage.getItem('ticker_notes') || '{}'); } catch(e) {}
    
    if (noteText === '') {
        delete notesData[ticker]; 
    } else {
        notesData[ticker] = noteText;
    }
    localStorage.setItem('ticker_notes', JSON.stringify(notesData));
    
    const btn = event.currentTarget;
    const oldText = btn.innerText;
    const oldBg = btn.style.background;
    
    btn.innerText = '☁️ SYNCING...';
    btn.style.pointerEvents = 'none';
    
    if (typeof google !== 'undefined' && google.script) {
        google.script.run
            .withSuccessHandler(function() {
                btn.innerText = '✅ SECURELY SAVED';
                btn.style.background = 'var(--success, #34c759)';
                setTimeout(() => {
                    btn.innerText = oldText;
                    btn.style.background = oldBg;
                    btn.style.pointerEvents = 'auto';
                    closeTickerNotesSheet();
                    
                    // Ricarica le card per far apparire l'icona 📝
                    if (typeof fetchWatchlist === 'function') fetchWatchlist(true);
                    if (typeof fetchLivePortfolio === 'function') fetchLivePortfolio(true);
                }, 800);
            })
            .withFailureHandler(function(err) {
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
        setTimeout(() => closeTickerNotesSheet(), 500);
    }
}

function hasTickerNote(ticker) {
    try {
        const notesData = JSON.parse(localStorage.getItem('ticker_notes') || '{}');
        return !!notesData[ticker];
    } catch(e) {
        return false;
    }
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