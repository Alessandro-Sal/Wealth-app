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

function saveNoteToCloud(ticker, noteText) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("App_Notes");
    
    // Se il foglio App_Notes non esiste, lo crea da solo
    if (!sheet) {
        sheet = ss.insertSheet("App_Notes");
        sheet.appendRow(["Ticker", "Note", "Last Update"]);
        sheet.getRange("A1:C1").setFontWeight("bold");
    }
    
    const data = sheet.getDataRange().getValues();
    let found = false;
    
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
    
    if (!found && noteText !== "") {
        sheet.appendRow([ticker, noteText, new Date()]);
    }

    return true;
}

// FIX (1.15): rimossa la SECONDA definizione duplicata e identica di
// getNotesFromCloud() che si trovava qui sotto. Due funzioni omonime nello
// stesso progetto GAS rendono il deploy non deterministico. Resta solo la
// definizione in alto (righe ~23-38).