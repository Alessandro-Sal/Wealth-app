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