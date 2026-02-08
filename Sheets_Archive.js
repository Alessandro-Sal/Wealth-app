/**
 * Converts formulas to values for the current month's column at the end of the month.
 * Includes automatic backup and cleanup of old files.
 */
function congelaSoloFineMese() {
  const SS = SpreadsheetApp.getActiveSpreadsheet();
  const SHEET = SS.getSheetByName("NW Analitico"); // Ensure this matches your actual sheet name
  if (!SHEET) return;

  const MAX_BACKUP = 5; // Number of backups to keep
  const OGGI = new Date();
  const DOMANI = new Date(OGGI);
  DOMANI.setDate(OGGI.getDate() + 1);

  // --- 1. END OF MONTH CHECK ---
  // If tomorrow is not the 1st of the month, today is not the last day.
  if (DOMANI.getDate() !== 1) {
    console.log("Not end of month. Script terminated.");
    return;
  }

  // --- 2. DRIVE BACKUP LOGIC ---
  const folderName = "Backup NW AnaliticoFineMese";
  const fileName = "Backup_NW_" + Utilities.formatDate(OGGI, SS.getSpreadsheetTimeZone(), "yyyy-MM-dd_HHmm");
  
  let folders = DriveApp.getFoldersByName(folderName);
  let folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

  // Create backup file
  let backupFile = SpreadsheetApp.create(fileName);
  SHEET.copyTo(backupFile).setName("Dati Congelati");
  let defaultSheet = backupFile.getSheetByName("Foglio1") || backupFile.getSheetByName("Sheet1");
  if (defaultSheet) backupFile.deleteSheet(defaultSheet);

  // Move to folder
  DriveApp.getFileById(backupFile.getId()).moveTo(folder);
  console.log("Backup created: " + fileName);

  // --- 3. CLEAN UP OLD BACKUPS ---
  eliminaBackupVecchi(folder, MAX_BACKUP);

  // --- 4. DATA FREEZE LOGIC ---
  const SETTINGS = {
    rowDate: 1,
    rowStart: 2,
    rowEnd: 200,
    colStart: 3,
    colEnd: 14
  };

  SpreadsheetApp.flush(); 
  
  const headers = SHEET.getRange(SETTINGS.rowDate, SETTINGS.colStart, 1, SETTINGS.colEnd - SETTINGS.colStart + 1).getValues()[0];
  const dOggi = OGGI.getDate();
  const mOggi = OGGI.getMonth();
  const aOggi = OGGI.getFullYear();

  headers.forEach((dateVal, i) => {
    let d = (dateVal instanceof Date) ? dateVal : new Date(dateVal);
    
    if (!isNaN(d.getTime())) {
      // Check if the column date is TODAY
      if (d.getDate() === dOggi && d.getMonth() === mOggi && d.getFullYear() === aOggi) {
        let col = i + SETTINGS.colStart;
        let range = SHEET.getRange(SETTINGS.rowStart, col, SETTINGS.rowEnd - SETTINGS.rowStart + 1, 1);
        range.copyTo(range, {contentsOnly: true});
        console.log("✅ Data frozen for column: " + col);
      }
    }
  });
}

/**
 * Helper function to delete the oldest files in a folder
 */
function eliminaBackupVecchi(folder, maxFiles) {
  let files = [];
  let iter = folder.getFiles();
  
  while (iter.hasNext()) {
    let file = iter.next();
    files.push({
      id: file.getId(),
      date: file.getDateCreated()
    });
  }

  // Sort files from newest to oldest
  files.sort((a, b) => b.date - a.date);

  // If exceeding the limit, delete the oldest ones
  if (files.length > maxFiles) {
    for (let i = maxFiles; i < files.length; i++) {
      DriveApp.getFileById(files[i].id).setTrashed(true);
      console.log("Deleted old backup ID: " + files[i].id);
    }
  }
}