/**
 * --- AUTOMATIC NIGHTLY BACKUP ---
 * Creates a complete copy of the file (App + Data) in a "Backups" folder.
 * Keeps only the last 7 backups to save space.
 */
function createNightlyBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Gets formatted date for the filename
  const timeZone = ss.getSpreadsheetTimeZone();
  const dateStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
  const backupName = ss.getName() + "_BACKUP_" + dateStr;
  
  // 1. Search for or create the folder "Backup MyAlFinancialStatements"
  const folderName = "Backup MyAlFinancialStatements";
  const folders = DriveApp.getFoldersByName(folderName);
  let folder;
  
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }
  
  // 2. Creates an exact copy of the current file (Data + Script + HTML)
  const file = DriveApp.getFileById(ss.getId());
  file.makeCopy(backupName, folder);
  
  console.log("✅ Backup created successfully: " + backupName);
  
  // 3. Cleanup old backups (keeps only the last 7)
  const files = folder.getFiles();
  let fileList = [];
  
  while (files.hasNext()) {
    let f = files.next();
    // Saves ID, date, and file object for sorting
    fileList.push({ id: f.getId(), date: f.getDateCreated(), file: f });
  }
  
  // Sorts from newest to oldest
  fileList.sort((a, b) => b.date - a.date);
  
  // If there are more than 7, delete the oldest ones (from index 7 onwards)
  if (fileList.length > 7) {
    for (let i = 7; i < fileList.length; i++) {
      fileList[i].file.setTrashed(true); // Moves to trash
      console.log("🗑️ Deleted old backup: " + fileList[i].file.getName());
    }
  }
}