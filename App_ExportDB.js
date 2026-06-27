/**
 * Esporta tutti i fogli della cartella di lavoro in un unico file JSON su Google Drive.
 * Funzione da lanciare manualmente dall'editor di Apps Script per il backup/esportazione SPA.
 */
function exportDBToDrive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const dbExport = {};

  Logger.log("Inizio esportazione dei fogli...");

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const dataRange = sheet.getDataRange();
    
    const rawValues = dataRange.getValues();
    const displayValues = dataRange.getDisplayValues();
    
    dbExport[sheetName] = {
      rawValues: rawValues,
      displayValues: displayValues
    };
    
    Logger.log("Elaborato foglio: " + sheetName);
  });

  const jsonString = JSON.stringify(dbExport);
  
  const fileName = "SPA_Database_Export_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd_HH-mm-ss") + ".json";
  
  const file = DriveApp.createFile(fileName, jsonString, MimeType.PLAIN_TEXT);
  
  Logger.log("Esportazione completata con successo!");
  Logger.log("URL del file (salvato nella cartella principale del tuo Drive): " + file.getUrl());
}
