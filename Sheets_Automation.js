// --- CONFIGURATION ---
const CONFIG = {
  targetSheetFragment: "Expenses Tracker",
  firstDataRow: 20,
  triggerColumn: 2, // Column B (Category/Description)
  dateColumn: 1     // Column A (Date)
};

/**
 * MANUAL Trigger: Fires if you edit the sheet manually.
 */
function onEdit(e) {
  if (!e) return;
  handleUpdate(e.range.getSheet(), e.range.getRow());
}

/**
 * APPSHEET Trigger: Fires when AppSheet adds a row.
 * (Ensure you set the trigger to "On change")
 */
function onChangeTrigger(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  // AppSheet adds data to the last row, so we check that one
  handleUpdate(sheet, sheet.getLastRow());
}

/**
 * CORE Function: Checks only for the presence of the DATE.
 */
function handleUpdate(sheet, rowIndex) {
  // 1. Sheet and row validation
  if (!sheet.getName().includes(CONFIG.targetSheetFragment)) return;
  if (rowIndex < CONFIG.firstDataRow) return;

  const triggerCell = sheet.getRange(rowIndex, CONFIG.triggerColumn);
  const dateCell = sheet.getRange(rowIndex, CONFIG.dateColumn);
  const triggerValue = triggerCell.getValue();

  // 2. Cleanup Logic: If column B is empty, clear the date (optional)
  if (triggerValue === "" || triggerValue === null) {
    // If you want the date to remain even if B is deleted, comment out the line below
    dateCell.clearContent();
    return;
  }

  // 3. Automatic Date Insertion
  const dateValue = dateCell.getValue();

  // If the date cell is empty or invalid, set today's date
  if (!(dateValue instanceof Date) || dateValue === "") {
    dateCell.setValue(new Date());
  }
}