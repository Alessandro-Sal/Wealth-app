/**
 * Converts a 1-based column index into an Excel-style letter.
 * Example: 1 -> A, 27 -> AA
 * @param {number} column - The column index.
 * @return {string} The corresponding column letter.
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

