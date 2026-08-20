// Small shared helpers used by every CRUD page.

// Generates the next sequential ID for a prefix, e.g. nextId(rows, 'SUP') -> 'SUP004'
// based on the highest existing number found in column A of the given rows
// (rows includes the header row as element 0).
function nextId(rows, prefix, padLength = 3) {
  let max = 0;
  for (let i = 1; i < rows.length; i++) {
    const cell = (rows[i][0] || '').toString();
    if (cell.startsWith(prefix)) {
      const n = parseInt(cell.slice(prefix.length), 10);
      if (!isNaN(n) && n > max) max = n;
    }
  }
  return prefix + String(max + 1).padStart(padLength, '0');
}

// Converts a 1-based column number to a spreadsheet column letter (1 -> A, 27 -> AA).
function colLetter(n) {
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function escapeHtml(s) {
  return (s ?? '').toString()
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Display label for a stored UNIT value — the sheet keeps the short code
// ('kg', 'pc') so existing rows and dropdown values don't need to change.
function unitLabel(u) {
  return u === 'pc' ? 'Pieces' : (u || '');
}
