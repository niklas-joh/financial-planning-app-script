// In-memory spreadsheet fake with the few Range/Sheet methods the writers use.
// Like the real thing, setValues only overwrites its target range and the whole-tab
// clearContents throws, so tests prove the code never empties a tab.
function fakeSheetBook() {
  const sheets = {};
  const makeSheet = () => {
    const s = { grid: [], formats: [], frozen: 0, cleared: [] };
    s.getLastRow = () => s.grid.length;
    s.getLastColumn = () => s.grid.reduce((w, r) => Math.max(w, r.length), 0);
    s.clearContents = () => { throw new Error('whole-tab clear is not allowed'); };
    s.setFrozenRows = (n) => { s.frozen = n; };
    s.getRange = (row, col, rows, cols) => ({
      getValues: () => s.grid.slice(row - 1, row - 1 + rows).map((r) => r.slice(col - 1, col - 1 + cols)),
      setValues: (m) => {
        m.forEach((r, i) => {
          const target = s.grid[row - 1 + i] || (s.grid[row - 1 + i] = []);
          r.forEach((v, j) => { target[col - 1 + j] = v; });
        });
      },
      setNumberFormat: (f) => { s.formats.push([col, f]); },
      clearContent: () => {
        s.cleared.push([row, col, rows, cols]);
        for (let r = row - 1; r < row - 1 + rows; r++) for (let c = col - 1; c < col - 1 + cols; c++) if (s.grid[r]) s.grid[r][c] = '';
        while (s.grid.length && s.grid[s.grid.length - 1].every((v) => v === '')) s.grid.pop();
      }
    });
    return s;
  };
  return {
    sheets,
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => (sheets[n] = makeSheet())
  };
}

module.exports = { fakeSheetBook };
