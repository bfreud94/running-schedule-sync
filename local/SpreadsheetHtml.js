function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function formatCellValue(value) {
  if (value instanceof Date) {
    return value.toLocaleDateString('en-US', { month: 'long', day: 'numeric' });
  }
  return value;
}

function renderSheet(name, sheet, changes) {
  const changedCells = new Set(
    changes
      .filter(change => change.sheet === name)
      .map(change => `${change.row},${change.column}`)
  );
  const columnCount = Math.max(9, ...sheet.values.map(row => row.length));
  const columns = Array.from({ length: columnCount }, (_, index) => String.fromCharCode(65 + index));
  const rows = sheet.values.map((row, rowIndex) => {
    const cells = columns.map((_, columnIndex) => {
      const key = `${rowIndex + 1},${columnIndex + 1}`;
      const style = sheet.styles?.[key] || {};
      const inlineStyle = [
        style.background ? `background:${style.background}` : '',
        style.fontColor ? `color:${style.fontColor}` : '',
        style.wrap ? 'white-space:pre-wrap' : '',
        style.verticalAlignment ? `vertical-align:${style.verticalAlignment}` : '',
        style.horizontalAlignment ? `text-align:${style.horizontalAlignment}` : ''
      ].filter(Boolean).join(';');
      const changedClass = changedCells.has(key) ? ' class="changed"' : '';
      return `<td${changedClass} style="${inlineStyle}">${escapeHtml(formatCellValue(row[columnIndex]))}</td>`;
    }).join('');

    return `<tr><th class="row-number">${rowIndex + 1}</th>${cells}</tr>`;
  }).join('');

  return `
    <section>
      <h2>${escapeHtml(name)}</h2>
      <div class="table-scroll">
        <table>
          <colgroup><col class="row-number-column">${columns.map((_, index) => `<col style="width:${sheet.columnWidths?.[index + 1] || 100}px">`).join('')}</colgroup>
          <thead><tr><th class="corner"></th>${columns.map(column => `<th>${column}</th>`).join('')}</tr></thead>
          <tbody>${rows}</tbody>
        </table>
      </div>
    </section>`;
}

function renderSpreadsheetHtml(workbook, changes) {
  const sheets = Object.entries(workbook.sheets)
    .map(([name, sheet]) => renderSheet(name, sheet, changes))
    .join('');

  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Local Spreadsheet Preview</title>
  <style>
    :root { color-scheme: light; --ink: #202124; --grid: #dadce0; --header: #f1f3f4; --accent: #1a73e8; }
    * { box-sizing: border-box; }
    body { margin: 0; color: var(--ink); background: #f8f9fa; font-family: Aptos, Calibri, sans-serif; }
    header { padding: 24px 32px 16px; border-bottom: 1px solid var(--grid); background: white; }
    h1 { margin: 0; font-family: Georgia, serif; font-size: 24px; letter-spacing: 0; }
    header p { margin: 6px 0 0; color: #5f6368; }
    main { padding: 24px 32px 40px; display: grid; gap: 28px; }
    section { min-width: 0; }
    h2 { margin: 0 0 10px; font-size: 15px; letter-spacing: 0; }
    .table-scroll { overflow-x: auto; border: 1px solid var(--grid); background: white; }
    table { border-collapse: collapse; min-width: 940px; width: 100%; table-layout: fixed; }
    th, td { border-right: 1px solid var(--grid); border-bottom: 1px solid var(--grid); height: 42px; padding: 8px 10px; text-align: left; font-size: 13px; white-space: pre-wrap; overflow-wrap: anywhere; }
    thead th, .row-number { background: var(--header); color: #5f6368; text-align: center; font-weight: 500; }
    .corner, .row-number { width: 42px; }
    tr:last-child td, tr:last-child th { border-bottom: 0; }
    th:last-child, td:last-child { border-right: 0; }
    @media (max-width: 700px) { header, main { padding-left: 16px; padding-right: 16px; } }
  </style>
</head>
<body>
  <header>
    <h1>Local Spreadsheet Preview</h1>
    <p>${changes.length} changes from the latest sync.</p>
  </header>
  <main>${sheets}</main>
</body>
</html>`;
}

module.exports = { renderSpreadsheetHtml };