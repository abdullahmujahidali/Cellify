import {
  csvToWorkbook,
  sheetToCsv,
  Workbook,
  workbookToXlsxBlob,
  xlsxBlobToWorkbook,
  initXlsxWasm,
  isXlsxWasmReady,
} from 'cellify';

window.Workbook = Workbook;
window.workbookToXlsxBlob = workbookToXlsxBlob;
window.xlsxBlobToWorkbook = xlsxBlobToWorkbook;
window.csvToWorkbook = csvToWorkbook;
window.sheetToCsv = sheetToCsv;

window.currentWorkbook = null;
window.currentSheetIndex = 0;

/**
 * Escape HTML special characters to prevent XSS
 */
function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function log(message, type = '') {
  const logEl = document.getElementById('log');
  const line = document.createElement('div');
  line.className = type;
  line.textContent = `[${new Date().toLocaleTimeString()}] ${message}`;
  logEl.appendChild(line);
  logEl.scrollTop = logEl.scrollHeight;
}
window.log = log;

log('Cellify loaded successfully', 'success');

// Initialize WASM for faster XLSX parsing
initXlsxWasm().then(wasmEnabled => {
  if (wasmEnabled) {
    log('WASM parser enabled - faster XLSX imports!', 'success');
  } else {
    log('WASM not available - using JavaScript parser', 'warning');
  }
});


function download(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
window.download = download;


function exportBasicExample() {
  log('Creating basic example...');

  const workbook = new Workbook();
  workbook.title = 'Basic Example';

  const sheet = workbook.addSheet('Data');


  sheet.cell('A1').value = 'Name';
  sheet.cell('B1').value = 'Age';
  sheet.cell('C1').value = 'City';


  const data = [
    ['Alice', 28, 'New York'],
    ['Bob', 34, 'San Francisco'],
    ['Charlie', 22, 'Chicago'],
    ['Diana', 31, 'Boston'],
  ];

  data.forEach((row, i) => {
    row.forEach((value, j) => {
      sheet.cell(i + 1, j).value = value;
    });
  });

  const blob = workbookToXlsxBlob(workbook);
  download(blob, 'basic-example.xlsx');
  log('Downloaded basic-example.xlsx', 'success');
}

function exportStyledExample() {
  log('Creating styled example...');

  const workbook = new Workbook();
  workbook.title = 'Styled Report';
  workbook.author = 'Cellify Demo';

  const sheet = workbook.addSheet('Sales Report');


  sheet.cell('A1').value = 'Q4 2024 Sales Report';
  sheet.cell('A1').style = {
    font: { bold: true, size: 18, color: '#FFFFFF' },
    fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#2F5496' },
    alignment: { horizontal: 'center' },
  };
  sheet.mergeCells('A1:D1');


  const headers = ['Product', 'Units', 'Price', 'Revenue'];
  headers.forEach((header, i) => {
    const cell = sheet.cell(2, i);
    cell.value = header;
    cell.style = {
      font: { bold: true, color: '#FFFFFF' },
      fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#4472C4' },
      alignment: { horizontal: 'center' },
      borders: {
        bottom: { style: 'medium', color: '#000000' },
      },
    };
  });


  const data = [
    ['Laptop Pro', 150, 1299.99, 194998.50],
    ['Wireless Mouse', 500, 49.99, 24995.00],
    ['Mechanical Keyboard', 300, 129.99, 38997.00],
    ['USB-C Hub', 450, 79.99, 35995.50],
    ['Monitor 27"', 200, 399.99, 79998.00],
  ];

  data.forEach((row, i) => {
    row.forEach((value, j) => {
      const cell = sheet.cell(i + 3, j);
      cell.value = value;
      if (j >= 2) {
        cell.style = {
          numberFormat: { formatCode: '#,##0.00' },
        };
      }
    });
  });


  sheet.cell('A8').value = 'Total';
  sheet.cell('A8').style = { font: { bold: true } };
  sheet.cell('D8').value = 374984.00;
  sheet.cell('D8').style = {
    font: { bold: true },
    numberFormat: { formatCode: '#,##0.00' },
    fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#E2EFDA' },
  };


  sheet.setColumnWidth(0, 20);
  sheet.setColumnWidth(1, 10);
  sheet.setColumnWidth(2, 12);
  sheet.setColumnWidth(3, 15);


  sheet.freeze(3, 0);

  const blob = workbookToXlsxBlob(workbook);
  download(blob, 'styled-report.xlsx');
  log('Downloaded styled-report.xlsx', 'success');
}

function exportFormulaExample() {
  log('Creating formula example...');

  const workbook = new Workbook();
  const sheet = workbook.addSheet('Calculations');


  sheet.cell('A1').value = 'Item';
  sheet.cell('B1').value = 'Quantity';
  sheet.cell('C1').value = 'Price';
  sheet.cell('D1').value = 'Total';


  for (let i = 0; i < 4; i++) {
    sheet.cell(0, i).style = {
      font: { bold: true },
      fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#D9E2F3' },
    };
  }


  const items = [
    ['Widget A', 10, 25.00],
    ['Widget B', 5, 50.00],
    ['Widget C', 20, 15.00],
  ];

  items.forEach((item, i) => {
    sheet.cell(i + 1, 0).value = item[0];
    sheet.cell(i + 1, 1).value = item[1];
    sheet.cell(i + 1, 2).value = item[2];
    sheet.cell(i + 1, 3).setFormula(`B${i + 2}*C${i + 2}`);
  });


  sheet.cell('A5').value = 'Grand Total:';
  sheet.cell('A5').style = { font: { bold: true } };
  sheet.cell('D5').setFormula('SUM(D2:D4)');
  sheet.cell('D5').style = {
    font: { bold: true },
    fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#C6EFCE' },
  };


  sheet.setColumnWidth(0, 15);
  sheet.setColumnWidth(1, 12);
  sheet.setColumnWidth(2, 12);
  sheet.setColumnWidth(3, 12);

  const blob = workbookToXlsxBlob(workbook);
  download(blob, 'formulas.xlsx');
  log('Downloaded formulas.xlsx (open in Excel to see calculated values)', 'success');
}

function exportMultiSheetExample() {
  log('Creating multi-sheet example...');

  const workbook = new Workbook();
  workbook.title = 'Multi-Sheet Workbook';


  const summary = workbook.addSheet('Summary');
  summary.cell('A1').value = 'Department Summary';
  summary.cell('A1').style = { font: { bold: true, size: 16 } };
  summary.mergeCells('A1:C1');

  summary.cell('A3').value = 'Department';
  summary.cell('B3').value = 'Employees';
  summary.cell('C3').value = 'Budget';

  const depts = [
    ['Engineering', 50, 5000000],
    ['Marketing', 25, 2000000],
    ['Sales', 40, 3000000],
    ['HR', 10, 500000],
  ];

  depts.forEach((d, i) => {
    summary.cell(i + 3, 0).value = d[0];
    summary.cell(i + 3, 1).value = d[1];
    summary.cell(i + 3, 2).value = d[2];
  });


  const eng = workbook.addSheet('Engineering');
  eng.cell('A1').value = 'Engineering Team';
  eng.cell('A1').style = { font: { bold: true } };


  const mkt = workbook.addSheet('Marketing');
  mkt.cell('A1').value = 'Marketing Team';
  mkt.cell('A1').style = { font: { bold: true } };

  const blob = workbookToXlsxBlob(workbook);
  download(blob, 'multi-sheet.xlsx');
  log('Downloaded multi-sheet.xlsx', 'success');
}

function exportCommentsExample() {
  log('Creating comments example...');

  const workbook = new Workbook();
  workbook.title = 'Comments Example';

  const sheet = workbook.addSheet('Reviews');


  const headers = ['Product', 'Rating', 'Reviewer', 'Status'];
  headers.forEach((header, i) => {
    const cell = sheet.cell(0, i);
    cell.value = header;
    cell.style = {
      font: { bold: true, color: '#FFFFFF' },
      fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#059669' },
      alignment: { horizontal: 'center' },
    };
  });


  const reviews = [
    { product: 'Laptop Pro', rating: 5, reviewer: 'Alice', status: 'Approved', comment: 'Excellent product! Fast shipping and great quality.' },
    { product: 'Wireless Mouse', rating: 4, reviewer: 'Bob', status: 'Pending', comment: 'Good mouse but battery drains quickly. Needs follow-up.' },
    { product: 'USB Hub', rating: 3, reviewer: 'Charlie', status: 'Flagged', comment: 'Product received damaged. Customer requesting refund.' },
  ];

  reviews.forEach((review, i) => {
    const row = i + 1;
    sheet.cell(row, 0).value = review.product;

    const ratingCell = sheet.cell(row, 1);
    ratingCell.value = review.rating;
    ratingCell.setComment(`Rating: ${review.rating}/5 stars`, review.reviewer);

    sheet.cell(row, 2).value = review.reviewer;

    const statusCell = sheet.cell(row, 3);
    statusCell.value = review.status;
    statusCell.setComment(review.comment, 'System');

    if (review.status === 'Approved') {
      statusCell.style = { fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#D1FAE5' } };
    } else if (review.status === 'Flagged') {
      statusCell.style = { fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#FEE2E2' } };
    } else {
      statusCell.style = { fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#FEF3C7' } };
    }
  });

  sheet.setColumnWidth(0, 18);
  sheet.setColumnWidth(1, 10);
  sheet.setColumnWidth(2, 12);
  sheet.setColumnWidth(3, 12);

  const blob = workbookToXlsxBlob(workbook);
  download(blob, 'with-comments.xlsx');
  log('Downloaded with-comments.xlsx (hover over cells to see comments)', 'success');
}

function exportHyperlinksExample() {
  log('Creating hyperlinks example...');

  const workbook = new Workbook();
  workbook.title = 'Hyperlinks Example';

  const sheet = workbook.addSheet('Resources');

  sheet.cell('A1').value = 'Useful Resources';
  sheet.cell('A1').style = {
    font: { bold: true, size: 16, color: '#FFFFFF' },
    fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#3B82F6' },
    alignment: { horizontal: 'center' },
  };
  sheet.mergeCells('A1:C1');

  const headers = ['Name', 'Description', 'Link Type'];
  headers.forEach((header, i) => {
    const cell = sheet.cell(2, i);
    cell.value = header;
    cell.style = {
      font: { bold: true, color: '#FFFFFF' },
      fill: { type: 'pattern', pattern: 'solid', foregroundColor: '#6366F1' },
      alignment: { horizontal: 'center' },
    };
  });

  const links = [
    { name: 'Cellify GitHub', desc: 'Source code repository', url: 'https://github.com/abdullahmujahidali/Cellify', type: 'External URL' },
    { name: 'Documentation', desc: 'Official documentation', url: 'https://abdullahmujahidali.github.io/Cellify/', type: 'External URL' },
    { name: 'npm Package', desc: 'Install from npm', url: 'https://www.npmjs.com/package/cellify', type: 'External URL' },
    { name: 'Email Support', desc: 'Contact the author', url: 'mailto:support@example.com', type: 'Email' },
    { name: 'Go to Summary', desc: 'Jump to summary section', url: '#Summary!A1', type: 'Internal' },
  ];

  links.forEach((link, i) => {
    const row = i + 3;
    const nameCell = sheet.cell(row, 0);
    nameCell.value = link.name;
    nameCell.setHyperlink(link.url, link.desc);
    nameCell.style = {
      font: { color: '#2563EB', underline: 'single' },
    };

    sheet.cell(row, 1).value = link.desc;
    sheet.cell(row, 2).value = link.type;
  });

  const summary = workbook.addSheet('Summary');
  summary.cell('A1').value = 'Summary Page';
  summary.cell('A1').style = { font: { bold: true, size: 14 } };
  summary.cell('A2').value = 'This page is linked from the Resources sheet';
  summary.cell('A3').value = 'Go back to Resources';
  summary.cell('A3').setHyperlink('#Resources!A1', 'Back to Resources');
  summary.cell('A3').style = { font: { color: '#2563EB', underline: 'single' } };

  sheet.setColumnWidth(0, 20);
  sheet.setColumnWidth(1, 25);
  sheet.setColumnWidth(2, 15);

  const blob = workbookToXlsxBlob(workbook);
  download(blob, 'with-hyperlinks.xlsx');
  log('Downloaded with-hyperlinks.xlsx (click links in Excel to navigate)', 'success');
}


const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');

dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('dragover');
});

dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('dragover');
});

dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});

fileInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (file) handleFile(file);
});

async function handleFile(file) {
  log(`Importing ${file.name}...`, 'info');

  try {
    let result;

    const lowerName = file.name.toLowerCase();
    if (lowerName.endsWith('.xlsx')) {
      result = await xlsxBlobToWorkbook(file);
    } else if (lowerName.endsWith('.csv')) {
      const text = await file.text();
      result = { workbook: csvToWorkbook(text), stats: { sheetCount: 1 }, warnings: [] };
    } else {
      throw new Error('Unsupported file type. Please use .xlsx or .csv');
    }

    window.currentWorkbook = result.workbook;
    window.currentSheetIndex = 0;
    undoStack.length = 0;
    activeFilters = {};
    clipboard = null;

    displayImportResult(result);
    log(`Successfully imported ${file.name}`, 'success');
  } catch (error) {
    log(`Error: ${error.message}`, 'error');
    console.error(error);
  }
}

function displayImportResult(result) {
  const { workbook, stats, warnings } = result;

  document.getElementById('importStats').style.display = 'block';


  let statsHtml = `
        <div class="stat-item">
          <div class="stat-value">${workbook.sheetCount}</div>
          <div class="stat-label">Sheets</div>
        </div>
        <div class="stat-item">
          <div class="stat-value">${(stats.totalCells || 0).toLocaleString()}</div>
          <div class="stat-label">Total Cells</div>
        </div>
      `;
  if (stats.formulaCells) {
    statsHtml += `
          <div class="stat-item">
            <div class="stat-value">${stats.formulaCells.toLocaleString()}</div>
            <div class="stat-label">Formulas</div>
          </div>
        `;
  }
  if (stats.mergedRanges) {
    statsHtml += `
          <div class="stat-item">
            <div class="stat-value">${stats.mergedRanges}</div>
            <div class="stat-label">Merged</div>
          </div>
        `;
  }
  if (stats.durationMs) {
    statsHtml += `
          <div class="stat-item">
            <div class="stat-value">${stats.durationMs}ms</div>
            <div class="stat-label">Import Time</div>
          </div>
        `;
  }
  document.getElementById('statsContent').innerHTML = statsHtml;


  const warningsEl = document.getElementById('warningsContent');
  if (warnings && warnings.length > 0) {
    warningsEl.style.display = 'block';
    warningsEl.innerHTML = `
          <p class="warnings-title">⚠️ Warnings</p>
          <ul>${warnings.map(w => `<li>${escapeHtml(w.message)}</li>`).join('')}</ul>
        `;
  } else {
    warningsEl.style.display = 'none';
  }


  const tabsEl = document.getElementById('sheetTabs');
  tabsEl.innerHTML = workbook.sheets.map((sheet, i) =>
    `<button class="sheet-tab ${i === 0 ? 'active' : ''}" data-sheet-index="${i}">${escapeHtml(sheet.name)}</button>`
  ).join('');


  tabsEl.querySelectorAll('.sheet-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      selectSheet(parseInt(tab.dataset.sheetIndex));
    });
  });


  displaySheet(workbook.sheets[0]);
}

function selectSheet(index) {
  window.currentSheetIndex = index;
  const sheet = window.currentWorkbook.sheets[index];


  undoStack.length = 0;
  activeFilters = {};


  document.querySelectorAll('.sheet-tab').forEach((tab, i) => {
    tab.classList.toggle('active', i === index);
  });

  displaySheet(sheet);
}

function displaySheet(sheet) {
  const dims = sheet.dimensions;

  if (!dims) {
    document.getElementById('preview').innerHTML = `
          <div class="empty-state">
            <div class="empty-state-icon">📭</div>
            <p>This sheet is empty</p>
          </div>
        `;
    return;
  }


  const rowsToShow = Math.min(dims.endRow - dims.startRow + 1, 50);
  const colsToShow = Math.min(dims.endCol - dims.startCol + 1, 20);

  const mergeStarts = new Map();
  const coveredCells = new Set();

  for (const merge of (sheet.merges || [])) {
    const rowSpan = merge.endRow - merge.startRow + 1;
    const colSpan = merge.endCol - merge.startCol + 1;
    const key = `${merge.startRow},${merge.startCol}`;
    mergeStarts.set(key, { rowSpan, colSpan });

    for (let r = merge.startRow; r <= merge.endRow; r++) {
      for (let c = merge.startCol; c <= merge.endCol; c++) {
        if (r !== merge.startRow || c !== merge.startCol) {
          coveredCells.add(`${r},${c}`);
        }
      }
    }
  }

  let html = '<table id="previewTable"><colgroup><col style="width:50px">';
  for (let c = dims.startCol; c < dims.startCol + colsToShow; c++) {
    html += `<col style="width:120px" data-col="${c}">`;
  }
  html += '</colgroup><thead><tr><th></th>';


  for (let c = dims.startCol; c < dims.startCol + colsToShow; c++) {
    html += `<th data-col="${c}">${columnLetter(c)}<button class="filter-btn" data-col="${c}" title="Filter column">▼</button><span class="resize-handle" data-col="${c}"></span></th>`;
  }
  html += '</tr></thead><tbody>';


  for (let r = dims.startRow; r < dims.startRow + rowsToShow; r++) {
    html += `<tr data-row="${r}"><th>${r + 1}<span class="row-resize-handle" data-row="${r}"></span></th>`;
    for (let c = dims.startCol; c < dims.startCol + colsToShow; c++) {
      const cellKey = `${r},${c}`;

      if (coveredCells.has(cellKey)) {
        continue;
      }

      const cell = sheet.getCell(r, c);
      let value = cell?.value ?? '';


      if (value instanceof Date) {
        value = value.toLocaleDateString();
      } else if (typeof value === 'number') {
        value = value.toLocaleString();
      }


      const formula = cell?.formula;
      const comment = cell?.comment;
      const hyperlink = cell?.hyperlink;
      let titleParts = [];
      if (formula) titleParts.push(`Formula: =${escapeHtml(formula.formula)}`);
      if (comment) {
        const commentText = typeof comment.text === 'string' ? comment.text : comment.text.plainText || '';
        titleParts.push(`Comment: ${escapeHtml(commentText)}${comment.author ? ` - ${escapeHtml(comment.author)}` : ''}`);
      }
      if (hyperlink) {
        titleParts.push(`Link: ${escapeHtml(hyperlink.target)}${hyperlink.tooltip ? ` (${escapeHtml(hyperlink.tooltip)})` : ''}`);
      }
      let title = titleParts.join('\n');
      let style = cellStyleToCss(cell?.style);
      let cssClass = comment ? 'has-comment' : '';
      if (hyperlink) cssClass += ' has-hyperlink';


      if (formula && !cell?.style?.font?.color) {
        style += 'color: #059669; font-style: italic;';
      }

      if (hyperlink && !cell?.style?.font?.color) {
        style += 'color: #2563eb; text-decoration: underline; cursor: pointer;';
      }


      const mergeInfo = mergeStarts.get(cellKey);
      let spanAttrs = '';
      if (mergeInfo) {
        if (mergeInfo.rowSpan > 1) spanAttrs += ` rowspan="${mergeInfo.rowSpan}"`;
        if (mergeInfo.colSpan > 1) spanAttrs += ` colspan="${mergeInfo.colSpan}"`;
      }

      html += `<td${spanAttrs} data-row="${r}" data-col="${c}" title="${title}" style="${style}" class="${cssClass}">${escapeHtml(String(value))}</td>`;
    }
    html += '</tr>';
  }

  html += '</tbody></table>';

  const totalRows = dims.endRow - dims.startRow + 1;
  const totalCols = dims.endCol - dims.startCol + 1;
  if (totalRows > rowsToShow || totalCols > colsToShow) {
    html += `<div class="preview-info">
          Showing ${rowsToShow} of ${totalRows.toLocaleString()} rows, ${colsToShow} of ${totalCols.toLocaleString()} columns
        </div>`;
  }

  html += `<div class="edit-hint">
        <kbd>Click</kbd> select · <kbd>Double-click</kbd>/<kbd>Enter</kbd> edit · <kbd>Ctrl+Click</kbd> open link · <kbd>Right-click</kbd> menu · <kbd>Ctrl+C/X/V</kbd> copy/cut/paste · <kbd>Ctrl+B/I</kbd> bold/italic · <kbd>Ctrl+Z</kbd> undo · <kbd>Del</kbd> clear
      </div>`;

  document.getElementById('preview').innerHTML = html;
  initResizeHandlers();
  initCellEditHandlers();
}

function initResizeHandlers() {
  const preview = document.getElementById('preview');
  const table = document.getElementById('previewTable');
  if (!table) return;


  if (window.resizeListeners) {
    preview.removeEventListener('mousedown', window.resizeListeners.previewMouseDown);
    document.removeEventListener('mousemove', window.resizeListeners.docMouseMove);
    document.removeEventListener('mouseup', window.resizeListeners.docMouseUp);
  }

  let isResizing = false;
  let currentCol = null;
  let currentRow = null;
  let startX = 0;
  let startY = 0;
  let startWidth = 0;
  let startHeight = 0;


  const previewMouseDown = (e) => {
    if (e.target.classList.contains('resize-handle')) {
      isResizing = true;
      currentCol = e.target.dataset.col;
      startX = e.pageX;
      const colEl = table.querySelector(`colgroup col[data-col="${currentCol}"]`);
      startWidth = colEl ? parseInt(colEl.style.width) || 120 : 120;
      e.target.classList.add('resizing');
      document.body.classList.add('col-resizing');
      e.preventDefault();
    } else if (e.target.classList.contains('row-resize-handle')) {
      isResizing = true;
      currentRow = e.target.dataset.row;
      startY = e.pageY;
      const rowEl = table.querySelector(`tr[data-row="${currentRow}"]`);
      startHeight = rowEl ? rowEl.offsetHeight : 30;
      e.target.classList.add('resizing');
      document.body.classList.add('row-resizing');
      e.preventDefault();
    }
  };

  const docMouseMove = (e) => {
    if (!isResizing) return;

    if (currentCol !== null) {
      const diff = e.pageX - startX;
      const newWidth = Math.max(60, startWidth + diff);
      const colEl = table.querySelector(`colgroup col[data-col="${currentCol}"]`);
      if (colEl) {
        colEl.style.width = newWidth + 'px';
      }
    }

    if (currentRow !== null) {
      const diff = e.pageY - startY;
      const newHeight = Math.max(25, startHeight + diff);
      const rowEl = table.querySelector(`tr[data-row="${currentRow}"]`);
      if (rowEl) {
        rowEl.style.height = newHeight + 'px';
      }
    }
  };

  const docMouseUp = () => {
    if (isResizing) {
      isResizing = false;
      document.body.classList.remove('col-resizing', 'row-resizing');
      document.querySelectorAll('.resize-handle.resizing, .row-resize-handle.resizing').forEach(el => {
        el.classList.remove('resizing');
      });
      currentCol = null;
      currentRow = null;
    }
  };

  preview.addEventListener('mousedown', previewMouseDown);
  document.addEventListener('mousemove', docMouseMove);
  document.addEventListener('mouseup', docMouseUp);


  window.resizeListeners = { previewMouseDown, docMouseMove, docMouseUp };
}

let selectedRow = null;
let selectedCol = null;
let isEditing = false;

let selectedColumn = null;  // Column index when whole column is selected
let selectedRowHeader = null;  // Row index when whole row is selected

const undoStack = [];
const MAX_UNDO_STACK = 50;


let clipboard = null;


const contextMenu = document.getElementById('contextMenu');
const commentTooltip = document.getElementById('commentTooltip');
let contextMenuRow = null;
let contextMenuCol = null;

function showContextMenu(e, row, col) {
  e.preventDefault();
  contextMenuRow = row;
  contextMenuCol = col;


  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const cell = sheet.getCell(row, col);
  const hasComment = cell?.comment;


  const addCommentItem = contextMenu.querySelector('[data-action="add-comment"]');
  const editCommentItem = contextMenu.querySelector('[data-action="edit-comment"]');
  const deleteCommentItem = contextMenu.querySelector('[data-action="delete-comment"]');

  if (hasComment) {
    addCommentItem.style.display = 'none';
    editCommentItem.style.display = 'flex';
    deleteCommentItem.style.display = 'flex';
  } else {
    addCommentItem.style.display = 'flex';
    editCommentItem.style.display = 'none';
    deleteCommentItem.style.display = 'none';
  }


  const pasteItem = contextMenu.querySelector('[data-action="paste"]');
  if (clipboard) {
    pasteItem.classList.remove('disabled');
  } else {
    pasteItem.classList.add('disabled');
  }


  const x = e.clientX;
  const y = e.clientY;

  contextMenu.style.left = x + 'px';
  contextMenu.style.top = y + 'px';
  contextMenu.classList.add('visible');


  const rect = contextMenu.getBoundingClientRect();
  if (rect.right > window.innerWidth) {
    contextMenu.style.left = (x - rect.width) + 'px';
  }
  if (rect.bottom > window.innerHeight) {
    contextMenu.style.top = (y - rect.height) + 'px';
  }


  selectCell(row, col);
}

function hideContextMenu() {
  contextMenu.classList.remove('visible');
  contextMenuRow = null;
  contextMenuCol = null;
}


document.addEventListener('click', (e) => {
  if (!contextMenu.contains(e.target)) {
    hideContextMenu();
  }
});


document.addEventListener('scroll', hideContextMenu, true);


contextMenu.addEventListener('click', (e) => {
  const item = e.target.closest('.context-menu-item');
  if (!item || item.classList.contains('disabled')) return;

  const action = item.dataset.action;
  if (!action) return;

  switch (action) {
    case 'copy':
      copyCell();
      break;
    case 'cut':
      cutCell();
      break;
    case 'paste':
      pasteCell();
      break;
    case 'bold':
      toggleBold();
      break;
    case 'italic':
      toggleItalic();
      break;
    case 'add-comment':
    case 'edit-comment':
      showCommentDialog();
      break;
    case 'delete-comment':
      deleteComment();
      break;
    case 'clear':
      clearCell();
      break;
  }

  hideContextMenu();
});


document.getElementById('fillColorPicker').addEventListener('click', (e) => {
  const swatch = e.target.closest('.color-swatch');
  if (!swatch) return;

  const color = swatch.dataset.color;
  setFillColor(color);
  hideContextMenu();
});

document.getElementById('textColorPicker').addEventListener('click', (e) => {
  const swatch = e.target.closest('.color-swatch');
  if (!swatch) return;

  const color = swatch.dataset.color;
  setTextColor(color);
  hideContextMenu();
});


function copyCell() {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const cell = sheet.getCell(selectedRow, selectedCol);

  clipboard = {
    row: selectedRow,
    col: selectedCol,
    value: cell?.value,
    formula: cell?.formula?.formula,
    style: cell?.style ? JSON.parse(JSON.stringify(cell.style)) : null,
    comment: cell?.comment ? JSON.parse(JSON.stringify(cell.comment)) : null,
    isCut: false
  };

  log(`Copied cell ${columnLetter(selectedCol)}${selectedRow + 1}`, 'info');
}


function cutCell() {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const cell = sheet.getCell(selectedRow, selectedCol);

  clipboard = {
    row: selectedRow,
    col: selectedCol,
    value: cell?.value,
    formula: cell?.formula?.formula,
    style: cell?.style ? JSON.parse(JSON.stringify(cell.style)) : null,
    comment: cell?.comment ? JSON.parse(JSON.stringify(cell.comment)) : null,
    isCut: true
  };


  const td = document.querySelector(`td[data-row="${selectedRow}"][data-col="${selectedCol}"]`);
  if (td) {
    td.style.opacity = '0.5';
    td.style.border = '2px dashed var(--primary)';
  }

  log(`Cut cell ${columnLetter(selectedCol)}${selectedRow + 1}`, 'info');
}


function pasteCell() {
  if (!clipboard || selectedRow === null || selectedCol === null) return;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const cell = sheet.cell(selectedRow, selectedCol);
  const oldValue = cell.value;


  undoStack.push({ row: selectedRow, col: selectedCol, oldValue, newValue: clipboard.value });
  if (undoStack.length > MAX_UNDO_STACK) undoStack.shift();


  if (clipboard.formula) {
    cell.setFormula(clipboard.formula);
  } else if (clipboard.value !== undefined && clipboard.value !== null) {
    cell.value = clipboard.value;
  }


  if (clipboard.style) {
    cell.style = JSON.parse(JSON.stringify(clipboard.style));
  }

  if (clipboard.comment) {
    const text = typeof clipboard.comment.text === 'string'
      ? clipboard.comment.text
      : clipboard.comment.text?.plainText || '';
    cell.setComment(text, clipboard.comment.author);
  }

  if (clipboard.isCut) {
    const sourceSheet = window.currentWorkbook.sheets[window.currentSheetIndex];
    const sourceCell = sourceSheet.cell(clipboard.row, clipboard.col);
    sourceCell.clear();

    const sourceTd = document.querySelector(`td[data-row="${clipboard.row}"][data-col="${clipboard.col}"]`);
    if (sourceTd) {
      sourceTd.textContent = '';
      sourceTd.style.opacity = '';
      sourceTd.style.border = '';
      sourceTd.style.backgroundColor = '';
      sourceTd.classList.remove('has-comment');
    }

    clipboard = null;
  }

  updateCellDisplay(selectedRow, selectedCol);
  log(`Pasted to cell ${columnLetter(selectedCol)}${selectedRow + 1}`, 'info');
}

function setFillColor(color) {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const cell = sheet.cell(selectedRow, selectedCol);

  if (!cell.style) cell.style = {};

  if (color) {
    cell.style.fill = {
      type: 'pattern',
      pattern: 'solid',
      foregroundColor: color
    };
  } else {
    delete cell.style.fill;
  }

  updateCellDisplay(selectedRow, selectedCol);
  log(`Set fill color to ${color || 'none'} for cell ${columnLetter(selectedCol)}${selectedRow + 1}`, 'info');
}

function setTextColor(color) {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const cell = sheet.cell(selectedRow, selectedCol);

  if (!cell.style) cell.style = {};
  if (!cell.style.font) cell.style.font = {};

  cell.style.font.color = color;

  updateCellDisplay(selectedRow, selectedCol);
  log(`Set text color to ${color} for cell ${columnLetter(selectedCol)}${selectedRow + 1}`, 'info');
}

function toggleBold() {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const cell = sheet.cell(selectedRow, selectedCol);

  if (!cell.style) cell.style = {};
  if (!cell.style.font) cell.style.font = {};

  cell.style.font.bold = !cell.style.font.bold;

  updateCellDisplay(selectedRow, selectedCol);
  log(`${cell.style.font.bold ? 'Added' : 'Removed'} bold for cell ${columnLetter(selectedCol)}${selectedRow + 1}`, 'info');
}

function toggleItalic() {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const cell = sheet.cell(selectedRow, selectedCol);

  if (!cell.style) cell.style = {};
  if (!cell.style.font) cell.style.font = {};

  cell.style.font.italic = !cell.style.font.italic;

  updateCellDisplay(selectedRow, selectedCol);
  log(`${cell.style.font.italic ? 'Added' : 'Removed'} italic for cell ${columnLetter(selectedCol)}${selectedRow + 1}`, 'info');
}

function showCommentDialog() {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const cell = sheet.getCell(selectedRow, selectedCol);
  const existingComment = cell?.comment;

  let currentText = '';
  let currentAuthor = '';
  if (existingComment) {
    currentText = typeof existingComment.text === 'string'
      ? existingComment.text
      : existingComment.text?.plainText || '';
    currentAuthor = existingComment.author || '';
  }

  const text = prompt('Enter comment:', currentText);
  if (text === null) return;

  if (text.trim() === '') {
    deleteComment();
    return;
  }

  const author = prompt('Author (optional):', currentAuthor) || undefined;

  const targetCell = sheet.cell(selectedRow, selectedCol);
  targetCell.setComment(text, author);

  updateCellDisplay(selectedRow, selectedCol);
  log(`${existingComment ? 'Updated' : 'Added'} comment for cell ${columnLetter(selectedCol)}${selectedRow + 1}`, 'info');
}


function deleteComment() {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const cell = sheet.cell(selectedRow, selectedCol);

  if (cell.comment) {
    cell.comment = undefined;
    updateCellDisplay(selectedRow, selectedCol);
    log(`Deleted comment from cell ${columnLetter(selectedCol)}${selectedRow + 1}`, 'info');
  }
}


function clearCell() {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const cell = sheet.cell(selectedRow, selectedCol);
  const oldValue = cell.value;


  undoStack.push({ row: selectedRow, col: selectedCol, oldValue, newValue: null });
  if (undoStack.length > MAX_UNDO_STACK) undoStack.shift();

  cell.clear();
  updateCellDisplay(selectedRow, selectedCol);
  log(`Cleared cell ${columnLetter(selectedCol)}${selectedRow + 1}`, 'info');
}


function updateCellDisplay(row, col) {
  const td = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
  if (!td) return;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const cell = sheet.getCell(row, col);


  let displayValue = '';
  if (cell?.value instanceof Date) {
    displayValue = cell.value.toLocaleDateString();
  } else if (cell?.value !== undefined && cell?.value !== null) {
    displayValue = typeof cell.value === 'number' ? cell.value.toLocaleString() : String(cell.value);
  }
  td.textContent = displayValue;


  td.style.cssText = cellStyleToCss(cell?.style);


  if (cell?.formula && !cell?.style?.font?.color) {
    td.style.color = '#059669';
    td.style.fontStyle = 'italic';
  }


  if (cell?.comment) {
    td.classList.add('has-comment');
    const commentText = typeof cell.comment.text === 'string'
      ? cell.comment.text
      : cell.comment.text?.plainText || '';
    let titleParts = [];
    if (cell.formula) titleParts.push(`Formula: =${escapeHtml(cell.formula.formula)}`);
    titleParts.push(`Comment: ${escapeHtml(commentText)}${cell.comment.author ? ` - ${escapeHtml(cell.comment.author)}` : ''}`);
    td.title = titleParts.join('\n');
  } else {
    td.classList.remove('has-comment');
    if (cell?.formula) {
      td.title = `Formula: =${escapeHtml(cell.formula.formula)}`;
    } else {
      td.title = '';
    }
  }

  td.classList.add('modified');
}


function showCommentTooltip(e, cell) {
  if (!cell?.comment) return;

  const commentText = typeof cell.comment.text === 'string'
    ? cell.comment.text
    : cell.comment.text?.plainText || '';

  const authorEl = commentTooltip.querySelector('.comment-tooltip-author');
  const textEl = commentTooltip.querySelector('.comment-tooltip-text');

  authorEl.textContent = cell.comment.author || 'Comment';
  textEl.textContent = commentText;

  commentTooltip.style.left = (e.clientX + 10) + 'px';
  commentTooltip.style.top = (e.clientY + 10) + 'px';
  commentTooltip.classList.add('visible');
}

function hideCommentTooltip() {
  commentTooltip.classList.remove('visible');
}


const filterDropdown = document.getElementById('filterDropdown');
const filterSearchInput = document.getElementById('filterSearchInput');
const filterOptions = document.getElementById('filterOptions');
let activeFilters = {};
let currentFilterCol = null;
let currentFilterValues = new Map();
let pendingFilterSelections = new Set();

function showFilterDropdown(e, colIndex) {
  e.stopPropagation();
  currentFilterCol = colIndex;


  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const dims = sheet.dimensions;
  if (!dims) return;

  currentFilterValues.clear();
  const rowsToShow = Math.min(dims.endRow - dims.startRow + 1, 50);

  for (let r = dims.startRow; r < dims.startRow + rowsToShow; r++) {
    const cell = sheet.getCell(r, colIndex);
    let value = cell?.value;

    if (value instanceof Date) {
      value = value.toLocaleDateString();
    } else if (value === null || value === undefined) {
      value = '(Empty)';
    } else {
      value = String(value);
    }

    currentFilterValues.set(value, (currentFilterValues.get(value) || 0) + 1);
  }


  if (activeFilters[colIndex]) {
    pendingFilterSelections = new Set(activeFilters[colIndex]);
  } else {
    pendingFilterSelections = new Set(currentFilterValues.keys());
  }


  renderFilterOptions('');


  const btn = e.target;
  const rect = btn.getBoundingClientRect();
  filterDropdown.style.left = Math.min(rect.left, window.innerWidth - 250) + 'px';
  filterDropdown.style.top = (rect.bottom + 5) + 'px';
  filterDropdown.classList.add('visible');


  filterSearchInput.value = '';
  filterSearchInput.focus();


  filterDropdown.querySelector('.filter-dropdown-title').textContent =
    `Filter: ${columnLetter(colIndex)}`;
}

function renderFilterOptions(searchTerm) {
  const search = searchTerm.toLowerCase();
  let html = '';


  const sortedValues = Array.from(currentFilterValues.entries())
    .filter(([value]) => value.toLowerCase().includes(search))
    .sort((a, b) => a[0].localeCompare(b[0]));


  const allSelected = sortedValues.every(([value]) => pendingFilterSelections.has(value));
  html += `
        <div class="filter-option" data-value="__select_all__">
          <input type="checkbox" ${allSelected ? 'checked' : ''}>
          <span class="filter-option-label">(Select All)</span>
        </div>
      `;

  for (const [value, count] of sortedValues) {
    const checked = pendingFilterSelections.has(value) ? 'checked' : '';
    html += `
          <div class="filter-option" data-value="${escapeHtml(value)}">
            <input type="checkbox" ${checked}>
            <span class="filter-option-label">${escapeHtml(value)}</span>
            <span class="filter-option-count">(${count})</span>
          </div>
        `;
  }

  filterOptions.innerHTML = html;
}

function hideFilterDropdown() {
  filterDropdown.classList.remove('visible');
  currentFilterCol = null;
}

function applyFilter() {
  if (currentFilterCol === null) return;


  if (pendingFilterSelections.size === currentFilterValues.size) {
    delete activeFilters[currentFilterCol];
  } else {
    activeFilters[currentFilterCol] = new Set(pendingFilterSelections);
  }

  applyAllFilters();
  updateFilterButtonStates();
  hideFilterDropdown();

  const count = Object.keys(activeFilters).length;
  if (count > 0) {
    log(`Applied filter on column ${columnLetter(currentFilterCol)} (${count} active filter${count > 1 ? 's' : ''})`, 'info');
  } else {
    log('All filters cleared', 'info');
  }
}

function clearCurrentFilter() {
  if (currentFilterCol === null) return;
  pendingFilterSelections = new Set(currentFilterValues.keys());
  renderFilterOptions(filterSearchInput.value);
}

function applyAllFilters() {
  const table = document.getElementById('previewTable');
  if (!table) return;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const tbody = table.querySelector('tbody');
  const rows = tbody.querySelectorAll('tr');

  rows.forEach(tr => {
    const rowIndex = parseInt(tr.dataset.row);
    let visible = true;


    for (const [colIndex, allowedValues] of Object.entries(activeFilters)) {
      const cell = sheet.getCell(rowIndex, parseInt(colIndex));
      let value = cell?.value;

      if (value instanceof Date) {
        value = value.toLocaleDateString();
      } else if (value === null || value === undefined) {
        value = '(Empty)';
      } else {
        value = String(value);
      }

      if (!allowedValues.has(value)) {
        visible = false;
        break;
      }
    }

    if (visible) {
      tr.classList.remove('filtered-out');
    } else {
      tr.classList.add('filtered-out');
    }
  });
}

function updateFilterButtonStates() {
  const table = document.getElementById('previewTable');
  if (!table) return;

  const filterBtns = table.querySelectorAll('th .filter-btn');
  filterBtns.forEach(btn => {
    const col = parseInt(btn.dataset.col);
    if (activeFilters[col]) {
      btn.classList.add('active');
      btn.closest('th').classList.add('has-filter');
    } else {
      btn.classList.remove('active');
      btn.closest('th').classList.remove('has-filter');
    }
  });
}

function clearAllFilters() {
  activeFilters = {};
  applyAllFilters();
  updateFilterButtonStates();
  log('All filters cleared', 'info');
}


filterSearchInput.addEventListener('input', (e) => {
  renderFilterOptions(e.target.value);
});

filterOptions.addEventListener('click', (e) => {
  const option = e.target.closest('.filter-option');
  if (!option) return;

  const value = option.dataset.value;
  const checkbox = option.querySelector('input[type="checkbox"]');

  if (value === '__select_all__') {

    const search = filterSearchInput.value.toLowerCase();
    const visibleValues = Array.from(currentFilterValues.keys())
      .filter(v => v.toLowerCase().includes(search));

    if (checkbox.checked) {
      visibleValues.forEach(v => pendingFilterSelections.delete(v));
    } else {
      visibleValues.forEach(v => pendingFilterSelections.add(v));
    }
  } else {
    if (pendingFilterSelections.has(value)) {
      pendingFilterSelections.delete(value);
    } else {
      pendingFilterSelections.add(value);
    }
  }

  renderFilterOptions(filterSearchInput.value);
});

document.getElementById('filterApplyBtn').addEventListener('click', applyFilter);
document.getElementById('filterCancelBtn').addEventListener('click', hideFilterDropdown);
document.getElementById('filterClearBtn').addEventListener('click', clearCurrentFilter);


document.addEventListener('click', (e) => {
  if (!filterDropdown.contains(e.target) && !e.target.closest('.filter-btn')) {
    hideFilterDropdown();
  }
});



function initCellEditHandlers() {
  const preview = document.getElementById('preview');
  const table = document.getElementById('previewTable');
  if (!table) return;


  if (window.cellEditListeners) {
    preview.removeEventListener('click', window.cellEditListeners.onClick);
    preview.removeEventListener('dblclick', window.cellEditListeners.onDblClick);
    preview.removeEventListener('contextmenu', window.cellEditListeners.onContextMenu);
    preview.removeEventListener('mousemove', window.cellEditListeners.onMouseMove);
    preview.removeEventListener('mouseleave', window.cellEditListeners.onMouseLeave);
    document.removeEventListener('keydown', window.cellEditListeners.onKeyDown);
  }


  selectedRow = null;
  selectedCol = null;
  isEditing = false;

  const onClick = (e) => {
    // Check if clicking on filter button (don't select column)
    if (e.target.closest('.filter-btn') || e.target.closest('.resize-handle') || e.target.closest('.row-resize-handle')) {
      return;
    }

    // Check for column header click (thead th with data-col)
    const colHeader = e.target.closest('thead th[data-col]');
    if (colHeader && window.currentWorkbook) {
      const col = parseInt(colHeader.dataset.col);
      selectColumn(col);
      return;
    }

    // Check for row header click (tbody th inside tr with data-row)
    const rowHeader = e.target.closest('tbody tr[data-row] th');
    if (rowHeader && window.currentWorkbook) {
      const tr = rowHeader.closest('tr[data-row]');
      if (tr) {
        const row = parseInt(tr.dataset.row);
        selectRow(row);
        return;
      }
    }

    // Regular cell click
    const td = e.target.closest('td[data-row][data-col]');
    if (td && !isEditing) {
      const row = parseInt(td.dataset.row);
      const col = parseInt(td.dataset.col);

      // Ctrl+Click or Cmd+Click on hyperlink opens the link
      if ((e.ctrlKey || e.metaKey) && td.classList.contains('has-hyperlink')) {
        const sheet = window.currentWorkbook?.sheets[window.currentSheetIndex];
        const cell = sheet?.getCell(row, col);
        if (cell?.hyperlink?.target) {
          window.open(cell.hyperlink.target, '_blank');
          return;
        }
      }

      selectCell(row, col);
    }
  };

  const onDblClick = (e) => {
    const td = e.target.closest('td[data-row][data-col]');
    if (td) {
      const row = parseInt(td.dataset.row);
      const col = parseInt(td.dataset.col);
      selectCell(row, col);
      startEditing(row, col);
    }
  };

  const onContextMenu = (e) => {
    const td = e.target.closest('td[data-row][data-col]');
    if (td && window.currentWorkbook) {
      const row = parseInt(td.dataset.row);
      const col = parseInt(td.dataset.col);
      showContextMenu(e, row, col);
    }
  };

  const onMouseMove = (e) => {
    const td = e.target.closest('td[data-row][data-col]');
    if (td && td.classList.contains('has-comment')) {
      const row = parseInt(td.dataset.row);
      const col = parseInt(td.dataset.col);
      const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
      const cell = sheet.getCell(row, col);
      showCommentTooltip(e, cell);
    } else {
      hideCommentTooltip();
    }
  };

  const onMouseLeave = () => {
    hideCommentTooltip();
  };

  const onKeyDown = (e) => {
    handleCellKeyDown(e);
  };

  const onFilterClick = (e) => {
    const filterBtn = e.target.closest('.filter-btn');
    if (filterBtn && window.currentWorkbook) {
      const col = parseInt(filterBtn.dataset.col);
      showFilterDropdown(e, col);
    }
  };

  preview.addEventListener('click', onClick);
  preview.addEventListener('dblclick', onDblClick);
  preview.addEventListener('contextmenu', onContextMenu);
  preview.addEventListener('mousemove', onMouseMove);
  preview.addEventListener('mouseleave', onMouseLeave);
  document.addEventListener('keydown', onKeyDown);


  const filterBtns = preview.querySelectorAll('.filter-btn');
  filterBtns.forEach(btn => btn.addEventListener('click', onFilterClick));

  window.cellEditListeners = { onClick, onDblClick, onContextMenu, onMouseMove, onMouseLeave, onKeyDown, onFilterClick };
}

function clearAllSelections() {
  document.querySelectorAll('td.selected, td.col-selected, td.row-selected').forEach(el => {
    el.classList.remove('selected', 'col-selected', 'row-selected');
  });
  document.querySelectorAll('th.col-selected, th.row-selected').forEach(el => {
    el.classList.remove('col-selected', 'row-selected');
  });
  selectedColumn = null;
  selectedRowHeader = null;
}

function selectCell(row, col) {
  // Clear all selections
  clearAllSelections();

  // If clicking same cell while editing, don't change selection
  if (isEditing && row === selectedRow && col === selectedCol) {
    return;
  }

  // If editing another cell, save first
  if (isEditing) {
    const input = document.querySelector('td.editing input');
    if (input) {
      saveEdit(selectedRow, selectedCol, input.value);
    }
  }

  selectedRow = row;
  selectedCol = col;

  // Find and highlight the cell
  const td = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
  if (td) {
    td.classList.add('selected');
  }
}

function selectColumn(colIndex) {
  // Clear all selections first
  clearAllSelections();
  selectedColumn = colIndex;
  selectedRow = null;
  selectedCol = null;

  // Highlight column header
  const th = document.querySelector(`thead th[data-col="${colIndex}"]`);
  if (th) {
    th.classList.add('col-selected');
  }

  // Highlight all cells in the column
  document.querySelectorAll(`td[data-col="${colIndex}"]`).forEach(td => {
    td.classList.add('col-selected');
  });

  log(`Selected column ${columnLetter(colIndex)}`, 'info');
}

function selectRow(rowIndex) {
  // Clear all selections first
  clearAllSelections();
  selectedRowHeader = rowIndex;
  selectedRow = null;
  selectedCol = null;

  // Highlight row header
  const th = document.querySelector(`tbody tr[data-row="${rowIndex}"] th`);
  if (th) {
    th.classList.add('row-selected');
  }

  // Highlight all cells in the row
  document.querySelectorAll(`td[data-row="${rowIndex}"]`).forEach(td => {
    td.classList.add('row-selected');
  });

  log(`Selected row ${rowIndex + 1}`, 'info');
}

function startEditing(row, col) {
  if (isEditing) return;

  const td = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
  if (!td) return;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const cell = sheet.getCell(row, col);


  let editValue = '';
  if (cell?.formula) {
    editValue = '=' + cell.formula.formula;
  } else if (cell?.value !== null && cell?.value !== undefined) {
    if (cell.value instanceof Date) {
      editValue = cell.value.toISOString().split('T')[0];
    } else {
      editValue = String(cell.value);
    }
  }

  isEditing = true;
  td.classList.add('editing');


  td.dataset.originalContent = td.innerHTML;


  const input = document.createElement('input');
  input.type = 'text';
  input.autocomplete = 'off';
  input.setAttribute('data-form-type', 'other');
  input.value = editValue;
  td.innerHTML = '';
  td.appendChild(input);


  const minWidth = td.offsetWidth;
  const textWidth = getTextWidth(editValue, getComputedStyle(input).font);
  input.style.width = Math.max(minWidth, textWidth + 30) + 'px';

  input.focus();
  input.select();

  input.scrollLeft = 0;


  input.addEventListener('blur', () => {
    if (isEditing) {
      saveEdit(row, col, input.value);
    }
  });

  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      saveEdit(row, col, input.value);

      moveSelection(1, 0);
    } else if (e.key === 'Escape') {
      e.preventDefault();
      cancelEdit(row, col);
    } else if (e.key === 'Tab') {
      e.preventDefault();
      saveEdit(row, col, input.value);
      moveSelection(0, e.shiftKey ? -1 : 1);
    }
  });
}

function saveEdit(row, col, newValue) {
  if (!isEditing) return;

  const td = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
  if (!td) return;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const cell = sheet.cell(row, col);
  const oldValue = cell.value;


  const parsed = parseInputValue(newValue);


  if (parsed.isFormula) {
    cell.setFormula(parsed.value);
    log(`Cell ${columnLetter(col)}${row + 1}: Set formula =${parsed.value}`, 'info');
  } else {
    cell.value = parsed.value;
    const displayOld = oldValue instanceof Date ? oldValue.toLocaleDateString() : oldValue;
    const displayNew = parsed.value instanceof Date ? parsed.value.toLocaleDateString() : parsed.value;
    if (displayOld !== displayNew) {
      log(`Cell ${columnLetter(col)}${row + 1}: "${displayOld}" → "${displayNew}"`, 'info');
    }
  }


  undoStack.push({ row, col, oldValue, newValue: parsed.value, wasFormula: parsed.isFormula });
  if (undoStack.length > MAX_UNDO_STACK) {
    undoStack.shift();
  }


  isEditing = false;
  td.classList.remove('editing');

  let displayValue = '';
  if (cell.formula) {
    displayValue = cell.value !== null && cell.value !== undefined ? cell.value : '';
    td.style.color = '#059669';
    td.style.fontStyle = 'italic';
    td.style.fontWeight = '500';
    td.title = `Formula: =${escapeHtml(cell.formula.formula)}`;
  } else {
    if (cell.value instanceof Date) {
      displayValue = cell.value.toLocaleDateString();
    } else if (typeof cell.value === 'number') {
      displayValue = cell.value.toLocaleString();
    } else {
      displayValue = cell.value ?? '';
    }
    td.style.color = '';
    td.style.fontStyle = '';
    td.style.fontWeight = '';
    td.title = '';
  }

  td.textContent = String(displayValue);
  td.classList.add('modified');
  delete td.dataset.originalContent;
}

function cancelEdit(row, col) {
  if (!isEditing) return;

  const td = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
  if (!td) return;

  isEditing = false;
  td.classList.remove('editing');


  if (td.dataset.originalContent !== undefined) {
    td.innerHTML = td.dataset.originalContent;
    delete td.dataset.originalContent;
  }

  log(`Edit cancelled for cell ${columnLetter(col)}${row + 1}`, 'info');
}

function undoLastEdit() {
  if (undoStack.length === 0) {
    log('Nothing to undo', 'info');
    return;
  }

  const action = undoStack.pop();
  const { row, col, oldValue } = action;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const cell = sheet.cell(row, col);

  if (oldValue !== undefined && oldValue !== null) {

    if (cell.formula) {
      cell.clearFormula();
    }
    cell.value = oldValue;
  } else {
    cell.clear();
  }


  const td = document.querySelector(`td[data-row="${row}"][data-col="${col}"]`);
  if (td) {
    let displayValue = '';
    if (oldValue instanceof Date) {
      displayValue = oldValue.toLocaleDateString();
    } else if (oldValue !== undefined && oldValue !== null) {
      displayValue = String(oldValue);
    }
    td.textContent = displayValue;
  }


  selectCell(row, col);

  const displayOld = oldValue instanceof Date ? oldValue.toLocaleDateString() : oldValue;
  log(`Undo: Cell ${columnLetter(col)}${row + 1} restored to "${displayOld}"`, 'info');
}

function parseInputValue(str) {
  const trimmed = str.trim();


  if (trimmed.startsWith('=')) {
    return { isFormula: true, value: trimmed.slice(1) };
  }


  if (trimmed === '') {
    return { isFormula: false, value: '' };
  }


  if (trimmed.toLowerCase() === 'true') {
    return { isFormula: false, value: true };
  }
  if (trimmed.toLowerCase() === 'false') {
    return { isFormula: false, value: false };
  }


  const num = Number(trimmed);
  if (!isNaN(num) && trimmed !== '') {
    return { isFormula: false, value: num };
  }


  const dateMatch = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (dateMatch) {
    const date = new Date(trimmed);
    if (!isNaN(date.getTime())) {
      return { isFormula: false, value: date };
    }
  }


  return { isFormula: false, value: trimmed };
}

function handleCellKeyDown(e) {

  if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') {
    return;
  }

  const preview = document.getElementById('preview');
  if (!preview || !preview.contains(e.target)) {
    return;
  }


  if (selectedRow === null || selectedCol === null) return;

  if (isEditing) {

    return;
  }


  if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
    e.preventDefault();
    undoLastEdit();
    return;
  }


  if ((e.ctrlKey || e.metaKey) && e.key === 'c') {
    e.preventDefault();
    copyCell();
    return;
  }


  if ((e.ctrlKey || e.metaKey) && e.key === 'x') {
    e.preventDefault();
    cutCell();
    return;
  }


  if ((e.ctrlKey || e.metaKey) && e.key === 'v') {
    e.preventDefault();
    pasteCell();
    return;
  }


  if ((e.ctrlKey || e.metaKey) && e.key === 'b') {
    e.preventDefault();
    toggleBold();
    return;
  }


  if ((e.ctrlKey || e.metaKey) && e.key === 'i') {
    e.preventDefault();
    toggleItalic();
    return;
  }


  if (e.key === 'Delete' || e.key === 'Backspace') {
    e.preventDefault();
    clearCell();
    return;
  }


  if (e.key === 'Enter' || e.key === 'F2') {
    e.preventDefault();
    startEditing(selectedRow, selectedCol);
    return;
  }


  if (e.key === 'ArrowUp') {
    e.preventDefault();
    moveSelection(-1, 0);
  } else if (e.key === 'ArrowDown') {
    e.preventDefault();
    moveSelection(1, 0);
  } else if (e.key === 'ArrowLeft') {
    e.preventDefault();
    moveSelection(0, -1);
  } else if (e.key === 'ArrowRight') {
    e.preventDefault();
    moveSelection(0, 1);
  } else if (e.key === 'Tab') {
    e.preventDefault();
    moveSelection(0, e.shiftKey ? -1 : 1);
  }


  if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
    e.preventDefault();
    startEditing(selectedRow, selectedCol);

    const input = document.querySelector('td.editing input');
    if (input) {
      input.value = e.key;

      input.setSelectionRange(input.value.length, input.value.length);
    }
  }
}

function moveSelection(rowDelta, colDelta) {
  if (selectedRow === null || selectedCol === null) return;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  const dims = sheet.dimensions;
  if (!dims) return;


  const table = document.getElementById('previewTable');
  if (!table) return;

  const newRow = selectedRow + rowDelta;
  const newCol = selectedCol + colDelta;


  const newTd = document.querySelector(`td[data-row="${newRow}"][data-col="${newCol}"]`);
  if (newTd) {
    selectCell(newRow, newCol);
  }
}

function columnLetter(index) {
  let letter = '';
  while (index >= 0) {
    letter = String.fromCharCode(65 + (index % 26)) + letter;
    index = Math.floor(index / 26) - 1;
  }
  return letter;
}

const textWidthCanvas = document.createElement('canvas');
const textWidthCtx = textWidthCanvas.getContext('2d');
function getTextWidth(text, font) {
  textWidthCtx.font = font;
  return textWidthCtx.measureText(text).width;
}


function cellStyleToCss(style) {
  if (!style) return '';

  let css = '';


  if (style.font) {
    const font = style.font;
    if (font.bold) css += 'font-weight: bold;';
    if (font.italic) css += 'font-style: italic;';
    const textDecorations = [];
    if (font.underline) textDecorations.push('underline');
    if (font.strikethrough) textDecorations.push('line-through');
    if (textDecorations.length > 0) css += `text-decoration: ${textDecorations.join(' ')};`;
    if (font.color) css += `color: ${font.color};`;
    if (font.size) css += `font-size: ${font.size}pt;`;
    if (font.name) css += `font-family: "${font.name}", sans-serif;`;
  }

  if (style.fill) {
    const fill = style.fill;
    if (fill.type === 'pattern' && fill.pattern === 'solid' && fill.foregroundColor) {
      css += `background-color: ${fill.foregroundColor};`;
    }
  }


  if (style.alignment) {
    const align = style.alignment;
    if (align.horizontal) {
      css += `text-align: ${align.horizontal};`;
    }
    if (align.vertical) {
      const vAlign = align.vertical === 'middle' ? 'middle' :
        align.vertical === 'top' ? 'top' : 'bottom';
      css += `vertical-align: ${vAlign};`;
    }
    if (align.wrapText) {
      css += 'white-space: pre-wrap;';
    }
    if (align.textRotation) {
      css += `writing-mode: vertical-rl; transform: rotate(${align.textRotation}deg);`;
    }
  }


  if (style.borders) {
    const borders = style.borders;
    if (borders.top) {
      css += `border-top: ${borderStyleToCss(borders.top)};`;
    }
    if (borders.right) {
      css += `border-right: ${borderStyleToCss(borders.right)};`;
    }
    if (borders.bottom) {
      css += `border-bottom: ${borderStyleToCss(borders.bottom)};`;
    }
    if (borders.left) {
      css += `border-left: ${borderStyleToCss(borders.left)};`;
    }
  }

  return css;
}


function borderStyleToCss(border) {
  if (!border) return 'none';

  const styleMap = {
    'thin': '1px solid',
    'medium': '2px solid',
    'thick': '3px solid',
    'dashed': '1px dashed',
    'dotted': '1px dotted',
    'double': '3px double',
    'hair': '1px solid',
    'dashDot': '1px dashed',
    'dashDotDot': '1px dashed',
    'mediumDashed': '2px dashed',
    'mediumDashDot': '2px dashed',
    'mediumDashDotDot': '2px dashed',
    'slantDashDot': '2px dashed'
  };

  const cssStyle = styleMap[border.style] || '1px solid';
  const color = border.color || '#000000';

  return `${cssStyle} ${color}`;
}

function reExportXlsx() {
  if (!window.currentWorkbook) return;

  log('Re-exporting as XLSX...');
  const blob = workbookToXlsxBlob(window.currentWorkbook);
  download(blob, 're-exported.xlsx');
  log('Downloaded re-exported.xlsx', 'success');
}

function reExportCsv() {
  if (!window.currentWorkbook) return;

  const sheet = window.currentWorkbook.sheets[window.currentSheetIndex];
  log(`Exporting sheet "${sheet.name}" as CSV...`);

  const csv = sheetToCsv(sheet);
  const blob = new Blob([csv], { type: 'text/csv' });
  download(blob, `${sheet.name}.csv`);
  log(`Downloaded ${sheet.name}.csv`, 'success');
}

document.getElementById('btnBasic').addEventListener('click', exportBasicExample);
document.getElementById('btnStyled').addEventListener('click', exportStyledExample);
document.getElementById('btnFormula').addEventListener('click', exportFormulaExample);
document.getElementById('btnMultiSheet').addEventListener('click', exportMultiSheetExample);
document.getElementById('btnComments').addEventListener('click', exportCommentsExample);
document.getElementById('btnHyperlinks').addEventListener('click', exportHyperlinksExample);
document.getElementById('btnReExportXlsx').addEventListener('click', reExportXlsx);
document.getElementById('btnReExportCsv').addEventListener('click', reExportCsv);
