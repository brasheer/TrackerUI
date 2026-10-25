// --- STATE ---
let workingFileHandle = null;
let parentDirHandle = null;
let workbook = null;
let allRows = []; // Flattened index across all MM-DD-YYYY tabs
let dateSheets = []; // Ordered list of MM-DD-YYYY sheet names
let activeKeyFilter = null;
let editingTarget = null; // { sheetName, rowNumber } or null for new

const DATE_TAB_REGEX = /^\d{2}-\d{2}-\d{4}$/;

// --- DOM ELEMENTS ---
const statusPill = document.getElementById('statusPill');
const sheetSelect = document.getElementById('sheetSelect');
const scopeSelect = document.getElementById('scopeSelect');
const searchInput = document.getElementById('searchInput');
const listBody = document.getElementById('listBody');
const filterBanner = document.getElementById('filterBanner');
const filterKeyText = document.getElementById('filterKeyText');
const editModal = document.getElementById('editModal');

// --- 1. LINK FILE / DIRECTORY ---
document.getElementById('btnLinkFile').addEventListener('click', async () => {
  try {
    // Pick the agent's working .xlsx file
    [workingFileHandle] = await window.showOpenFilePicker({
      types: [{ description: 'Excel Workbook', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }],
      multiple: false
    });
    await loadWorkbookFromDisk();
  } catch (err) {
    if (err.name !== 'AbortError') alert('Could not open file: ' + err.message);
  }
});

// --- 2. READ WORKBOOK & INDEX ALL DAILY TABS ---
async function loadWorkbookFromDisk(preserveSelectedSheet = null) {
  if (!workingFileHandle) return;
  setStatus('Reading...', '');
  const file = await workingFileHandle.getFile();
  const arrayBuffer = await file.arrayBuffer();

  workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);

  allRows = [];
  dateSheets = [];

  workbook.eachSheet((worksheet) => {
    const name = worksheet.name.trim();
    if (DATE_TAB_REGEX.test(name)) {
      dateSheets.push(name);
      indexWorksheetRows(worksheet, name);
    }
  });

  // Sort date tabs newest first
  dateSheets.sort((a, b) => {
    const [m1, d1, y1] = a.split('-').map(Number);
    const [m2, d2, y2] = b.split('-').map(Number);
    return new Date(y2, m2 - 1, d2) - new Date(y1, m1 - 1, d1);
  });

  populateSheetDropdown(preserveSelectedSheet);
  sheetSelect.disabled = false;
  scopeSelect.disabled = false;
  searchInput.disabled = false;
  document.getElementById('btnNewEntry').style.display = 'inline-block';
  document.getElementById('btnRollover').style.display = 'inline-block';

  setStatus(`Synced: ${file.name}`, 'ready');
  renderList();
}

function cellText(row, colIndex) {
  const val = row.getCell(colIndex).value;
  if (val === null || val === undefined) return '';
  if (typeof val === 'object') {
    if (val.result !== undefined) return String(val.result);
    if (val.text !== undefined) return String(val.text);
    if (val.richText) return val.richText.map(t => t.text).join('');
  }
  return String(val).trim();
}

function indexWorksheetRows(worksheet, sheetName) {
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return; // Skip header row
    const subject = cellText(row, 1);
    const ticket = cellText(row, 2);
    const broker = cellText(row, 3);
    const status = cellText(row, 10);

    // Skip blank template rows
    if (!subject && !ticket && !broker && !status) return;

    let key = cellText(row, 11);
    if (!key) {
      // Fallback hash if Column K was empty: YYYYMMDD-Broker-Row
      const [mm, dd, yyyy] = sheetName.split('-');
      const cleanBroker = (broker || 'UNK').replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
      key = ticket || `${yyyy}${mm}${dd}-${cleanBroker}-${rowNumber}`;
    }

    allRows.push({
      sheetName,
      rowNumber,
      subject,
      ticket,
      broker,
      master: cellText(row, 4),
      bidder: cellText(row, 5),
      action: cellText(row, 6),
      yard: cellText(row, 7),
      lots: cellText(row, 8),
      notes: cellText(row, 9),
      status: status || 'Pending',
      key
    });
  });
}

function populateSheetDropdown(preserveSheet) {
  sheetSelect.innerHTML = '';
  dateSheets.forEach(name => {
    const opt = document.createElement('option');
    opt.value = name;
    opt.textContent = name;
    sheetSelect.appendChild(opt);
  });
  if (preserveSheet && dateSheets.includes(preserveSheet)) {
    sheetSelect.value = preserveSheet;
  } else if (dateSheets.length > 0) {
    sheetSelect.value = dateSheets[0];
  }
}

// --- 3. RENDER LIST VIEW WITH COLLAPSIBLE NOTES & KEY HISTORY ---
function renderList() {
  const query = searchInput.value.trim().toLowerCase();
  const scope = scopeSelect.value;
  const currentSheet = sheetSelect.value;

  let filtered = allRows.filter(item => {
    // If a unique key filter is active, show ALL occurrences across all dates
    if (activeKeyFilter) {
      return item.key.toLowerCase() === activeKeyFilter.toLowerCase();
    }
    // Otherwise respect scope dropdown
    if (scope === 'current' && item.sheetName !== currentSheet) {
      return false;
    }
    if (!query) return true;

    const haystack = [
      item.subject, item.ticket, item.broker, item.master,
      item.bidder, item.action, item.yard, item.lots, item.notes, item.key
    ].join(' ').toLowerCase();
    return haystack.includes(query);
  });

  listBody.innerHTML = '';
  if (filtered.length === 0) {
    listBody.innerHTML = `<tr><td colspan="11" style="text-align:center; color:var(--text-muted); padding: 24px;">No matching requests found.</td></tr>`;
    return;
  }

  filtered.forEach((item, idx) => {
    const tr = document.createElement('tr');
    const isComplete = item.status.toLowerCase() === 'complete';
    const noteText = item.notes || '—';
    const isLongNote = noteText.length > 45 || noteText.includes('\n');

    tr.innerHTML = `
      <td style="white-space:nowrap; font-weight:500;">${escapeHtml(item.sheetName)}</td>
      <td>
        <span class="key-chip" title="Click to pull all history for this Key across all dates" data-key="${escapeHtml(item.key)}">
          🔍 ${escapeHtml(item.key)}
        </span>
      </td>
      <td>${escapeHtml(item.subject)}</td>
      <td>${escapeHtml(item.ticket)}</td>
      <td>${escapeHtml(item.broker)}</td>
      <td style="font-family:monospace; font-size:12px;">
        M: ${escapeHtml(item.master || '—')}<br>B: ${escapeHtml(item.bidder || '—')}
      </td>
      <td>${escapeHtml(item.action)}</td>
      <td style="font-size:12px;">
        Y: ${escapeHtml(item.yard || '—')}<br>L: ${escapeHtml(item.lots || '—')}
      </td>
      <td class="notes-cell">
        <div class="notes-preview" id="note-${idx}">${escapeHtml(noteText)}</div>
        ${isLongNote ? `<span class="notes-toggle" data-target="note-${idx}">[+ Expand]</span>` : ''}
      </td>
      <td>
        <span class="badge ${isComplete ? 'badge-complete' : 'badge-pending'}">
          ${escapeHtml(item.status)}
        </span>
      </td>
      <td>
        <button class="btn-edit" data-sheet="${escapeHtml(item.sheetName)}" data-row="${item.rowNumber}">Edit</button>
      </td>
    `;
    listBody.appendChild(tr);
  });
}

// Delegate clicks for Key Chip filter, Collapsible Notes, and Edit button
listBody.addEventListener('click', (e) => {
  const keyChip = e.target.closest('.key-chip');
  if (keyChip) {
    activeKeyFilter = keyChip.dataset.key;
    filterKeyText.textContent = activeKeyFilter;
    filterBanner.style.display = 'flex';
    renderList();
    return;
  }

  const noteToggle = e.target.closest('.notes-toggle');
  if (noteToggle) {
    const targetDiv = document.getElementById(noteToggle.dataset.target);
    const expanded = targetDiv.classList.toggle('expanded');
    noteToggle.textContent = expanded ? '[- Collapse]' : '[+ Expand]';
    return;
  }

  const editBtn = e.target.closest('.btn-edit');
  if (editBtn) {
    openEditModal(editBtn.dataset.sheet, Number(editBtn.dataset.row));
  }
});

document.getElementById('btnClearKeyFilter').addEventListener('click', () => {
  activeKeyFilter = null;
  filterBanner.style.display = 'none';
  renderList();
});

sheetSelect.addEventListener('change', renderList);
scopeSelect.addEventListener('change', renderList);
searchInput.addEventListener('input', () => {
  // Automatically search across all tabs when typing if user wants global lookup
  renderList();
});

// --- 4. EDIT & CREATE MODAL ---
document.getElementById('btnNewEntry').addEventListener('click', () => {
  if (!sheetSelect.value) return alert('Create or select a daily tab first.');
  editingTarget = null;
  document.getElementById('modalTitle').textContent = `New Request (${sheetSelect.value})`;
  document.getElementById('modalKeyDisplay').textContent = 'Auto-Generated on Save';
  document.getElementById('entryForm').reset();
  document.getElementById('fStatus').value = 'Pending';
  editModal.classList.add('open');
});

function openEditModal(sheetName, rowNumber) {
  const item = allRows.find(r => r.sheetName === sheetName && r.rowNumber === rowNumber);
  if (!item) return;

  editingTarget = { sheetName, rowNumber, key: item.key };
  document.getElementById('modalTitle').textContent = `Edit Request (${sheetName} • Row ${rowNumber})`;
  document.getElementById('modalKeyDisplay').textContent = item.key;

  document.getElementById('fSubject').value = item.subject;
  document.getElementById('fTicket').value = item.ticket;
  document.getElementById('fBroker').value = item.broker;
  document.getElementById('fMaster').value = item.master;
  document.getElementById('fBidder').value = item.bidder;
  document.getElementById('fAction').value = item.action;
  document.getElementById('fYard').value = item.yard;
  document.getElementById('fLots').value = item.lots;
  document.getElementById('fNotes').value = item.notes;
  document.getElementById('fStatus').value = item.status;

  editModal.classList.add('open');
}

document.getElementById('btnCancelModal').addEventListener('click', () => {
  editModal.classList.remove('open');
});

document.getElementById('entryForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const currentSheetName = editingTarget ? editingTarget.sheetName : sheetSelect.value;

  // Reload fresh copy from disk right before mutating to prevent stale overwrites
  await loadWorkbookFromDisk(currentSheetName);
  const worksheet = workbook.getWorksheet(currentSheetName);
  if (!worksheet) return alert('Target worksheet not found!');

  let targetRowNum;
  let trackingKey;

  const ticketVal = document.getElementById('fTicket').value.trim();
  const brokerVal = document.getElementById('fBroker').value.trim();

  if (editingTarget) {
    targetRowNum = editingTarget.rowNumber;
    trackingKey = ticketVal || editingTarget.key;
  } else {
    targetRowNum = getNextEmptyRow(worksheet);
    const [mm, dd, yyyy] = currentSheetName.split('-');
    const cleanBroker = (brokerVal || 'UNK').replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
    trackingKey = ticketVal || `${yyyy}${mm}${dd}-${cleanBroker}-${targetRowNum}`;
  }

  const row = worksheet.getRow(targetRowNum);
  row.getCell(1).value = document.getElementById('fSubject').value.trim();
  row.getCell(2).value = ticketVal;
  row.getCell(3).value = brokerVal;
  row.getCell(4).value = document.getElementById('fMaster').value.trim();
  row.getCell(5).value = document.getElementById('fBidder').value.trim();
  row.getCell(6).value = document.getElementById('fAction').value.trim();
  row.getCell(7).value = document.getElementById('fYard').value.trim();
  row.getCell(8).value = document.getElementById('fLots').value.trim();
  row.getCell(9).value = document.getElementById('fNotes').value.trim();
  row.getCell(10).value = document.getElementById('fStatus').value;
  row.getCell(11).value = trackingKey; // Hidden Column K persistent ID
  row.commit();

  const saved = await saveWorkbookToDisk();
  if (saved) {
    editModal.classList.remove('open');
    await loadWorkbookFromDisk(currentSheetName);
  }
});

function getNextEmptyRow(worksheet) {
  let lastRow = 1;
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    const hasData = [1, 2, 3, 10].some(col => cellText(row, col) !== '');
    if (hasData && rowNumber > lastRow) lastRow = rowNumber;
  });
  return lastRow + 1;
}

// --- 5. DAILY ROLLOVER ENGINE (CARRY FORWARD UNFINISHED) ---
document.getElementById('btnRollover').addEventListener('click', async () => {
  if (!workbook || dateSheets.length === 0) return;

  const todayStr = getTodayFormatted();
  const targetTabName = prompt('Enter new daily tab name (MM-DD-YYYY):', todayStr);
  if (!targetTabName || !DATE_TAB_REGEX.test(targetTabName.trim())) {
    return alert('Please use exact MM-DD-YYYY format.');
  }

  await loadWorkbookFromDisk();
  if (workbook.getWorksheet(targetTabName)) {
    return alert(`Tab ${targetTabName} already exists!`);
  }

  const sourceTabName = dateSheets[0]; // Most recent MM-DD-YYYY sheet
  const sourceSheet = workbook.getWorksheet(sourceTabName);
  const templateSheet = workbook.getWorksheet('template') || sourceSheet;

  // Create new sheet and clone headers/column structure from template
  const newSheet = workbook.addWorksheet(targetTabName);
  templateSheet.columns.forEach((col, idx) => {
    newSheet.getColumn(idx + 1).width = col.width || 15;
    newSheet.getColumn(idx + 1).hidden = col.hidden || false;
  });

  // Copy header row (Row 1) with styles
  const headerRow = templateSheet.getRow(1);
  const newHeader = newSheet.getRow(1);
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const targetCell = newHeader.getCell(colNumber);
    targetCell.value = cell.value;
    targetCell.style = JSON.parse(JSON.stringify(cell.style || {}));
  });
  newHeader.commit();

  // Ensure Column K is hidden
  newSheet.getColumn(11).hidden = true;

  // Filter unfinished rows from most recent daily sheet
  const unfinished = allRows.filter(
    r => r.sheetName === sourceTabName && r.status.toLowerCase() !== 'complete'
  );

  let writeRowIdx = 2;
  unfinished.forEach(item => {
    const r = newSheet.getRow(writeRowIdx++);
    r.getCell(1).value = item.subject;
    r.getCell(2).value = item.ticket;
    r.getCell(3).value = item.broker;
    r.getCell(4).value = item.master;
    r.getCell(5).value = item.bidder;
    r.getCell(6).value = item.action;
    r.getCell(7).value = item.yard;
    r.getCell(8).value = item.lots;
    r.getCell(9).value = item.notes;
    r.getCell(10).value = item.status;
    r.getCell(11).value = item.key; // Preserves original tracking key across days!
    r.commit();
  });

  const saved = await saveWorkbookToDisk();
  if (saved) {
    await loadWorkbookFromDisk(targetTabName);
    alert(`Created ${targetTabName} and rolled over ${unfinished.length} unfinished request(s) from ${sourceTabName}.`);
  }
});

// --- 6. SAFE WRITE + LOCK DETECTION ---
async function saveWorkbookToDisk() {
  try {
    setStatus('Saving to Excel...', '');
    const buffer = await workbook.xlsx.writeBuffer();
    const writable = await workingFileHandle.createWritable();
    await writable.write(buffer);
    await writable.close();
    setStatus('Saved & Synced', 'ready');
    return true;
  } catch (err) {
    console.error(err);
    setStatus('File Locked! Close Excel & Retry', 'locked');
    alert(
      'SAVE BLOCKED BY WINDOWS:\n\n' +
      'The workbook is currently open in Desktop Excel. Close the file in Excel and click Save again.'
    );
    return false;
  }
}

function getTodayFormatted() {
  const d = new Date();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  const yyyy = d.getFullYear();
  return `${mm}-${dd}-${yyyy}`;
}

function setStatus(text, stateClass) {
  statusPill.textContent = text;
  statusPill.className = 'status-pill ' + (stateClass || '');
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

