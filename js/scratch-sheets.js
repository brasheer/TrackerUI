// ============================================================================
// scratch-sheets.js — In-App Workbook Scratch Sheet Tabs, Excel-Like Range
//                     Selection, Auto-Scroll, Copy/Cut/Paste, Undo/Redo & Grid
// ============================================================================

// --- SCRATCH SHEET EXTENDED INTERACTION STATE (v1.15.0) ---
const SCRATCH_MAX_UNDO_STEPS = 50;
const scratchSelectionRangeMap = new Map(); // sheetName -> { anchorRow, anchorCol, focusRow, focusCol }
const scratchUndoStackMap = new Map(); // sheetName -> Array<{ changes, selectionBefore, selectionAfter, coalescedKey? }>
const scratchRedoStackMap = new Map(); // sheetName -> Array<{ changes, selectionBefore, selectionAfter }>
const scratchPendingColWidths = new Map(); // sheetName -> Map<colNumber, excelWidth>
let scratchClipboardRangeState = null; // { sheetName, minRow, maxRow, minCol, maxCol } | null
let scratchDragState = null; // { mode: 'cells'|'rows'|'cols', sheetName, clientX, clientY, rafId, lastExpandAt } | null
let scratchColResizeState = null; // { sheetName, col, thEl, resizerEl, startX, startWidthPx, currentWidthPx } | null
let lastScratchNativeCopyAt = 0;
let lastFormulaCoalesceMeta = null; // { sheetName, row, col, at }

// --- 6D. IN-APP WORKBOOK SCRATCH SHEET TABS & EDITABLE SPREADSHEET ENGINE ---
function colIndexToExcelLetter(colNumber) {
  let n = Math.max(1, Math.floor(colNumber));
  let letters = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

function getCellCoordString(row, col) {
  return `${colIndexToExcelLetter(col)}${row}`;
}

function getRangeCoordString(range) {
  if (!range) return 'A1';
  if (range.minRow === range.maxRow && range.minCol === range.maxCol) {
    return getCellCoordString(range.minRow, range.minCol);
  }
  return `${getCellCoordString(range.minRow, range.minCol)}:${getCellCoordString(range.maxRow, range.maxCol)}`;
}

function formatExcelDateDisplay(dateObj) {
  if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) return '';
  const yyyy = dateObj.getUTCFullYear();
  const mm = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(dateObj.getUTCDate()).padStart(2, '0');
  return `${mm}/${dd}/${yyyy}`;
}

function extractScratchCellInfo(ws, rNum, cNum) {
  if (!ws) return { display: '', rawEdit: '', isFormula: false };
  const row = ws.findRow ? ws.findRow(rNum) : ws.getRow(rNum);
  if (!row) return { display: '', rawEdit: '', isFormula: false };
  const cell = row.findCell ? row.findCell(cNum) : row.getCell(cNum);
  if (!cell) return { display: '', rawEdit: '', isFormula: false };

  const val = cell.value;
  if (val === null || val === undefined) {
    return { display: '', rawEdit: '', isFormula: false };
  }
  if (val instanceof Date) {
    const dStr = formatExcelDateDisplay(val);
    return { display: dStr, rawEdit: dStr, isFormula: false };
  }
  if (typeof val === 'object') {
    if (val.formula !== undefined || val.sharedFormula !== undefined) {
      let resStr = '';
      if (val.result instanceof Date) {
        resStr = formatExcelDateDisplay(val.result);
      } else if (val.result !== undefined && val.result !== null && typeof val.result !== 'object') {
        resStr = String(val.result);
      } else if (val.result && typeof val.result === 'object' && val.result.error) {
        resStr = String(val.result.error);
      }
      const rawFormula = val.formula
        ? `=${String(val.formula).replace(/^_xlfn\./i, '')}`
        : resStr;
      return { display: resStr, rawEdit: rawFormula, isFormula: true };
    }
    if (Array.isArray(val.richText)) {
      const joined = val.richText.map(t => t.text || '').join('');
      return { display: joined, rawEdit: joined, isFormula: false };
    }
    if (val.text !== undefined && val.text !== null) {
      const s = String(val.text);
      return { display: s, rawEdit: s, isFormula: false };
    }
    if (val.error) {
      const eStr = String(val.error);
      return { display: eStr, rawEdit: eStr, isFormula: false };
    }
    return { display: '', rawEdit: '', isFormula: false };
  }

  const str = String(val);
  return { display: str, rawEdit: str, isFormula: false };
}

function measureScratchWorksheetBounds(ws) {
  let nonEmptyRows = 0;
  let maxRow = 0;
  let maxCol = 0;
  if (!ws) return { nonEmptyRows: 0, maxRow: 0, maxCol: 0 };

  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    let rowHasValue = false;
    row.eachCell({ includeEmpty: false }, (_cell, colNumber) => {
      const info = extractScratchCellInfo(ws, rowNumber, colNumber);
      if (info.display.trim() !== '' || info.isFormula) {
        rowHasValue = true;
        if (rowNumber > maxRow) maxRow = rowNumber;
        if (colNumber > maxCol) maxCol = colNumber;
      }
    });
    if (rowHasValue) {
      nonEmptyRows++;
    }
  });

  return { nonEmptyRows, maxRow, maxCol };
}

function applyRawEditToExcelCell(ws, rNum, cNum, rawText) {
  if (!ws) return;
  const normalized = String(rawText ?? '').replace(/\r/g, '');
  const row = ws.getRow(rNum);
  const cell = row.getCell(cNum);

  if (normalized === '') {
    cell.value = null;
  } else if (normalized.startsWith('=') && normalized.length > 1) {
    cell.value = { formula: normalized.slice(1) };
  } else {
    const trimmed = normalized.trim();
    // Preserve leading-zero codes and 11+ digit carrier tracking numbers (FedEx/UPS/USPS) as exact strings
    // so Excel never formats them into floating-point scientific notation (e.g., 7.74829E+11).
    const isSafeNumber =
      /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(trimmed) &&
      trimmed.replace(/[-.]/g, '').length <= 10;
    cell.value = isSafeNumber ? Number(trimmed) : normalized;
  }
  row.commit();
}

function hasAnyPendingScratchMutations(sheetName = null) {
  if (sheetName) {
    const hasCellEdits = scratchPendingEdits.has(sheetName) && scratchPendingEdits.get(sheetName).size > 0;
    const hasColEdits = scratchPendingColWidths.has(sheetName) && scratchPendingColWidths.get(sheetName).size > 0;
    return hasCellEdits || hasColEdits;
  }
  for (const m of scratchPendingEdits.values()) {
    if (m.size > 0) return true;
  }
  for (const m of scratchPendingColWidths.values()) {
    if (m.size > 0) return true;
  }
  return false;
}

function getPendingScratchMutationCount(sheetName) {
  const cellCount = scratchPendingEdits.has(sheetName) ? scratchPendingEdits.get(sheetName).size : 0;
  const colCount = scratchPendingColWidths.has(sheetName) ? scratchPendingColWidths.get(sheetName).size : 0;
  return cellCount + colCount;
}

function reapplyPendingScratchEditsToWorkbook(wb) {
  if (!wb) return;
  if (scratchPendingEdits.size > 0) {
    for (const [sheetName, cellMap] of scratchPendingEdits.entries()) {
      const ws = wb.getWorksheet(sheetName);
      if (!ws) continue;
      for (const [coordKey, rawVal] of cellMap.entries()) {
        const [rStr, cStr] = coordKey.split('::');
        const r = parseInt(rStr, 10);
        const c = parseInt(cStr, 10);
        if (r >= 1 && c >= 1) {
          applyRawEditToExcelCell(ws, r, c, rawVal);
        }
      }
    }
  }
  if (scratchPendingColWidths.size > 0) {
    for (const [sheetName, colMap] of scratchPendingColWidths.entries()) {
      const ws = wb.getWorksheet(sheetName);
      if (!ws) continue;
      for (const [colNum, excelW] of colMap.entries()) {
        if (colNum >= 1 && excelW > 0) {
          ws.getColumn(colNum).width = excelW;
        }
      }
    }
  }
}

function getPersistedOpenScratchTabsMap() {
  try {
    const raw = localStorage.getItem(SCRATCH_OPEN_TABS_STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return (parsed && typeof parsed === 'object') ? parsed : {};
  } catch (_) {
    return {};
  }
}

function persistOpenScratchTabsForWorkbook() {
  if (!workingFileHandle || !workingFileHandle.name) return;
  try {
    const map = getPersistedOpenScratchTabsMap();
    map[workingFileHandle.name] = [...openScratchTabs];
    localStorage.setItem(SCRATCH_OPEN_TABS_STORAGE_KEY, JSON.stringify(map));
  } catch (e) {
    console.warn('Could not persist open scratch tabs:', e);
  }
}

function restoreOpenScratchTabsForWorkbook() {
  if (!workingFileHandle || !workingFileHandle.name) {
    openScratchTabs = [];
    activeWorkspaceTab = DAILY_TRACKER_WS_TAB;
    return;
  }
  const map = getPersistedOpenScratchTabsMap();
  const savedList = Array.isArray(map[workingFileHandle.name]) ? map[workingFileHandle.name] : [];
  const validNamesByLower = new Map(scratchSheets.map(s => [s.name.toLowerCase(), s.name]));

  const restored = [];
  savedList.forEach(name => {
    const canonical = validNamesByLower.get(String(name || '').toLowerCase());
    if (canonical && !restored.includes(canonical)) {
      restored.push(canonical);
    }
  });

  // Ensure currently active scratch tab stays open if it still exists in the workbook
  if (activeWorkspaceTab !== DAILY_TRACKER_WS_TAB) {
    const activeCanonical = validNamesByLower.get(activeWorkspaceTab.toLowerCase());
    if (activeCanonical) {
      activeWorkspaceTab = activeCanonical;
      if (!restored.includes(activeCanonical)) {
        restored.push(activeCanonical);
      }
    } else {
      activeWorkspaceTab = DAILY_TRACKER_WS_TAB;
    }
  }

  openScratchTabs = restored;
  persistOpenScratchTabsForWorkbook();
}

function closeScratchMenu() {
  if (scratchSheetDropdown) {
    scratchSheetDropdown.classList.remove('open');
  }
  if (btnOpenScratchMenu) {
    btnOpenScratchMenu.setAttribute('aria-expanded', 'false');
  }
}

function renderScratchMenuDropdown() {
  if (!scratchSheetMenuList) return;
  if (scratchAvailableCountBadge) {
    scratchAvailableCountBadge.textContent = String(scratchSheets.length);
  }

  scratchSheetMenuList.innerHTML = '';
  if (scratchSheets.length === 0) {
    scratchSheetMenuList.innerHTML = `
      <div class="scratch-menu-empty">
        No custom scratch sheets in this workbook yet.<br>
        Click <strong>Create New Scratch Sheet</strong> below to add one (e.g., for Title Tracking).
      </div>
    `;
    return;
  }

  scratchSheets.forEach(info => {
    const isOpen = openScratchTabs.includes(info.name);
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'scratch-menu-item';
    btn.setAttribute('role', 'menuitem');
    btn.innerHTML = `
      <span class="scratch-menu-item-name">${escapeHtml(info.name)}</span>
      <span class="scratch-menu-item-meta">
        <span>${info.nonEmptyRows} row${info.nonEmptyRows === 1 ? '' : 's'}</span>
        ${isOpen ? '<span class="scratch-open-pill">Open</span>' : ''}
      </span>
    `;
    btn.addEventListener('click', () => {
      closeScratchMenu();
      openScratchSheetInAppTab(info.name);
    });
    scratchSheetMenuList.appendChild(btn);
  });
}

function updateScratchSaveBadgeUI() {
  if (!scratchSaveStateBadge) return;
  const activeSheetDirty =
    activeWorkspaceTab !== DAILY_TRACKER_WS_TAB &&
    hasAnyPendingScratchMutations(activeWorkspaceTab);

  if (scratchIsSaving) {
    scratchSaveStateBadge.textContent = 'Saving...';
    scratchSaveStateBadge.className = 'scratch-save-badge dirty';
  } else if (activeSheetDirty) {
    const editCount = getPendingScratchMutationCount(activeWorkspaceTab);
    scratchSaveStateBadge.textContent = `Unsaved (${editCount}) • Auto-saving...`;
    scratchSaveStateBadge.className = 'scratch-save-badge dirty';
  } else {
    scratchSaveStateBadge.textContent = 'Saved';
    scratchSaveStateBadge.className = 'scratch-save-badge saved';
  }
}

function updateScratchUndoRedoButtonsUI() {
  const isScratch = activeWorkspaceTab !== DAILY_TRACKER_WS_TAB;
  const undoStack = isScratch ? (scratchUndoStackMap.get(activeWorkspaceTab) || []) : [];
  const redoStack = isScratch ? (scratchRedoStackMap.get(activeWorkspaceTab) || []) : [];

  if (btnScratchUndo) {
    btnScratchUndo.disabled = undoStack.length === 0;
    btnScratchUndo.title = undoStack.length > 0
      ? `Undo last edit (${undoStack.length} available • Ctrl+Z)`
      : 'Undo last edit (Ctrl+Z)';
  }
  if (btnScratchRedo) {
    btnScratchRedo.disabled = redoStack.length === 0;
    btnScratchRedo.title = redoStack.length > 0
      ? `Redo edit (${redoStack.length} available • Ctrl+Y)`
      : 'Redo edit (Ctrl+Y / Ctrl+Shift+Z)';
  }
}

function renderWorkspaceTabsBar() {
  if (!workspaceTabsBar || !workspaceTabsList) return;
  if (!workbook) {
    workspaceTabsBar.style.display = 'none';
    return;
  }

  workspaceTabsBar.style.display = 'flex';
  renderScratchMenuDropdown();

  const openRequestsCount = allRows.filter(r => !isCompletedStatus(r.status)).length;
  workspaceTabsList.innerHTML = '';

  // 1. Primary Pinned Daily Tracker Tab
  const dailyBtn = document.createElement('button');
  dailyBtn.type = 'button';
  dailyBtn.className = `ws-app-tab${activeWorkspaceTab === DAILY_TRACKER_WS_TAB ? ' active' : ''}`;
  dailyBtn.setAttribute('role', 'tab');
  dailyBtn.setAttribute('aria-selected', String(activeWorkspaceTab === DAILY_TRACKER_WS_TAB));
  dailyBtn.innerHTML = `
    <svg class="ui-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
      <rect x="2.5" y="3" width="15" height="14" rx="2"></rect>
      <line x1="2.5" y1="7.5" x2="17.5" y2="7.5"></line>
      <line x1="7" y1="11.5" x2="13.5" y2="11.5"></line>
      <line x1="7" y1="14.2" x2="11.5" y2="14.2"></line>
    </svg>
    <span>Daily Tracker</span>
    <span class="ws-tab-pill-count" title="${openRequestsCount} open requests across daily tabs">${openRequestsCount}</span>
  `;
  dailyBtn.addEventListener('click', () => {
    switchToWorkspaceTab(DAILY_TRACKER_WS_TAB);
  });
  workspaceTabsList.appendChild(dailyBtn);

  // 2. Open Custom / Scratch Sheet App Tabs
  openScratchTabs.forEach(sheetName => {
    const isActive = activeWorkspaceTab === sheetName;
    const isDirty = hasAnyPendingScratchMutations(sheetName);
    const meta = scratchSheets.find(s => s.name === sheetName);
    const rowCount = meta ? meta.nonEmptyRows : 0;

    const tabBtn = document.createElement('button');
    tabBtn.type = 'button';
    tabBtn.className = `ws-app-tab${isActive ? ' active' : ''}`;
    tabBtn.setAttribute('role', 'tab');
    tabBtn.setAttribute('aria-selected', String(isActive));
    tabBtn.title = `Scratch Sheet: ${sheetName} (${rowCount} non-empty rows)`;
    tabBtn.innerHTML = `
      <svg class="ui-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <rect x="3" y="3" width="14" height="14" rx="2"></rect>
        <line x1="3" y1="8" x2="17" y2="8"></line>
        <line x1="8" y1="8" x2="8" y2="17"></line>
      </svg>
      <span>${escapeHtml(sheetName)}</span>
      ${isDirty ? '<span class="ws-tab-dirty-dot" title="Unsaved edits pending"></span>' : `<span class="ws-tab-pill-count">${rowCount}</span>`}
      <span class="ws-tab-close-btn" data-close-scratch-tab="${escapeHtml(sheetName)}" title="Close ${escapeHtml(sheetName)} tab (sheet stays saved in workbook)" aria-label="Close ${escapeHtml(sheetName)} tab">×</span>
    `;

    tabBtn.addEventListener('click', (e) => {
      const closeTarget = e.target.closest('[data-close-scratch-tab]');
      if (closeTarget) {
        e.stopPropagation();
        closeScratchSheetAppTab(sheetName);
        return;
      }
      switchToWorkspaceTab(sheetName);
    });
    workspaceTabsList.appendChild(tabBtn);
  });
}

function syncWorkspaceTabsUI() {
  stopScratchPointerDrag();
  renderWorkspaceTabsBar();

  const isDailyActive = activeWorkspaceTab === DAILY_TRACKER_WS_TAB;
  if (dailyTrackerWorkspace) {
    dailyTrackerWorkspace.style.display = isDailyActive ? 'flex' : 'none';
  }
  if (scratchSheetWorkspace) {
    scratchSheetWorkspace.style.display = isDailyActive ? 'none' : 'flex';
  }

  if (!isDailyActive) {
    renderActiveScratchSheetGrid();
  }
  updateScratchUndoRedoButtonsUI();
  updateStatusBarTelemetry();
}

async function switchToWorkspaceTab(targetTab) {
  stopScratchPointerDrag();
  commitActiveScratchInlineEdit();
  if (targetTab === activeWorkspaceTab) return;

  if (hasAnyPendingScratchMutations()) {
    flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: true });
  }

  activeWorkspaceTab = targetTab;
  syncWorkspaceTabsUI();
}

function openScratchSheetInAppTab(sheetName) {
  if (!workbook) return;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;

  if (!openScratchTabs.includes(ws.name)) {
    openScratchTabs.push(ws.name);
    persistOpenScratchTabsForWorkbook();
  }
  switchToWorkspaceTab(ws.name);
}

async function closeScratchSheetAppTab(sheetName) {
  stopScratchPointerDrag();
  commitActiveScratchInlineEdit();
  if (hasAnyPendingScratchMutations(sheetName)) {
    await flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: true });
  }

  const idx = openScratchTabs.indexOf(sheetName);
  if (idx !== -1) {
    openScratchTabs.splice(idx, 1);
    persistOpenScratchTabsForWorkbook();
  }

  if (activeWorkspaceTab === sheetName) {
    const fallbackTab = openScratchTabs[idx] || openScratchTabs[idx - 1] || DAILY_TRACKER_WS_TAB;
    activeWorkspaceTab = fallbackTab;
  }
  syncWorkspaceTabsUI();
}

function getScratchSheetGridDimensions(ws, sheetName) {
  const bounds = measureScratchWorksheetBounds(ws);
  const existing = scratchGridDimensionsMap.get(sheetName) || { rows: 0, cols: 0 };
  const rows = Math.max(bounds.maxRow + 12, existing.rows, 30);
  const cols = Math.max(bounds.maxCol + 3, existing.cols, 10);
  const nextDims = { rows, cols };
  scratchGridDimensionsMap.set(sheetName, nextDims);
  return { ...nextDims, nonEmptyRows: bounds.nonEmptyRows, dataMaxRow: bounds.maxRow, dataMaxCol: bounds.maxCol };
}

function getActiveScratchCell(sheetName) {
  const cur = scratchActiveCellMap.get(sheetName);
  if (cur && cur.row >= 1 && cur.col >= 1) return cur;
  const initial = { row: 1, col: 1 };
  scratchActiveCellMap.set(sheetName, initial);
  return initial;
}

function normalizeScratchRangeObj(raw) {
  const anchorRow = Math.max(1, Math.floor(raw.anchorRow || 1));
  const anchorCol = Math.max(1, Math.floor(raw.anchorCol || 1));
  const focusRow = Math.max(1, Math.floor(raw.focusRow || anchorRow));
  const focusCol = Math.max(1, Math.floor(raw.focusCol || anchorCol));
  const minRow = Math.min(anchorRow, focusRow);
  const maxRow = Math.max(anchorRow, focusRow);
  const minCol = Math.min(anchorCol, focusCol);
  const maxCol = Math.max(anchorCol, focusCol);
  const rowCount = maxRow - minRow + 1;
  const colCount = maxCol - minCol + 1;
  return {
    anchorRow,
    anchorCol,
    focusRow,
    focusCol,
    minRow,
    maxRow,
    minCol,
    maxCol,
    rowCount,
    colCount,
    cellCount: rowCount * colCount
  };
}

function getScratchSelectionRange(sheetName) {
  const existing = scratchSelectionRangeMap.get(sheetName);
  if (existing && existing.anchorRow >= 1 && existing.anchorCol >= 1) {
    return normalizeScratchRangeObj(existing);
  }
  const active = getActiveScratchCell(sheetName);
  const initial = {
    anchorRow: active.row,
    anchorCol: active.col,
    focusRow: active.row,
    focusCol: active.col
  };
  scratchSelectionRangeMap.set(sheetName, initial);
  return normalizeScratchRangeObj(initial);
}

function formatScratchStatNumber(num) {
  if (!Number.isFinite(num)) return '0';
  const rounded = Math.round(num * 100) / 100;
  return rounded.toLocaleString(undefined, { maximumFractionDigits: 2 });
}

function updateScratchSelectionStatsUI() {
  if (!scratchSelectionStats) return;
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) {
    scratchSelectionStats.style.display = 'none';
    return;
  }
  const range = getScratchSelectionRange(activeWorkspaceTab);
  if (range.cellCount <= 1) {
    scratchSelectionStats.style.display = 'none';
    return;
  }

  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) {
    scratchSelectionStats.style.display = 'none';
    return;
  }

  const parts = [`${range.rowCount}R × ${range.colCount}C`];
  if (range.cellCount <= 25000) {
    let nonEmptyCount = 0;
    let numCount = 0;
    let numSum = 0;

    for (let r = range.minRow; r <= range.maxRow; r++) {
      const row = ws.findRow ? ws.findRow(r) : null;
      if (!row) continue;
      for (let c = range.minCol; c <= range.maxCol; c++) {
        const info = extractScratchCellInfo(ws, r, c);
        const trimmed = info.display.trim();
        if (trimmed === '') continue;
        nonEmptyCount++;

        const stripped = trimmed.replace(/^\$/, '').replace(/,/g, '');
        if (
          stripped !== '' &&
          /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(stripped) &&
          stripped.replace(/[-.]/g, '').length <= 12
        ) {
          const n = Number(stripped);
          if (Number.isFinite(n)) {
            numCount++;
            numSum += n;
          }
        }
      }
    }

    if (nonEmptyCount > 0) {
      parts.push(`Count: ${nonEmptyCount.toLocaleString()}`);
    }
    if (numCount >= 2) {
      parts.push(`Sum: ${formatScratchStatNumber(numSum)}`);
      parts.push(`Avg: ${formatScratchStatNumber(numSum / numCount)}`);
    }
  } else {
    parts.push(`${range.cellCount.toLocaleString()} cells`);
  }

  scratchSelectionStats.textContent = parts.join(' • ');
  scratchSelectionStats.style.display = 'inline-block';
}

function applyScratchSelectionDOM() {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook || !scratchGridBody) return;
  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) return;

  const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
  const range = getScratchSelectionRange(activeWorkspaceTab);
  const clip =
    scratchClipboardRangeState && scratchClipboardRangeState.sheetName === activeWorkspaceTab
      ? scratchClipboardRangeState
      : null;

  const prevClasses = [
    'selected-cell',
    'in-range-cell',
    'range-edge-top',
    'range-edge-bottom',
    'range-edge-left',
    'range-edge-right',
    'copied-edge-top',
    'copied-edge-bottom',
    'copied-edge-left',
    'copied-edge-right'
  ];
  const selector = prevClasses.map(cls => `.${cls}`).join(',');
  scratchGridBody.querySelectorAll(selector).forEach(el => {
    el.classList.remove(...prevClasses);
  });
  scratchGridBody.querySelectorAll('.active-row-hdr').forEach(el => el.classList.remove('active-row-hdr'));

  if (scratchGridHead) {
    scratchGridHead.querySelectorAll('.active-col-hdr').forEach(el => el.classList.remove('active-col-hdr'));
    const cornerTh = scratchGridHead.querySelector('.scratch-corner-th');
    if (cornerTh) {
      const isWholeSheet =
        range.minRow === 1 &&
        range.maxRow >= dims.rows &&
        range.minCol === 1 &&
        range.maxCol >= dims.cols;
      cornerTh.classList.toggle('active-corner-hdr', isWholeSheet);
    }
    const hdrRow = scratchGridHead.rows[0];
    if (hdrRow) {
      const maxHdrCol = Math.min(range.maxCol, hdrRow.cells.length - 1);
      for (let c = range.minCol; c <= maxHdrCol; c++) {
        if (hdrRow.cells[c]) hdrRow.cells[c].classList.add('active-col-hdr');
      }
    }
  }

  const trs = scratchGridBody.rows;
  for (let i = 0; i < trs.length; i++) {
    const tr = trs[i];
    const rowHdr = tr.cells[0];
    if (!rowHdr) continue;
    const r = Number(rowHdr.dataset.rowHdr);
    if (!r) continue;

    if (r >= range.minRow && r <= range.maxRow) {
      rowHdr.classList.add('active-row-hdr');
      const maxCellIdx = Math.min(range.maxCol, tr.cells.length - 1);
      if (range.cellCount === 1) {
        const td = tr.cells[range.minCol];
        if (td) td.classList.add('selected-cell');
      } else {
        for (let c = range.minCol; c <= maxCellIdx; c++) {
          const td = tr.cells[c];
          if (!td) continue;
          td.classList.add('in-range-cell');
          if (r === range.anchorRow && c === range.anchorCol) {
            td.classList.add('selected-cell');
          }
          if (r === range.minRow) td.classList.add('range-edge-top');
          if (r === range.maxRow) td.classList.add('range-edge-bottom');
          if (c === range.minCol) td.classList.add('range-edge-left');
          if (c === range.maxCol) td.classList.add('range-edge-right');
        }
      }
    }

    if (clip && r >= clip.minRow && r <= clip.maxRow) {
      const maxClipCol = Math.min(clip.maxCol, tr.cells.length - 1);
      for (let c = clip.minCol; c <= maxClipCol; c++) {
        const td = tr.cells[c];
        if (!td) continue;
        if (r === clip.minRow) td.classList.add('copied-edge-top');
        if (r === clip.maxRow) td.classList.add('copied-edge-bottom');
        if (c === clip.minCol) td.classList.add('copied-edge-left');
        if (c === clip.maxCol) td.classList.add('copied-edge-right');
      }
    }
  }

  updateScratchSelectionStatsUI();
}

function setScratchSelectionRange(
  sheetName,
  rawRange,
  { scrollIntoView = false, scrollTarget = 'focus' } = {}
) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB) return;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;

  const norm = normalizeScratchRangeObj(rawRange);
  const dims = getScratchSheetGridDimensions(ws, sheetName);

  let needsReRender = false;
  let targetRows = dims.rows;
  let targetCols = dims.cols;
  if (norm.maxRow > dims.rows) {
    targetRows = norm.maxRow + 9;
    needsReRender = true;
  }
  if (norm.maxCol > dims.cols) {
    targetCols = norm.maxCol + 2;
    needsReRender = true;
  }
  if (needsReRender) {
    scratchGridDimensionsMap.set(sheetName, { rows: targetRows, cols: targetCols });
  }

  scratchSelectionRangeMap.set(sheetName, {
    anchorRow: norm.anchorRow,
    anchorCol: norm.anchorCol,
    focusRow: norm.focusRow,
    focusCol: norm.focusCol
  });
  scratchActiveCellMap.set(sheetName, { row: norm.anchorRow, col: norm.anchorCol });

  if (activeWorkspaceTab === sheetName) {
    if (needsReRender) {
      renderActiveScratchSheetGrid();
    } else {
      applyScratchSelectionDOM();
      syncScratchFormulaBarFromActiveCell();
    }

    if (scrollIntoView && scratchGridBody) {
      const targetR = scrollTarget === 'anchor' ? norm.anchorRow : norm.focusRow;
      const targetC = scrollTarget === 'anchor' ? norm.anchorCol : norm.focusCol;
      const td = scratchGridBody.querySelector(`td[data-r="${targetR}"][data-c="${targetC}"]`);
      if (td) {
        td.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      }
    }
  }
}

function selectScratchCell(rNum, cNum, { scrollIntoView = false, extendSelection = false } = {}) {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  const clampedRow = Math.max(1, rNum);
  const clampedCol = Math.max(1, cNum);

  if (extendSelection) {
    const cur = getScratchSelectionRange(activeWorkspaceTab);
    setScratchSelectionRange(
      activeWorkspaceTab,
      {
        anchorRow: cur.anchorRow,
        anchorCol: cur.anchorCol,
        focusRow: clampedRow,
        focusCol: clampedCol
      },
      { scrollIntoView, scrollTarget: 'focus' }
    );
  } else {
    setScratchSelectionRange(
      activeWorkspaceTab,
      {
        anchorRow: clampedRow,
        anchorCol: clampedCol,
        focusRow: clampedRow,
        focusCol: clampedCol
      },
      { scrollIntoView, scrollTarget: 'anchor' }
    );
  }
}

function commitActiveScratchInlineEdit() {
  if (!scratchInlineEditState) return;
  const state = scratchInlineEditState;
  scratchInlineEditState = null;

  const newRaw = state.inputEl ? state.inputEl.value : state.originalRaw;
  if (state.inputEl && state.inputEl.parentNode) {
    state.inputEl.remove();
  }

  if (newRaw !== state.originalRaw) {
    recordScratchCellEdit(state.sheetName, state.row, state.col, newRaw);
  } else {
    refreshSingleScratchCellDOM(state.sheetName, state.row, state.col);
  }
}

function cancelActiveScratchInlineEdit() {
  if (!scratchInlineEditState) return;
  const state = scratchInlineEditState;
  scratchInlineEditState = null;

  if (state.inputEl && state.inputEl.parentNode) {
    state.inputEl.remove();
  }
  refreshSingleScratchCellDOM(state.sheetName, state.row, state.col);
  syncScratchFormulaBarFromActiveCell();
}

function scheduleScratchAutoSave() {
  if (scratchAutoSaveTimer) {
    clearTimeout(scratchAutoSaveTimer);
  }
  scratchAutoSaveTimer = setTimeout(() => {
    scratchAutoSaveTimer = null;
    flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: true });
  }, 1200);
}

async function flushPendingScratchEditsToDisk({ reloadAfter = false, silentToast = false } = {}) {
  if (scratchAutoSaveTimer) {
    clearTimeout(scratchAutoSaveTimer);
    scratchAutoSaveTimer = null;
  }
  commitActiveScratchInlineEdit();

  if (!workbook || !workingFileHandle) return false;
  if (!hasAnyPendingScratchMutations() && silentToast) return true;
  if (scratchIsSaving) return false;

  scratchIsSaving = true;
  updateScratchSaveBadgeUI();

  const saved = await saveWorkbookToDisk();
  scratchIsSaving = false;

  if (saved) {
    scratchPendingEdits.clear();
    scratchPendingColWidths.clear();
    updateScratchSaveBadgeUI();
    renderWorkspaceTabsBar();
    if (reloadAfter) {
      await loadWorkbookFromDisk(sheetSelect.value, {
        preserveStatus: saved === 'mirror-locked',
        skipFlushPendingScratch: true
      });
    }
    if (!silentToast) {
      showToast('Scratch sheet saved to workbook & mirror.', 'success');
    }
    return true;
  } else {
    updateScratchSaveBadgeUI();
    return false;
  }
}

function refreshSingleScratchCellDOM(sheetName, rNum, cNum) {
  if (activeWorkspaceTab !== sheetName || !scratchGridBody || !workbook) return;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;
  const td = scratchGridBody.querySelector(`td[data-r="${rNum}"][data-c="${cNum}"]`);
  if (!td) return;

  const info = extractScratchCellInfo(ws, rNum, cNum);
  td.classList.toggle('formula-cell', info.isFormula);
  td.title = info.isFormula ? `${info.rawEdit} → ${info.display}` : info.display;

  let span = td.querySelector('.scratch-cell-text');
  if (!span) {
    span = document.createElement('span');
    span.className = 'scratch-cell-text';
    td.innerHTML = '';
    td.appendChild(span);
  }
  span.textContent = info.display;
}

// --- BATCH CELL MUTATION ENGINE WITH 50-STEP UNDO / REDO ---
function applyScratchCellChanges(
  sheetName,
  editsList,
  {
    recordUndo = true,
    coalesceKey = null,
    selectionAfter = null
  } = {}
) {
  if (!workbook || !sheetName || !Array.isArray(editsList) || editsList.length === 0) return 0;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return 0;

  const selectionBefore = getScratchSelectionRange(sheetName);
  const recordedChanges = [];
  let maxTargetRow = 0;
  let maxTargetCol = 0;

  if (!scratchPendingEdits.has(sheetName)) {
    scratchPendingEdits.set(sheetName, new Map());
  }
  const pendingSheetMap = scratchPendingEdits.get(sheetName);

  editsList.forEach(item => {
    const r = Math.max(1, Math.floor(item.r));
    const c = Math.max(1, Math.floor(item.c));
    const afterRaw = String(item.rawText ?? '').replace(/\r/g, '');
    const beforeRaw = extractScratchCellInfo(ws, r, c).rawEdit;

    if (r > maxTargetRow) maxTargetRow = r;
    if (c > maxTargetCol) maxTargetCol = c;

    if (beforeRaw === afterRaw) return;

    applyRawEditToExcelCell(ws, r, c, afterRaw);
    pendingSheetMap.set(`${r}::${c}`, afterRaw);
    recordedChanges.push({ r, c, beforeRaw, afterRaw });
  });

  if (recordedChanges.length === 0) {
    if (selectionAfter) {
      setScratchSelectionRange(sheetName, selectionAfter);
    }
    return 0;
  }

  if (recordUndo) {
    if (!scratchUndoStackMap.has(sheetName)) {
      scratchUndoStackMap.set(sheetName, []);
    }
    const undoStack = scratchUndoStackMap.get(sheetName);
    const now = Date.now();
    const canCoalesce =
      coalesceKey &&
      recordedChanges.length === 1 &&
      undoStack.length > 0 &&
      undoStack[undoStack.length - 1].coalesceKey === coalesceKey &&
      lastFormulaCoalesceMeta &&
      lastFormulaCoalesceMeta.key === coalesceKey &&
      now - lastFormulaCoalesceMeta.at < 2500;

    if (canCoalesce) {
      const top = undoStack[undoStack.length - 1];
      top.changes[0].afterRaw = recordedChanges[0].afterRaw;
      lastFormulaCoalesceMeta.at = now;
    } else {
      undoStack.push({
        changes: recordedChanges,
        selectionBefore: {
          anchorRow: selectionBefore.anchorRow,
          anchorCol: selectionBefore.anchorCol,
          focusRow: selectionBefore.focusRow,
          focusCol: selectionBefore.focusCol
        },
        selectionAfter: selectionAfter || {
          anchorRow: selectionBefore.anchorRow,
          anchorCol: selectionBefore.anchorCol,
          focusRow: selectionBefore.focusRow,
          focusCol: selectionBefore.focusCol
        },
        coalesceKey: coalesceKey || null
      });
      if (undoStack.length > SCRATCH_MAX_UNDO_STEPS) {
        undoStack.shift();
      }
      lastFormulaCoalesceMeta = coalesceKey ? { key: coalesceKey, at: now } : null;
    }
    scratchRedoStackMap.set(sheetName, []);
  }

  // Update grid dimensions if paste/edit exceeded current bounds
  const curDims = getScratchSheetGridDimensions(ws, sheetName);
  const needsExpand = maxTargetRow > curDims.rows || maxTargetCol > curDims.cols;
  if (needsExpand) {
    scratchGridDimensionsMap.set(sheetName, {
      rows: Math.max(curDims.rows, maxTargetRow + 5),
      cols: Math.max(curDims.cols, maxTargetCol + 2)
    });
  }

  // Update cached non-empty row count metadata
  const bounds = measureScratchWorksheetBounds(ws);
  const meta = scratchSheets.find(s => s.name === sheetName);
  if (meta) {
    meta.nonEmptyRows = bounds.nonEmptyRows;
    meta.maxRow = bounds.maxRow;
    meta.maxCol = bounds.maxCol;
  }

  if (selectionAfter) {
    scratchSelectionRangeMap.set(sheetName, {
      anchorRow: selectionAfter.anchorRow,
      anchorCol: selectionAfter.anchorCol,
      focusRow: selectionAfter.focusRow,
      focusCol: selectionAfter.focusCol
    });
    scratchActiveCellMap.set(sheetName, {
      row: selectionAfter.anchorRow,
      col: selectionAfter.anchorCol
    });
  }

  const hasActiveSearch = Boolean(scratchSearchInput && scratchSearchInput.value.trim().length > 0);
  if (activeWorkspaceTab === sheetName) {
    if (needsExpand || recordedChanges.length > 40 || hasActiveSearch) {
      renderActiveScratchSheetGrid();
    } else {
      recordedChanges.forEach(ch => refreshSingleScratchCellDOM(sheetName, ch.r, ch.c));
      applyScratchSelectionDOM();
      syncScratchFormulaBarFromActiveCell();
      if (scratchStatsMeta) {
        const nextDims = getScratchSheetGridDimensions(ws, sheetName);
        scratchStatsMeta.textContent = `${bounds.nonEmptyRows} active row${bounds.nonEmptyRows === 1 ? '' : 's'} (${nextDims.rows} × ${nextDims.cols} grid)`;
      }
    }
  }

  updateScratchUndoRedoButtonsUI();
  updateScratchSaveBadgeUI();
  renderWorkspaceTabsBar();
  scheduleScratchAutoSave();
  return recordedChanges.length;
}

function recordScratchCellEdit(sheetName, rNum, cNum, rawText, { coalesceKey = null } = {}) {
  applyScratchCellChanges(sheetName, [{ r: rNum, c: cNum, rawText }], {
    recordUndo: true,
    coalesceKey
  });
}

function undoScratchEdit() {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  commitActiveScratchInlineEdit();
  const sheetName = activeWorkspaceTab;
  const undoStack = scratchUndoStackMap.get(sheetName);
  if (!undoStack || undoStack.length === 0) return;

  const entry = undoStack.pop();
  lastFormulaCoalesceMeta = null;

  const reverseEdits = entry.changes.map(ch => ({
    r: ch.r,
    c: ch.c,
    rawText: ch.beforeRaw
  }));

  applyScratchCellChanges(sheetName, reverseEdits, {
    recordUndo: false,
    selectionAfter: entry.selectionBefore
  });

  if (!scratchRedoStackMap.has(sheetName)) {
    scratchRedoStackMap.set(sheetName, []);
  }
  scratchRedoStackMap.get(sheetName).push(entry);
  updateScratchUndoRedoButtonsUI();
  syncScratchFormulaBarFromActiveCell();
}

function redoScratchEdit() {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  commitActiveScratchInlineEdit();
  const sheetName = activeWorkspaceTab;
  const redoStack = scratchRedoStackMap.get(sheetName);
  if (!redoStack || redoStack.length === 0) return;

  const entry = redoStack.pop();
  lastFormulaCoalesceMeta = null;

  const forwardEdits = entry.changes.map(ch => ({
    r: ch.r,
    c: ch.c,
    rawText: ch.afterRaw
  }));

  applyScratchCellChanges(sheetName, forwardEdits, {
    recordUndo: false,
    selectionAfter: entry.selectionAfter
  });

  if (!scratchUndoStackMap.has(sheetName)) {
    scratchUndoStackMap.set(sheetName, []);
  }
  scratchUndoStackMap.get(sheetName).push(entry);
  updateScratchUndoRedoButtonsUI();
  syncScratchFormulaBarFromActiveCell();
}

function syncScratchFormulaBarFromActiveCell() {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) return;
  const range = getScratchSelectionRange(activeWorkspaceTab);
  if (scratchCellCoord) {
    scratchCellCoord.textContent = getRangeCoordString(range);
  }
  if (scratchFormulaInput && document.activeElement !== scratchFormulaInput) {
    const info = extractScratchCellInfo(ws, range.anchorRow, range.anchorCol);
    scratchFormulaInput.value = info.rawEdit;
  }
}

function startScratchInlineEdit(rNum, cNum, initialChar = null) {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook || !scratchGridBody) return;
  commitActiveScratchInlineEdit();
  selectScratchCell(rNum, cNum);

  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) return;
  const td = scratchGridBody.querySelector(`td[data-r="${rNum}"][data-c="${cNum}"]`);
  if (!td) return;

  const info = extractScratchCellInfo(ws, rNum, cNum);
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'scratch-inline-editor';
  input.value = initialChar !== null ? initialChar : info.rawEdit;

  scratchInlineEditState = {
    sheetName: activeWorkspaceTab,
    row: rNum,
    col: cNum,
    originalRaw: info.rawEdit,
    inputEl: input,
    tdEl: td
  };

  if (scratchFormulaInput) {
    scratchFormulaInput.value = input.value;
  }

  input.addEventListener('input', () => {
    if (scratchFormulaInput) {
      scratchFormulaInput.value = input.value;
    }
  });

  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      e.stopPropagation();
      commitActiveScratchInlineEdit();
      selectScratchCell(e.shiftKey ? Math.max(1, rNum - 1) : rNum + 1, cNum, { scrollIntoView: true });
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
    } else if (e.key === 'Tab') {
      e.preventDefault();
      e.stopPropagation();
      commitActiveScratchInlineEdit();
      selectScratchCell(rNum, e.shiftKey ? Math.max(1, cNum - 1) : cNum + 1, { scrollIntoView: true });
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
    } else if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      cancelActiveScratchInlineEdit();
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
    }
  });

  input.addEventListener('blur', () => {
    commitActiveScratchInlineEdit();
  });

  td.appendChild(input);
  input.focus();
  if (initialChar !== null) {
    input.setSelectionRange(input.value.length, input.value.length);
  } else {
    input.select();
  }
}

function renderActiveScratchSheetGrid() {
  if (!scratchGridHead || !scratchGridBody || !workbook || activeWorkspaceTab === DAILY_TRACKER_WS_TAB) {
    return;
  }
  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) {
    activeWorkspaceTab = DAILY_TRACKER_WS_TAB;
    syncWorkspaceTabsUI();
    return;
  }

  const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
  const range = getScratchSelectionRange(activeWorkspaceTab);
  const query = scratchSearchInput ? scratchSearchInput.value.trim().toLowerCase() : '';
  const filterRowsOnly = Boolean(chkScratchFilterRows && chkScratchFilterRows.checked && query.length > 0);

  if (scratchActiveSheetBadge) {
    scratchActiveSheetBadge.textContent = ws.name;
  }
  if (scratchStatsMeta) {
    scratchStatsMeta.textContent = `${dims.nonEmptyRows} active row${dims.nonEmptyRows === 1 ? '' : 's'} (${dims.rows} × ${dims.cols} grid)`;
  }
  updateScratchSaveBadgeUI();
  updateScratchUndoRedoButtonsUI();

  // Build Sticky Column Header Row (with clickable headers + drag-to-resize handles)
  let headHtml = `<tr><th class="scratch-corner-th" title="Click to select all cells (Ctrl+A)">#</th>`;
  for (let c = 1; c <= dims.cols; c++) {
    const excelCol = ws.getColumn(c);
    const colW = excelCol && excelCol.width
      ? Math.max(68, Math.min(500, Math.round(excelCol.width * 7.6)))
      : 132;
    const isActiveCol = c >= range.minCol && c <= range.maxCol;
    const colLetter = colIndexToExcelLetter(c);
    headHtml += `<th data-col-hdr="${c}" class="${isActiveCol ? 'active-col-hdr' : ''}" style="width:${colW}px; min-width:${colW}px;" title="Click or drag to select Column ${colLetter}"><span class="scratch-col-label">${colLetter}</span><span class="scratch-col-resizer" data-resize-col="${c}" title="Drag to resize Column ${colLetter} • Double-click to auto-fit"></span></th>`;
  }
  headHtml += `</tr>`;
  scratchGridHead.innerHTML = headHtml;

  // Check whether Row 1 looks like a header row
  let row1HasValues = false;
  for (let c = 1; c <= dims.cols; c++) {
    if (extractScratchCellInfo(ws, 1, c).display.trim() !== '') {
      row1HasValues = true;
      break;
    }
  }

  let matchCellCount = 0;
  const rowsHtml = [];

  for (let r = 1; r <= dims.rows; r++) {
    let rowMatchesSearch = false;
    const rowCellsHtml = [];

    for (let c = 1; c <= dims.cols; c++) {
      const info = extractScratchCellInfo(ws, r, c);
      const isHit = query.length > 0 && (
        info.display.toLowerCase().includes(query) ||
        info.rawEdit.toLowerCase().includes(query)
      );
      if (isHit) {
        rowMatchesSearch = true;
        matchCellCount++;
      }

      const classes = ['scratch-cell'];
      if (r === 1 && row1HasValues) classes.push('header-like-cell');
      if (info.isFormula) classes.push('formula-cell');
      if (isHit) classes.push('search-hit');

      const titleText = info.isFormula ? `${info.rawEdit} → ${info.display}` : info.display;
      rowCellsHtml.push(
        `<td data-r="${r}" data-c="${c}" class="${classes.join(' ')}" title="${escapeHtml(titleText)}"><span class="scratch-cell-text">${escapeHtml(info.display)}</span></td>`
      );
    }

    // Keep Row 1 header visible when filtering rows, plus any matching row
    if (filterRowsOnly && r > 1 && !rowMatchesSearch) {
      continue;
    }

    rowsHtml.push(
      `<tr><th data-row-hdr="${r}" class="scratch-row-hdr" title="Click or drag to select Row ${r}">${r}</th>${rowCellsHtml.join('')}</tr>`
    );
  }

  scratchGridBody.innerHTML = rowsHtml.join('');

  if (scratchSearchMatchCount) {
    if (query.length > 0) {
      scratchSearchMatchCount.style.display = 'inline-block';
      scratchSearchMatchCount.textContent = `${matchCellCount} match${matchCellCount === 1 ? '' : 'es'}`;
    } else {
      scratchSearchMatchCount.style.display = 'none';
    }
  }

  applyScratchSelectionDOM();
  syncScratchFormulaBarFromActiveCell();
}

// --- POINTER DRAG RANGE SELECTION + SMOOTH EDGE AUTO-SCROLL ---
function calcScratchAutoScrollVelocity(distanceIntoZone) {
  return Math.min(26, Math.max(3, Math.round(distanceIntoZone * 0.3)));
}

function updateScratchDragSelectionFromPoint(clientX, clientY) {
  if (!scratchDragState || !scratchGridWrap || !workbook) return;
  if (activeWorkspaceTab !== scratchDragState.sheetName) {
    stopScratchPointerDrag();
    return;
  }
  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) return;

  const rect = scratchGridWrap.getBoundingClientRect();
  const headerH = scratchGridHead ? (scratchGridHead.offsetHeight || 28) : 28;
  const rowHdrW = 48;
  const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
  const curRange = getScratchSelectionRange(activeWorkspaceTab);

  if (scratchDragState.mode === 'rows') {
    const probeX = Math.max(rect.left + 12, Math.min(rect.right - 8, rect.left + 20));
    const probeY = Math.max(rect.top + headerH + 3, Math.min(rect.bottom - 3, clientY));
    const hit = document.elementFromPoint(probeX, probeY);
    if (!hit) return;
    const rowEl = hit.closest('th[data-row-hdr], td[data-r]');
    if (!rowEl) return;
    const targetRow = Number(rowEl.dataset.rowHdr || rowEl.dataset.r);
    if (targetRow >= 1 && (targetRow !== curRange.focusRow || curRange.focusCol !== dims.cols)) {
      setScratchSelectionRange(activeWorkspaceTab, {
        anchorRow: curRange.anchorRow,
        anchorCol: 1,
        focusRow: targetRow,
        focusCol: dims.cols
      });
    }
    return;
  }

  if (scratchDragState.mode === 'cols') {
    const probeX = Math.max(rect.left + rowHdrW + 3, Math.min(rect.right - 3, clientX));
    const probeY = Math.max(rect.top + 6, Math.min(rect.bottom - 6, rect.top + 12));
    const hit = document.elementFromPoint(probeX, probeY);
    if (!hit) return;
    const colEl = hit.closest('th[data-col-hdr], td[data-c]');
    if (!colEl) return;
    const targetCol = Number(colEl.dataset.colHdr || colEl.dataset.c);
    if (targetCol >= 1 && (targetCol !== curRange.focusCol || curRange.focusRow !== dims.rows)) {
      setScratchSelectionRange(activeWorkspaceTab, {
        anchorRow: 1,
        anchorCol: curRange.anchorCol,
        focusRow: dims.rows,
        focusCol: targetCol
      });
    }
    return;
  }

  // Standard rectangular cell drag ('cells')
  const probeX = Math.max(rect.left + rowHdrW + 3, Math.min(rect.right - 3, clientX));
  const probeY = Math.max(rect.top + headerH + 3, Math.min(rect.bottom - 3, clientY));
  const hit = document.elementFromPoint(probeX, probeY);
  if (!hit) return;
  const td = hit.closest('td[data-r][data-c]');
  if (!td) return;
  const r = Number(td.dataset.r);
  const c = Number(td.dataset.c);
  if (r >= 1 && c >= 1 && (r !== curRange.focusRow || c !== curRange.focusCol)) {
    setScratchSelectionRange(activeWorkspaceTab, {
      anchorRow: curRange.anchorRow,
      anchorCol: curRange.anchorCol,
      focusRow: r,
      focusCol: c
    });
  }
}

function runScratchDragAutoScrollTick() {
  if (!scratchDragState || !scratchGridWrap || !workbook) return;
  const rect = scratchGridWrap.getBoundingClientRect();
  const headerH = scratchGridHead ? (scratchGridHead.offsetHeight || 28) : 28;
  const rowHdrW = 48;
  const edgeZone = 40;

  let vx = 0;
  let vy = 0;
  const { clientX, clientY, mode } = scratchDragState;

  if (mode !== 'cols') {
    const topZoneEdge = rect.top + headerH + edgeZone;
    const bottomZoneEdge = rect.bottom - edgeZone;
    if (clientY < topZoneEdge) {
      vy = -calcScratchAutoScrollVelocity(topZoneEdge - clientY);
    } else if (clientY > bottomZoneEdge) {
      vy = calcScratchAutoScrollVelocity(clientY - bottomZoneEdge);
    }
  }

  if (mode !== 'rows') {
    const leftZoneEdge = rect.left + rowHdrW + edgeZone;
    const rightZoneEdge = rect.right - edgeZone;
    if (clientX < leftZoneEdge) {
      vx = -calcScratchAutoScrollVelocity(leftZoneEdge - clientX);
    } else if (clientX > rightZoneEdge) {
      vx = calcScratchAutoScrollVelocity(clientX - rightZoneEdge);
    }
  }

  if (vx !== 0 || vy !== 0) {
    const prevTop = scratchGridWrap.scrollTop;
    const prevLeft = scratchGridWrap.scrollLeft;
    scratchGridWrap.scrollTop += vy;
    scratchGridWrap.scrollLeft += vx;

    // Auto-expand grid smoothly when dragging past the bottom/right scroll boundary
    const now = Date.now();
    if (now - scratchDragState.lastExpandAt > 200) {
      const ws = workbook.getWorksheet(activeWorkspaceTab);
      if (ws) {
        const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
        const atBottomLimit =
          vy > 0 &&
          clientY > rect.bottom + 8 &&
          scratchGridWrap.scrollTop === prevTop &&
          scratchGridWrap.scrollTop + scratchGridWrap.clientHeight >= scratchGridWrap.scrollHeight - 3;
        const atRightLimit =
          vx > 0 &&
          clientX > rect.right + 8 &&
          scratchGridWrap.scrollLeft === prevLeft &&
          scratchGridWrap.scrollLeft + scratchGridWrap.clientWidth >= scratchGridWrap.scrollWidth - 3;

        if (atBottomLimit || atRightLimit) {
          scratchDragState.lastExpandAt = now;
          scratchGridDimensionsMap.set(activeWorkspaceTab, {
            rows: atBottomLimit ? dims.rows + 5 : dims.rows,
            cols: atRightLimit ? dims.cols + 1 : dims.cols
          });
          renderActiveScratchSheetGrid();
          if (atBottomLimit) scratchGridWrap.scrollTop += 40;
          if (atRightLimit) scratchGridWrap.scrollLeft += 60;
        }
      }
    }

    updateScratchDragSelectionFromPoint(clientX, clientY);
  }

  if (scratchDragState) {
    scratchDragState.rafId = requestAnimationFrame(runScratchDragAutoScrollTick);
  }
}

function startScratchPointerDrag(mode, clientX, clientY) {
  stopScratchPointerDrag();
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB) return;
  scratchDragState = {
    mode,
    sheetName: activeWorkspaceTab,
    clientX,
    clientY,
    rafId: null,
    lastExpandAt: 0
  };
  if (scratchGridWrap) {
    scratchGridWrap.classList.add('is-drag-selecting');
  }
  scratchDragState.rafId = requestAnimationFrame(runScratchDragAutoScrollTick);
}

function stopScratchPointerDrag() {
  if (scratchDragState) {
    if (scratchDragState.rafId) {
      cancelAnimationFrame(scratchDragState.rafId);
    }
    scratchDragState = null;
  }
  if (scratchGridWrap) {
    scratchGridWrap.classList.remove('is-drag-selecting');
  }
}

// --- COLUMN WIDTH DRAG-RESIZE & DOUBLE-CLICK AUTO-FIT ---
function recordScratchColumnWidthChange(sheetName, colNum, widthPx) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB) return;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;
  const clampedPx = Math.max(68, Math.min(500, Math.round(widthPx)));
  const excelWidth = +(clampedPx / 7.6).toFixed(1);
  ws.getColumn(colNum).width = excelWidth;

  if (!scratchPendingColWidths.has(sheetName)) {
    scratchPendingColWidths.set(sheetName, new Map());
  }
  scratchPendingColWidths.get(sheetName).set(colNum, excelWidth);
  updateScratchSaveBadgeUI();
  renderWorkspaceTabsBar();
  scheduleScratchAutoSave();
}

function autoFitScratchColumnWidth(sheetName, colNum) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB || !scratchGridHead) return;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;
  const dims = getScratchSheetGridDimensions(ws, sheetName);

  let maxLen = colIndexToExcelLetter(colNum).length + 2;
  for (let r = 1; r <= dims.rows; r++) {
    const disp = extractScratchCellInfo(ws, r, colNum).display;
    if (disp.length > maxLen) {
      maxLen = disp.length;
    }
  }
  const fittedPx = Math.max(80, Math.min(440, Math.round(maxLen * 7.8 + 26)));
  const th = scratchGridHead.querySelector(`th[data-col-hdr="${colNum}"]`);
  if (th) {
    th.style.width = `${fittedPx}px`;
    th.style.minWidth = `${fittedPx}px`;
  }
  recordScratchColumnWidthChange(sheetName, colNum, fittedPx);
}

// --- KEYBOARD DATA-BOUNDARY JUMP (Ctrl+Arrow) & CLIPBOARD HELPERS ---
function computeScratchCtrlArrowTarget(ws, dims, startRow, startCol, dr, dc) {
  const hasVal = (r, c) => extractScratchCellInfo(ws, r, c).display.trim() !== '';
  const inBounds = (r, c) => r >= 1 && r <= dims.rows && c >= 1 && c <= dims.cols;

  let nextR = startRow + dr;
  let nextC = startCol + dc;
  if (!inBounds(nextR, nextC)) {
    return { row: startRow, col: startCol };
  }

  if (hasVal(startRow, startCol) && hasVal(nextR, nextC)) {
    // Walk to the last non-empty cell of the current contiguous block
    let curR = nextR;
    let curC = nextC;
    while (inBounds(curR + dr, curC + dc) && hasVal(curR + dr, curC + dc)) {
      curR += dr;
      curC += dc;
    }
    return { row: curR, col: curC };
  } else {
    // Skip empty cells until landing on the next non-empty cell (or grid edge)
    let curR = nextR;
    let curC = nextC;
    while (inBounds(curR, curC) && !hasVal(curR, curC)) {
      if (!inBounds(curR + dr, curC + dc)) {
        return { row: curR, col: curC };
      }
      curR += dr;
      curC += dc;
    }
    return { row: curR, col: curC };
  }
}

function serializeScratchSelectionToTSV(ws, range) {
  const lines = [];
  for (let r = range.minRow; r <= range.maxRow; r++) {
    const rowVals = [];
    for (let c = range.minCol; c <= range.maxCol; c++) {
      const disp = extractScratchCellInfo(ws, r, c).display.replace(/[\t\r\n]+/g, ' ');
      rowVals.push(disp);
    }
    lines.push(rowVals.join('\t'));
  }
  return lines.join('\r\n');
}

function clearScratchSelectionCells(sheetName, range) {
  if (!workbook) return 0;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return 0;
  const edits = [];
  for (let r = range.minRow; r <= range.maxRow; r++) {
    for (let c = range.minCol; c <= range.maxCol; c++) {
      edits.push({ r, c, rawText: '' });
    }
  }
  return applyScratchCellChanges(sheetName, edits, { recordUndo: true });
}

function parseClipboardTSVMatrix(rawText) {
  const normalized = String(rawText || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  const lines = normalized.split('\n');
  if (lines.length > 1 && lines[lines.length - 1] === '') {
    lines.pop();
  }
  return lines.map(line => line.split('\t'));
}

function migrateScratchSheetRuntimeMaps(oldName, newName) {
  const maps = [
    scratchGridDimensionsMap,
    scratchActiveCellMap,
    scratchSelectionRangeMap,
    scratchUndoStackMap,
    scratchRedoStackMap,
    scratchPendingColWidths
  ];
  maps.forEach(m => {
    if (m.has(oldName)) {
      if (newName) m.set(newName, m.get(oldName));
      m.delete(oldName);
    }
  });
  if (scratchClipboardRangeState && scratchClipboardRangeState.sheetName === oldName) {
    if (newName) {
      scratchClipboardRangeState.sheetName = newName;
    } else {
      scratchClipboardRangeState = null;
    }
  }
}

function resetScratchWorkbookRuntimeState() {
  stopScratchPointerDrag();
  scratchPendingEdits.clear();
  scratchPendingColWidths.clear();
  scratchGridDimensionsMap.clear();
  scratchActiveCellMap.clear();
  scratchSelectionRangeMap.clear();
  scratchUndoStackMap.clear();
  scratchRedoStackMap.clear();
  scratchClipboardRangeState = null;
  lastFormulaCoalesceMeta = null;
}

function validateScratchSheetNameInput(rawVal, ignoreCurrentName = null) {
  const trimmed = String(rawVal || '').trim();
  if (!trimmed) return 'Please enter a sheet name.';
  if (trimmed.length > 31) return 'Excel worksheet names cannot exceed 31 characters.';
  if (/[:\\/?*[\]]/.test(trimmed)) {
    return 'Sheet names cannot contain : \\ / ? * [ or ].';
  }
  if (trimmed.startsWith("'") || trimmed.endsWith("'")) {
    return 'Sheet names cannot start or end with an apostrophe.';
  }
  if (DATE_TAB_REGEX.test(trimmed)) {
    return 'MM-DD-YYYY names are reserved for Daily Tracker tabs. Use a descriptive name (e.g., "Title Tracking").';
  }
  if (isProtectedReferenceSheetName(trimmed) || isTemplatePlaceholderTab(trimmed)) {
    return `"${trimmed}" is a reserved tracker reference sheet name.`;
  }
  if (workbook) {
    let duplicate = false;
    workbook.eachSheet(ws => {
      if (
        (!ignoreCurrentName || ws.name.toLowerCase() !== ignoreCurrentName.toLowerCase()) &&
        ws.name.trim().toLowerCase() === trimmed.toLowerCase()
      ) {
        duplicate = true;
      }
    });
    if (duplicate) {
      return `A worksheet named "${trimmed}" already exists in this workbook.`;
    }
  }
  return null;
}

async function promptCreateNewScratchSheet() {
  closeScratchMenu();
  closeToolsMenu();
  if (!workbook || !workingFileHandle) {
    return showModalAlert({
      title: 'No Active Workbook',
      message: 'Please connect an Agent Folder and open a workbook first.'
    });
  }

  const defaultSuggested = scratchSheets.some(s => s.name.toLowerCase() === 'title tracking')
    ? `Scratch ${scratchSheets.length + 1}`
    : 'Title Tracking';

  const result = await showModalPrompt({
    title: 'Create New Scratch Sheet',
    badge: workingFileHandle.name,
    message: 'Create a custom worksheet inside your workbook and open it in an app tab (saved directly to Excel & mirrored to View-Only):',
    selectLabel: 'Starter Layout',
    selectOptions: [
      {
        value: 'title_headers',
        label: 'Title Tracking Columns (Lot #, Yard #, Broker / Master #, Title Status, Carrier / Tracking #, Notes)'
      },
      {
        value: 'blank',
        label: 'Blank Spreadsheet Grid'
      }
    ],
    defaultSelectValue: 'title_headers',
    inputLabel: 'Scratch Sheet Name',
    defaultValue: defaultSuggested,
    placeholder: 'e.g., Title Tracking, Pending Titles, Scratch',
    confirmText: 'Create & Open Tab',
    validate: (val) => validateScratchSheetNameInput(val)
  });

  if (!result) return;
  const newSheetName = (typeof result === 'object' ? result.value : result).trim();
  const starterLayout = (typeof result === 'object' && result.selectValue) ? result.selectValue : 'blank';

  await loadWorkbookFromDisk(sheetSelect.value);
  const recheckErr = validateScratchSheetNameInput(newSheetName);
  if (recheckErr) {
    return showModalAlert({ title: 'Cannot Create Scratch Sheet', message: recheckErr });
  }

  const newWs = workbook.addWorksheet(newSheetName);

  if (starterLayout === 'title_headers') {
    const starterHeaders = [
      { header: 'Lot #', width: 16 },
      { header: 'Yard #', width: 11 },
      { header: 'Broker / Master #', width: 26 },
      { header: 'Title Status', width: 18 },
      { header: 'Carrier / Tracking #', width: 26 },
      { header: 'Date Updated', width: 15 },
      { header: 'Notes', width: 36 }
    ];
    const headerRow = newWs.getRow(1);
    headerRow.height = 22;
    starterHeaders.forEach((colDef, idx) => {
      newWs.getColumn(idx + 1).width = colDef.width;
      const cell = headerRow.getCell(idx + 1);
      cell.value = colDef.header;
      cell.font = { bold: true, color: { argb: 'FF0F172A' } };
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE2E8F0' }
      };
    });
    headerRow.commit();
    scratchActiveCellMap.set(newSheetName, { row: 2, col: 1 });
    scratchSelectionRangeMap.set(newSheetName, { anchorRow: 2, anchorCol: 1, focusRow: 2, focusCol: 1 });
  } else {
    for (let c = 1; c <= 10; c++) {
      newWs.getColumn(c).width = 18;
    }
    scratchActiveCellMap.set(newSheetName, { row: 1, col: 1 });
    scratchSelectionRangeMap.set(newSheetName, { anchorRow: 1, anchorCol: 1, focusRow: 1, focusCol: 1 });
  }

  const saved = await saveWorkbookToDisk();
  if (saved) {
    if (!openScratchTabs.includes(newSheetName)) {
      openScratchTabs.push(newSheetName);
      persistOpenScratchTabsForWorkbook();
    }
    activeWorkspaceTab = newSheetName;
    await loadWorkbookFromDisk(sheetSelect.value, {
      preserveStatus: saved === 'mirror-locked',
      skipFlushPendingScratch: true
    });
    showToast(`Created scratch sheet "${newSheetName}" and opened in tab.`, 'success');
  }
}

async function handleRenameActiveScratchSheet() {
  if (!workbook || activeWorkspaceTab === DAILY_TRACKER_WS_TAB) return;
  const currentName = activeWorkspaceTab;
  if (!isCustomScratchSheetName(currentName)) return;

  stopScratchPointerDrag();
  commitActiveScratchInlineEdit();
  if (hasAnyPendingScratchMutations()) {
    await flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: true });
  }

  const rawNew = await showModalPrompt({
    title: 'Rename Scratch Sheet',
    badge: currentName,
    message: `Enter a new worksheet name for "${currentName}":`,
    inputLabel: 'New Scratch Sheet Name',
    defaultValue: currentName,
    confirmText: 'Rename Sheet',
    validate: (val) => validateScratchSheetNameInput(val, currentName)
  });
  if (!rawNew) return;
  const nextName = (typeof rawNew === 'object' ? rawNew.value : rawNew).trim();
  if (nextName === currentName) return;

  await loadWorkbookFromDisk(sheetSelect.value);
  const ws = workbook.getWorksheet(currentName);
  if (!ws) return;

  ws.name = nextName;
  const saved = await saveWorkbookToDisk();
  if (saved) {
    openScratchTabs = openScratchTabs.map(n => (n === currentName ? nextName : n));
    persistOpenScratchTabsForWorkbook();
    migrateScratchSheetRuntimeMaps(currentName, nextName);
    activeWorkspaceTab = nextName;
    await loadWorkbookFromDisk(sheetSelect.value, {
      preserveStatus: saved === 'mirror-locked',
      skipFlushPendingScratch: true
    });
    showToast(`Renamed scratch sheet "${currentName}" → "${nextName}".`, 'success');
  }
}

async function handleDeleteActiveScratchSheet() {
  if (!workbook || activeWorkspaceTab === DAILY_TRACKER_WS_TAB) return;
  const targetName = activeWorkspaceTab;
  if (!isCustomScratchSheetName(targetName)) return;

  stopScratchPointerDrag();
  commitActiveScratchInlineEdit();
  const meta = scratchSheets.find(s => s.name === targetName);
  const rowCount = meta ? meta.nonEmptyRows : 0;

  const confirmed = await showModalConfirm({
    title: `Delete Scratch Sheet "${targetName}"?`,
    badge: `${rowCount} non-empty row(s)`,
    message:
      `Are you sure you want to permanently delete worksheet "${targetName}" from "${workingFileHandle.name}"?\n\n` +
      `(Tip: If you only want to hide the tab in the app without deleting the Excel sheet, click the "×" button on the tab instead.)`,
    confirmText: 'Delete Worksheet Permanently',
    variant: 'danger'
  });
  if (!confirmed) return;

  scratchPendingEdits.delete(targetName);
  scratchPendingColWidths.delete(targetName);
  await loadWorkbookFromDisk(sheetSelect.value, { skipFlushPendingScratch: true });
  const ws = workbook.getWorksheet(targetName);
  if (!ws) return;

  if (workbook.worksheets.length <= 1) {
    return showModalAlert({
      title: 'Cannot Delete Only Worksheet',
      message: 'Excel workbooks must contain at least one worksheet.'
    });
  }

  workbook.removeWorksheet(ws.id);
  const saved = await saveWorkbookToDisk();
  if (saved) {
    openScratchTabs = openScratchTabs.filter(n => n !== targetName);
    persistOpenScratchTabsForWorkbook();
    migrateScratchSheetRuntimeMaps(targetName, null);
    activeWorkspaceTab = openScratchTabs[0] || DAILY_TRACKER_WS_TAB;
    await loadWorkbookFromDisk(sheetSelect.value, {
      preserveStatus: saved === 'mirror-locked',
      skipFlushPendingScratch: true
    });
    showToast(`Deleted scratch worksheet "${targetName}".`, 'success');
  }
}

// ============================================================================
// EVENT LISTENERS — SCRATCH MENU, TOOLBAR, FORMULA BAR, DRAG GRID & CLIPBOARD
// ============================================================================
if (btnOpenScratchMenu) {
  btnOpenScratchMenu.addEventListener('click', (e) => {
    e.stopPropagation();
    closeToolsMenu();
    const isOpen = scratchSheetDropdown && scratchSheetDropdown.classList.contains('open');
    if (isOpen) {
      closeScratchMenu();
    } else if (scratchSheetDropdown) {
      renderScratchMenuDropdown();
      scratchSheetDropdown.classList.add('open');
      btnOpenScratchMenu.setAttribute('aria-expanded', 'true');
    }
  });
}

document.addEventListener('click', (e) => {
  if (scratchSheetMenuWrap && !scratchSheetMenuWrap.contains(e.target)) {
    closeScratchMenu();
  }
});

if (btnNewScratchSheetFromBar) {
  btnNewScratchSheetFromBar.addEventListener('click', promptCreateNewScratchSheet);
}
if (menuNewScratchSheet) {
  menuNewScratchSheet.addEventListener('click', promptCreateNewScratchSheet);
}

if (scratchSearchInput) {
  scratchSearchInput.addEventListener('input', () => {
    commitActiveScratchInlineEdit();
    renderActiveScratchSheetGrid();
  });
}
if (chkScratchFilterRows) {
  chkScratchFilterRows.addEventListener('change', () => {
    commitActiveScratchInlineEdit();
    renderActiveScratchSheetGrid();
  });
}

if (btnScratchUndo) {
  btnScratchUndo.addEventListener('click', () => {
    undoScratchEdit();
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
  });
}
if (btnScratchRedo) {
  btnScratchRedo.addEventListener('click', () => {
    redoScratchEdit();
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
  });
}

if (btnScratchAddRows) {
  btnScratchAddRows.addEventListener('click', () => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    commitActiveScratchInlineEdit();
    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;
    const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
    scratchGridDimensionsMap.set(activeWorkspaceTab, { rows: dims.rows + 10, cols: dims.cols });
    renderActiveScratchSheetGrid();
  });
}

if (btnScratchAddCol) {
  btnScratchAddCol.addEventListener('click', () => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    commitActiveScratchInlineEdit();
    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;
    const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
    scratchGridDimensionsMap.set(activeWorkspaceTab, { rows: dims.rows, cols: dims.cols + 1 });
    renderActiveScratchSheetGrid();
  });
}

if (btnScratchRenameSheet) {
  btnScratchRenameSheet.addEventListener('click', handleRenameActiveScratchSheet);
}
if (btnScratchDeleteSheet) {
  btnScratchDeleteSheet.addEventListener('click', handleDeleteActiveScratchSheet);
}
if (btnScratchSaveNow) {
  btnScratchSaveNow.addEventListener('click', () => {
    flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: false });
  });
}

if (scratchFormulaInput) {
  scratchFormulaInput.addEventListener('input', () => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    const { row, col } = getActiveScratchCell(activeWorkspaceTab);
    if (scratchInlineEditState && scratchInlineEditState.inputEl) {
      scratchInlineEditState.inputEl.value = scratchFormulaInput.value;
    } else {
      recordScratchCellEdit(activeWorkspaceTab, row, col, scratchFormulaInput.value, {
        coalesceKey: `fx::${activeWorkspaceTab}::${row}::${col}`
      });
    }
  });

  scratchFormulaInput.addEventListener('keydown', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB) return;
    const { row, col } = getActiveScratchCell(activeWorkspaceTab);
    if (e.key === 'Enter') {
      e.preventDefault();
      lastFormulaCoalesceMeta = null;
      selectScratchCell(e.shiftKey ? Math.max(1, row - 1) : row + 1, col, { scrollIntoView: true });
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
    } else if (e.key === 'Tab') {
      e.preventDefault();
      lastFormulaCoalesceMeta = null;
      selectScratchCell(row, e.shiftKey ? Math.max(1, col - 1) : col + 1, { scrollIntoView: true });
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
    } else if (e.key === 'Escape') {
      e.preventDefault();
      lastFormulaCoalesceMeta = null;
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
      syncScratchFormulaBarFromActiveCell();
    }
  });
}

// Column Header Selection + Column Width Drag-Resize + Corner Select-All
if (scratchGridHead) {
  scratchGridHead.addEventListener('mousedown', (e) => {
    if (e.button !== 0 || activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;

    // 1. Column Width Drag-Resize Handle
    const resizer = e.target.closest('.scratch-col-resizer[data-resize-col]');
    if (resizer) {
      e.preventDefault();
      e.stopPropagation();
      commitActiveScratchInlineEdit();
      const colNum = parseInt(resizer.dataset.resizeCol, 10);
      const thEl = resizer.closest('th[data-col-hdr]');
      if (!thEl || !colNum) return;
      const startW = thEl.getBoundingClientRect().width;
      resizer.classList.add('resizing');
      scratchColResizeState = {
        sheetName: activeWorkspaceTab,
        col: colNum,
        thEl,
        resizerEl: resizer,
        startX: e.clientX,
        startWidthPx: startW,
        currentWidthPx: startW
      };
      return;
    }

    // 2. Top-Left Corner '#' Select-All Button
    const cornerTh = e.target.closest('th.scratch-corner-th');
    if (cornerTh) {
      e.preventDefault();
      commitActiveScratchInlineEdit();
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
      const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
      setScratchSelectionRange(activeWorkspaceTab, {
        anchorRow: 1,
        anchorCol: 1,
        focusRow: dims.rows,
        focusCol: dims.cols
      });
      return;
    }

    // 3. Whole-Column Click / Shift+Click / Drag Selection
    const colTh = e.target.closest('th[data-col-hdr]');
    if (colTh) {
      e.preventDefault();
      commitActiveScratchInlineEdit();
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
      const c = parseInt(colTh.dataset.colHdr, 10);
      if (!c) return;
      const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
      const anchorCol = e.shiftKey ? getScratchSelectionRange(activeWorkspaceTab).anchorCol : c;
      setScratchSelectionRange(activeWorkspaceTab, {
        anchorRow: 1,
        anchorCol,
        focusRow: dims.rows,
        focusCol: c
      });
      startScratchPointerDrag('cols', e.clientX, e.clientY);
    }
  });

  scratchGridHead.addEventListener('dblclick', (e) => {
    const resizer = e.target.closest('.scratch-col-resizer[data-resize-col]');
    if (!resizer || activeWorkspaceTab === DAILY_TRACKER_WS_TAB) return;
    e.preventDefault();
    e.stopPropagation();
    const colNum = parseInt(resizer.dataset.resizeCol, 10);
    if (colNum >= 1) {
      autoFitScratchColumnWidth(activeWorkspaceTab, colNum);
    }
  });
}

// Cell & Row-Header Selection + Drag Highlighting + Double-Click Inline Edit
if (scratchGridBody) {
  scratchGridBody.addEventListener('mousedown', (e) => {
    if (e.button !== 0 || activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;

    // 1. Whole-Row Header Click / Shift+Click / Drag
    const rowTh = e.target.closest('th[data-row-hdr]');
    if (rowTh) {
      e.preventDefault();
      commitActiveScratchInlineEdit();
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
      const r = parseInt(rowTh.dataset.rowHdr, 10);
      if (!r) return;
      const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
      const anchorRow = e.shiftKey ? getScratchSelectionRange(activeWorkspaceTab).anchorRow : r;
      setScratchSelectionRange(activeWorkspaceTab, {
        anchorRow,
        anchorCol: 1,
        focusRow: r,
        focusCol: dims.cols
      });
      startScratchPointerDrag('rows', e.clientX, e.clientY);
      return;
    }

    // 2. Grid Cell Click / Shift+Click / Drag Range Selection
    const td = e.target.closest('td[data-r][data-c]');
    if (!td) return;
    if (scratchInlineEditState && scratchInlineEditState.tdEl === td) return;

    e.preventDefault();
    commitActiveScratchInlineEdit();
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });

    const r = parseInt(td.dataset.r, 10);
    const c = parseInt(td.dataset.c, 10);
    selectScratchCell(r, c, { extendSelection: e.shiftKey });
    startScratchPointerDrag('cells', e.clientX, e.clientY);
  });

  scratchGridBody.addEventListener('dblclick', (e) => {
    const td = e.target.closest('td[data-r][data-c]');
    if (!td) return;
    const r = parseInt(td.dataset.r, 10);
    const c = parseInt(td.dataset.c, 10);
    startScratchInlineEdit(r, c);
  });
}

// Global Pointer Tracking for Cell/Header Dragging & Column Width Resizing
window.addEventListener('mousemove', (e) => {
  if (scratchColResizeState) {
    const deltaX = e.clientX - scratchColResizeState.startX;
    const nextW = Math.max(68, Math.min(500, Math.round(scratchColResizeState.startWidthPx + deltaX)));
    scratchColResizeState.currentWidthPx = nextW;
    scratchColResizeState.thEl.style.width = `${nextW}px`;
    scratchColResizeState.thEl.style.minWidth = `${nextW}px`;
    return;
  }

  if (scratchDragState) {
    scratchDragState.clientX = e.clientX;
    scratchDragState.clientY = e.clientY;
    updateScratchDragSelectionFromPoint(e.clientX, e.clientY);
  }
});

window.addEventListener('mouseup', () => {
  if (scratchColResizeState) {
    const st = scratchColResizeState;
    scratchColResizeState = null;
    if (st.resizerEl) st.resizerEl.classList.remove('resizing');
    if (Math.abs(st.currentWidthPx - st.startWidthPx) >= 2) {
      recordScratchColumnWidthChange(st.sheetName, st.col, st.currentWidthPx);
    }
  }
  if (scratchDragState) {
    stopScratchPointerDrag();
  }
});

window.addEventListener('blur', () => {
  if (scratchDragState) {
    stopScratchPointerDrag();
  }
});

// Keyboard Navigation, Shortcuts & Clipboard Copy/Cut/Paste on Grid
if (scratchGridWrap) {
  scratchGridWrap.addEventListener('keydown', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    if (scratchInlineEditState) return; // Handled by inline input

    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;
    const range = getScratchSelectionRange(activeWorkspaceTab);
    const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
    const keyLower = e.key.toLowerCase();
    const mod = e.ctrlKey || e.metaKey;

    // Save: Ctrl/Cmd + S
    if (mod && keyLower === 's') {
      e.preventDefault();
      flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: false });
      return;
    }

    // Undo: Ctrl/Cmd + Z (without Shift)
    if (mod && keyLower === 'z' && !e.shiftKey) {
      e.preventDefault();
      undoScratchEdit();
      return;
    }

    // Redo: Ctrl/Cmd + Y or Ctrl/Cmd + Shift + Z
    if (mod && (keyLower === 'y' || (keyLower === 'z' && e.shiftKey))) {
      e.preventDefault();
      redoScratchEdit();
      return;
    }

    // Select All: Ctrl/Cmd + A (Data bounds first, then entire grid on second press)
    if (mod && keyLower === 'a') {
      e.preventDefault();
      const hasData = dims.dataMaxRow >= 1 && dims.dataMaxCol >= 1;
      const alreadyCoversData =
        hasData &&
        range.minRow === 1 &&
        range.maxRow === dims.dataMaxRow &&
        range.minCol === 1 &&
        range.maxCol === dims.dataMaxCol;

      if (hasData && !alreadyCoversData) {
        setScratchSelectionRange(activeWorkspaceTab, {
          anchorRow: 1,
          anchorCol: 1,
          focusRow: dims.dataMaxRow,
          focusCol: dims.dataMaxCol
        });
      } else {
        setScratchSelectionRange(activeWorkspaceTab, {
          anchorRow: 1,
          anchorCol: 1,
          focusRow: dims.rows,
          focusCol: dims.cols
        });
      }
      return;
    }

    // Copy / Cut fallback if browser does not dispatch native copy/cut event
    if (mod && (keyLower === 'c' || keyLower === 'x')) {
      const isCut = keyLower === 'x';
      const snapRange = { ...range };
      const snapSheet = activeWorkspaceTab;
      setTimeout(() => {
        if (Date.now() - lastScratchNativeCopyAt < 80) return;
        const targetWs = workbook && workbook.getWorksheet(snapSheet);
        if (!targetWs) return;
        const tsv = serializeScratchSelectionToTSV(targetWs, snapRange);
        if (navigator.clipboard && navigator.clipboard.writeText) {
          navigator.clipboard.writeText(tsv).catch(() => {});
        }
        if (isCut) {
          scratchClipboardRangeState = null;
          clearScratchSelectionCells(snapSheet, snapRange);
          if (snapRange.cellCount > 1) {
            showToast(`Cut ${snapRange.cellCount} cells (${getRangeCoordString(snapRange)}).`, 'info');
          }
        } else {
          scratchClipboardRangeState = {
            sheetName: snapSheet,
            minRow: snapRange.minRow,
            maxRow: snapRange.maxRow,
            minCol: snapRange.minCol,
            maxCol: snapRange.maxCol
          };
          applyScratchSelectionDOM();
          if (snapRange.cellCount > 1) {
            showToast(`Copied ${snapRange.cellCount} cells (${getRangeCoordString(snapRange)}).`, 'info');
          }
        }
      }, 15);
      return;
    }

    // Arrow Navigation (Single move, Shift+Arrow range extend, Ctrl+Arrow data jump, Ctrl+Shift+Arrow)
    if (
      e.key === 'ArrowUp' ||
      e.key === 'ArrowDown' ||
      e.key === 'ArrowLeft' ||
      e.key === 'ArrowRight'
    ) {
      e.preventDefault();
      const dr = e.key === 'ArrowUp' ? -1 : e.key === 'ArrowDown' ? 1 : 0;
      const dc = e.key === 'ArrowLeft' ? -1 : e.key === 'ArrowRight' ? 1 : 0;
      const baseRow = e.shiftKey ? range.focusRow : range.anchorRow;
      const baseCol = e.shiftKey ? range.focusCol : range.anchorCol;

      let nextRow = Math.max(1, baseRow + dr);
      let nextCol = Math.max(1, baseCol + dc);
      if (mod) {
        const jumped = computeScratchCtrlArrowTarget(ws, dims, baseRow, baseCol, dr, dc);
        nextRow = jumped.row;
        nextCol = jumped.col;
      }

      selectScratchCell(nextRow, nextCol, {
        scrollIntoView: true,
        extendSelection: e.shiftKey
      });
      return;
    }

    switch (e.key) {
      case 'Tab':
        e.preventDefault();
        selectScratchCell(
          range.anchorRow,
          e.shiftKey ? Math.max(1, range.anchorCol - 1) : range.anchorCol + 1,
          { scrollIntoView: true }
        );
        return;
      case 'Enter':
      case 'F2':
        e.preventDefault();
        startScratchInlineEdit(range.anchorRow, range.anchorCol);
        return;
      case 'Delete':
      case 'Backspace': {
        e.preventDefault();
        const cleared = clearScratchSelectionCells(activeWorkspaceTab, range);
        syncScratchFormulaBarFromActiveCell();
        if (range.cellCount > 1 && cleared > 1) {
          showToast(`Cleared ${cleared} cells (${getRangeCoordString(range)}).`, 'info');
        }
        return;
      }
      case 'Escape':
        if (scratchClipboardRangeState) {
          e.preventDefault();
          scratchClipboardRangeState = null;
          applyScratchSelectionDOM();
          return;
        }
        if (range.cellCount > 1) {
          e.preventDefault();
          selectScratchCell(range.anchorRow, range.anchorCol);
          return;
        }
        break;
      default:
        break;
    }

    // Printable character starts immediate inline cell entry on the anchor cell (Excel-style)
    if (e.key.length === 1 && !mod && !e.altKey) {
      e.preventDefault();
      startScratchInlineEdit(range.anchorRow, range.anchorCol, e.key);
    }
  });

  // Multi-Cell TSV Copy Handler (Ctrl+C)
  scratchGridWrap.addEventListener('copy', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    if (scratchInlineEditState) return;

    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;
    const range = getScratchSelectionRange(activeWorkspaceTab);
    const tsv = serializeScratchSelectionToTSV(ws, range);

    if (e.clipboardData) {
      e.preventDefault();
      e.clipboardData.setData('text/plain', tsv);
      lastScratchNativeCopyAt = Date.now();
    }

    scratchClipboardRangeState = {
      sheetName: activeWorkspaceTab,
      minRow: range.minRow,
      maxRow: range.maxRow,
      minCol: range.minCol,
      maxCol: range.maxCol
    };
    applyScratchSelectionDOM();
    if (range.cellCount > 1) {
      showToast(`Copied ${range.cellCount} cells (${getRangeCoordString(range)}).`, 'info');
    }
  });

  // Multi-Cell TSV Cut Handler (Ctrl+X)
  scratchGridWrap.addEventListener('cut', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    if (scratchInlineEditState) return;

    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;
    const range = getScratchSelectionRange(activeWorkspaceTab);
    const tsv = serializeScratchSelectionToTSV(ws, range);

    if (e.clipboardData) {
      e.preventDefault();
      e.clipboardData.setData('text/plain', tsv);
      lastScratchNativeCopyAt = Date.now();
    }

    scratchClipboardRangeState = null;
    clearScratchSelectionCells(activeWorkspaceTab, range);
    syncScratchFormulaBarFromActiveCell();
    if (range.cellCount > 1) {
      showToast(`Cut ${range.cellCount} cells (${getRangeCoordString(range)}).`, 'info');
    }
  });

  // Multi-Cell TSV / Excel / Email Clipboard Paste Handler (Ctrl+V + 1-Cell Repeat-Fill)
  scratchGridWrap.addEventListener('paste', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    if (scratchInlineEditState) return; // Let inline editor handle single-cell text paste

    const text = (e.clipboardData || window.clipboardData).getData('text/plain');
    if (text === undefined || text === null || text === '') return;
    e.preventDefault();

    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;

    const matrix = parseClipboardTSVMatrix(text);
    if (matrix.length === 0) return;

    const curRange = getScratchSelectionRange(activeWorkspaceTab);
    const isSingleValueClipboard = matrix.length === 1 && matrix[0].length === 1;
    const edits = [];
    let nextSelection = null;

    if (isSingleValueClipboard && curRange.cellCount > 1) {
      // Excel-style Repeat-Fill: fill entire highlighted multi-cell range with the single copied value
      const fillVal = matrix[0][0];
      for (let r = curRange.minRow; r <= curRange.maxRow; r++) {
        for (let c = curRange.minCol; c <= curRange.maxCol; c++) {
          edits.push({ r, c, rawText: fillVal });
        }
      }
      nextSelection = {
        anchorRow: curRange.anchorRow,
        anchorCol: curRange.anchorCol,
        focusRow: curRange.focusRow,
        focusCol: curRange.focusCol
      };
    } else {
      // Paste 2D TSV block starting from top-left of current selection
      const startRow = curRange.minRow;
      const startCol = curRange.minCol;
      let maxR = startRow;
      let maxC = startCol;

      matrix.forEach((cols, rOffset) => {
        cols.forEach((cellVal, cOffset) => {
          const r = startRow + rOffset;
          const c = startCol + cOffset;
          if (r > maxR) maxR = r;
          if (c > maxC) maxC = c;
          edits.push({ r, c, rawText: cellVal });
        });
      });

      nextSelection = {
        anchorRow: startRow,
        anchorCol: startCol,
        focusRow: maxR,
        focusCol: maxC
      };
    }

    scratchClipboardRangeState = null;
    applyScratchCellChanges(activeWorkspaceTab, edits, {
      recordUndo: true,
      selectionAfter: nextSelection
    });

    if (edits.length > 1) {
      showToast(`Pasted ${edits.length} cells into "${activeWorkspaceTab}".`, 'info');
    }
  });
}

// Global Ctrl+S shortcut while viewing a Scratch Sheet Tab
document.addEventListener('keydown', (e) => {
  if (
    activeWorkspaceTab !== DAILY_TRACKER_WS_TAB &&
    (e.ctrlKey || e.metaKey) &&
    e.key.toLowerCase() === 's'
  ) {
    e.preventDefault();
    flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: false });
  }
});
