// ============================================================================
// scratch-sheets.js — In-App Workbook Scratch Sheet Tabs, Editable Spreadsheet
//                     Grid, Formula Bar, Keyboard Nav & Multi-Cell TSV Paste
// ============================================================================

// --- 6D. IN-APP WORKBOOK SCRATCH SHEET TABS & EDITABLE SPREADSHEET ENGINE (v1.14.0) ---
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

function reapplyPendingScratchEditsToWorkbook(wb) {
  if (!wb || scratchPendingEdits.size === 0) return;
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
    scratchPendingEdits.has(activeWorkspaceTab) &&
    scratchPendingEdits.get(activeWorkspaceTab).size > 0;

  if (scratchIsSaving) {
    scratchSaveStateBadge.textContent = 'Saving...';
    scratchSaveStateBadge.className = 'scratch-save-badge dirty';
  } else if (activeSheetDirty) {
    const editCount = scratchPendingEdits.get(activeWorkspaceTab).size;
    scratchSaveStateBadge.textContent = `Unsaved (${editCount}) • Auto-saving...`;
    scratchSaveStateBadge.className = 'scratch-save-badge dirty';
  } else {
    scratchSaveStateBadge.textContent = 'Saved';
    scratchSaveStateBadge.className = 'scratch-save-badge saved';
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
    const isDirty = scratchPendingEdits.has(sheetName) && scratchPendingEdits.get(sheetName).size > 0;
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
  updateStatusBarTelemetry();
}

async function switchToWorkspaceTab(targetTab) {
  commitActiveScratchInlineEdit();
  if (targetTab === activeWorkspaceTab) return;

  if (scratchPendingEdits.size > 0) {
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
  commitActiveScratchInlineEdit();
  if (scratchPendingEdits.has(sheetName) && scratchPendingEdits.get(sheetName).size > 0) {
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
  return { ...nextDims, nonEmptyRows: bounds.nonEmptyRows, dataMaxCol: bounds.maxCol };
}

function getActiveScratchCell(sheetName) {
  const cur = scratchActiveCellMap.get(sheetName);
  if (cur && cur.row >= 1 && cur.col >= 1) return cur;
  const initial = { row: 1, col: 1 };
  scratchActiveCellMap.set(sheetName, initial);
  return initial;
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
  if (scratchPendingEdits.size === 0 && silentToast) return true;
  if (scratchIsSaving) return false;

  scratchIsSaving = true;
  updateScratchSaveBadgeUI();

  const saved = await saveWorkbookToDisk();
  scratchIsSaving = false;

  if (saved) {
    scratchPendingEdits.clear();
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

function recordScratchCellEdit(sheetName, rNum, cNum, rawText) {
  if (!workbook) return;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;

  applyRawEditToExcelCell(ws, rNum, cNum, rawText);

  if (!scratchPendingEdits.has(sheetName)) {
    scratchPendingEdits.set(sheetName, new Map());
  }
  scratchPendingEdits.get(sheetName).set(`${rNum}::${cNum}`, rawText);

  // Update cached non-empty row count metadata
  const bounds = measureScratchWorksheetBounds(ws);
  const meta = scratchSheets.find(s => s.name === sheetName);
  if (meta) {
    meta.nonEmptyRows = bounds.nonEmptyRows;
    meta.maxRow = bounds.maxRow;
    meta.maxCol = bounds.maxCol;
  }

  refreshSingleScratchCellDOM(sheetName, rNum, cNum);
  if (scratchStatsMeta && activeWorkspaceTab === sheetName) {
    const dims = getScratchSheetGridDimensions(ws, sheetName);
    scratchStatsMeta.textContent = `${bounds.nonEmptyRows} active row${bounds.nonEmptyRows === 1 ? '' : 's'} (${dims.rows} × ${dims.cols} grid)`;
  }
  updateScratchSaveBadgeUI();
  renderWorkspaceTabsBar();
  scheduleScratchAutoSave();
}

function syncScratchFormulaBarFromActiveCell() {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) return;
  const { row, col } = getActiveScratchCell(activeWorkspaceTab);
  if (scratchCellCoord) {
    scratchCellCoord.textContent = getCellCoordString(row, col);
  }
  if (scratchFormulaInput && document.activeElement !== scratchFormulaInput) {
    const info = extractScratchCellInfo(ws, row, col);
    scratchFormulaInput.value = info.rawEdit;
  }
}

function selectScratchCell(rNum, cNum, { scrollIntoView = false } = {}) {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook || !scratchGridBody) return;
  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) return;

  const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
  let clampedRow = Math.max(1, rNum);
  let clampedCol = Math.max(1, cNum);

  // Auto-expand grid if navigating just past bottom or right edge
  let needsReRender = false;
  if (clampedRow > dims.rows) {
    scratchGridDimensionsMap.set(activeWorkspaceTab, { rows: clampedRow + 9, cols: dims.cols });
    needsReRender = true;
  }
  if (clampedCol > dims.cols) {
    const curRows = (scratchGridDimensionsMap.get(activeWorkspaceTab) || dims).rows;
    scratchGridDimensionsMap.set(activeWorkspaceTab, { rows: curRows, cols: clampedCol + 1 });
    needsReRender = true;
  }

  scratchActiveCellMap.set(activeWorkspaceTab, { row: clampedRow, col: clampedCol });

  if (needsReRender) {
    renderActiveScratchSheetGrid();
  } else {
    scratchGridBody.querySelectorAll('.selected-cell').forEach(el => el.classList.remove('selected-cell'));
    scratchGridBody.querySelectorAll('.active-row-hdr').forEach(el => el.classList.remove('active-row-hdr'));
    if (scratchGridHead) {
      scratchGridHead.querySelectorAll('.active-col-hdr').forEach(el => el.classList.remove('active-col-hdr'));
    }

    const td = scratchGridBody.querySelector(`td[data-r="${clampedRow}"][data-c="${clampedCol}"]`);
    if (td) {
      td.classList.add('selected-cell');
      if (scrollIntoView) {
        td.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      }
    }
    const rowHdr = scratchGridBody.querySelector(`th[data-row-hdr="${clampedRow}"]`);
    if (rowHdr) rowHdr.classList.add('active-row-hdr');
    if (scratchGridHead) {
      const colHdr = scratchGridHead.querySelector(`th[data-col-hdr="${clampedCol}"]`);
      if (colHdr) colHdr.classList.add('active-col-hdr');
    }
  }

  syncScratchFormulaBarFromActiveCell();
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
  const activePos = getActiveScratchCell(activeWorkspaceTab);
  const query = scratchSearchInput ? scratchSearchInput.value.trim().toLowerCase() : '';
  const filterRowsOnly = Boolean(chkScratchFilterRows && chkScratchFilterRows.checked && query.length > 0);

  if (scratchActiveSheetBadge) {
    scratchActiveSheetBadge.textContent = ws.name;
  }
  if (scratchStatsMeta) {
    scratchStatsMeta.textContent = `${dims.nonEmptyRows} active row${dims.nonEmptyRows === 1 ? '' : 's'} (${dims.rows} × ${dims.cols} grid)`;
  }
  updateScratchSaveBadgeUI();

  // Build Sticky Column Header Row
  let headHtml = `<tr><th class="scratch-corner-th" title="Row / Column">#</th>`;
  for (let c = 1; c <= dims.cols; c++) {
    const excelCol = ws.getColumn(c);
    const colW = excelCol && excelCol.width
      ? Math.max(88, Math.min(340, Math.round(excelCol.width * 7.6)))
      : 132;
    const isActiveCol = c === activePos.col;
    headHtml += `<th data-col-hdr="${c}" class="${isActiveCol ? 'active-col-hdr' : ''}" style="width:${colW}px; min-width:${colW}px;">${colIndexToExcelLetter(c)}</th>`;
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
      if (r === activePos.row && c === activePos.col) classes.push('selected-cell');

      const titleText = info.isFormula ? `${info.rawEdit} → ${info.display}` : info.display;
      rowCellsHtml.push(
        `<td data-r="${r}" data-c="${c}" class="${classes.join(' ')}" title="${escapeHtml(titleText)}"><span class="scratch-cell-text">${escapeHtml(info.display)}</span></td>`
      );
    }

    // Keep Row 1 header visible when filtering rows, plus any matching row
    if (filterRowsOnly && r > 1 && !rowMatchesSearch) {
      continue;
    }

    const isActiveRow = r === activePos.row;
    rowsHtml.push(
      `<tr><th data-row-hdr="${r}" class="scratch-row-hdr${isActiveRow ? ' active-row-hdr' : ''}">${r}</th>${rowCellsHtml.join('')}</tr>`
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

  syncScratchFormulaBarFromActiveCell();
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
  } else {
    for (let c = 1; c <= 10; c++) {
      newWs.getColumn(c).width = 18;
    }
    scratchActiveCellMap.set(newSheetName, { row: 1, col: 1 });
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

  commitActiveScratchInlineEdit();
  if (scratchPendingEdits.size > 0) {
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
    if (scratchGridDimensionsMap.has(currentName)) {
      scratchGridDimensionsMap.set(nextName, scratchGridDimensionsMap.get(currentName));
      scratchGridDimensionsMap.delete(currentName);
    }
    if (scratchActiveCellMap.has(currentName)) {
      scratchActiveCellMap.set(nextName, scratchActiveCellMap.get(currentName));
      scratchActiveCellMap.delete(currentName);
    }
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
    scratchGridDimensionsMap.delete(targetName);
    scratchActiveCellMap.delete(targetName);
    activeWorkspaceTab = openScratchTabs[0] || DAILY_TRACKER_WS_TAB;
    await loadWorkbookFromDisk(sheetSelect.value, {
      preserveStatus: saved === 'mirror-locked',
      skipFlushPendingScratch: true
    });
    showToast(`Deleted scratch worksheet "${targetName}".`, 'success');
  }
}

// Event Listeners for Scratch Sheet Menu, Toolbar, Formula Bar & Grid
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
      recordScratchCellEdit(activeWorkspaceTab, row, col, scratchFormulaInput.value);
    }
  });

  scratchFormulaInput.addEventListener('keydown', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB) return;
    const { row, col } = getActiveScratchCell(activeWorkspaceTab);
    if (e.key === 'Enter') {
      e.preventDefault();
      selectScratchCell(e.shiftKey ? Math.max(1, row - 1) : row + 1, col, { scrollIntoView: true });
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
    } else if (e.key === 'Tab') {
      e.preventDefault();
      selectScratchCell(row, e.shiftKey ? Math.max(1, col - 1) : col + 1, { scrollIntoView: true });
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
    }
  });
}

if (scratchGridBody) {
  scratchGridBody.addEventListener('mousedown', (e) => {
    const td = e.target.closest('td[data-r][data-c]');
    if (!td) return;
    if (scratchInlineEditState && scratchInlineEditState.tdEl === td) return;
    commitActiveScratchInlineEdit();
    const r = parseInt(td.dataset.r, 10);
    const c = parseInt(td.dataset.c, 10);
    selectScratchCell(r, c);
  });

  scratchGridBody.addEventListener('dblclick', (e) => {
    const td = e.target.closest('td[data-r][data-c]');
    if (!td) return;
    const r = parseInt(td.dataset.r, 10);
    const c = parseInt(td.dataset.c, 10);
    startScratchInlineEdit(r, c);
  });
}

if (scratchGridWrap) {
  scratchGridWrap.addEventListener('keydown', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    if (scratchInlineEditState) return; // Handled by inline input

    const { row, col } = getActiveScratchCell(activeWorkspaceTab);

    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: false });
      return;
    }

    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
      const ws = workbook.getWorksheet(activeWorkspaceTab);
      if (ws && navigator.clipboard && navigator.clipboard.writeText) {
        const info = extractScratchCellInfo(ws, row, col);
        navigator.clipboard.writeText(info.display).catch(() => {});
      }
      return;
    }

    switch (e.key) {
      case 'ArrowUp':
        e.preventDefault();
        selectScratchCell(Math.max(1, row - 1), col, { scrollIntoView: true });
        return;
      case 'ArrowDown':
        e.preventDefault();
        selectScratchCell(row + 1, col, { scrollIntoView: true });
        return;
      case 'ArrowLeft':
        e.preventDefault();
        selectScratchCell(row, Math.max(1, col - 1), { scrollIntoView: true });
        return;
      case 'ArrowRight':
        e.preventDefault();
        selectScratchCell(row, col + 1, { scrollIntoView: true });
        return;
      case 'Tab':
        e.preventDefault();
        selectScratchCell(row, e.shiftKey ? Math.max(1, col - 1) : col + 1, { scrollIntoView: true });
        return;
      case 'Enter':
      case 'F2':
        e.preventDefault();
        startScratchInlineEdit(row, col);
        return;
      case 'Delete':
      case 'Backspace':
        e.preventDefault();
        recordScratchCellEdit(activeWorkspaceTab, row, col, '');
        syncScratchFormulaBarFromActiveCell();
        return;
      default:
        break;
    }

    // Printable character starts immediate inline cell entry (Excel-style)
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      e.preventDefault();
      startScratchInlineEdit(row, col, e.key);
    }
  });

  // Multi-cell TSV / Excel / Email Clipboard Paste Handler
  scratchGridWrap.addEventListener('paste', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    if (scratchInlineEditState) return; // Let inline editor handle single-cell text paste unless grid selected

    const text = (e.clipboardData || window.clipboardData).getData('text/plain');
    if (!text) return;
    e.preventDefault();

    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;

    const rawLines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
    if (rawLines.length > 1 && rawLines[rawLines.length - 1] === '') {
      rawLines.pop();
    }

    const startPos = getActiveScratchCell(activeWorkspaceTab);
    let maxTargetRow = startPos.row;
    let maxTargetCol = startPos.col;
    let pastedCells = 0;

    rawLines.forEach((line, rOffset) => {
      const cols = line.split('\t');
      cols.forEach((cellVal, cOffset) => {
        const r = startPos.row + rOffset;
        const c = startPos.col + cOffset;
        if (r > maxTargetRow) maxTargetRow = r;
        if (c > maxTargetCol) maxTargetCol = c;

        applyRawEditToExcelCell(ws, r, c, cellVal);
        if (!scratchPendingEdits.has(activeWorkspaceTab)) {
          scratchPendingEdits.set(activeWorkspaceTab, new Map());
        }
        scratchPendingEdits.get(activeWorkspaceTab).set(`${r}::${c}`, cellVal);
        pastedCells++;
      });
    });

    const curDims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
    scratchGridDimensionsMap.set(activeWorkspaceTab, {
      rows: Math.max(curDims.rows, maxTargetRow + 5),
      cols: Math.max(curDims.cols, maxTargetCol + 2)
    });

    const bounds = measureScratchWorksheetBounds(ws);
    const meta = scratchSheets.find(s => s.name === activeWorkspaceTab);
    if (meta) {
      meta.nonEmptyRows = bounds.nonEmptyRows;
      meta.maxRow = bounds.maxRow;
      meta.maxCol = bounds.maxCol;
    }

    renderActiveScratchSheetGrid();
    renderWorkspaceTabsBar();
    scheduleScratchAutoSave();
    if (pastedCells > 1) {
      showToast(`Pasted ${pastedCells} cells into "${activeWorkspaceTab}".`, 'info');
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

