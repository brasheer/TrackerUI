// ============================================================================
// scratch-sheets.js — Multi-Tracker Profile & Scratch Sheet Workspace Tabs,
//                     Spreadsheet Grid DOM, Pointer Selection, Inline Editor,
//                     Undo/Redo, Fill Palette, Context Menu & Wizards
// ============================================================================

// --- SCRATCH SHEET EXTENDED INTERACTION STATE (v1.16.0 - v1.19.8) ---
const SCRATCH_MAX_UNDO_STEPS = 50;
const SCRATCH_STRIPED_STORAGE_KEY = 'trackerui_scratch_striped_rows';
const SCRATCH_AUTO_INCREMENT_STORAGE_KEY = 'trackerui_scratch_auto_increment';
const SCRATCH_LAST_FILL_STORAGE_KEY = 'trackerui_scratch_last_fill_id';
const SCRATCH_ROW1_UNBOLD_STORAGE_KEY = 'trackerui_scratch_row1_unbold_v1';
const scratchSelectionRangeMap = new Map(); // sheetName -> { anchorRow, anchorCol, focusRow, focusCol }
const scratchUndoStackMap = new Map(); // sheetName -> Array<{ changes, selectionBefore, selectionAfter, coalesceKey? }>
const scratchRedoStackMap = new Map(); // sheetName -> Array<{ changes, selectionBefore, selectionAfter }>
const scratchPendingColWidths = new Map(); // sheetName -> Map<colNumber, excelWidth>
const scratchPendingFills = new Map(); // sheetName -> Map<'row::col', { fill, font, numFmt? }>
let scratchClipboardRangeState = null; // { sheetName, minRow, maxRow, minCol, maxCol } | null
let scratchInternalClipboardText = '';
let scratchInternalClipboardState = null; // { tsv, html, styleMatrix, rawEditMatrix, sourceRange } | null
let scratchDragState = null; // { mode: 'cells'|'rows'|'cols'|'fill', sheetName, clientX, clientY, rafId, lastExpandAt, sourceRange?, fillTargetRange? } | null
let scratchColResizeState = null; // { sheetName, col, thEl, resizerEl, startX, startWidthPx, currentWidthPx } | null
let lastScratchNativeCopyAt = 0;
let lastFormulaCoalesceMeta = null; // { key, at }
let scratchMutationVersion = 0;
let scratchStripedRowsEnabled = localStorage.getItem(SCRATCH_STRIPED_STORAGE_KEY) !== 'false';
let scratchAutoIncrementEnabled = localStorage.getItem(SCRATCH_AUTO_INCREMENT_STORAGE_KEY) === 'true';
let activeScratchFillPresetId = localStorage.getItem(SCRATCH_LAST_FILL_STORAGE_KEY) || 'gold_80';
let activeTextToColumnsState = null; // { sheetName, col, minRow, maxRow, sourceRows: [{ r, raw }] } | null
let activeScratchAutoFilterState = null; // { sheetName, headers: [{ col, letter, label }] } | null

// --- PENDING SCRATCH MUTATIONS & PERSISTED OPEN TABS ---
function hasAnyPendingScratchMutations(sheetName = null) {
  if (sheetName) {
    const hasCellEdits = scratchPendingEdits.has(sheetName) && scratchPendingEdits.get(sheetName).size > 0;
    const hasColEdits = scratchPendingColWidths.has(sheetName) && scratchPendingColWidths.get(sheetName).size > 0;
    const hasFillEdits = scratchPendingFills.has(sheetName) && scratchPendingFills.get(sheetName).size > 0;
    return hasCellEdits || hasColEdits || hasFillEdits;
  }
  for (const m of scratchPendingEdits.values()) {
    if (m.size > 0) return true;
  }
  for (const m of scratchPendingColWidths.values()) {
    if (m.size > 0) return true;
  }
  for (const m of scratchPendingFills.values()) {
    if (m.size > 0) return true;
  }
  return false;
}

function getPendingScratchMutationCount(sheetName) {
  const cellCount = scratchPendingEdits.has(sheetName) ? scratchPendingEdits.get(sheetName).size : 0;
  const colCount = scratchPendingColWidths.has(sheetName) ? scratchPendingColWidths.get(sheetName).size : 0;
  const fillCount = scratchPendingFills.has(sheetName) ? scratchPendingFills.get(sheetName).size : 0;
  return cellCount + colCount + fillCount;
}

function getScratchRow1UnboldStorageKey(sheetName) {
  const wbKey = typeof getScratchPersistenceWorkbookKey === 'function' ? getScratchPersistenceWorkbookKey() : '';
  const cleanSheet = String(sheetName || '').trim();
  if (!wbKey || !cleanSheet) return '';
  return `${wbKey}::${cleanSheet}`;
}

function getPersistedRow1UnboldMap() {
  try {
    const raw = localStorage.getItem(SCRATCH_ROW1_UNBOLD_STORAGE_KEY);
    if (!raw) return {};
    const parsed = JSON.parse(raw);
    return (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) ? parsed : {};
  } catch (_) {
    return {};
  }
}

function getPersistedRow1UnboldCols(sheetName) {
  const key = getScratchRow1UnboldStorageKey(sheetName);
  if (!key) return [];
  const map = getPersistedRow1UnboldMap();
  const cols = map[key];
  return Array.isArray(cols) ? cols.map(n => Number(n)).filter(n => Number.isInteger(n) && n >= 1) : [];
}

function setPersistedRow1UnboldCol(sheetName, colNum, isUnbold) {
  const key = getScratchRow1UnboldStorageKey(sheetName);
  const col = Math.floor(Number(colNum) || 0);
  if (!key || col < 1) return;
  try {
    const map = getPersistedRow1UnboldMap();
    const existing = new Set(
      Array.isArray(map[key]) ? map[key].map(n => Number(n)).filter(n => Number.isInteger(n) && n >= 1) : []
    );
    if (isUnbold) existing.add(col);
    else existing.delete(col);
    if (existing.size > 0) {
      map[key] = Array.from(existing).sort((a, b) => a - b);
    } else {
      delete map[key];
    }
    localStorage.setItem(SCRATCH_ROW1_UNBOLD_STORAGE_KEY, JSON.stringify(map));
  } catch (_) { /* ignore */ }
}

function applyPersistedRow1UnboldToWorksheet(ws, sheetName) {
  if (!ws || !sheetName) return;
  const cols = getPersistedRow1UnboldCols(sheetName);
  if (!cols || cols.length === 0) return;
  const pendingFillMap = scratchPendingFills.get(sheetName);
  const row1 = ws.getRow(1);
  let touched = false;
  cols.forEach(c => {
    if (pendingFillMap && pendingFillMap.has(`1::${c}`)) return;
    const cell = row1.getCell(c);
    const curFont = cell.style && cell.style.font ? cell.style.font : cell.font;
    if (curFont && curFont.bold === false) return;
    if (!curFont || curFont.bold !== true) {
      const nextFont = curFont ? JSON.parse(JSON.stringify(curFont)) : {};
      nextFont.bold = false;
      cell.style = { ...(cell.style || {}), font: nextFont };
      touched = true;
    }
  });
  if (touched && row1.commit) row1.commit();
}

function reapplyPendingScratchEditsToWorkbook(wb) {
  if (!wb) return;
  if (typeof wb.eachSheet === 'function') {
    wb.eachSheet(ws => {
      if (ws && ws.name) {
        applyPersistedRow1UnboldToWorksheet(ws, ws.name);
      }
    });
  }
  // Apply style/format snapshots first so applyRawEditToExcelCell sees each cell's numFmt
  if (scratchPendingFills.size > 0) {
    for (const [sheetName, fillMap] of scratchPendingFills.entries()) {
      const ws = wb.getWorksheet(sheetName);
      if (!ws) continue;
      for (const [coordKey, styleSnap] of fillMap.entries()) {
        const [rStr, cStr] = coordKey.split('::');
        const r = parseInt(rStr, 10);
        const c = parseInt(cStr, 10);
        if (r >= 1 && c >= 1) {
          applyCellStyleSnapshotToExcelCell(ws, r, c, styleSnap);
        }
      }
    }
  }
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

function getScratchPersistenceWorkbookKey() {
  if (!workingFileHandle || !workingFileHandle.name) return '';
  if (!activeProfileId || activeProfileId === DEFAULT_TRACKER_PROFILE_ID) {
    return workingFileHandle.name;
  }
  return `${activeProfileId}::${workingFileHandle.name}`;
}

function persistOpenScratchTabsForWorkbook() {
  const storageWbKey = getScratchPersistenceWorkbookKey();
  if (!storageWbKey) return;
  try {
    const map = getPersistedOpenScratchTabsMap();
    map[storageWbKey] = [...openScratchTabs];
    localStorage.setItem(SCRATCH_OPEN_TABS_STORAGE_KEY, JSON.stringify(map));
  } catch (e) {
    console.warn('Could not persist open scratch tabs:', e);
  }
}

function restoreOpenScratchTabsForWorkbook() {
  const storageWbKey = getScratchPersistenceWorkbookKey();
  if (!storageWbKey) {
    openScratchTabs = [];
    activeWorkspaceTab = DAILY_TRACKER_WS_TAB;
    return;
  }
  const map = getPersistedOpenScratchTabsMap();
  const savedList = Array.isArray(map[storageWbKey])
    ? map[storageWbKey]
    : (Array.isArray(map[workingFileHandle.name]) ? map[workingFileHandle.name] : []);
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

function renderTrackerProfilesBar() {
  if (!trackerProfilesBar || !trackerProfilesList) return;
  const profilesToRender = Array.isArray(trackerProfiles) && trackerProfiles.length > 0
    ? trackerProfiles
    : [{ id: DEFAULT_TRACKER_PROFILE_ID, name: 'Daily Tracker', mirrorEnabled: true }];

  trackerProfilesBar.style.display = 'flex';
  trackerProfilesList.innerHTML = '';

  profilesToRender.forEach(prof => {
    const isCurrentProf = prof.id === activeProfileId;

    let openCount = null;
    let isCustomDept = false;
    if (isCurrentProf && workbook) {
      openCount = allRows.filter(r => !isCompletedStatus(r.status)).length;
      isCustomDept = Boolean(activeSchemaConfig && !activeSchemaConfig.isStandardSchema);
    } else if (!isCurrentProf && trackerProfileRuntimeCache.has(prof.id)) {
      const snap = trackerProfileRuntimeCache.get(prof.id);
      if (snap && snap.workbook && Array.isArray(snap.allRows)) {
        openCount = snap.allRows.filter(r => !isCompletedStatus(r.status)).length;
        isCustomDept = Boolean(snap.activeSchemaConfig && !snap.activeSchemaConfig.isStandardSchema);
      }
    }

    const canCloseProfile = prof.id !== DEFAULT_TRACKER_PROFILE_ID;
    const countBadgeHtml = openCount !== null
      ? `<span class="ws-tab-pill-count" title="${openCount} open request(s)">${openCount}</span>`
      : '';
    const customTagHtml = isCustomDept
      ? `<span class="ws-tab-schema-tag" title="Dynamic custom column schema active">Custom</span>`
      : '';
    const closeBtnHtml = canCloseProfile
      ? `<span class="ws-tab-close-btn" data-close-profile-tab="${escapeHtml(prof.id)}" title="Remove ${escapeHtml(prof.name)} workbook tab (Excel files stay untouched)" aria-label="Remove ${escapeHtml(prof.name)} tab">×</span>`
      : '';

    const profBtn = document.createElement('button');
    profBtn.type = 'button';
    profBtn.className = `wb-profile-tab${isCurrentProf ? ' active' : ''}`;
    profBtn.setAttribute('role', 'tab');
    profBtn.setAttribute('aria-selected', String(isCurrentProf));
    profBtn.title = `${prof.name}${prof.mirrorEnabled === false ? ' (Mirror Off)' : ''} — Double-click to rename`;
    profBtn.innerHTML = `
      <svg class="ui-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
        <path d="M2.5 5.5A1.5 1.5 0 0 1 4 4h3.9a1.5 1.5 0 0 1 1.1.47l1.1 1.06a1.5 1.5 0 0 0 1.05.47H16a1.5 1.5 0 0 1 1.5 1.5v7A1.5 1.5 0 0 1 16 16H4a1.5 1.5 0 0 1-1.5-1.5v-9Z"></path>
      </svg>
      <span>${escapeHtml(prof.name)}</span>
      ${customTagHtml}
      ${countBadgeHtml}
      ${closeBtnHtml}
    `;

    profBtn.addEventListener('click', (e) => {
      const closeProfileEl = e.target.closest('[data-close-profile-tab]');
      if (closeProfileEl) {
        e.stopPropagation();
        if (typeof removeTrackerProfileInteractive === 'function') {
          removeTrackerProfileInteractive(prof.id);
        }
        return;
      }
      if (isCurrentProf) {
        if (activeWorkspaceTab !== DAILY_TRACKER_WS_TAB) {
          switchToWorkspaceTab(DAILY_TRACKER_WS_TAB);
        }
      } else if (typeof switchTrackerProfile === 'function') {
        switchTrackerProfile(prof.id);
      }
    });

    profBtn.addEventListener('dblclick', (e) => {
      if (e.target.closest('[data-close-profile-tab]')) return;
      if (typeof renameActiveTrackerProfileInteractive === 'function') {
        renameActiveTrackerProfileInteractive(prof.id);
      }
    });

    trackerProfilesList.appendChild(profBtn);
  });
}

function renderWorkspaceTabsBar() {
  renderTrackerProfilesBar();

  if (!workspaceTabsBar || !workspaceTabsList) return;
  if (!workbook) {
    workspaceTabsBar.style.display = 'none';
    return;
  }

  workspaceTabsBar.style.display = 'flex';
  if (scratchSheetMenuWrap) {
    scratchSheetMenuWrap.style.display = 'inline-flex';
  }
  renderScratchMenuDropdown();

  workspaceTabsList.innerHTML = '';

  // 1. Active Workbook's Primary Daily Requests Sheet View Tab
  const isTrackerViewActive = activeWorkspaceTab === DAILY_TRACKER_WS_TAB;
  const openCount = allRows.filter(r => !isCompletedStatus(r.status)).length;
  const trackerViewBtn = document.createElement('button');
  trackerViewBtn.type = 'button';
  trackerViewBtn.className = `ws-app-tab${isTrackerViewActive ? ' active' : ''}`;
  trackerViewBtn.setAttribute('role', 'tab');
  trackerViewBtn.setAttribute('aria-selected', String(isTrackerViewActive));
  trackerViewBtn.title = `View daily MM-DD-YYYY request tabs (${openCount} open request(s))`;
  trackerViewBtn.innerHTML = `
    <svg class="ui-icon" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
      <rect x="2.5" y="3" width="15" height="14" rx="2"></rect>
      <line x1="2.5" y1="7.5" x2="17.5" y2="7.5"></line>
      <line x1="7" y1="11.5" x2="13.5" y2="11.5"></line>
      <line x1="7" y1="14.2" x2="11.5" y2="14.2"></line>
    </svg>
    <span>Daily Requests</span>
    <span class="ws-tab-pill-count" title="${openCount} open request(s)">${openCount}</span>
  `;
  trackerViewBtn.addEventListener('click', () => {
    switchToWorkspaceTab(DAILY_TRACKER_WS_TAB);
  });
  workspaceTabsList.appendChild(trackerViewBtn);

  // 2. Open Custom / Scratch Sheet App Tabs (belonging to this active workbook)
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

function closeScratchFillDropdown() {
  if (scratchFillDropdown) {
    scratchFillDropdown.classList.remove('open');
  }
  if (btnScratchFillDropdownToggle) {
    btnScratchFillDropdownToggle.setAttribute('aria-expanded', 'false');
  }
}

function closeScratchContextMenu() {
  if (scratchContextMenu) {
    scratchContextMenu.style.display = 'none';
  }
}

function syncScratchStripedRowsUI() {
  if (chkScratchStripedRows) {
    chkScratchStripedRows.checked = scratchStripedRowsEnabled;
  }
  if (scratchGridTable) {
    scratchGridTable.classList.toggle('striped-rows', scratchStripedRowsEnabled);
  }
}

function syncScratchAutoIncrementUI() {
  if (chkScratchAutoIncrement) {
    chkScratchAutoIncrement.checked = scratchAutoIncrementEnabled;
  }
}

function syncWorkspaceTabsUI() {
  stopScratchPointerDrag();
  closeScratchFillDropdown();
  closeScratchContextMenu();
  syncScratchStripedRowsUI();
  syncScratchAutoIncrementUI();
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
  closeScratchFillDropdown();
  closeScratchContextMenu();
  commitActiveScratchInlineEdit();
  if (targetTab === activeWorkspaceTab) return;

  if (hasAnyPendingScratchMutations()) {
    flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: true });
  }

  activeWorkspaceTab = targetTab;
  syncWorkspaceTabsUI();
}

function setActiveWorkspaceTab(targetTab) {
  return switchToWorkspaceTab(targetTab);
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

// --- CELL DATA TYPE / FORMAT SELECTOR & LIVE FORMULA / SPILL ENGINE (v1.19.0) ---
function syncScratchFormatSelectUI() {
  if (!scratchCellFormatSelect || activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) return;
  const range = getScratchSelectionRange(activeWorkspaceTab);
  const fmtId = getScratchCellFormatId(ws, range.anchorRow, range.anchorCol);
  scratchCellFormatSelect.value = SCRATCH_FORMAT_DEFS[fmtId] ? fmtId : 'auto';
}

function applyScratchFormatToSelection(sheetName, fmtId) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB) return 0;
  const def = SCRATCH_FORMAT_DEFS[fmtId] || SCRATCH_FORMAT_DEFS.auto;
  commitActiveScratchInlineEdit();
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return 0;

  const range = getScratchSelectionRange(sheetName);
  const dims = getScratchSheetGridDimensions(ws, sheetName);
  const isFullColumnSelection = range.minRow === 1 && range.maxRow >= dims.rows && range.colCount < dims.cols;
  const effectiveMaxRow = isFullColumnSelection ? Math.max(dims.dataMaxRow, dims.rows) : range.maxRow;

  if (isFullColumnSelection) {
    for (let c = range.minCol; c <= range.maxCol; c++) {
      ws.getColumn(c).numFmt = def.numFmt || undefined;
    }
  }

  const edits = [];
  for (let r = range.minRow; r <= effectiveMaxRow; r++) {
    for (let c = range.minCol; c <= range.maxCol; c++) {
      const beforeSnap = captureCellStyleSnapshot(ws, r, c);
      const nextSnap = {
        fill: beforeSnap.fill,
        font: beforeSnap.font,
        numFmt: def.numFmt,
        alignment: beforeSnap.alignment || null
      };
      const info = extractScratchCellInfo(ws, r, c, { skipSpillLookup: true });
      const editObj = { r, c, styleSnap: nextSnap };
      if (!info.isFormula && info.rawEdit !== '' && !(isFullColumnSelection && r === 1)) {
        editObj.rawText = info.rawEdit;
      }
      edits.push(editObj);
    }
  }

  const count = applyScratchCellChanges(sheetName, edits, { recordUndo: true });
  syncScratchFormatSelectUI();
  if (count > 0) {
    showToast(`Set cell type to ${def.label} on ${count} cell${count === 1 ? '' : 's'}.`, 'info');
  }
  return count;
}

function executeScratchFillHandleDrag(sheetName, sourceRange, fillTargetRange, { ctrlKey = false } = {}) {
  if (!workbook || !sheetName || !sourceRange || !fillTargetRange) return 0;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return 0;

  const edits = [];
  if (fillTargetRange.axis === 'vertical') {
    const dir = fillTargetRange.dir;
    const targetCount = dir > 0
      ? (fillTargetRange.maxRow - sourceRange.maxRow)
      : (sourceRange.minRow - fillTargetRange.minRow);
    if (targetCount <= 0) return 0;

    for (let c = sourceRange.minCol; c <= sourceRange.maxCol; c++) {
      const sourceItems = [];
      for (let r = sourceRange.minRow; r <= sourceRange.maxRow; r++) {
        const info = extractScratchCellInfo(ws, r, c);
        sourceItems.push({
          rawEdit: info.rawEdit,
          display: info.display,
          isFormula: info.isFormula,
          formatId: getScratchCellFormatId(ws, r, c),
          styleSnap: captureCellStyleSnapshot(ws, r, c)
        });
      }
      const generated = buildFillSeriesValues(sourceItems, targetCount, dir, {
        axis: 'vertical',
        ctrlKey,
        autoIncrement: scratchAutoIncrementEnabled
      });
      for (let m = 0; m < generated.length; m++) {
        const targetRow = dir > 0 ? (sourceRange.maxRow + 1 + m) : (sourceRange.minRow - 1 - m);
        edits.push({ r: targetRow, c, rawText: generated[m].rawText, styleSnap: generated[m].styleSnap });
      }
    }
  } else {
    const dir = fillTargetRange.dir;
    const targetCount = dir > 0
      ? (fillTargetRange.maxCol - sourceRange.maxCol)
      : (sourceRange.minCol - fillTargetRange.minCol);
    if (targetCount <= 0) return 0;

    for (let r = sourceRange.minRow; r <= sourceRange.maxRow; r++) {
      const sourceItems = [];
      for (let c = sourceRange.minCol; c <= sourceRange.maxCol; c++) {
        const info = extractScratchCellInfo(ws, r, c);
        sourceItems.push({
          rawEdit: info.rawEdit,
          display: info.display,
          isFormula: info.isFormula,
          formatId: getScratchCellFormatId(ws, r, c),
          styleSnap: captureCellStyleSnapshot(ws, r, c)
        });
      }
      const generated = buildFillSeriesValues(sourceItems, targetCount, dir, {
        axis: 'horizontal',
        ctrlKey,
        autoIncrement: scratchAutoIncrementEnabled
      });
      for (let m = 0; m < generated.length; m++) {
        const targetCol = dir > 0 ? (sourceRange.maxCol + 1 + m) : (sourceRange.minCol - 1 - m);
        edits.push({ r, c: targetCol, rawText: generated[m].rawText, styleSnap: generated[m].styleSnap });
      }
    }
  }

  if (edits.length === 0) return 0;
  scratchClipboardRangeState = null;
  const nextSelection = {
    anchorRow: fillTargetRange.minRow,
    anchorCol: fillTargetRange.minCol,
    focusRow: fillTargetRange.maxRow,
    focusCol: fillTargetRange.maxCol
  };
  return applyScratchCellChanges(sheetName, edits, {
    recordUndo: true,
    selectionAfter: nextSelection
  });
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
  const fillPreview =
    scratchDragState && scratchDragState.mode === 'fill' && scratchDragState.sheetName === activeWorkspaceTab
      ? scratchDragState.fillTargetRange
      : null;

  const prevClasses = [
    'selected-cell',
    'in-range-cell',
    'fill-handle-host',
    'range-edge-top',
    'range-edge-bottom',
    'range-edge-left',
    'range-edge-right',
    'copied-edge-top',
    'copied-edge-bottom',
    'copied-edge-left',
    'copied-edge-right',
    'in-fill-preview-cell',
    'fill-edge-top',
    'fill-edge-bottom',
    'fill-edge-left',
    'fill-edge-right'
  ];
  const selector = prevClasses.map(cls => `.${cls}`).join(',');
  scratchGridBody.querySelectorAll(selector).forEach(el => {
    el.classList.remove(...prevClasses);
  });
  scratchGridBody.querySelectorAll('.scratch-fill-handle').forEach(el => el.remove());
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

    if (fillPreview && r >= fillPreview.minRow && r <= fillPreview.maxRow) {
      const maxFillCol = Math.min(fillPreview.maxCol, tr.cells.length - 1);
      for (let c = fillPreview.minCol; c <= maxFillCol; c++) {
        const td = tr.cells[c];
        if (!td) continue;
        const inSrc = r >= range.minRow && r <= range.maxRow && c >= range.minCol && c <= range.maxCol;
        if (!inSrc) {
          td.classList.add('in-fill-preview-cell');
          if (r === fillPreview.minRow) td.classList.add('fill-edge-top');
          if (r === fillPreview.maxRow) td.classList.add('fill-edge-bottom');
          if (c === fillPreview.minCol) td.classList.add('fill-edge-left');
          if (c === fillPreview.maxCol) td.classList.add('fill-edge-right');
        }
      }
    }
  }

  // Attach Excel-style bottom-right corner Fill Handle square when not inline-editing
  if (!scratchInlineEditState && !fillPreview) {
    const cornerTd = scratchGridBody.querySelector(`td[data-r="${range.maxRow}"][data-c="${range.maxCol}"]`);
    if (cornerTd) {
      cornerTd.classList.add('fill-handle-host');
      const handleEl = document.createElement('span');
      handleEl.className = 'scratch-fill-handle';
      handleEl.dataset.fillHandle = 'true';
      handleEl.title = 'Drag corner square to copy/fill • Dates & Numbers auto-increment (hold Ctrl to copy without incrementing)';
      cornerTd.appendChild(handleEl);
    }
  }

  updateScratchSelectionStatsUI();
  syncScratchFormatSelectUI();
  syncScratchTextFormattingUI();
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

function commitActiveScratchInlineEdit({ restoreGridFocus = false } = {}) {
  if (!scratchInlineEditState) return;
  const state = scratchInlineEditState;
  const inputHadFocus = Boolean(state.inputEl && document.activeElement === state.inputEl);
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

  if (
    (restoreGridFocus || inputHadFocus) &&
    activeWorkspaceTab !== DAILY_TRACKER_WS_TAB &&
    scratchGridWrap &&
    (!document.activeElement || document.activeElement === document.body)
  ) {
    scratchGridWrap.focus({ preventScroll: true });
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
  if (
    activeWorkspaceTab !== DAILY_TRACKER_WS_TAB &&
    scratchGridWrap &&
    (!document.activeElement || document.activeElement === document.body)
  ) {
    scratchGridWrap.focus({ preventScroll: true });
  }
}

function isUserActivelyInteractingWithScratch() {
  if (scratchInlineEditState) return true;
  if (scratchDragState || scratchColResizeState) return true;
  if (scratchFormulaInput && document.activeElement === scratchFormulaInput) return true;
  return false;
}

function scheduleScratchAutoSave(delayMs = 1800) {
  if (typeof isLockedByOtherTab !== 'undefined' && isLockedByOtherTab) {
    return;
  }
  if (scratchAutoSaveTimer) {
    clearTimeout(scratchAutoSaveTimer);
  }
  scratchAutoSaveTimer = setTimeout(() => {
    scratchAutoSaveTimer = null;
    if (typeof isLockedByOtherTab !== 'undefined' && isLockedByOtherTab) {
      return;
    }
    if (isUserActivelyInteractingWithScratch()) {
      // Never steal focus or commit mid-typing; wait until the user finishes editing the cell
      scheduleScratchAutoSave(1500);
      return;
    }
    flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: true, commitInline: false });
  }, delayMs);
}

async function flushPendingScratchEditsToDisk({
  reloadAfter = false,
  silentToast = false,
  commitInline = true
} = {}) {
  if (typeof isLockedByOtherTab !== 'undefined' && isLockedByOtherTab && silentToast) {
    return false;
  }
  if (commitInline) {
    if (scratchAutoSaveTimer) {
      clearTimeout(scratchAutoSaveTimer);
      scratchAutoSaveTimer = null;
    }
    commitActiveScratchInlineEdit({ restoreGridFocus: true });
  } else if (isUserActivelyInteractingWithScratch()) {
    scheduleScratchAutoSave(1500);
    return false;
  }

  if (!workbook || !workingFileHandle) return false;
  if (!hasAnyPendingScratchMutations() && silentToast) return true;
  if (scratchIsSaving) {
    if (hasAnyPendingScratchMutations()) {
      scheduleScratchAutoSave(1500);
    }
    return false;
  }

  const versionAtSaveStart = scratchMutationVersion;
  scratchIsSaving = true;
  updateScratchSaveBadgeUI();

  const saved = await saveWorkbookToDisk({ silentLockAlert: silentToast });
  scratchIsSaving = false;

  // Restore grid keyboard focus if browser dropped activeElement to body during async file write
  if (
    activeWorkspaceTab !== DAILY_TRACKER_WS_TAB &&
    !scratchInlineEditState &&
    scratchGridWrap &&
    (!document.activeElement || document.activeElement === document.body)
  ) {
    scratchGridWrap.focus({ preventScroll: true });
  }

  if (saved) {
    if (scratchMutationVersion === versionAtSaveStart) {
      scratchPendingEdits.clear();
      scratchPendingFills.clear();
      scratchPendingColWidths.clear();
    } else {
      // User made additional edits while disk I/O was in flight; queue follow-up auto-save
      scheduleScratchAutoSave(1800);
    }
    updateScratchSaveBadgeUI();
    renderWorkspaceTabsBar();
    if (reloadAfter && saved !== 'staged-primary') {
      await loadWorkbookFromDisk(sheetSelect.value, {
        preserveStatus: saved === 'mirror-locked',
        skipFlushPendingScratch: true
      });
    }
    if (!silentToast && saved !== 'staged-primary') {
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
  const visual = extractScratchCellVisualStyle(ws, rNum, cNum);
  const anchorSpills = scratchSpillStateMap.get(sheetName);
  const isArrayAnchor = Boolean(anchorSpills && anchorSpills.has(`${rNum}::${cNum}`));
  td.classList.toggle('formula-cell', info.isFormula);
  td.classList.toggle('formula-anchor-cell', isArrayAnchor);
  td.classList.toggle('spilled-array-cell', Boolean(info.isSpilled && info.anchorR > 0));
  applyVisualStyleToScratchTd(td, visual);
  if (info.isFormula) {
    td.title = `${info.rawEdit} → ${info.display}`;
  } else if (info.isSpilled && info.anchorR > 0) {
    td.title = `${info.display} (Spilled from ${getCellCoordString(info.anchorR, info.anchorC)})`;
  } else {
    td.title = info.display;
  }

  let span = td.querySelector('.scratch-cell-text');
  if (!span) {
    span = document.createElement('span');
    span.className = 'scratch-cell-text';
    td.innerHTML = '';
    td.appendChild(span);
  }
  span.textContent = info.display;
}

// --- BATCH CELL MUTATION ENGINE WITH 50-STEP UNDO / REDO (VALUES & EXCEL FILL COLORS) ---
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
  if (!scratchPendingFills.has(sheetName)) {
    scratchPendingFills.set(sheetName, new Map());
  }
  const pendingSheetMap = scratchPendingEdits.get(sheetName);
  const pendingFillMap = scratchPendingFills.get(sheetName);

  editsList.forEach(item => {
    const r = Math.max(1, Math.floor(item.r));
    const c = Math.max(1, Math.floor(item.c));
    const hasTextProp = Object.prototype.hasOwnProperty.call(item, 'rawText');
    const hasStyleProp = Object.prototype.hasOwnProperty.call(item, 'styleSnap');

    if (r > maxTargetRow) maxTargetRow = r;
    if (c > maxTargetCol) maxTargetCol = c;

    let textChanged = false;
    let beforeRaw = '';
    let afterRaw = '';
    if (hasTextProp) {
      beforeRaw = extractScratchCellInfo(ws, r, c, { skipSpillLookup: true }).rawEdit;
      afterRaw = String(item.rawText ?? '').replace(/\r/g, '');
      textChanged = beforeRaw !== afterRaw;
    }

    let styleChanged = false;
    let beforeStyle = null;
    let afterStyle = null;
    if (hasStyleProp) {
      beforeStyle = captureCellStyleSnapshot(ws, r, c);
      afterStyle = normalizeScratchStyleSnapshot(item.styleSnap);
      styleChanged = JSON.stringify(beforeStyle) !== JSON.stringify(afterStyle);
    }

    if (!textChanged && !styleChanged) return;

    // Apply style/format snapshot first so applyRawEditToExcelCell respects updated numFmt
    if (styleChanged) {
      applyCellStyleSnapshotToExcelCell(ws, r, c, afterStyle);
      pendingFillMap.set(`${r}::${c}`, afterStyle);
      if (r === 1) {
        setPersistedRow1UnboldCol(sheetName, c, Boolean(afterStyle && afterStyle.font && afterStyle.font.bold === false));
      }
    }
    if (textChanged) {
      applyRawEditToExcelCell(ws, r, c, afterRaw);
      pendingSheetMap.set(`${r}::${c}`, afterRaw);
    }

    recordedChanges.push({
      r,
      c,
      textChanged,
      beforeRaw,
      afterRaw,
      styleChanged,
      beforeStyle,
      afterStyle
    });
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
      recordedChanges[0].textChanged &&
      !recordedChanges[0].styleChanged &&
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

  // Re-evaluate any =FILTER / =SORT / =UNIQUE / summary formulas affected by these edits
  const formulaAffectedCoords = recalculateScratchWorksheetFormulas(ws, sheetName);
  formulaAffectedCoords.forEach(({ r, c }) => {
    if (r > maxTargetRow) maxTargetRow = r;
    if (c > maxTargetCol) maxTargetCol = c;
  });

  // Update cached non-empty row count & data bounds metadata (including live spills)
  const bounds = measureScratchWorksheetBounds(ws);
  if (bounds.maxRow > maxTargetRow) maxTargetRow = bounds.maxRow;
  if (bounds.maxCol > maxTargetCol) maxTargetCol = bounds.maxCol;

  // Update grid dimensions if paste/edit/spill exceeded previously rendered/cached bounds
  const prevRenderedRows = scratchGridBody ? scratchGridBody.rows.length : 0;
  const prevRenderedCols = (scratchGridHead && scratchGridHead.rows && scratchGridHead.rows[0])
    ? Math.max(0, scratchGridHead.rows[0].cells.length - 1)
    : 0;
  const prevCachedDims = scratchGridDimensionsMap.get(sheetName) || {
    rows: prevRenderedRows,
    cols: prevRenderedCols
  };
  const curDims = getScratchSheetGridDimensions(ws, sheetName);
  const targetRows = Math.max(curDims.rows, maxTargetRow + 5);
  const targetCols = Math.max(curDims.cols, maxTargetCol + 2);
  if (targetRows > curDims.rows || targetCols > curDims.cols) {
    scratchGridDimensionsMap.set(sheetName, {
      rows: targetRows,
      cols: targetCols
    });
  }
  const needsExpand =
    targetRows > prevRenderedRows ||
    targetCols > prevRenderedCols ||
    targetRows > prevCachedDims.rows ||
    targetCols > prevCachedDims.cols ||
    maxTargetRow > prevRenderedRows ||
    maxTargetCol > prevRenderedCols;
  const touchedRow1 = recordedChanges.some(ch => ch.r === 1);

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
    if (needsExpand || touchedRow1 || (recordedChanges.length + formulaAffectedCoords.length) > 40 || hasActiveSearch) {
      renderActiveScratchSheetGrid();
    } else {
      const toRefresh = new Map();
      recordedChanges.forEach(ch => toRefresh.set(`${ch.r}::${ch.c}`, { r: ch.r, c: ch.c }));
      formulaAffectedCoords.forEach(pt => toRefresh.set(`${pt.r}::${pt.c}`, pt));
      toRefresh.forEach(pt => refreshSingleScratchCellDOM(sheetName, pt.r, pt.c));
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
  scratchMutationVersion++;
  if (typeof window !== 'undefined') {
    window.__lastTrackerScratchEditAt = Date.now();
  }
  scheduleScratchAutoSave();
  return recordedChanges.length;
}

function recordScratchCellEdit(sheetName, rNum, cNum, rawText, { coalesceKey = null } = {}) {
  scratchClipboardRangeState = null;
  applyScratchCellChanges(sheetName, [{ r: rNum, c: cNum, rawText }], {
    recordUndo: true,
    coalesceKey
  });
}

function applyScratchFillPresetToSelection(sheetName, presetId) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB) return 0;
  commitActiveScratchInlineEdit();
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return 0;

  const range = getScratchSelectionRange(sheetName);
  const dims = getScratchSheetGridDimensions(ws, sheetName);
  // If entire columns were highlighted via header click, cap row fill to active data bounds (or at least row 1)
  // unless user explicitly has empty cells below they want filled; wait: if full column selected (rows === dims.rows),
  // capping to Math.max(range.anchorRow, dims.dataMaxRow, 1) prevents coloring 40 blank rows unintentionally unless whole column has data,
  // OR coloring exact range if range.rowCount < dims.rows! Let's check:
  const isFullColumnSelection = range.minRow === 1 && range.maxRow >= dims.rows && range.colCount < dims.cols;
  const effectiveMaxRow = isFullColumnSelection
    ? Math.max(1, dims.dataMaxRow)
    : range.maxRow;

  const edits = [];
  for (let r = range.minRow; r <= effectiveMaxRow; r++) {
    for (let c = range.minCol; c <= range.maxCol; c++) {
      const beforeSnap = captureCellStyleSnapshot(ws, r, c);
      const nextSnap = buildStyleSnapshotForFillPreset(beforeSnap, presetId);
      edits.push({ r, c, styleSnap: nextSnap });
    }
  }

  if (presetId && presetId !== 'none') {
    activeScratchFillPresetId = presetId;
    try {
      localStorage.setItem(SCRATCH_LAST_FILL_STORAGE_KEY, presetId);
    } catch (_) { /* ignore */ }
    syncScratchFillIndicatorUI();
  }

  const count = applyScratchCellChanges(sheetName, edits, { recordUndo: true });
  if (count > 0) {
    const presetObj = EXCEL_DEFAULT_FILL_PALETTE.find(p => p.id === presetId);
    const label = presetId === 'none' ? 'Cleared fill color' : `Applied "${presetObj ? presetObj.label : 'Fill'}"`;
    if (count > 1) {
      showToast(`${label} on ${count} cells (${getRangeCoordString(range)}).`, 'info');
    }
  }
  return count;
}

function syncScratchTextFormattingUI() {
  const isScratch = activeWorkspaceTab !== DAILY_TRACKER_WS_TAB && Boolean(workbook);
  const ws = isScratch ? workbook.getWorksheet(activeWorkspaceTab) : null;
  const syncToggleBtn = (btn, active, disabled) => {
    if (!btn) return;
    btn.disabled = Boolean(disabled);
    btn.classList.toggle('active', !disabled && Boolean(active));
    btn.setAttribute('aria-pressed', String(!disabled && Boolean(active)));
  };

  if (!ws) {
    [
      btnScratchFormatBold,
      btnScratchFormatItalic,
      btnScratchFormatUnderline,
      btnScratchAlignLeft,
      btnScratchAlignCenter,
      btnScratchAlignRight,
      btnScratchClearTextFormat
    ].forEach(btn => syncToggleBtn(btn, false, true));
    if (scratchContextMenu) {
      scratchContextMenu.querySelectorAll('.scratch-ctx-fmt-btn').forEach(btn => {
        btn.classList.remove('active');
        btn.setAttribute('aria-pressed', 'false');
      });
    }
    return;
  }

  const range = getScratchSelectionRange(activeWorkspaceTab);
  const visual = extractScratchCellVisualStyle(ws, range.anchorRow, range.anchorCol);
  let row1HasValues = false;
  if (range.anchorRow === 1) {
    const scanCols = Math.max(ws.columnCount || 0, range.anchorCol, 10);
    for (let c = 1; c <= scanCols; c++) {
      if (extractScratchCellInfo(ws, 1, c).display.trim() !== '') {
        row1HasValues = true;
        break;
      }
    }
  }
  const effectiveBold = visual.isExplicitUnbold ? false : (visual.isBold || (range.anchorRow === 1 && row1HasValues));

  syncToggleBtn(btnScratchFormatBold, effectiveBold, false);
  syncToggleBtn(btnScratchFormatItalic, visual.isItalic, false);
  syncToggleBtn(btnScratchFormatUnderline, visual.isUnderline, false);
  syncToggleBtn(btnScratchAlignLeft, visual.align === 'left', false);
  syncToggleBtn(btnScratchAlignCenter, visual.align === 'center', false);
  syncToggleBtn(btnScratchAlignRight, visual.align === 'right', false);
  if (btnScratchClearTextFormat) {
    btnScratchClearTextFormat.disabled = false;
  }

  if (scratchContextMenu) {
    const ctxStates = {
      'toggle-bold': effectiveBold,
      'toggle-italic': visual.isItalic,
      'toggle-underline': visual.isUnderline,
      'align-left': visual.align === 'left',
      'align-center': visual.align === 'center',
      'align-right': visual.align === 'right'
    };
    scratchContextMenu.querySelectorAll('.scratch-ctx-fmt-btn[data-ctx-action]').forEach(btn => {
      const act = btn.dataset.ctxAction;
      const isAct = Boolean(ctxStates[act]);
      btn.classList.toggle('active', isAct);
      btn.setAttribute('aria-pressed', String(isAct));
    });
  }
}

function applyScratchTextFormattingToSelection(sheetName, action) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB) return 0;
  commitActiveScratchInlineEdit();
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return 0;

  const range = getScratchSelectionRange(sheetName);
  const dims = getScratchSheetGridDimensions(ws, sheetName);
  const isWholeSheetSelection =
    range.minRow === 1 && range.maxRow >= dims.rows && range.minCol === 1 && range.maxCol >= dims.cols;
  const isFullColumnSelection = !isWholeSheetSelection && range.minRow === 1 && range.maxRow >= dims.rows && range.colCount < dims.cols;
  const isFullRowSelection = !isWholeSheetSelection && range.minCol === 1 && range.maxCol >= dims.cols && range.rowCount < dims.rows;

  const effectiveMaxRow = (isWholeSheetSelection || isFullColumnSelection)
    ? Math.max(1, dims.dataMaxRow, range.anchorRow)
    : range.maxRow;
  const effectiveMaxCol = (isWholeSheetSelection || isFullRowSelection)
    ? Math.max(1, dims.dataMaxCol, range.anchorCol)
    : range.maxCol;

  const anchorVisual = extractScratchCellVisualStyle(ws, range.anchorRow, range.anchorCol);
  let row1HasValues = false;
  if (range.minRow === 1) {
    const scanCols = Math.max(dims.cols, dims.dataMaxCol, effectiveMaxCol, 10);
    for (let c = 1; c <= scanCols; c++) {
      if (extractScratchCellInfo(ws, 1, c).display.trim() !== '') {
        row1HasValues = true;
        break;
      }
    }
  }
  const anchorEffectiveBold = anchorVisual.isExplicitUnbold
    ? false
    : (anchorVisual.isBold || (range.anchorRow === 1 && row1HasValues));

  const patchBase = {};
  let toastLabel = '';
  switch (action) {
    case 'toggle-bold':
      patchBase.bold = !anchorEffectiveBold;
      toastLabel = patchBase.bold ? 'Applied Bold' : 'Unbolded';
      break;
    case 'toggle-italic':
      patchBase.italic = !anchorVisual.isItalic;
      toastLabel = patchBase.italic ? 'Applied Italic' : 'Removed Italic';
      break;
    case 'toggle-underline':
      patchBase.underline = !anchorVisual.isUnderline;
      toastLabel = patchBase.underline ? 'Applied Underline' : 'Removed Underline';
      break;
    case 'align-left':
      patchBase.align = anchorVisual.align === 'left' ? null : 'left';
      toastLabel = patchBase.align ? 'Aligned Left' : 'Reset Alignment';
      break;
    case 'align-center':
      patchBase.align = anchorVisual.align === 'center' ? null : 'center';
      toastLabel = patchBase.align ? 'Aligned Center' : 'Reset Alignment';
      break;
    case 'align-right':
      patchBase.align = anchorVisual.align === 'right' ? null : 'right';
      toastLabel = patchBase.align ? 'Aligned Right' : 'Reset Alignment';
      break;
    case 'clear-text-format':
      patchBase.clearFormatting = true;
      toastLabel = 'Cleared text formatting / unbolded';
      break;
    default:
      return 0;
  }

  const edits = [];
  for (let r = range.minRow; r <= effectiveMaxRow; r++) {
    for (let c = range.minCol; c <= effectiveMaxCol; c++) {
      const beforeSnap = captureCellStyleSnapshot(ws, r, c);
      const cellPatch = {
        ...patchBase,
        explicitUnbold: r === 1 && (patchBase.bold === false || Boolean(patchBase.clearFormatting))
      };
      const nextSnap = buildStyleSnapshotForTextFormatting(beforeSnap, cellPatch);
      edits.push({ r, c, styleSnap: nextSnap });
    }
  }

  const count = applyScratchCellChanges(sheetName, edits, { recordUndo: true });
  syncScratchTextFormattingUI();
  if (count > 0 && toastLabel && (count > 1 || action === 'clear-text-format')) {
    showToast(`${toastLabel} on ${count} cell${count === 1 ? '' : 's'} (${getRangeCoordString(range)}).`, 'info');
  }
  return count;
}

function undoScratchEdit() {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  commitActiveScratchInlineEdit();
  const sheetName = activeWorkspaceTab;
  const undoStack = scratchUndoStackMap.get(sheetName);
  if (!undoStack || undoStack.length === 0) return;

  const entry = undoStack.pop();
  lastFormulaCoalesceMeta = null;

  const reverseEdits = entry.changes.map(ch => {
    const rev = { r: ch.r, c: ch.c };
    if (ch.textChanged || ch.textChanged === undefined) {
      rev.rawText = ch.beforeRaw;
    }
    if (ch.styleChanged) {
      rev.styleSnap = ch.beforeStyle;
    }
    return rev;
  });

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

  const forwardEdits = entry.changes.map(ch => {
    const fwd = { r: ch.r, c: ch.c };
    if (ch.textChanged || ch.textChanged === undefined) {
      fwd.rawText = ch.afterRaw;
    }
    if (ch.styleChanged) {
      fwd.styleSnap = ch.afterStyle;
    }
    return fwd;
  });

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
  scratchClipboardRangeState = null;
  commitActiveScratchInlineEdit();
  selectScratchCell(rNum, cNum);

  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) return;
  const td = scratchGridBody.querySelector(`td[data-r="${rNum}"][data-c="${cNum}"]`);
  if (!td) return;

  const info = extractScratchCellInfo(ws, rNum, cNum);
  const visual = extractScratchCellVisualStyle(ws, rNum, cNum);
  const input = document.createElement('input');
  input.type = 'text';
  input.className = 'scratch-inline-editor';
  input.value = initialChar !== null ? initialChar : (info.isSpilled ? '' : info.rawEdit);
  if (visual.isBold) input.style.fontWeight = '700';
  else if (visual.isExplicitUnbold) input.style.fontWeight = '400';
  if (visual.isItalic) input.style.fontStyle = 'italic';
  if (visual.isUnderline) input.style.textDecoration = 'underline';
  if (visual.align) input.style.textAlign = visual.align;

  scratchInlineEditState = {
    sheetName: activeWorkspaceTab,
    row: rNum,
    col: cNum,
    originalRaw: info.rawEdit,
    inputEl: input,
    tdEl: td
  };

  // Defer any pending background auto-save as soon as the user starts editing another cell
  if (scratchAutoSaveTimer) {
    scheduleScratchAutoSave(1800);
  }

  if (scratchFormulaInput) {
    scratchFormulaInput.value = input.value;
  }

  input.addEventListener('input', () => {
    if (scratchFormulaInput) {
      scratchFormulaInput.value = input.value;
    }
    if (scratchAutoSaveTimer) {
      scheduleScratchAutoSave(1800);
    }
  });

  input.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && !e.altKey) {
      const lower = e.key.toLowerCase();
      const fmtMap = { b: 'toggle-bold', i: 'toggle-italic', u: 'toggle-underline' };
      if (fmtMap[lower]) {
        e.preventDefault();
        e.stopPropagation();
        const curVal = input.value;
        const selStart = input.selectionStart;
        const selEnd = input.selectionEnd;
        const editSheet = activeWorkspaceTab;
        commitActiveScratchInlineEdit({ restoreGridFocus: false });
        applyScratchTextFormattingToSelection(editSheet, fmtMap[lower]);
        startScratchInlineEdit(rNum, cNum);
        if (scratchInlineEditState && scratchInlineEditState.inputEl) {
          scratchInlineEditState.inputEl.value = curVal;
          if (scratchFormulaInput) scratchFormulaInput.value = curVal;
          try {
            if (typeof selStart === 'number' && typeof selEnd === 'number') {
              scratchInlineEditState.inputEl.setSelectionRange(selStart, selEnd);
            }
          } catch (_) { /* ignore */ }
        }
        return;
      }
    }
    if (e.key === 'Enter') {
      e.preventDefault();
      e.stopPropagation();
      commitActiveScratchInlineEdit({ restoreGridFocus: true });
      selectScratchCell(e.shiftKey ? Math.max(1, rNum - 1) : rNum + 1, cNum, { scrollIntoView: true });
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
    } else if (e.key === 'Tab') {
      e.preventDefault();
      e.stopPropagation();
      commitActiveScratchInlineEdit({ restoreGridFocus: true });
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
  if (scratchInlineEditState) {
    commitActiveScratchInlineEdit();
  }
  const ws = workbook.getWorksheet(activeWorkspaceTab);
  if (!ws) {
    activeWorkspaceTab = DAILY_TRACKER_WS_TAB;
    syncWorkspaceTabsUI();
    return;
  }

  const prevScrollTop = scratchGridWrap ? scratchGridWrap.scrollTop : 0;
  const prevScrollLeft = scratchGridWrap ? scratchGridWrap.scrollLeft : 0;

  applyPersistedRow1UnboldToWorksheet(ws, activeWorkspaceTab);
  recalculateScratchWorksheetFormulas(ws, activeWorkspaceTab);
  const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
  const anchorSpills = scratchSpillStateMap.get(activeWorkspaceTab);
  const range = getScratchSelectionRange(activeWorkspaceTab);
  const query = scratchSearchInput ? scratchSearchInput.value.trim().toLowerCase() : '';
  const filterRowsOnly = Boolean(chkScratchFilterRows && chkScratchFilterRows.checked && query.length > 0);

  if (scratchActiveSheetBadge) {
    scratchActiveSheetBadge.textContent = ws.name;
  }
  if (scratchStatsMeta) {
    scratchStatsMeta.textContent = `${dims.nonEmptyRows} active row${dims.nonEmptyRows === 1 ? '' : 's'} (${dims.rows} × ${dims.cols} grid)`;
  }
  syncScratchStripedRowsUI();
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
      const visual = extractScratchCellVisualStyle(ws, r, c);
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
      if (visual.isBold) classes.push('bold-cell');
      if (visual.isExplicitUnbold) classes.push('unbold-cell');
      if (visual.isItalic) classes.push('italic-cell');
      if (visual.isUnderline) classes.push('underline-cell');
      if (visual.bgHex) classes.push('has-custom-fill');
      else if (visual.fgHex) classes.push('has-custom-fg');
      if (info.isFormula) classes.push('formula-cell');
      if (anchorSpills && anchorSpills.has(`${r}::${c}`)) classes.push('formula-anchor-cell');
      if (info.isSpilled && info.anchorR > 0) classes.push('spilled-array-cell');
      if (isHit) classes.push('search-hit');

      const styleParts = [];
      if (visual.bgHex) {
        styleParts.push(`--cell-fill-bg:${visual.bgHex}`, `--cell-fill-fg:${visual.fgHex || getContrastingTextHexForBg(visual.bgHex)}`);
      } else if (visual.fgHex) {
        styleParts.push(`--cell-fill-fg:${visual.fgHex}`);
      }
      if (visual.isItalic) styleParts.push('font-style:italic');
      if (visual.isUnderline) styleParts.push('text-decoration:underline');
      if (visual.align) styleParts.push(`text-align:${visual.align}`);
      const styleAttr = styleParts.length > 0 ? ` style="${styleParts.join('; ')};"` : '';
      const titleText = info.isFormula
        ? `${info.rawEdit} → ${info.display}`
        : (info.isSpilled && info.anchorR > 0 ? `${info.display} (Spilled from ${getCellCoordString(info.anchorR, info.anchorC)})` : info.display);
      rowCellsHtml.push(
        `<td data-r="${r}" data-c="${c}" class="${classes.join(' ')}"${styleAttr} title="${escapeHtml(titleText)}"><span class="scratch-cell-text">${escapeHtml(info.display)}</span></td>`
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

  if (scratchGridWrap) {
    scratchGridWrap.scrollTop = prevScrollTop;
    scratchGridWrap.scrollLeft = prevScrollLeft;
  }

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

  if (scratchDragState.mode === 'fill') {
    const probeX = Math.max(rect.left + rowHdrW + 3, Math.min(rect.right - 3, clientX));
    const probeY = Math.max(rect.top + headerH + 3, Math.min(rect.bottom - 3, clientY));
    const hit = document.elementFromPoint(probeX, probeY);
    if (!hit) return;
    const td = hit.closest('td[data-r][data-c]');
    if (!td) return;
    const r = Number(td.dataset.r);
    const c = Number(td.dataset.c);
    if (r >= 1 && c >= 1) {
      const nextTarget = computeScratchFillTargetRange(scratchDragState.sourceRange, r, c);
      const prevKey = scratchDragState.fillTargetRange ? JSON.stringify(scratchDragState.fillTargetRange) : '';
      const nextKey = nextTarget ? JSON.stringify(nextTarget) : '';
      if (prevKey !== nextKey) {
        scratchDragState.fillTargetRange = nextTarget;
        applyScratchSelectionDOM();
      }
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

function startScratchPointerDrag(mode, clientX, clientY, extraState = null) {
  stopScratchPointerDrag();
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB) return;
  scratchDragState = {
    mode,
    sheetName: activeWorkspaceTab,
    clientX,
    clientY,
    rafId: null,
    lastExpandAt: 0,
    ...(extraState || {})
  };
  if (scratchGridWrap) {
    scratchGridWrap.classList.add('is-drag-selecting');
    if (mode === 'fill') {
      scratchGridWrap.classList.add('is-fill-dragging');
    }
  }
  scratchDragState.rafId = requestAnimationFrame(runScratchDragAutoScrollTick);
}

function stopScratchPointerDrag() {
  const wasFill = Boolean(scratchDragState && scratchDragState.mode === 'fill');
  if (scratchDragState) {
    if (scratchDragState.rafId) {
      cancelAnimationFrame(scratchDragState.rafId);
    }
    scratchDragState = null;
  }
  if (scratchGridWrap) {
    scratchGridWrap.classList.remove('is-drag-selecting');
    scratchGridWrap.classList.remove('is-fill-dragging');
  }
  if (wasFill) {
    applyScratchSelectionDOM();
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
  scratchMutationVersion++;
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

function resolveEffectiveScratchCopyRange(ws, range) {
  if (!ws || !range) return range;
  const dims = getScratchSheetGridDimensions(ws, ws.name || activeWorkspaceTab);
  const fullHeight = range.minRow === 1 && range.maxRow >= dims.rows;
  const fullWidth = range.minCol === 1 && range.maxCol >= dims.cols;
  if (!fullHeight && !fullWidth) return range;

  let lastRow = range.minRow;
  let lastCol = range.minCol;
  for (let r = range.minRow; r <= range.maxRow; r++) {
    for (let c = range.minCol; c <= range.maxCol; c++) {
      const info = extractScratchCellInfo(ws, r, c);
      const visual = extractScratchCellVisualStyle(ws, r, c);
      if (info.display.trim() !== '' || visual.bgHex || visual.isBold) {
        if (r > lastRow) lastRow = r;
        if (c > lastCol) lastCol = c;
      }
    }
  }
  return {
    ...range,
    maxRow: fullHeight ? Math.max(range.minRow, lastRow) : range.maxRow,
    maxCol: fullWidth ? Math.max(range.minCol, lastCol) : range.maxCol
  };
}

function clearScratchSelectionCells(sheetName, range) {
  if (!workbook) return 0;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return 0;
  scratchClipboardRangeState = null;
  const edits = [];
  for (let r = range.minRow; r <= range.maxRow; r++) {
    for (let c = range.minCol; c <= range.maxCol; c++) {
      edits.push({ r, c, rawText: '' });
    }
  }
  return applyScratchCellChanges(sheetName, edits, { recordUndo: true });
}

function performScratchCopyOrCut(sheetName, range, { isCut = false, writeSystemClipboard = true } = {}) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB) return '';
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return '';

  const effRange = isCut ? range : resolveEffectiveScratchCopyRange(ws, range);
  const tsv = serializeScratchSelectionToTSV(ws, effRange);
  const html = serializeScratchSelectionToExcelHtml(ws, effRange);
  const styleMatrix = buildScratchStyleMatrix(ws, effRange);
  const rawEditMatrix = buildScratchRawEditMatrix(ws, effRange);
  scratchInternalClipboardText = tsv;
  scratchInternalClipboardState = {
    tsv,
    html,
    styleMatrix,
    rawEditMatrix,
    sourceRange: {
      minRow: effRange.minRow,
      maxRow: effRange.maxRow,
      minCol: effRange.minCol,
      maxCol: effRange.maxCol
    }
  };

  if (writeSystemClipboard) {
    writeScratchClipboardPayloadToSystem(tsv, html);
  }

  if (isCut) {
    scratchClipboardRangeState = null;
    clearScratchSelectionCells(sheetName, range);
    syncScratchFormulaBarFromActiveCell();
    showToast(
      range.cellCount > 1
        ? `Cut ${range.cellCount} cells (${getRangeCoordString(range)}).`
        : `Cut cell ${getRangeCoordString(range)}.`,
      'info'
    );
  } else {
    scratchClipboardRangeState = {
      sheetName,
      minRow: effRange.minRow,
      maxRow: effRange.maxRow,
      minCol: effRange.minCol,
      maxCol: effRange.maxCol
    };
    applyScratchSelectionDOM();
    if (effRange.cellCount > 1 || !writeSystemClipboard) {
      showToast(
        effRange.cellCount > 1
          ? `Copied ${effRange.cellCount} cells (${getRangeCoordString(effRange)}).`
          : `Copied cell ${getRangeCoordString(effRange)}.`,
        'info'
      );
    }
  }
  return tsv;
}

function applyScratchPasteFromText(
  sheetName,
  rawText,
  htmlText = '',
  { clearCopyMarquee = false, transpose = false } = {}
) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB) return 0;
  const hasHtmlTable = Boolean(htmlText && /<table[\s>]/i.test(htmlText));
  if ((rawText === undefined || rawText === null || rawText === '') && !hasHtmlTable) return 0;

  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return 0;

  const hasMatchingInternalClip = Boolean(
    rawText &&
    scratchInternalClipboardState &&
    scratchInternalClipboardState.tsv === rawText
  );
  let matrix = [];
  let styleMatrix = null;
  let rawEditMatrix = null;
  let srcRange = null;

  if (hasMatchingInternalClip) {
    if (Array.isArray(scratchInternalClipboardState.styleMatrix)) {
      styleMatrix = scratchInternalClipboardState.styleMatrix;
    }
    if (Array.isArray(scratchInternalClipboardState.rawEditMatrix)) {
      rawEditMatrix = scratchInternalClipboardState.rawEditMatrix;
      matrix = rawEditMatrix.map(row => row.map(cell => (cell ? cell.rawEdit : '')));
    } else {
      matrix = parseClipboardTSVMatrix(rawText);
    }
    srcRange = scratchInternalClipboardState.sourceRange || null;
  } else if (hasHtmlTable) {
    const parsedHtml = parseClipboardHtmlTableToMatrix(htmlText);
    if (parsedHtml && Array.isArray(parsedHtml.values) && parsedHtml.values.length > 0) {
      matrix = parsedHtml.values;
      styleMatrix = parsedHtml.styles || null;
    } else {
      matrix = parseClipboardTSVMatrix(rawText || '');
    }
  } else {
    matrix = parseClipboardTSVMatrix(rawText || '');
  }

  if (matrix.length === 0) return 0;

  if (transpose) {
    matrix = transpose2DMatrix(matrix, '');
    if (styleMatrix) styleMatrix = transpose2DMatrix(styleMatrix, null);
    if (rawEditMatrix) rawEditMatrix = transpose2DMatrix(rawEditMatrix, null);
  }

  const curRange = getScratchSelectionRange(sheetName);
  const isSingleValueClipboard = matrix.length === 1 && matrix[0].length === 1;
  const edits = [];
  let nextSelection = null;

  if (isSingleValueClipboard && curRange.cellCount > 1) {
    // Excel-style Repeat-Fill: fill entire highlighted multi-cell range with the single copied value
    const singleRawEntry = rawEditMatrix && rawEditMatrix[0] ? rawEditMatrix[0][0] : null;
    const baseVal = singleRawEntry ? singleRawEntry.rawEdit : matrix[0][0];
    const singleIsFormula = Boolean(singleRawEntry && singleRawEntry.isFormula && baseVal.startsWith('='));
    const singleStyle = styleMatrix && styleMatrix[0] ? styleMatrix[0][0] : undefined;
    for (let r = curRange.minRow; r <= curRange.maxRow; r++) {
      for (let c = curRange.minCol; c <= curRange.maxCol; c++) {
        const finalVal = (singleIsFormula && srcRange)
          ? shiftFormulaCellReferences(baseVal, r - srcRange.minRow, c - srcRange.minCol)
          : baseVal;
        const editObj = { r, c, rawText: finalVal };
        if (singleStyle !== undefined) {
          editObj.styleSnap = singleStyle || { fill: null, font: null, numFmt: null, alignment: null };
        }
        edits.push(editObj);
      }
    }
    nextSelection = {
      anchorRow: curRange.anchorRow,
      anchorCol: curRange.anchorCol,
      focusRow: curRange.focusRow,
      focusCol: curRange.focusCol
    };
  } else {
    // Paste 2D block starting from top-left of current selection
    const startRow = curRange.minRow;
    const startCol = curRange.minCol;
    let maxR = startRow;
    let maxC = startCol;

    matrix.forEach((cols, rOffset) => {
      const styleRow = styleMatrix && styleMatrix[rOffset] ? styleMatrix[rOffset] : null;
      const rawRow = rawEditMatrix && rawEditMatrix[rOffset] ? rawEditMatrix[rOffset] : null;
      cols.forEach((cellVal, cOffset) => {
        const r = startRow + rOffset;
        const c = startCol + cOffset;
        if (r > maxR) maxR = r;
        if (c > maxC) maxC = c;
        const rawEntry = rawRow && rawRow[cOffset] ? rawRow[cOffset] : null;
        let finalVal = rawEntry ? rawEntry.rawEdit : cellVal;
        if (rawEntry && rawEntry.isFormula && finalVal.startsWith('=') && srcRange) {
          const origR = transpose ? (srcRange.minRow + cOffset) : (srcRange.minRow + rOffset);
          const origC = transpose ? (srcRange.minCol + rOffset) : (srcRange.minCol + cOffset);
          finalVal = shiftFormulaCellReferences(finalVal, r - origR, c - origC);
        }
        const editObj = { r, c, rawText: finalVal };
        if (styleMatrix) {
          const cellStyle = styleRow && styleRow[cOffset] ? styleRow[cOffset] : null;
          editObj.styleSnap = cellStyle || { fill: null, font: null, numFmt: null, alignment: null };
        }
        edits.push(editObj);
      });
    });

    nextSelection = {
      anchorRow: startRow,
      anchorCol: startCol,
      focusRow: maxR,
      focusCol: maxC
    };
  }

  if (clearCopyMarquee) {
    scratchClipboardRangeState = null;
  }
  const changed = applyScratchCellChanges(sheetName, edits, {
    recordUndo: true,
    selectionAfter: nextSelection
  });

  if (clearCopyMarquee) {
    applyScratchSelectionDOM();
  }

  if (transpose && edits.length > 0) {
    showToast(
      `Pasted & transposed ${edits.length} cell${edits.length === 1 ? '' : 's'} (${matrix.length}×${matrix[0].length}) into "${sheetName}".`,
      'info'
    );
  } else if (edits.length > 1) {
    showToast(`Pasted ${edits.length} cells into "${sheetName}".`, 'info');
  }
  return changed;
}

async function handleScratchContextPaste({ transpose = false } = {}) {
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  const targetSheet = activeWorkspaceTab;
  let clipText = null;
  let clipHtml = '';

  if (navigator.clipboard && typeof navigator.clipboard.read === 'function') {
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        if (!clipText && item.types && item.types.includes('text/plain')) {
          const blob = await item.getType('text/plain');
          clipText = await blob.text();
        }
        if (!clipHtml && item.types && item.types.includes('text/html')) {
          const blob = await item.getType('text/html');
          clipHtml = await blob.text();
        }
      }
    } catch (_) {
      clipText = null;
    }
  }

  if ((clipText === null || clipText === '') && navigator.clipboard && typeof navigator.clipboard.readText === 'function') {
    try {
      clipText = await navigator.clipboard.readText();
    } catch (_) {
      clipText = null;
    }
  }

  if ((clipText === null || clipText === '') && !clipHtml && scratchInternalClipboardText !== '') {
    clipText = scratchInternalClipboardText;
    if (scratchInternalClipboardState && scratchInternalClipboardState.tsv === clipText) {
      clipHtml = scratchInternalClipboardState.html || '';
    }
  }

  const hasHtmlTable = Boolean(clipHtml && /<table[\s>]/i.test(clipHtml));
  if ((clipText !== null && clipText !== '') || hasHtmlTable) {
    applyScratchPasteFromText(targetSheet, clipText || '', clipHtml, { transpose });
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
  } else {
    showToast('Allow clipboard read permission when prompted, or press Ctrl+V to paste.', 'info');
  }
}

function transposeScratchSelection(sheetName, range) {
  if (!workbook || !sheetName || sheetName === DAILY_TRACKER_WS_TAB || !range) return 0;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return 0;

  const effRange = resolveEffectiveScratchCopyRange(ws, range);
  const rowCount = effRange.maxRow - effRange.minRow + 1;
  const colCount = effRange.maxCol - effRange.minCol + 1;
  if (rowCount * colCount <= 1) {
    showToast('Select 2 or more cells to transpose.', 'info');
    return 0;
  }

  const srcGrid = [];
  for (let r = effRange.minRow; r <= effRange.maxRow; r++) {
    const rowItems = [];
    for (let c = effRange.minCol; c <= effRange.maxCol; c++) {
      const info = extractScratchCellInfo(ws, r, c);
      rowItems.push({
        srcR: r,
        srcC: c,
        rawEdit: info.rawEdit,
        isFormula: info.isFormula,
        styleSnap: captureCellStyleSnapshot(ws, r, c)
      });
    }
    srcGrid.push(rowItems);
  }

  const edits = [];
  const writtenCoords = new Set();

  for (let rOff = 0; rOff < rowCount; rOff++) {
    for (let cOff = 0; cOff < colCount; cOff++) {
      const destR = effRange.minRow + cOff;
      const destC = effRange.minCol + rOff;
      const cellSnap = srcGrid[rOff][cOff];
      let nextRaw = cellSnap.rawEdit;
      if (cellSnap.isFormula && nextRaw.startsWith('=')) {
        nextRaw = shiftFormulaCellReferences(nextRaw, destR - cellSnap.srcR, destC - cellSnap.srcC);
      }
      edits.push({
        r: destR,
        c: destC,
        rawText: nextRaw,
        styleSnap: cellSnap.styleSnap
      });
      writtenCoords.add(`${destR}::${destC}`);
    }
  }

  for (let r = effRange.minRow; r <= effRange.maxRow; r++) {
    for (let c = effRange.minCol; c <= effRange.maxCol; c++) {
      if (!writtenCoords.has(`${r}::${c}`)) {
        edits.push({
          r,
          c,
          rawText: '',
          styleSnap: { fill: null, font: null, numFmt: null, alignment: null }
        });
      }
    }
  }

  scratchClipboardRangeState = null;
  const nextSelection = {
    anchorRow: effRange.minRow,
    anchorCol: effRange.minCol,
    focusRow: effRange.minRow + colCount - 1,
    focusCol: effRange.minCol + rowCount - 1
  };
  const changed = applyScratchCellChanges(sheetName, edits, {
    recordUndo: true,
    selectionAfter: nextSelection
  });
  syncScratchFormulaBarFromActiveCell();
  showToast(
    `Transposed ${rowCount}×${colCount} → ${colCount}×${rowCount} (${getRangeCoordString(normalizeScratchRangeObj(nextSelection))}).`,
    'success'
  );
  return changed;
}

// --- EXCEL DEFAULT FILL PALETTE & RIGHT-CLICK CONTEXT MENU ---
function syncScratchFillIndicatorUI() {
  const preset = EXCEL_DEFAULT_FILL_PALETTE.find(p => p.id === activeScratchFillPresetId) || EXCEL_DEFAULT_FILL_PALETTE[5];
  if (scratchFillActiveBar && preset) {
    scratchFillActiveBar.style.background = preset.hex;
  }
  if (btnScratchApplyFill && preset) {
    btnScratchApplyFill.title = `Apply "${preset.label}" fill to selected cell(s)`;
  }
  if (scratchFillDropdown) {
    scratchFillDropdown.querySelectorAll('.scratch-fill-swatch').forEach(btn => {
      btn.classList.toggle('active-swatch', btn.dataset.fillId === activeScratchFillPresetId);
    });
  }
}

function initScratchFillPaletteUI() {
  const makeSwatchBtn = (preset, clsName) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = clsName;
    btn.dataset.fillId = preset.id;
    btn.style.background = preset.hex;
    btn.title = preset.label;
    btn.setAttribute('aria-label', preset.label);
    return btn;
  };

  if (scratchFillPastelGrid) {
    scratchFillPastelGrid.innerHTML = '';
    EXCEL_DEFAULT_FILL_PALETTE.filter(p => p.group === 'pastel').forEach(preset => {
      scratchFillPastelGrid.appendChild(makeSwatchBtn(preset, 'scratch-fill-swatch'));
    });
  }
  if (scratchFillStandardGrid) {
    scratchFillStandardGrid.innerHTML = '';
    EXCEL_DEFAULT_FILL_PALETTE.filter(p => p.group === 'standard').forEach(preset => {
      scratchFillStandardGrid.appendChild(makeSwatchBtn(preset, 'scratch-fill-swatch'));
    });
  }
  if (scratchCtxQuickSwatches) {
    scratchCtxQuickSwatches.innerHTML = '';
    const noneBtn = document.createElement('button');
    noneBtn.type = 'button';
    noneBtn.className = 'scratch-ctx-swatch-btn scratch-swatch-none';
    noneBtn.dataset.fillId = 'none';
    noneBtn.title = 'No Fill (Clear Color)';
    noneBtn.setAttribute('aria-label', 'No Fill');
    scratchCtxQuickSwatches.appendChild(noneBtn);

    EXCEL_DEFAULT_FILL_PALETTE.filter(p => p.quick).slice(0, 7).forEach(preset => {
      scratchCtxQuickSwatches.appendChild(makeSwatchBtn(preset, 'scratch-ctx-swatch-btn'));
    });
  }

  syncScratchFillIndicatorUI();
  syncScratchStripedRowsUI();
  syncScratchAutoIncrementUI();
}

function openScratchContextMenuAt(clientX, clientY) {
  if (!scratchContextMenu || activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  const range = getScratchSelectionRange(activeWorkspaceTab);
  if (scratchCtxRangeLabel) {
    scratchCtxRangeLabel.textContent = `Selection • ${getRangeCoordString(range)}`;
  }
  const transposeSelBtn = scratchContextMenu.querySelector('[data-ctx-action="transpose-selection"]');
  if (transposeSelBtn) {
    transposeSelBtn.disabled = range.cellCount <= 1;
  }
  syncScratchTextFormattingUI();

  scratchContextMenu.style.display = 'block';
  scratchContextMenu.style.left = `${clientX}px`;
  scratchContextMenu.style.top = `${clientY}px`;

  const menuRect = scratchContextMenu.getBoundingClientRect();
  const maxLeft = Math.max(8, window.innerWidth - menuRect.width - 8);
  const maxTop = Math.max(8, window.innerHeight - menuRect.height - 8);
  scratchContextMenu.style.left = `${Math.min(clientX, maxLeft)}px`;
  scratchContextMenu.style.top = `${Math.min(clientY, maxTop)}px`;
}

// --- DATA TO COLUMNS (TEXT TO COLUMNS) WIZARD & ENGINE ---
function getActiveTextToColumnsConfig() {
  const delims = [];
  if (chkTtcTab && chkTtcTab.checked) delims.push('\t');
  if (chkTtcComma && chkTtcComma.checked) delims.push(',');
  if (chkTtcSemicolon && chkTtcSemicolon.checked) delims.push(';');
  if (chkTtcPipe && chkTtcPipe.checked) delims.push('|');
  if (chkTtcHyphen && chkTtcHyphen.checked) delims.push('-');
  if (chkTtcSpace && chkTtcSpace.checked) delims.push(' ');
  if (chkTtcNewline && chkTtcNewline.checked) delims.push('\n');
  if (chkTtcOther && chkTtcOther.checked && ttcOtherInput && ttcOtherInput.value.length > 0) {
    delims.push(ttcOtherInput.value);
  }
  return {
    delims,
    trim: Boolean(!chkTtcTrim || chkTtcTrim.checked),
    mergeConsecutive: Boolean(chkTtcConsecutive && chkTtcConsecutive.checked),
    direction: (radTtcDown && radTtcDown.checked) ? 'down' : 'across'
  };
}

function autoDetectDelimitersForSamples(samples) {
  const sampleSlice = Array.isArray(samples) ? samples.slice(0, 15) : [];
  const hasNewline = sampleSlice.some(s => /[\r\n]/.test(String(s || '')));
  const hasBlankLines = hasNewline && sampleSlice.some(s => /\r?\n[ \t]*\r?\n/.test(String(s || '')));
  const joined = sampleSlice.join('\n');
  const hasTab = joined.includes('\t');
  const hasComma = joined.includes(',');
  const hasPipe = joined.includes('|');
  const hasSemi = joined.includes(';');
  const hasHyphen = joined.includes(' - ') || /[A-Za-z0-9]-[A-Za-z0-9]/.test(joined);

  if (chkTtcTab) chkTtcTab.checked = false;
  if (chkTtcComma) chkTtcComma.checked = false;
  if (chkTtcSemicolon) chkTtcSemicolon.checked = false;
  if (chkTtcPipe) chkTtcPipe.checked = false;
  if (chkTtcHyphen) chkTtcHyphen.checked = false;
  if (chkTtcSpace) chkTtcSpace.checked = false;
  if (chkTtcNewline) chkTtcNewline.checked = false;
  if (chkTtcOther) chkTtcOther.checked = false;
  if (chkTtcConsecutive) chkTtcConsecutive.checked = false;
  if (radTtcAcross) radTtcAcross.checked = true;
  if (radTtcDown) radTtcDown.checked = false;

  if (hasNewline) {
    if (chkTtcNewline) chkTtcNewline.checked = true;
    if (hasBlankLines && chkTtcConsecutive) chkTtcConsecutive.checked = true;
    if (sampleSlice.length === 1 && radTtcDown) {
      radTtcDown.checked = true;
      if (radTtcAcross) radTtcAcross.checked = false;
    }
  } else if (hasTab) {
    if (chkTtcTab) chkTtcTab.checked = true;
  } else if (hasComma) {
    if (chkTtcComma) chkTtcComma.checked = true;
  } else if (hasPipe) {
    if (chkTtcPipe) chkTtcPipe.checked = true;
  } else if (hasSemi) {
    if (chkTtcSemicolon) chkTtcSemicolon.checked = true;
  } else if (hasHyphen) {
    if (chkTtcHyphen) chkTtcHyphen.checked = true;
  } else if (joined.includes(' ')) {
    if (chkTtcSpace) chkTtcSpace.checked = true;
    if (chkTtcConsecutive) chkTtcConsecutive.checked = true;
  } else {
    if (chkTtcComma) chkTtcComma.checked = true;
  }
}

function closeTextToColumnsModal() {
  activeTextToColumnsState = null;
  if (textToColumnsModal) {
    textToColumnsModal.classList.remove('open');
  }
}

function renderTextToColumnsPreview() {
  if (!activeTextToColumnsState || !workbook || !ttcPreviewBox) return;
  const { sheetName, col, sourceRows } = activeTextToColumnsState;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;

  const config = getActiveTextToColumnsConfig();
  const isDown = config.direction === 'down';
  const colLetter = colIndexToExcelLetter(col);

  if (ttcSubtitle) {
    ttcSubtitle.textContent = isDown
      ? `Splitting ${sourceRows.length} non-empty cell${sourceRows.length === 1 ? '' : 's'} from Column ${colLetter} downward into consecutive rows`
      : `Splitting ${sourceRows.length} non-empty cell${sourceRows.length === 1 ? '' : 's'} from Column ${colLetter} into adjacent columns to the right`;
  }
  if (btnConfirmTextToColumns) {
    btnConfirmTextToColumns.textContent = isDown ? 'Split into Rows' : 'Split into Columns';
  }

  if (config.delims.length === 0) {
    ttcPreviewBox.innerHTML = `<div class="ttc-preview-empty">Select at least one delimiter above to preview split ${isDown ? 'rows' : 'columns'}.</div>`;
    if (ttcOverwriteBanner) ttcOverwriteBanner.style.display = 'none';
    if (btnConfirmTextToColumns) btnConfirmTextToColumns.disabled = true;
    return;
  }

  if (btnConfirmTextToColumns) {
    btnConfirmTextToColumns.disabled = false;
  }

  if (isDown) {
    const flatParts = [];
    sourceRows.forEach(item => {
      splitScratchLineByDelimiters(item.raw, config).forEach(part => {
        flatParts.push(part);
      });
    });

    const startRow = sourceRows[0].r;
    const sourceRowSet = new Set(sourceRows.map(s => s.r));
    let overwriteCount = 0;
    let firstOverwrittenRow = 0;
    let lastOverwrittenRow = 0;

    for (let i = 0; i < flatParts.length; i++) {
      const targetR = startRow + i;
      if (sourceRowSet.has(targetR)) continue;
      const existingDisp = extractScratchCellInfo(ws, targetR, col).display.trim();
      if (existingDisp !== '') {
        overwriteCount++;
        if (!firstOverwrittenRow) firstOverwrittenRow = targetR;
        lastOverwrittenRow = targetR;
      }
    }

    if (ttcOverwriteBanner) {
      if (overwriteCount > 0) {
        const rowSpanLabel = firstOverwrittenRow === lastOverwrittenRow
          ? `Row ${firstOverwrittenRow}`
          : `Rows ${firstOverwrittenRow}–${lastOverwrittenRow}`;
        ttcOverwriteBanner.textContent = `Heads up: Splitting down will overwrite ${overwriteCount} existing non-empty cell${overwriteCount === 1 ? '' : 's'} in Column ${colLetter} (${rowSpanLabel}). You can undo anytime with Ctrl+Z.`;
        ttcOverwriteBanner.style.display = 'block';
      } else {
        ttcOverwriteBanner.style.display = 'none';
      }
    }

    if (ttcPreviewTitle) {
      ttcPreviewTitle.textContent = `Data Preview (${sourceRows.length} cell${sourceRows.length === 1 ? '' : 's'} → ${flatParts.length} row${flatParts.length === 1 ? '' : 's'} in Col ${colLetter})`;
    }

    const previewLimit = Math.min(12, flatParts.length);
    const headHtml = `<tr><th class="ttc-row-num-col">#</th><th>Col ${colLetter}</th></tr>`;
    const bodyRows = [];
    for (let i = 0; i < previewLimit; i++) {
      const val = flatParts[i] !== undefined ? flatParts[i] : '';
      bodyRows.push(`<tr><td class="ttc-row-num-col">${startRow + i}</td><td title="${escapeHtml(val)}">${escapeHtml(val)}</td></tr>`);
    }
    if (flatParts.length > previewLimit) {
      bodyRows.push(`<tr><td class="ttc-row-num-col">…</td><td style="color:var(--text-muted);">+${flatParts.length - previewLimit} more row(s)</td></tr>`);
    }
    ttcPreviewBox.innerHTML = `<table class="ttc-preview-table"><thead>${headHtml}</thead><tbody>${bodyRows.join('')}</tbody></table>`;
    return;
  }

  let maxCols = 1;
  const splitRows = sourceRows.map(item => {
    const parts = splitScratchLineByDelimiters(item.raw, config);
    if (parts.length > maxCols) maxCols = parts.length;
    return { r: item.r, parts };
  });

  // Check if splitting into adjacent columns (col + 1 .. col + maxCols - 1) overwrites existing non-empty cells
  let overwriteCount = 0;
  if (maxCols > 1) {
    splitRows.forEach(sr => {
      for (let cOff = 1; cOff < sr.parts.length; cOff++) {
        const existingDisp = extractScratchCellInfo(ws, sr.r, col + cOff).display.trim();
        if (existingDisp !== '') {
          overwriteCount++;
        }
      }
    });
  }

  if (ttcOverwriteBanner) {
    if (overwriteCount > 0) {
      const firstTargetLetter = colIndexToExcelLetter(col + 1);
      const lastTargetLetter = colIndexToExcelLetter(col + maxCols - 1);
      const colSpanLabel = firstTargetLetter === lastTargetLetter
        ? `Column ${firstTargetLetter}`
        : `Columns ${firstTargetLetter}–${lastTargetLetter}`;
      ttcOverwriteBanner.textContent = `Heads up: Splitting will overwrite ${overwriteCount} existing non-empty cell${overwriteCount === 1 ? '' : 's'} in ${colSpanLabel}. You can undo anytime with Ctrl+Z.`;
      ttcOverwriteBanner.style.display = 'block';
    } else {
      ttcOverwriteBanner.style.display = 'none';
    }
  }

  if (ttcPreviewTitle) {
    ttcPreviewTitle.textContent = `Data Preview (${sourceRows.length} non-empty row${sourceRows.length === 1 ? '' : 's'} → ${maxCols} column${maxCols === 1 ? '' : 's'})`;
  }

  const previewLimit = Math.min(8, splitRows.length);
  const displayCols = Math.min(maxCols, 12);
  let headHtml = `<tr><th class="ttc-row-num-col">#</th>`;
  for (let cOff = 0; cOff < displayCols; cOff++) {
    headHtml += `<th>Col ${colIndexToExcelLetter(col + cOff)}</th>`;
  }
  headHtml += `</tr>`;

  const bodyRows = [];
  for (let i = 0; i < previewLimit; i++) {
    const sr = splitRows[i];
    let cells = `<td class="ttc-row-num-col">${sr.r}</td>`;
    for (let cOff = 0; cOff < displayCols; cOff++) {
      const val = sr.parts[cOff] !== undefined ? sr.parts[cOff] : '';
      cells += `<td title="${escapeHtml(val)}">${escapeHtml(val)}</td>`;
    }
    bodyRows.push(`<tr>${cells}</tr>`);
  }

  ttcPreviewBox.innerHTML = `<table class="ttc-preview-table"><thead>${headHtml}</thead><tbody>${bodyRows.join('')}</tbody></table>`;
}

function openTextToColumnsModal() {
  closeScratchContextMenu();
  closeScratchFillDropdown();
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  commitActiveScratchInlineEdit();

  const sheetName = activeWorkspaceTab;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;

  const dims = getScratchSheetGridDimensions(ws, sheetName);
  const range = getScratchSelectionRange(sheetName);
  const col = range.minCol;

  let minRow = range.minRow;
  let maxRow = range.maxRow;
  if (range.minRow === 1 && range.maxRow >= dims.rows) {
    maxRow = Math.max(1, dims.dataMaxRow);
  }

  let sourceRows = [];
  for (let r = minRow; r <= maxRow; r++) {
    const disp = extractScratchCellInfo(ws, r, col).display;
    if (disp.trim() !== '') {
      sourceRows.push({ r, raw: disp });
    }
  }

  // Convenience: if user had only 1 empty cell selected, check if that column has data rows we can target
  if (sourceRows.length === 0 && range.cellCount === 1 && dims.dataMaxRow >= 1) {
    minRow = 1;
    maxRow = dims.dataMaxRow;
    for (let r = minRow; r <= maxRow; r++) {
      const disp = extractScratchCellInfo(ws, r, col).display;
      if (disp.trim() !== '') {
        sourceRows.push({ r, raw: disp });
      }
    }
  }

  if (sourceRows.length === 0) {
    showModalAlert({
      title: 'No Text to Split',
      badge: `Column ${colIndexToExcelLetter(col)}`,
      message: `Select one or more non-empty cells (or click Column ${colIndexToExcelLetter(col)} header) containing delimited text to split across columns.`
    });
    return;
  }

  activeTextToColumnsState = {
    sheetName,
    col,
    minRow: sourceRows[0].r,
    maxRow: sourceRows[sourceRows.length - 1].r,
    sourceRows
  };

  const coordBadge = sourceRows.length === 1
    ? getCellCoordString(sourceRows[0].r, col)
    : `${getCellCoordString(sourceRows[0].r, col)}:${getCellCoordString(sourceRows[sourceRows.length - 1].r, col)}`;

  if (ttcSourceBadge) {
    ttcSourceBadge.textContent = `Source • ${coordBadge}`;
  }
  if (ttcSubtitle) {
    ttcSubtitle.textContent = `Splitting ${sourceRows.length} non-empty cell${sourceRows.length === 1 ? '' : 's'} from Column ${colIndexToExcelLetter(col)} into adjacent columns to the right`;
  }

  autoDetectDelimitersForSamples(sourceRows.map(s => s.raw));
  renderTextToColumnsPreview();

  if (textToColumnsModal) {
    textToColumnsModal.classList.add('open');
  }
}

function executeTextToColumnsSplit() {
  if (!activeTextToColumnsState || !workbook) return;
  const { sheetName, col, minRow, maxRow, sourceRows } = activeTextToColumnsState;
  const config = getActiveTextToColumnsConfig();
  if (config.delims.length === 0) return;

  if (config.direction === 'down') {
    const flatParts = [];
    sourceRows.forEach(item => {
      splitScratchLineByDelimiters(item.raw, config).forEach(part => {
        flatParts.push(part);
      });
    });

    const startRow = sourceRows[0].r;
    const edits = flatParts.map((val, idx) => ({
      r: startRow + idx,
      c: col,
      rawText: val
    }));

    const endWrittenRow = startRow + flatParts.length - 1;
    sourceRows.forEach(item => {
      if (item.r > endWrittenRow) {
        edits.push({ r: item.r, c: col, rawText: '' });
      }
    });

    const endRow = startRow + Math.max(1, flatParts.length) - 1;
    const nextSelection = {
      anchorRow: startRow,
      anchorCol: col,
      focusRow: endRow,
      focusCol: col
    };

    closeTextToColumnsModal();
    applyScratchCellChanges(sheetName, edits, {
      recordUndo: true,
      selectionAfter: nextSelection
    });

    showToast(
      `Split ${sourceRows.length} cell${sourceRows.length === 1 ? '' : 's'} down into ${flatParts.length} row${flatParts.length === 1 ? '' : 's'} (${getRangeCoordString(normalizeScratchRangeObj(nextSelection))}).`,
      'success'
    );
    if (scratchGridWrap) {
      scratchGridWrap.focus({ preventScroll: true });
    }
    return;
  }

  const edits = [];
  let maxSplitCols = 1;

  sourceRows.forEach(item => {
    const parts = splitScratchLineByDelimiters(item.raw, config);
    if (parts.length > maxSplitCols) {
      maxSplitCols = parts.length;
    }
    parts.forEach((val, offset) => {
      edits.push({
        r: item.r,
        c: col + offset,
        rawText: val
      });
    });
  });

  const endCol = col + Math.max(1, maxSplitCols) - 1;
  const nextSelection = {
    anchorRow: minRow,
    anchorCol: col,
    focusRow: maxRow,
    focusCol: endCol
  };

  closeTextToColumnsModal();
  applyScratchCellChanges(sheetName, edits, {
    recordUndo: true,
    selectionAfter: nextSelection
  });

  showToast(
    `Split ${sourceRows.length} row${sourceRows.length === 1 ? '' : 's'} across ${maxSplitCols} column${maxSplitCols === 1 ? '' : 's'} (${getRangeCoordString(normalizeScratchRangeObj(nextSelection))}).`,
    'success'
  );
  if (scratchGridWrap) {
    scratchGridWrap.focus({ preventScroll: true });
  }
}

// --- AUTO-FILTER SIDE LISTS (=FILTER ARRAY BUILDER) ---
function closeScratchAutoFilterModal() {
  activeScratchAutoFilterState = null;
  if (scratchAutoFilterModal) {
    scratchAutoFilterModal.classList.remove('open');
  }
}

function buildAutoFilterPlanFromUI() {
  if (!activeScratchAutoFilterState || !workbook) return null;
  const { sheetName, headers } = activeScratchAutoFilterState;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return null;

  const presetMode = afPresetModeSelect ? afPresetModeSelect.value : 'three_bucket';
  let retStartCol = parseInt(afReturnStartCol ? afReturnStartCol.value : '1', 10) || 1;
  let retEndCol = parseInt(afReturnEndCol ? afReturnEndCol.value : '1', 10) || 1;
  if (retEndCol < retStartCol) {
    const tmp = retStartCol;
    retStartCol = retEndCol;
    retEndCol = tmp;
  }
  const critCol = parseInt(afCriteriaCol ? afCriteriaCol.value : '1', 10) || 1;
  const destStartCol = parseInt(afDestStartCol ? afDestStartCol.value : '6', 10) || 6;
  const startRow = Math.max(2, parseInt(afStartRowInput ? afStartRowInput.value : '2', 10) || 2);
  const maxRow = Math.max(startRow, Math.min(5000, parseInt(afMaxRowInput ? afMaxRowInput.value : '200', 10) || 200));

  const retStartLetter = colIndexToExcelLetter(retStartCol);
  const retEndLetter = colIndexToExcelLetter(retEndCol);
  const critLetter = colIndexToExcelLetter(critCol);
  const retRange = `${retStartLetter}${startRow}:${retEndLetter}${maxRow}`;
  const critRange = `${critLetter}${startRow}:${critLetter}${maxRow}`;
  const nonEmptyGuard = `(${retStartLetter}${startRow}:${retStartLetter}${maxRow}<>"")`;

  const getHeaderLabel = (c) => {
    const found = headers.find(h => h.col === c);
    return (found && found.rawLabel) ? found.rawLabel : `Col ${colIndexToExcelLetter(c)}`;
  };
  const baseHeaders = [];
  for (let c = retStartCol; c <= retEndCol; c++) {
    baseHeaders.push(getHeaderLabel(c));
  }
  const critHeader = getHeaderLabel(critCol);

  const buckets = [];
  let cursorCol = destStartCol;

  if (presetMode === 'three_bucket') {
    const keyword = (afYellowKeywordInput ? afYellowKeywordInput.value : 'Duplicate').trim() || 'Duplicate';
    const kwEsc = keyword.replace(/"/g, '""');
    const includeCritInGreen = Boolean(!afIncludeCriteriaInGreen || afIncludeCriteriaInGreen.checked);

    // 1. Red Bucket: Empty Criteria
    buckets.push({
      key: 'red',
      title: `Missing ${critHeader} (Red)`,
      badgeClass: 'af-red',
      fillPresetId: 'excel_bad',
      startCol: cursorCol,
      headers: baseHeaders.map((h, i) => (i === 0 ? `${h} — No ${critHeader}` : h)),
      formulas: [
        {
          colOffset: 0,
          formula: `=FILTER(${retRange}, ${nonEmptyGuard}*(${critRange}=""), "")`
        }
      ]
    });
    cursorCol += baseHeaders.length + 1;

    // 2. Yellow Bucket: Keyword Match (e.g. "Duplicate")
    buckets.push({
      key: 'yellow',
      title: `${keyword} ${critHeader} (Yellow)`,
      badgeClass: 'af-yellow',
      fillPresetId: 'excel_neutral',
      startCol: cursorCol,
      headers: baseHeaders.map((h, i) => (i === 0 ? `${h} — ${keyword}` : h)),
      formulas: [
        {
          colOffset: 0,
          formula: `=FILTER(${retRange}, ${critRange}="${kwEsc}", "")`
        }
      ]
    });
    cursorCol += baseHeaders.length + 1;

    // 3. Green Bucket: Occupied Criteria & Not Keyword
    const greenCond = `(${critRange}<>"")*(${critRange}<>"${kwEsc}")`;
    const critAlreadyIncluded = critCol >= retStartCol && critCol <= retEndCol;
    const greenHeaders = [...baseHeaders];
    const greenFormulas = [];

    if (includeCritInGreen && !critAlreadyIncluded && critCol === retEndCol + 1) {
      const greenRange = `${retStartLetter}${startRow}:${critLetter}${maxRow}`;
      greenHeaders.push(critHeader);
      greenFormulas.push({
        colOffset: 0,
        formula: `=FILTER(${greenRange}, ${greenCond}, "")`
      });
    } else if (includeCritInGreen && !critAlreadyIncluded) {
      greenFormulas.push({
        colOffset: 0,
        formula: `=FILTER(${retRange}, ${greenCond}, "")`
      });
      greenFormulas.push({
        colOffset: baseHeaders.length,
        formula: `=FILTER(${critRange}, ${greenCond}, "")`
      });
      greenHeaders.push(critHeader);
    } else {
      greenFormulas.push({
        colOffset: 0,
        formula: `=FILTER(${retRange}, ${greenCond}, "")`
      });
    }

    buckets.push({
      key: 'green',
      title: `${critHeader} Available (Green)`,
      badgeClass: 'af-green',
      fillPresetId: 'excel_good',
      startCol: cursorCol,
      headers: greenHeaders.map((h, i) => (i === 0 ? `${h} — Available` : h)),
      formulas: greenFormulas
    });
  } else {
    const op = afCustomOperatorSelect ? afCustomOperatorSelect.value : 'isempty';
    const val = (afCustomValueInput ? afCustomValueInput.value : '').trim();
    const valEsc = val.replace(/"/g, '""');
    const colorId = afCustomColorSelect ? afCustomColorSelect.value : 'none';

    let condExpr = `${nonEmptyGuard}*(${critRange}="")`;
    let suffix = 'Filtered';
    if (op === 'isempty') {
      condExpr = `${nonEmptyGuard}*(${critRange}="")`;
      suffix = `Missing ${critHeader}`;
    } else if (op === 'isnotempty') {
      condExpr = `${critRange}<>""`;
      suffix = `Has ${critHeader}`;
    } else if (op === 'equals') {
      condExpr = `${critRange}="${valEsc}"`;
      suffix = `${critHeader} = ${val}`;
    } else if (op === 'notequals') {
      condExpr = `${nonEmptyGuard}*(${critRange}<>"${valEsc}")`;
      suffix = `${critHeader} <> ${val}`;
    } else if (op === 'contains') {
      condExpr = `ISNUMBER(SEARCH("${valEsc}", ${critRange}))`;
      suffix = `${critHeader} ~ "${val}"`;
    }

    buckets.push({
      key: 'custom',
      title: `Filtered (${suffix})`,
      badgeClass: colorId === 'excel_bad' ? 'af-red' : (colorId === 'excel_neutral' ? 'af-yellow' : (colorId === 'excel_good' ? 'af-green' : '')),
      fillPresetId: colorId,
      startCol: cursorCol,
      headers: baseHeaders.map((h, i) => (i === 0 ? `${h} — ${suffix}` : h)),
      formulas: [
        {
          colOffset: 0,
          formula: `=FILTER(${retRange}, ${condExpr}, "")`
        }
      ]
    });
  }

  return { sheetName, startRow, maxRow, buckets };
}

function renderScratchAutoFilterPreview() {
  if (!activeScratchAutoFilterState || !afFormulaPreviewBox) return;
  const presetMode = afPresetModeSelect ? afPresetModeSelect.value : 'three_bucket';
  if (afThreeBucketWrap) afThreeBucketWrap.style.display = presetMode === 'three_bucket' ? 'block' : 'none';
  if (afSingleCustomWrap) afSingleCustomWrap.style.display = presetMode === 'single_custom' ? 'block' : 'none';

  const plan = buildAutoFilterPlanFromUI();
  if (!plan) return;

  afFormulaPreviewBox.innerHTML = plan.buckets.map(b => {
    const anchorCoord = getCellCoordString(plan.startRow, b.startCol);
    const colSpan = `${colIndexToExcelLetter(b.startCol)}:${colIndexToExcelLetter(b.startCol + b.headers.length - 1)}`;
    const formulasHtml = b.formulas.map(f => `<code>${escapeHtml(f.formula)}</code>`).join(' ');
    return `
      <div class="af-preview-row">
        <span class="af-legend-chip ${b.badgeClass}">${escapeHtml(b.title)} (${escapeHtml(colSpan)} • @${escapeHtml(anchorCoord)})</span>
        ${formulasHtml}
      </div>
    `;
  }).join('');
}

function openScratchAutoFilterModal() {
  closeScratchContextMenu();
  closeScratchFillDropdown();
  if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
  commitActiveScratchInlineEdit();

  const sheetName = activeWorkspaceTab;
  const ws = workbook.getWorksheet(sheetName);
  if (!ws) return;
  const dims = getScratchSheetGridDimensions(ws, sheetName);
  const scanCols = Math.max(dims.dataMaxCol, 10);
  const headers = [];
  for (let c = 1; c <= scanCols; c++) {
    const rawLabel = extractScratchCellInfo(ws, 1, c, { skipSpillLookup: true }).display.trim();
    const letter = colIndexToExcelLetter(c);
    headers.push({
      col: c,
      letter,
      rawLabel,
      label: rawLabel ? `Col ${letter} — ${rawLabel}` : `Col ${letter}`
    });
  }

  activeScratchAutoFilterState = { sheetName, headers };
  if (afSheetBadge) afSheetBadge.textContent = sheetName;

  const optsHtml = headers.map(h => `<option value="${h.col}">${escapeHtml(h.label)}</option>`).join('');
  [afReturnStartCol, afReturnEndCol, afCriteriaCol].forEach(sel => {
    if (sel) sel.innerHTML = optsHtml;
  });

  if (afDestStartCol) {
    const destHeaders = [];
    const maxDest = Math.max(scanCols + 8, 20);
    for (let c = 1; c <= maxDest; c++) {
      const found = headers.find(h => h.col === c);
      const lbl = found && found.rawLabel ? `Col ${found.letter} (${found.rawLabel})` : `Col ${colIndexToExcelLetter(c)}`;
      destHeaders.push(`<option value="${c}">${escapeHtml(lbl)}</option>`);
    }
    afDestStartCol.innerHTML = destHeaders.join('');
  }

  // Smart column auto-detection
  let detectedCritCol = Math.max(2, Math.min(dims.dataMaxCol || 4, scanCols));
  const critMatch = headers.find(h => /track|carrier|mail|status|title\s*#/i.test(h.rawLabel));
  if (critMatch) {
    detectedCritCol = critMatch.col;
  }
  const retEndDefault = detectedCritCol > 1 ? detectedCritCol - 1 : 1;
  const destDefault = Math.max(dims.dataMaxCol + 2, detectedCritCol + 2, 6);

  if (afReturnStartCol) afReturnStartCol.value = '1';
  if (afReturnEndCol) afReturnEndCol.value = String(retEndDefault);
  if (afCriteriaCol) afCriteriaCol.value = String(detectedCritCol);
  if (afDestStartCol) afDestStartCol.value = String(destDefault);
  if (afStartRowInput) afStartRowInput.value = '2';
  if (afMaxRowInput) afMaxRowInput.value = String(Math.max(200, dims.dataMaxRow + 50));

  renderScratchAutoFilterPreview();
  if (scratchAutoFilterModal) {
    scratchAutoFilterModal.classList.add('open');
  }
}

function executeScratchAutoFilterSetup() {
  const plan = buildAutoFilterPlanFromUI();
  if (!plan || !workbook) return;
  const ws = workbook.getWorksheet(plan.sheetName);
  if (!ws) return;

  const dims = getScratchSheetGridDimensions(ws, plan.sheetName);
  const headerRowNum = Math.max(1, plan.startRow - 1);
  const tintEndRow = Math.max(30, Math.min(plan.maxRow, Math.max(dims.dataMaxRow + 15, 35)));
  const edits = [];
  let maxEndCol = plan.buckets[0].startCol;

  plan.buckets.forEach(bucket => {
    const colCount = bucket.headers.length;
    const endCol = bucket.startCol + colCount - 1;
    if (endCol > maxEndCol) maxEndCol = endCol;

    // 1. Headers on row above startRow + fill tint down tintEndRow
    bucket.headers.forEach((hdrText, cOff) => {
      const c = bucket.startCol + cOff;
      const baseSnap = captureCellStyleSnapshot(ws, headerRowNum, c);
      let hdrSnap = baseSnap;
      if (bucket.fillPresetId && bucket.fillPresetId !== 'none') {
        hdrSnap = buildStyleSnapshotForFillPreset(baseSnap, bucket.fillPresetId);
      }
      hdrSnap = {
        fill: hdrSnap.fill,
        font: { ...(hdrSnap.font || {}), bold: true },
        numFmt: hdrSnap.numFmt || null
      };
      edits.push({ r: headerRowNum, c, rawText: hdrText, styleSnap: hdrSnap });

      // Ensure formula / data rows start clean (empty values except formula anchor) & tinted
      for (let r = plan.startRow; r <= tintEndRow; r++) {
        const cellBefore = captureCellStyleSnapshot(ws, r, c);
        let cellSnap = cellBefore;
        if (bucket.fillPresetId && bucket.fillPresetId !== 'none') {
          cellSnap = buildStyleSnapshotForFillPreset(cellBefore, bucket.fillPresetId);
        }
        const formulaSpec = (r === plan.startRow)
          ? bucket.formulas.find(f => f.colOffset === cOff)
          : null;
        edits.push({
          r,
          c,
          rawText: formulaSpec ? formulaSpec.formula : '',
          styleSnap: cellSnap
        });
      }
    });
  });

  const firstBucket = plan.buckets[0];
  const nextSelection = {
    anchorRow: plan.startRow,
    anchorCol: firstBucket.startCol,
    focusRow: plan.startRow,
    focusCol: maxEndCol
  };

  closeScratchAutoFilterModal();
  applyScratchCellChanges(plan.sheetName, edits, {
    recordUndo: true,
    selectionAfter: nextSelection
  });

  showToast(
    `Created ${plan.buckets.length} auto-filtered side list${plan.buckets.length === 1 ? '' : 's'} using live =FILTER formulas.`,
    'success'
  );
  if (scratchGridWrap) {
    scratchGridWrap.focus({ preventScroll: true });
  }
}

function migrateScratchSheetRuntimeMaps(oldName, newName) {
  const maps = [
    scratchGridDimensionsMap,
    scratchActiveCellMap,
    scratchSelectionRangeMap,
    scratchUndoStackMap,
    scratchRedoStackMap,
    scratchPendingColWidths,
    scratchPendingFills,
    scratchSpillStateMap,
    scratchSpilledCellsMap
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

function captureScratchRuntimeSnapshot() {
  const cloneNestedMap = (outerMap) => {
    const out = new Map();
    if (!(outerMap instanceof Map)) return out;
    for (const [k, inner] of outerMap.entries()) {
      if (inner instanceof Map) {
        out.set(k, new Map(JSON.parse(JSON.stringify(Array.from(inner.entries())))));
      } else {
        out.set(k, JSON.parse(JSON.stringify(inner)));
      }
    }
    return out;
  };
  return {
    gridDimensions: cloneNestedMap(scratchGridDimensionsMap),
    activeCells: cloneNestedMap(scratchActiveCellMap),
    selectionRanges: cloneNestedMap(scratchSelectionRangeMap),
    undoStacks: cloneNestedMap(scratchUndoStackMap),
    redoStacks: cloneNestedMap(scratchRedoStackMap),
    pendingEdits: cloneNestedMap(scratchPendingEdits),
    pendingFills: cloneNestedMap(scratchPendingFills),
    pendingColWidths: cloneNestedMap(scratchPendingColWidths),
    spillStates: cloneNestedMap(scratchSpillStateMap),
    spilledCells: cloneNestedMap(scratchSpilledCellsMap),
    clipboardRange: scratchClipboardRangeState ? { ...scratchClipboardRangeState } : null
  };
}

function restoreScratchRuntimeSnapshot(snap) {
  resetScratchWorkbookRuntimeState();
  if (!snap || typeof snap !== 'object') return;
  const copyMapInto = (srcMap, destMap) => {
    if (!(srcMap instanceof Map)) return;
    for (const [k, v] of srcMap.entries()) {
      if (v instanceof Map) {
        destMap.set(k, new Map(JSON.parse(JSON.stringify(Array.from(v.entries())))));
      } else {
        destMap.set(k, JSON.parse(JSON.stringify(v)));
      }
    }
  };
  copyMapInto(snap.gridDimensions, scratchGridDimensionsMap);
  copyMapInto(snap.activeCells, scratchActiveCellMap);
  copyMapInto(snap.selectionRanges, scratchSelectionRangeMap);
  copyMapInto(snap.undoStacks, scratchUndoStackMap);
  copyMapInto(snap.redoStacks, scratchRedoStackMap);
  copyMapInto(snap.pendingEdits, scratchPendingEdits);
  copyMapInto(snap.pendingFills, scratchPendingFills);
  copyMapInto(snap.pendingColWidths, scratchPendingColWidths);
  copyMapInto(snap.spillStates, scratchSpillStateMap);
  copyMapInto(snap.spilledCells, scratchSpilledCellsMap);
  scratchClipboardRangeState = snap.clipboardRange ? { ...snap.clipboardRange } : null;
}

function resetScratchWorkbookRuntimeState() {
  stopScratchPointerDrag();
  closeScratchFillDropdown();
  closeScratchContextMenu();
  scratchPendingEdits.clear();
  scratchPendingFills.clear();
  scratchPendingColWidths.clear();
  scratchGridDimensionsMap.clear();
  scratchActiveCellMap.clear();
  scratchSelectionRangeMap.clear();
  scratchUndoStackMap.clear();
  scratchRedoStackMap.clear();
  scratchSpillStateMap.clear();
  scratchSpilledCellsMap.clear();
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
        value: 'title_buckets',
        label: 'Title Tracking + Auto-Filtered Red / Yellow / Green Side Lists (Lot #, Buyer #, LLY Date, Tracking #)'
      },
      {
        value: 'title_headers',
        label: 'Title Tracking Columns (Lot #, Yard #, Broker / Master #, Title Status, Carrier / Tracking #, Notes)'
      },
      {
        value: 'blank',
        label: 'Blank Spreadsheet Grid'
      }
    ],
    defaultSelectValue: 'title_buckets',
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

  if (starterLayout === 'title_buckets') {
    const mainHeaders = [
      { header: 'Lot #', width: 15, numFmt: '@' },
      { header: 'Buyer #', width: 14, numFmt: '@' },
      { header: 'LLY Date', width: 14, numFmt: 'mm/dd/yyyy' },
      { header: 'Tracking #', width: 24, numFmt: '@' }
    ];
    const headerRow = newWs.getRow(1);
    headerRow.height = 22;
    mainHeaders.forEach((colDef, idx) => {
      const colObj = newWs.getColumn(idx + 1);
      colObj.width = colDef.width;
      colObj.numFmt = colDef.numFmt;
      const cell = headerRow.getCell(idx + 1);
      cell.value = colDef.header;
      cell.font = { bold: true, color: { argb: 'FF0F172A' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2E8F0' } };
    });

    const bucketDefs = [
      {
        startCol: 6,
        fillArgb: 'FFFFC7CE',
        fontArgb: 'FF9C0006',
        headers: ['Lot # — No Tracking', 'Buyer #', 'LLY Date'],
        formula: '=_xlfn._xlws.FILTER(A2:C200, (A2:A200<>"")*(D2:D200=""), "")'.slice(1)
      },
      {
        startCol: 10,
        fillArgb: 'FFFFEB9C',
        fontArgb: 'FF9C6500',
        headers: ['Lot # — Duplicate', 'Buyer #', 'LLY Date'],
        formula: '=_xlfn._xlws.FILTER(A2:C200, D2:D200="Duplicate", "")'.slice(1)
      },
      {
        startCol: 14,
        fillArgb: 'FFC6EFCE',
        fontArgb: 'FF006100',
        headers: ['Lot # — Available', 'Buyer #', 'LLY Date', 'Tracking #'],
        formula: '=_xlfn._xlws.FILTER(A2:D200, (D2:D200<>"")*(D2:D200<>"Duplicate"), "")'.slice(1)
      }
    ];
    bucketDefs.forEach(b => {
      b.headers.forEach((hText, offset) => {
        const c = b.startCol + offset;
        newWs.getColumn(c).width = offset === 3 ? 24 : 16;
        const hCell = headerRow.getCell(c);
        hCell.value = hText;
        hCell.font = { bold: true, color: { argb: b.fontArgb } };
        hCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: b.fillArgb } };
        for (let r = 2; r <= 30; r++) {
          const rowObj = newWs.getRow(r);
          const dCell = rowObj.getCell(c);
          dCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: b.fillArgb } };
          dCell.font = { color: { argb: b.fontArgb } };
          if (r === 2 && offset === 0) {
            dCell.value = { formula: b.formula };
          }
          rowObj.commit();
        }
      });
    });
    headerRow.commit();
    scratchActiveCellMap.set(newSheetName, { row: 2, col: 1 });
    scratchSelectionRangeMap.set(newSheetName, { anchorRow: 2, anchorCol: 1, focusRow: 2, focusCol: 1 });
  } else if (starterLayout === 'title_headers') {
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
  scratchPendingFills.delete(targetName);
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
// EVENT LISTENERS — SCRATCH MENU, TOOLBAR, FILL PALETTE, CONTEXT MENU & GRID
// ============================================================================
initScratchFillPaletteUI();

if (btnOpenScratchMenu) {
  btnOpenScratchMenu.addEventListener('click', (e) => {
    e.stopPropagation();
    closeToolsMenu();
    closeScratchFillDropdown();
    closeScratchContextMenu();
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
  if (scratchFillMenuWrap && !scratchFillMenuWrap.contains(e.target)) {
    closeScratchFillDropdown();
  }
  if (scratchContextMenu && !scratchContextMenu.contains(e.target)) {
    closeScratchContextMenu();
  }
});

if (chkScratchStripedRows) {
  chkScratchStripedRows.addEventListener('change', () => {
    scratchStripedRowsEnabled = Boolean(chkScratchStripedRows.checked);
    try {
      localStorage.setItem(SCRATCH_STRIPED_STORAGE_KEY, String(scratchStripedRowsEnabled));
    } catch (_) { /* ignore */ }
    syncScratchStripedRowsUI();
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
  });
}

if (chkScratchAutoIncrement) {
  chkScratchAutoIncrement.addEventListener('change', () => {
    scratchAutoIncrementEnabled = Boolean(chkScratchAutoIncrement.checked);
    try {
      localStorage.setItem(SCRATCH_AUTO_INCREMENT_STORAGE_KEY, String(scratchAutoIncrementEnabled));
    } catch (_) { /* ignore */ }
    syncScratchAutoIncrementUI();
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
  });
}

if (btnScratchApplyFill) {
  btnScratchApplyFill.addEventListener('click', () => {
    closeScratchFillDropdown();
    closeScratchContextMenu();
    applyScratchFillPresetToSelection(activeWorkspaceTab, activeScratchFillPresetId);
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
  });
}

if (btnScratchFillDropdownToggle) {
  btnScratchFillDropdownToggle.addEventListener('click', (e) => {
    e.stopPropagation();
    closeToolsMenu();
    closeScratchMenu();
    closeScratchContextMenu();
    const isOpen = scratchFillDropdown && scratchFillDropdown.classList.contains('open');
    if (isOpen) {
      closeScratchFillDropdown();
    } else if (scratchFillDropdown) {
      syncScratchFillIndicatorUI();
      scratchFillDropdown.classList.add('open');
      btnScratchFillDropdownToggle.setAttribute('aria-expanded', 'true');
    }
  });
}

if (scratchFillDropdown) {
  scratchFillDropdown.addEventListener('click', (e) => {
    const fillBtn = e.target.closest('[data-fill-id]');
    if (!fillBtn) return;
    const presetId = fillBtn.dataset.fillId;
    closeScratchFillDropdown();
    applyScratchFillPresetToSelection(activeWorkspaceTab, presetId);
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
  });
}

if (scratchCellFormatSelect) {
  scratchCellFormatSelect.addEventListener('change', () => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    applyScratchFormatToSelection(activeWorkspaceTab, scratchCellFormatSelect.value);
    if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
  });
}

if (btnScratchTextToColumns) {
  btnScratchTextToColumns.addEventListener('click', openTextToColumnsModal);
}

if (btnScratchAutoFilter) {
  btnScratchAutoFilter.addEventListener('click', openScratchAutoFilterModal);
}

[
  [btnScratchFormatBold, 'toggle-bold'],
  [btnScratchFormatItalic, 'toggle-italic'],
  [btnScratchFormatUnderline, 'toggle-underline'],
  [btnScratchAlignLeft, 'align-left'],
  [btnScratchAlignCenter, 'align-center'],
  [btnScratchAlignRight, 'align-right'],
  [btnScratchClearTextFormat, 'clear-text-format']
].forEach(([btnEl, fmtAction]) => {
  if (btnEl) {
    btnEl.addEventListener('click', () => {
      closeScratchFillDropdown();
      closeScratchContextMenu();
      applyScratchTextFormattingToSelection(activeWorkspaceTab, fmtAction);
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
    });
  }
});

// Auto-Filter Side Lists Modal Listeners
[
  afPresetModeSelect,
  afReturnStartCol,
  afReturnEndCol,
  afCriteriaCol,
  afDestStartCol,
  afStartRowInput,
  afMaxRowInput,
  afYellowKeywordInput,
  afIncludeCriteriaInGreen,
  afCustomOperatorSelect,
  afCustomValueInput,
  afCustomColorSelect
].forEach(el => {
  if (el) {
    el.addEventListener('change', renderScratchAutoFilterPreview);
    if (el.tagName === 'INPUT' && el.type !== 'checkbox') {
      el.addEventListener('input', renderScratchAutoFilterPreview);
    }
  }
});

if (btnCancelAutoFilter) {
  btnCancelAutoFilter.addEventListener('click', closeScratchAutoFilterModal);
}

if (scratchAutoFilterModal) {
  scratchAutoFilterModal.addEventListener('click', (e) => {
    if (e.target === scratchAutoFilterModal) {
      closeScratchAutoFilterModal();
    }
  });
}

if (scratchAutoFilterForm) {
  scratchAutoFilterForm.addEventListener('submit', (e) => {
    e.preventDefault();
    executeScratchAutoFilterSetup();
  });
}

// Data to Columns Modal Listeners
[chkTtcComma, chkTtcTab, chkTtcSemicolon, chkTtcPipe, chkTtcHyphen, chkTtcSpace, chkTtcNewline, chkTtcOther, chkTtcTrim, chkTtcConsecutive, radTtcAcross, radTtcDown].forEach(el => {
  if (el) {
    el.addEventListener('change', renderTextToColumnsPreview);
  }
});

if (ttcOtherInput) {
  ttcOtherInput.addEventListener('input', () => {
    if (chkTtcOther && ttcOtherInput.value.length > 0 && !chkTtcOther.checked) {
      chkTtcOther.checked = true;
    }
    renderTextToColumnsPreview();
  });
}

if (btnCancelTextToColumns) {
  btnCancelTextToColumns.addEventListener('click', closeTextToColumnsModal);
}

if (textToColumnsModal) {
  textToColumnsModal.addEventListener('click', (e) => {
    if (e.target === textToColumnsModal) {
      closeTextToColumnsModal();
    }
  });
}

if (textToColumnsForm) {
  textToColumnsForm.addEventListener('submit', (e) => {
    e.preventDefault();
    executeTextToColumnsSplit();
  });
}

// Right-Click Context Menu Actions
if (scratchContextMenu) {
  scratchContextMenu.addEventListener('click', (e) => {
    const swatchBtn = e.target.closest('[data-fill-id]');
    if (swatchBtn) {
      const presetId = swatchBtn.dataset.fillId;
      closeScratchContextMenu();
      applyScratchFillPresetToSelection(activeWorkspaceTab, presetId);
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
      return;
    }

    const actionBtn = e.target.closest('[data-ctx-action]');
    if (!actionBtn || activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    const action = actionBtn.dataset.ctxAction;
    const range = getScratchSelectionRange(activeWorkspaceTab);
    closeScratchContextMenu();

    switch (action) {
      case 'cut':
        performScratchCopyOrCut(activeWorkspaceTab, range, { isCut: true, writeSystemClipboard: true });
        if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
        break;
      case 'copy':
        performScratchCopyOrCut(activeWorkspaceTab, range, { isCut: false, writeSystemClipboard: true });
        if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
        break;
      case 'paste':
        handleScratchContextPaste();
        break;
      case 'paste-transpose':
        handleScratchContextPaste({ transpose: true });
        break;
      case 'transpose-selection':
        transposeScratchSelection(activeWorkspaceTab, range);
        if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
        break;
      case 'text-to-columns':
        openTextToColumnsModal();
        break;
      case 'auto-filter-lists':
        openScratchAutoFilterModal();
        break;
      case 'clear-contents': {
        const cleared = clearScratchSelectionCells(activeWorkspaceTab, range);
        syncScratchFormulaBarFromActiveCell();
        if (cleared > 1) {
          showToast(`Cleared ${cleared} cells (${getRangeCoordString(range)}).`, 'info');
        }
        if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
        break;
      }
      case 'clear-fill':
        applyScratchFillPresetToSelection(activeWorkspaceTab, 'none');
        if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
        break;
      case 'toggle-bold':
      case 'toggle-italic':
      case 'toggle-underline':
      case 'align-left':
      case 'align-center':
      case 'align-right':
      case 'clear-text-format':
        applyScratchTextFormattingToSelection(activeWorkspaceTab, action);
        if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
        break;
      default:
        break;
    }
  });
}

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
    closeScratchContextMenu();
    closeScratchFillDropdown();
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
    closeScratchContextMenu();
    closeScratchFillDropdown();
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

    // 2. Corner Fill Handle Drag
    const fillHandle = e.target.closest('.scratch-fill-handle');
    if (fillHandle) {
      e.preventDefault();
      e.stopPropagation();
      commitActiveScratchInlineEdit();
      if (scratchGridWrap) scratchGridWrap.focus({ preventScroll: true });
      const curRange = getScratchSelectionRange(activeWorkspaceTab);
      startScratchPointerDrag('fill', e.clientX, e.clientY, {
        sourceRange: { ...curRange },
        fillTargetRange: null
      });
      return;
    }

    // 3. Grid Cell Click / Shift+Click / Drag Range Selection
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
    const fillHandle = e.target.closest('.scratch-fill-handle');
    if (fillHandle && activeWorkspaceTab !== DAILY_TRACKER_WS_TAB && workbook) {
      e.preventDefault();
      e.stopPropagation();
      const ws = workbook.getWorksheet(activeWorkspaceTab);
      if (!ws) return;
      const curRange = getScratchSelectionRange(activeWorkspaceTab);
      const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
      // Find contiguous filled rows in adjacent left (or right) column
      const probeCols = [curRange.minCol - 1, curRange.maxCol + 1].filter(col => col >= 1 && col <= dims.cols);
      let targetEndRow = curRange.maxRow;
      for (const pCol of probeCols) {
        let r = curRange.maxRow + 1;
        while (r <= Math.max(dims.dataMaxRow, dims.rows) && extractScratchCellInfo(ws, r, pCol).display.trim() !== '') {
          if (r > targetEndRow) targetEndRow = r;
          r++;
        }
        if (targetEndRow > curRange.maxRow) break;
      }
      if (targetEndRow > curRange.maxRow) {
        executeScratchFillHandleDrag(
          activeWorkspaceTab,
          curRange,
          {
            axis: 'vertical',
            dir: 1,
            minRow: curRange.minRow,
            maxRow: targetEndRow,
            minCol: curRange.minCol,
            maxCol: curRange.maxCol
          },
          { ctrlKey: Boolean(e.ctrlKey || e.metaKey) }
        );
      }
      return;
    }

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

window.addEventListener('mouseup', (e) => {
  if (scratchColResizeState) {
    const st = scratchColResizeState;
    scratchColResizeState = null;
    if (st.resizerEl) st.resizerEl.classList.remove('resizing');
    if (Math.abs(st.currentWidthPx - st.startWidthPx) >= 2) {
      recordScratchColumnWidthChange(st.sheetName, st.col, st.currentWidthPx);
    }
  }
  if (scratchDragState) {
    const dragSnap = scratchDragState;
    if (dragSnap.mode === 'fill' && dragSnap.fillTargetRange) {
      const src = dragSnap.sourceRange;
      const target = dragSnap.fillTargetRange;
      const sheet = dragSnap.sheetName;
      stopScratchPointerDrag();
      executeScratchFillHandleDrag(sheet, src, target, {
        ctrlKey: Boolean(e && (e.ctrlKey || e.metaKey))
      });
    } else {
      stopScratchPointerDrag();
    }
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

    // Text formatting shortcuts: Ctrl/Cmd + B (Bold/Unbold), Ctrl/Cmd + I (Italic), Ctrl/Cmd + U (Underline)
    if (mod && !e.altKey && (keyLower === 'b' || keyLower === 'i' || keyLower === 'u')) {
      e.preventDefault();
      const fmtMap = { b: 'toggle-bold', i: 'toggle-italic', u: 'toggle-underline' };
      applyScratchTextFormattingToSelection(activeWorkspaceTab, fmtMap[keyLower]);
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
        performScratchCopyOrCut(snapSheet, snapRange, { isCut, writeSystemClipboard: true });
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
        e.preventDefault();
        if (
          scratchClipboardRangeState &&
          scratchInternalClipboardText !== ''
        ) {
          applyScratchPasteFromText(
            activeWorkspaceTab,
            scratchInternalClipboardText,
            scratchInternalClipboardState ? scratchInternalClipboardState.html : '',
            { clearCopyMarquee: true }
          );
          return;
        }
        selectScratchCell(
          e.shiftKey ? Math.max(1, range.anchorRow - 1) : range.anchorRow + 1,
          range.anchorCol,
          { scrollIntoView: true }
        );
        return;
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

  // Right-Click Context Menu on Scratch Grid (Cut / Copy / Paste / Fill / Data to Columns)
  scratchGridWrap.addEventListener('contextmenu', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    // Allow native browser right-click menu when actively editing text inside an inline cell input
    if (e.target.closest('.scratch-inline-editor')) return;

    e.preventDefault();
    closeToolsMenu();
    closeScratchMenu();
    closeScratchFillDropdown();
    commitActiveScratchInlineEdit();

    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;
    const dims = getScratchSheetGridDimensions(ws, activeWorkspaceTab);
    const curRange = getScratchSelectionRange(activeWorkspaceTab);

    const td = e.target.closest('td[data-r][data-c]');
    const rowTh = e.target.closest('th[data-row-hdr]');
    const colTh = e.target.closest('th[data-col-hdr]');
    const cornerTh = e.target.closest('th.scratch-corner-th');

    if (td) {
      const r = parseInt(td.dataset.r, 10);
      const c = parseInt(td.dataset.c, 10);
      const insideSelection =
        r >= curRange.minRow &&
        r <= curRange.maxRow &&
        c >= curRange.minCol &&
        c <= curRange.maxCol;
      if (!insideSelection && r >= 1 && c >= 1) {
        selectScratchCell(r, c);
      }
    } else if (rowTh) {
      const r = parseInt(rowTh.dataset.rowHdr, 10);
      const rowAlreadySelected =
        curRange.minCol === 1 &&
        curRange.maxCol >= dims.cols &&
        r >= curRange.minRow &&
        r <= curRange.maxRow;
      if (!rowAlreadySelected && r >= 1) {
        setScratchSelectionRange(activeWorkspaceTab, {
          anchorRow: r,
          anchorCol: 1,
          focusRow: r,
          focusCol: dims.cols
        });
      }
    } else if (colTh) {
      const c = parseInt(colTh.dataset.colHdr, 10);
      const colAlreadySelected =
        curRange.minRow === 1 &&
        curRange.maxRow >= dims.rows &&
        c >= curRange.minCol &&
        c <= curRange.maxCol;
      if (!colAlreadySelected && c >= 1) {
        setScratchSelectionRange(activeWorkspaceTab, {
          anchorRow: 1,
          anchorCol: c,
          focusRow: dims.rows,
          focusCol: c
        });
      }
    } else if (cornerTh) {
      setScratchSelectionRange(activeWorkspaceTab, {
        anchorRow: 1,
        anchorCol: 1,
        focusRow: dims.rows,
        focusCol: dims.cols
      });
    }

    scratchGridWrap.focus({ preventScroll: true });
    openScratchContextMenuAt(e.clientX, e.clientY);
  });

  scratchGridWrap.addEventListener('scroll', () => {
    closeScratchContextMenu();
  });

  // Multi-Cell Excel HTML + TSV Copy Handler (Ctrl+C)
  scratchGridWrap.addEventListener('copy', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    if (scratchInlineEditState) return;

    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;
    const range = getScratchSelectionRange(activeWorkspaceTab);
    const effRange = resolveEffectiveScratchCopyRange(ws, range);
    const tsv = serializeScratchSelectionToTSV(ws, effRange);
    const html = serializeScratchSelectionToExcelHtml(ws, effRange);
    const styleMatrix = buildScratchStyleMatrix(ws, effRange);
    const rawEditMatrix = buildScratchRawEditMatrix(ws, effRange);
    scratchInternalClipboardText = tsv;
    scratchInternalClipboardState = {
      tsv,
      html,
      styleMatrix,
      rawEditMatrix,
      sourceRange: {
        minRow: effRange.minRow,
        maxRow: effRange.maxRow,
        minCol: effRange.minCol,
        maxCol: effRange.maxCol
      }
    };

    if (e.clipboardData) {
      e.preventDefault();
      e.clipboardData.setData('text/plain', tsv);
      if (html) e.clipboardData.setData('text/html', html);
      lastScratchNativeCopyAt = Date.now();
    }

    scratchClipboardRangeState = {
      sheetName: activeWorkspaceTab,
      minRow: effRange.minRow,
      maxRow: effRange.maxRow,
      minCol: effRange.minCol,
      maxCol: effRange.maxCol
    };
    applyScratchSelectionDOM();
    if (effRange.cellCount > 1) {
      showToast(`Copied ${effRange.cellCount} cells (${getRangeCoordString(effRange)}).`, 'info');
    }
  });

  // Multi-Cell Excel HTML + TSV Cut Handler (Ctrl+X)
  scratchGridWrap.addEventListener('cut', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    if (scratchInlineEditState) return;

    const ws = workbook.getWorksheet(activeWorkspaceTab);
    if (!ws) return;
    const range = getScratchSelectionRange(activeWorkspaceTab);
    const tsv = serializeScratchSelectionToTSV(ws, range);
    const html = serializeScratchSelectionToExcelHtml(ws, range);
    const styleMatrix = buildScratchStyleMatrix(ws, range);
    const rawEditMatrix = buildScratchRawEditMatrix(ws, range);
    scratchInternalClipboardText = tsv;
    scratchInternalClipboardState = {
      tsv,
      html,
      styleMatrix,
      rawEditMatrix,
      sourceRange: {
        minRow: range.minRow,
        maxRow: range.maxRow,
        minCol: range.minCol,
        maxCol: range.maxCol
      }
    };

    if (e.clipboardData) {
      e.preventDefault();
      e.clipboardData.setData('text/plain', tsv);
      if (html) e.clipboardData.setData('text/html', html);
      lastScratchNativeCopyAt = Date.now();
    }

    scratchClipboardRangeState = null;
    clearScratchSelectionCells(activeWorkspaceTab, range);
    syncScratchFormulaBarFromActiveCell();
    if (range.cellCount > 1) {
      showToast(`Cut ${range.cellCount} cells (${getRangeCoordString(range)}).`, 'info');
    }
  });

  // Multi-Cell TSV / Excel / HTML Clipboard Paste Handler (Ctrl+V + 1-Cell Repeat-Fill)
  function isScratchPasteBlockedByInputEl(el) {
    if (!el || !el.tagName) return false;
    if (el.isContentEditable) return true;
    const tag = el.tagName.toUpperCase();
    if (tag === 'TEXTAREA' || tag === 'SELECT') return true;
    if (tag === 'INPUT') {
      const type = String(el.type || 'text').toLowerCase();
      return !['checkbox', 'radio', 'button', 'submit', 'reset', 'range', 'color', 'file'].includes(type);
    }
    return false;
  }

  function handleScratchGridPasteEvent(e) {
    const cb = e.clipboardData || window.clipboardData;
    if (!cb) return false;
    const text = cb.getData('text/plain');
    const html = cb.getData('text/html') || '';
    const hasHtmlTable = Boolean(html && /<table[\s>]/i.test(html));
    if ((text === undefined || text === null || text === '') && !hasHtmlTable) return false;
    e.preventDefault();

    applyScratchPasteFromText(activeWorkspaceTab, text || '', html);
    return true;
  }

  scratchGridWrap.addEventListener('paste', (e) => {
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    if (scratchInlineEditState) return; // Let inline editor handle single-cell text paste
    handleScratchGridPasteEvent(e);
  });

  // Document-level paste fallback when Scratch Sheet is active but focus slipped onto toolbar/body
  document.addEventListener('paste', (e) => {
    if (e.defaultPrevented) return;
    if (activeWorkspaceTab === DAILY_TRACKER_WS_TAB || !workbook) return;
    if (scratchInlineEditState) return;
    if (typeof document.querySelector === 'function' && document.querySelector('.modal-backdrop.open')) return;
    if (isScratchPasteBlockedByInputEl(e.target) || isScratchPasteBlockedByInputEl(document.activeElement)) return;
    handleScratchGridPasteEvent(e);
  });
}

// Global Ctrl+S & Escape handlers while viewing a Scratch Sheet Tab
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    if (scratchAutoFilterModal && scratchAutoFilterModal.classList.contains('open')) {
      e.preventDefault();
      closeScratchAutoFilterModal();
      return;
    }
    if (textToColumnsModal && textToColumnsModal.classList.contains('open')) {
      e.preventDefault();
      closeTextToColumnsModal();
      return;
    }
    if (scratchContextMenu && scratchContextMenu.style.display !== 'none') {
      e.preventDefault();
      closeScratchContextMenu();
      return;
    }
    if (scratchFillDropdown && scratchFillDropdown.classList.contains('open')) {
      e.preventDefault();
      closeScratchFillDropdown();
      return;
    }
  }
  if (
    activeWorkspaceTab !== DAILY_TRACKER_WS_TAB &&
    (e.ctrlKey || e.metaKey) &&
    e.key.toLowerCase() === 's'
  ) {
    e.preventDefault();
    flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: false });
  }
});

// Closeable / Acknowledgeable Scratch Sheet Formula Warning Persistence
const SCRATCH_FORMULA_NOTICE_ACK_KEY = 'trackerui_scratch_formula_notice_ack_v1';

function syncScratchFormulaNoticeVisibility() {
  if (!scratchFormulaNotice) return;
  const isAcked = localStorage.getItem(SCRATCH_FORMULA_NOTICE_ACK_KEY) === 'true';
  scratchFormulaNotice.style.display = isAcked ? 'none' : 'flex';
}

if (btnAckScratchFormulaNotice) {
  btnAckScratchFormulaNotice.addEventListener('click', () => {
    localStorage.setItem(SCRATCH_FORMULA_NOTICE_ACK_KEY, 'true');
    syncScratchFormulaNoticeVisibility();
    showToast('Scratch sheet formula warning acknowledged and hidden.', 'info');
  });
}

syncScratchFormulaNoticeVisibility();
renderTrackerProfilesBar();
