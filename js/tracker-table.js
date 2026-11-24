// ============================================================================
// tracker-table.js — Worksheet Row Indexing, UUIDv4 Backfill, Sheet Stepper &
//                    Tab Bar, Daily Request Table, Search, Sort, Like-Item
//                    Filters, Permanent Row Selection, Batch Status & Delete
// ============================================================================

// --- INDEX WORKSHEET ROWS & LEGACY GUID BACKFILL ---
function indexWorksheetRows(worksheet, sheetName) {
  const colMap = getWorksheetColumnMap(worksheet);
  if (colMap.isCustomSchema && Array.isArray(colMap.columns)) {
    colMap.columns.forEach(c => {
      worksheet.getColumn(c.colIndex).hidden = false;
    });
  }
  worksheet.getColumn(colMap.trackingCol).hidden = true;
  if (colMap.emailSubjCol) {
    worksheet.getColumn(colMap.emailSubjCol).hidden = false;
  }

  const defaultStatusVal = getDefaultStatus();

  // Pre-scan explicit Tracking IDs already present on this sheet so backfilled rows never collide on the same sheet
  const usedKeysOnSheet = new Set();
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const existingKey = cellText(row, colMap.trackingCol).toLowerCase();
    if (existingKey && UUID_REGEX.test(existingKey)) {
      usedKeysOnSheet.add(existingKey);
    } else if (colMap.isCustomSchema && colMap.statusCol) {
      const legacyColUuid = cellText(row, colMap.statusCol);
      if (UUID_REGEX.test(legacyColUuid)) {
        usedKeysOnSheet.add(legacyColUuid.toLowerCase());
      }
    } else if (colMap.isTemplateLayout) {
      const col11Text = cellText(row, 11);
      if (UUID_REGEX.test(col11Text)) {
        usedKeysOnSheet.add(col11Text.toLowerCase());
      }
    }
  });

  if (colMap.isCustomSchema) {
    const customCols = Array.isArray(colMap.columns) ? colMap.columns : [];
    const inferredActionColObj = (activeSchemaConfig && activeSchemaConfig.actionCol)
      ? (customCols.find(c => c.colIndex === activeSchemaConfig.actionCol) || null)
      : customCols.find(c =>
          c.colIndex !== colMap.primaryCol &&
          c.colIndex !== colMap.statusCol &&
          c.colIndex !== colMap.notesCol &&
          /\b(action|subtype|issue|category|reason|request\s*type|type)\b/i.test(c.excelHeader || c.header)
        );

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return;

      const customValues = {};
      let hasAnyVisibleValue = false;
      let legacyUuidFromVisibleCol = '';
      const uuidPollutedColIndices = [];

      customCols.forEach(c => {
        let v = cellText(row, c.colIndex);
        if (v && UUID_REGEX.test(v)) {
          if (!legacyUuidFromVisibleCol) {
            legacyUuidFromVisibleCol = v.toLowerCase();
          }
          uuidPollutedColIndices.push(c.colIndex);
          v = '';
        } else if (c.colIndex === colMap.ticketCol && v) {
          v = v.toUpperCase();
        }
        customValues[`col_${c.colIndex}`] = v;
        if (v !== '') hasAnyVisibleValue = true;
      });

      let statusRaw = colMap.statusCol ? (customValues[`col_${colMap.statusCol}`] || cellText(row, colMap.statusCol)) : '';
      if (statusRaw && UUID_REGEX.test(statusRaw)) {
        if (!legacyUuidFromVisibleCol) {
          legacyUuidFromVisibleCol = statusRaw.toLowerCase();
        }
        if (colMap.statusCol && !uuidPollutedColIndices.includes(colMap.statusCol)) {
          uuidPollutedColIndices.push(colMap.statusCol);
        }
        statusRaw = '';
        if (colMap.statusCol) {
          customValues[`col_${colMap.statusCol}`] = '';
        }
      }

      let notesRaw = colMap.notesCol ? (customValues[`col_${colMap.notesCol}`] || cellText(row, colMap.notesCol)) : '';
      if (notesRaw && UUID_REGEX.test(notesRaw)) {
        if (!legacyUuidFromVisibleCol) {
          legacyUuidFromVisibleCol = notesRaw.toLowerCase();
        }
        if (colMap.notesCol && !uuidPollutedColIndices.includes(colMap.notesCol)) {
          uuidPollutedColIndices.push(colMap.notesCol);
        }
        notesRaw = '';
        if (colMap.notesCol) {
          customValues[`col_${colMap.notesCol}`] = '';
        }
      }

      if (statusRaw !== '' || notesRaw !== '') hasAnyVisibleValue = true;
      if (!hasAnyVisibleValue) {
        // Clear any stray UUID on an otherwise completely empty row
        if (uuidPollutedColIndices.length > 0) {
          uuidPollutedColIndices.forEach(idx => {
            row.getCell(idx).value = null;
          });
          row.commit();
        }
        return;
      }

      const subject = colMap.primaryCol ? (customValues[`col_${colMap.primaryCol}`] || '') : '';
      const ticket = colMap.ticketCol ? (customValues[`col_${colMap.ticketCol}`] || '').toUpperCase() : '';
      const broker = colMap.secondaryCol ? (customValues[`col_${colMap.secondaryCol}`] || '') : '';
      const action = inferredActionColObj ? (customValues[`col_${inferredActionColObj.colIndex}`] || '') : '';

      let key = cellText(row, colMap.trackingCol).toLowerCase();
      let rowMutated = false;

      if (!key || !UUID_REGEX.test(key)) {
        const cacheKey = `${workingFileHandle ? workingFileHandle.name : 'wb'}::${sheetName}::${rowNumber}`;
        const candidateRow = {
          subject,
          ticket,
          broker,
          master: '',
          bidder: '',
          action,
          yard: '',
          lots: '',
          customValues,
          isCustomSchema: true,
          key: legacyUuidFromVisibleCol || ''
        };

        let inheritedPriorKey = '';
        for (let i = allRows.length - 1; i >= 0; i--) {
          const prior = allRows[i];
          if (prior.sheetName === sheetName) continue;
          if (usedKeysOnSheet.has(prior.key)) continue;
          if (areRowsSameUniqueRequest(prior, candidateRow)) {
            inheritedPriorKey = prior.key;
            break;
          }
        }

        if (inheritedPriorKey) {
          key = inheritedPriorKey;
          legacyBackfillMap.set(cacheKey, key);
        } else if (legacyUuidFromVisibleCol) {
          key = legacyUuidFromVisibleCol;
          legacyBackfillMap.set(cacheKey, key);
        } else if (legacyBackfillMap.has(cacheKey) && !usedKeysOnSheet.has(legacyBackfillMap.get(cacheKey))) {
          key = legacyBackfillMap.get(cacheKey);
        } else {
          key = generateTrackingGuid();
          legacyBackfillMap.set(cacheKey, key);
        }

        usedKeysOnSheet.add(key);
        row.getCell(colMap.trackingCol).value = key;
        rowMutated = true;
      } else {
        usedKeysOnSheet.add(key);
      }

      // If statusRaw was blank or held a legacy UUID (from before Status was added in Col J),
      // inherit the latest real status from a prior sheet for this same Tracking ID, or fall back to defaultStatusVal
      let resolvedStatus = statusRaw;
      if (!resolvedStatus && colMap.statusCol) {
        for (let i = allRows.length - 1; i >= 0; i--) {
          const prior = allRows[i];
          if (prior.key === key && prior.status && !UUID_REGEX.test(prior.status)) {
            resolvedStatus = prior.status;
            break;
          }
        }
      }
      const finalStatus = colMap.statusCol
        ? (resolvedStatus || defaultStatusVal)
        : (resolvedStatus || 'Open');
      if (colMap.statusCol) {
        customValues[`col_${colMap.statusCol}`] = finalStatus;
      }

      // Auto-heal any visible cell (such as Col J Status) that held a retroactive UUIDv4
      if (uuidPollutedColIndices.length > 0) {
        uuidPollutedColIndices.forEach(idx => {
          row.getCell(idx).value = (idx === colMap.statusCol) ? finalStatus : '';
        });
        rowMutated = true;
      }
      if (rowMutated) {
        row.commit();
      }

      allRows.push({
        sheetName,
        rowNumber,
        subject,
        ticket,
        broker,
        master: '',
        bidder: '',
        action,
        yard: '',
        lots: '',
        notes: notesRaw,
        status: finalStatus,
        key,
        customValues,
        isCustomSchema: true
      });
    });
    return;
  }

  const coreStdSet = new Set([1, 2, 3, 4, 5, 6, 7, 8, colMap.notesCol, colMap.statusCol, colMap.trackingCol]);
  const extraStdCols = (activeSchemaConfig && activeSchemaConfig.isStandardSchema && Array.isArray(activeSchemaConfig.columns))
    ? activeSchemaConfig.columns.filter(c => !coreStdSet.has(c.colIndex))
    : [];

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;

    const subject = cellText(row, 1);
    const ticket = cellText(row, 2);
    const master = cellText(row, 4);
    const rawBroker = cellText(row, 3);
    const broker = resolveBrokerFromMaster(master, rawBroker);
    const bidder = cellText(row, 5);
    const action = cellText(row, 6);
    const yard = cellText(row, 7);
    const lots = cellText(row, 8);

    let status = cellText(row, colMap.statusCol);
    let notes = cellText(row, colMap.notesCol);

    // Auto-heal rows on a Template-layout sheet that were previously saved with Status in Col 10 and Notes in Col 9
    if (
      colMap.isTemplateLayout &&
      notes &&
      KNOWN_STATUS_VALUES.has(notes.toLowerCase()) &&
      (!status || !KNOWN_STATUS_VALUES.has(status.toLowerCase()))
    ) {
      const tmp = status;
      status = notes;
      notes = tmp;
    }

    const extraCustomValues = {};
    let hasExtraStdData = false;
    extraStdCols.forEach(c => {
      let ev = cellText(row, c.colIndex);
      if (ev && UUID_REGEX.test(ev)) ev = '';
      extraCustomValues[`col_${c.colIndex}`] = ev;
      if (ev !== '') hasExtraStdData = true;
    });

    const hasData = [subject, ticket, broker, master, bidder, action, yard, lots, notes, status]
      .some(val => val !== '') || hasExtraStdData;

    if (!hasData) {
      // Clean up any blank template row that was previously polluted by the old [object Object] bug
      const rawCell3 = row.getCell(3).value;
      if (rawCell3 === '[object Object]' && colMap.isTemplateLayout) {
        setBrokerCellWithFormula(row, rowNumber, '', '', colMap);
      }
      if (colMap.isTemplateLayout) {
        const strayCol11 = cellText(row, 11);
        if (UUID_REGEX.test(strayCol11)) {
          row.getCell(11).value = null;
        }
      }
      return;
    }

    // Keep Column C formula result synced with resolved broker
    if (colMap.isTemplateLayout) {
      setBrokerCellWithFormula(row, rowNumber, broker, master, colMap);
    }

    let key = cellText(row, colMap.trackingCol).toLowerCase();
    // Migrate UUID from Col 11 -> Col 12 if an earlier version wrote a UUID into "Sent Email Subject"
    if (!key && colMap.isTemplateLayout) {
      const col11Text = cellText(row, 11);
      if (UUID_REGEX.test(col11Text)) {
        key = col11Text.toLowerCase();
        row.getCell(11).value = null;
        row.getCell(colMap.trackingCol).value = key;
        row.commit();
      }
    }

    if (!key) {
      const cacheKey = `${workingFileHandle ? workingFileHandle.name : 'wb'}::${sheetName}::${rowNumber}`;
      const candidateRow = { subject, ticket, broker, master, bidder, action, yard, lots, key: '' };

      // Check prior indexed daily sheets (newest prior first) so carried-over rows without Col K inherit the same Tracking ID
      let inheritedPriorKey = '';
      for (let i = allRows.length - 1; i >= 0; i--) {
        const prior = allRows[i];
        if (prior.sheetName === sheetName) continue;
        if (usedKeysOnSheet.has(prior.key)) continue;
        if (areRowsSameUniqueRequest(prior, candidateRow)) {
          inheritedPriorKey = prior.key;
          break;
        }
      }

      if (inheritedPriorKey) {
        key = inheritedPriorKey;
        legacyBackfillMap.set(cacheKey, key);
      } else if (legacyBackfillMap.has(cacheKey) && !usedKeysOnSheet.has(legacyBackfillMap.get(cacheKey))) {
        key = legacyBackfillMap.get(cacheKey);
      } else {
        key = generateTrackingGuid();
        legacyBackfillMap.set(cacheKey, key);
      }

      usedKeysOnSheet.add(key);
      row.getCell(colMap.trackingCol).value = key;
      row.commit();
    } else {
      usedKeysOnSheet.add(key);
    }

    allRows.push({
      sheetName,
      rowNumber,
      subject,
      ticket,
      broker,
      master,
      bidder,
      action,
      yard,
      lots,
      notes,
      status: status || defaultStatusVal,
      key,
      ...(extraStdCols.length > 0 ? { customValues: extraCustomValues } : {})
    });
  });
}

// --- SHEET DROPDOWN, PREV/NEXT STEPPER & QUICK TAB BAR ---
function populateSheetDropdown(preserveSheet) {
  sheetSelect.innerHTML = '';
  if (dateSheets.length === 0) {
    const opt = document.createElement('option');
    opt.value = '';
    opt.textContent = 'No MM-DD-YYYY tabs yet';
    sheetSelect.appendChild(opt);
    return;
  }
  const todayTime = parseDateTab(getTodayFormatted());
  dateSheets.forEach(name => {
    const opt = document.createElement('option');
    opt.value = name;
    const isFuture = parseDateTab(name) > todayTime;
    opt.textContent = isFuture ? `${name} (Future)` : name;
    sheetSelect.appendChild(opt);
  });
  if (preserveSheet && dateSheets.includes(preserveSheet)) {
    sheetSelect.value = preserveSheet;
  } else {
    sheetSelect.value = dateSheets[0];
  }
}

function updateSheetStepperState() {
  if (dateSheets.length <= 1) {
    btnPrevSheet.disabled = true;
    btnNextSheet.disabled = true;
    return;
  }
  const currentIdx = dateSheets.indexOf(sheetSelect.value);
  // dateSheets is sorted newest first (0 = newest, length - 1 = oldest)
  // Prev (◀) steps to older tab (currentIdx + 1); Next (▶) steps to newer tab (currentIdx - 1)
  btnPrevSheet.disabled = currentIdx === -1 || currentIdx >= dateSheets.length - 1;
  btnNextSheet.disabled = currentIdx <= 0;

  if (!btnPrevSheet.disabled) {
    btnPrevSheet.title = `Older Day Tab: ${dateSheets[currentIdx + 1]}`;
  } else {
    btnPrevSheet.title = 'Oldest Day Tab reached';
  }

  if (!btnNextSheet.disabled) {
    btnNextSheet.title = `Newer Day Tab: ${dateSheets[currentIdx - 1]}`;
  } else {
    btnNextSheet.title = 'Newest Day Tab reached';
  }
}

// Natural Filter Auto-Clearing when navigating tabs, steppers, scope, or workbooks
function clearTransientFiltersOnNavigation({ resetScopeToCurrent = true } = {}) {
  activeKeyFilter = null;
  activeFieldFilter = null;
  if (searchInput && searchInput.value) {
    searchInput.value = '';
  }
  if (resetScopeToCurrent && scopeSelect && scopeSelect.value !== 'current') {
    scopeSelect.value = 'current';
  }
  if (selectedForDelete.size > 0) {
    selectedForDelete.clear();
    updateDeleteSelectionUI();
  }
}

function renderSheetTabsBar() {
  if (dateSheets.length === 0) {
    sheetTabsBar.style.display = 'none';
    sheetTabsBar.innerHTML = '';
    return;
  }

  sheetTabsBar.style.display = 'flex';
  sheetTabsBar.innerHTML = '';
  const currentTab = sheetSelect.value;
  const todayTime = parseDateTab(getTodayFormatted());

  // Invert tab strip order (oldest on left -> newest on right) so ◀ (Older) moves left and ▶ (Newer) moves right
  const chronologicalTabs = [...dateSheets].reverse();
  let activePillEl = null;

  chronologicalTabs.forEach(tabName => {
    const count = allRows.filter(r => r.sheetName === tabName).length;
    const isFuture = parseDateTab(tabName) > todayTime;
    const isActive = tabName === currentTab;
    const pill = document.createElement('button');
    pill.type = 'button';
    pill.className = 'sheet-tab-pill' + (isActive ? ' active' : '');
    pill.innerHTML = `
      <span>${escapeHtml(tabName)}${isFuture ? ' • Future' : ''}</span>
      <span class="sheet-tab-count">${count}</span>
    `;
    pill.addEventListener('click', () => {
      clearTransientFiltersOnNavigation({ resetScopeToCurrent: true });
      sheetSelect.value = tabName;
      onActiveSheetChanged();
    });
    sheetTabsBar.appendChild(pill);
    if (isActive) activePillEl = pill;
  });

  if (activePillEl && typeof activePillEl.scrollIntoView === 'function') {
    activePillEl.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'nearest' });
  }
}

function onActiveSheetChanged() {
  updateSheetStepperState();
  renderSheetTabsBar();
  if (!dockEditingTarget) {
    if (typeof hasUnsavedNewDockDraft === 'function' && hasUnsavedNewDockDraft()) {
      saveNewDockDraftToStorage();
    }
    setBottomDockNewMode(false, { restoreSavedDraft: true });
  }
  renderList();
}

btnPrevSheet.addEventListener('click', () => {
  const idx = dateSheets.indexOf(sheetSelect.value);
  if (idx !== -1 && idx < dateSheets.length - 1) {
    clearTransientFiltersOnNavigation({ resetScopeToCurrent: true });
    sheetSelect.value = dateSheets[idx + 1];
    onActiveSheetChanged();
  }
});

btnNextSheet.addEventListener('click', () => {
  const idx = dateSheets.indexOf(sheetSelect.value);
  if (idx > 0) {
    clearTransientFiltersOnNavigation({ resetScopeToCurrent: true });
    sheetSelect.value = dateSheets[idx - 1];
    onActiveSheetChanged();
  }
});

sheetSelect.addEventListener('change', () => {
  clearTransientFiltersOnNavigation({ resetScopeToCurrent: true });
  onActiveSheetChanged();
});

// --- 5. DYNAMIC TABLE HEADER, LIST VIEW, GLOBAL SEARCH & BATCH SELECTION ---
function getRowCompositeId(sheetName, rowNumber) {
  return `${sheetName}::${rowNumber}`;
}

function getVisibleCustomTableColumns() {
  if (!activeSchemaConfig || !Array.isArray(activeSchemaConfig.columns)) return [];
  const visible = activeSchemaConfig.columns.filter(c => c.showInTable !== false);
  return visible.length > 0 ? visible : activeSchemaConfig.columns;
}

function getStandardTableColumnsDescriptor() {
  const cols = (activeSchemaConfig && Array.isArray(activeSchemaConfig.columns) && activeSchemaConfig.columns.length > 0)
    ? activeSchemaConfig.columns
    : [];
  const byIdx = new Map();
  cols.forEach(c => byIdx.set(c.colIndex, c));
  const stColIdx = (activeSchemaConfig && activeSchemaConfig.statusCol) || 10;
  const nColIdx = (activeSchemaConfig && activeSchemaConfig.notesCol) || 9;
  const trkColIdx = (activeSchemaConfig && activeSchemaConfig.trackingCol) || 11;
  const isVis = (idx, defaultVis = true) => {
    const c = byIdx.get(idx);
    return c ? (c.showInTable !== false) : defaultVis;
  };
  const getLbl = (idx, fallback) => {
    const c = byIdx.get(idx);
    return (c && c.header) ? c.header : fallback;
  };
  const getCompactMasterBidderLbl = (idx, fallback) => {
    const c = byIdx.get(idx);
    if (c && c.label) return c.label;
    const raw = (c && c.header) ? c.header : fallback;
    const stripped = String(raw || '').replace(/^broker\s+/i, '').trim();
    return stripped || fallback;
  };

  const slots = [];
  if (isVis(1, true)) slots.push({ kind: 'subject', colIndex: 1, sortKey: 'subject', label: getLbl(1, 'Subject') });
  if (isVis(2, true)) slots.push({ kind: 'ticket', colIndex: 2, sortKey: 'ticket', label: getLbl(2, 'Ticket #') });
  if (isVis(3, true)) slots.push({ kind: 'broker', colIndex: 3, sortKey: 'broker', label: getLbl(3, 'Broker') });

  const showMaster = isVis(4, true);
  const showBidder = isVis(5, true);
  if (showMaster || showBidder) {
    slots.push({
      kind: 'master_bidder',
      showMaster,
      showBidder,
      masterLabel: getCompactMasterBidderLbl(4, 'Master #'),
      bidderLabel: getCompactMasterBidderLbl(5, 'Bidder #')
    });
  }

  if (isVis(6, true)) slots.push({ kind: 'action', colIndex: 6, sortKey: 'action', label: getLbl(6, 'Action') });

  const showYard = isVis(7, true);
  const showLots = isVis(8, true);
  if (showYard || showLots) {
    slots.push({
      kind: 'yard_lots',
      showYard,
      showLots,
      yardLabel: getLbl(7, 'Yard'),
      lotsLabel: getLbl(8, 'Lots')
    });
  }

  if (isVis(nColIdx, true)) slots.push({ kind: 'notes', colIndex: nColIdx, sortKey: 'notes', label: getLbl(nColIdx, 'Action Notes') });
  if (isVis(stColIdx, true)) slots.push({ kind: 'status', colIndex: stColIdx, sortKey: 'status', label: getLbl(stColIdx, 'Status') });

  const coreStdSet = new Set([1, 2, 3, 4, 5, 6, 7, 8, nColIdx, stColIdx, trkColIdx]);
  cols.forEach(c => {
    if (!coreStdSet.has(c.colIndex) && c.showInTable === true) {
      slots.push({
        kind: 'extra',
        colIndex: c.colIndex,
        colObj: c,
        sortKey: `col_${c.colIndex}`,
        label: c.header
      });
    }
  });

  if (slots.length === 0) {
    slots.push({ kind: 'subject', colIndex: 1, sortKey: 'subject', label: getLbl(1, 'Subject') });
  }
  return { slots, stColIdx, nColIdx };
}

function getActiveTableColSpan() {
  if (activeSchemaConfig && !activeSchemaConfig.isStandardSchema && Array.isArray(activeSchemaConfig.columns) && activeSchemaConfig.columns.length > 0) {
    return 4 + getVisibleCustomTableColumns().length;
  }
  return 4 + getStandardTableColumnsDescriptor().slots.length;
}

let trackerColResizeState = null; // { colKey, thEl, resizerEl, tableEl, startX, startWidthPx, currentWidthPx } | null
let lastTrackerColResizeAt = 0;

function getActiveTrackerTableColWidthsKey() {
  const profId = (typeof activeProfileId === 'string' && activeProfileId.trim()) ? activeProfileId.trim() : 'default';
  const seriesKey = (typeof getWorkbookSeriesKey === 'function' && workingFileHandle && workingFileHandle.name)
    ? getWorkbookSeriesKey(workingFileHandle.name)
    : 'default';
  const modeKey = (activeSchemaConfig && !activeSchemaConfig.isStandardSchema) ? 'custom' : 'standard';
  return `${profId}::${seriesKey}::${modeKey}`;
}

function getSavedTrackerTableColWidths() {
  if (typeof tableColWidthsStore !== 'object' || !tableColWidthsStore) return {};
  const key = getActiveTrackerTableColWidthsKey();
  const entry = tableColWidthsStore[key];
  return (entry && typeof entry === 'object' && !Array.isArray(entry)) ? entry : {};
}

function setSavedTrackerTableColWidth(colKey, widthPx) {
  if (!colKey || typeof tableColWidthsStore !== 'object' || !tableColWidthsStore) return;
  const storeKey = getActiveTrackerTableColWidthsKey();
  if (!tableColWidthsStore[storeKey] || typeof tableColWidthsStore[storeKey] !== 'object') {
    tableColWidthsStore[storeKey] = {};
  }
  if (widthPx === null || widthPx === undefined) {
    delete tableColWidthsStore[storeKey][colKey];
    if (Object.keys(tableColWidthsStore[storeKey]).length === 0) {
      delete tableColWidthsStore[storeKey];
    }
  } else {
    const clamped = Math.max(64, Math.min(680, Math.round(Number(widthPx) || 0)));
    tableColWidthsStore[storeKey][colKey] = clamped;
  }
  if (typeof saveTableColWidthsStore === 'function') {
    saveTableColWidthsStore();
  }
}

function applySavedTrackerTableColumnWidths() {
  if (!tableHead) return;
  const tableEl = tableHead.closest('table');
  const saved = getSavedTrackerTableColWidths();
  const savedKeys = Object.keys(saved);
  const hasAny = savedKeys.length > 0;
  if (tableEl) {
    tableEl.classList.toggle('has-custom-col-widths', hasAny);
  }
  tableHead.querySelectorAll('th[data-col-width-key]').forEach(th => {
    const k = th.dataset.colWidthKey;
    const w = saved[k];
    if (Number.isFinite(Number(w)) && Number(w) >= 64) {
      const clamped = Math.max(64, Math.min(680, Math.round(Number(w))));
      th.style.width = `${clamped}px`;
      th.style.minWidth = `${clamped}px`;
    } else {
      th.style.width = '';
      th.style.minWidth = '';
    }
  });
}

function resetActiveTrackerTableColWidths() {
  if (typeof closeToolsMenu === 'function') closeToolsMenu();
  const storeKey = getActiveTrackerTableColWidthsKey();
  if (typeof tableColWidthsStore === 'object' && tableColWidthsStore && tableColWidthsStore[storeKey]) {
    delete tableColWidthsStore[storeKey];
    if (typeof saveTableColWidthsStore === 'function') saveTableColWidthsStore();
    applySavedTrackerTableColumnWidths();
    showToast('Reset Daily Tracker table column widths to automatic defaults.', 'info');
  } else {
    applySavedTrackerTableColumnWidths();
    showToast('Daily Tracker columns are already at default widths.', 'info');
  }
}

function autoFitTrackerTableColumnWidth(thEl) {
  if (!thEl || !tableHead) return;
  const colKey = thEl.dataset.colWidthKey;
  if (!colKey) return;
  const hdrRow = tableHead.rows && tableHead.rows[0];
  if (!hdrRow) return;
  const cellIndex = Array.from(hdrRow.cells).indexOf(thEl);
  if (cellIndex < 0) return;

  let maxChars = (thEl.textContent || '').trim().length + 4;
  if (listBody) {
    listBody.querySelectorAll('tr.tracker-main-row').forEach(tr => {
      const td = tr.cells[cellIndex];
      if (!td) return;
      const txt = (td.textContent || '').replace(/\s+/g, ' ').trim();
      const firstLineLen = Math.min(80, txt.length);
      if (firstLineLen > maxChars) maxChars = firstLineLen;
    });
  }
  const fittedPx = Math.max(88, Math.min(460, Math.round(maxChars * 7.6 + 34)));
  const saved = getSavedTrackerTableColWidths();
  // Double-clicking an already auto-fitted column toggles it back to automatic default width
  if (saved[colKey] && Math.abs(saved[colKey] - fittedPx) <= 6) {
    setSavedTrackerTableColWidth(colKey, null);
  } else {
    setSavedTrackerTableColWidth(colKey, fittedPx);
  }
  applySavedTrackerTableColumnWidths();
}

function renderTableHeader() {
  if (!tableHead) return;
  const selectColClass = isDeleteMode ? 'delete-col' : 'select-col';
  const resizerSpan = (colKey) => `<span class="tracker-col-resizer" data-resize-col-key="${escapeHtml(colKey)}" title="Drag to resize column • Double-click to auto-fit"></span>`;

  if (activeSchemaConfig && !activeSchemaConfig.isStandardSchema && Array.isArray(activeSchemaConfig.columns) && activeSchemaConfig.columns.length > 0) {
    const visibleCols = getVisibleCustomTableColumns();
    const visibleSortKeys = new Set(visibleCols.map(c => `col_${c.colIndex}`));
    const hasCustomStatusCol = Boolean(
      activeSchemaConfig.statusCol &&
      visibleCols.some(c => c.colIndex === activeSchemaConfig.statusCol)
    );
    if (sortState.column && !['date', 'key', 'accessed'].includes(sortState.column) && !visibleSortKeys.has(sortState.column)) {
      sortState = { column: null, direction: 'asc' };
    }
    const customThsHtml = visibleCols.map(c => {
      const sortKey = `col_${c.colIndex}`;
      if (hasCustomStatusCol && c.colIndex === activeSchemaConfig.statusCol) {
        return `
          <th class="sortable-th tracker-status-th" data-sort="${sortKey}" data-col-width-key="${sortKey}" title="Click ${escapeHtml(c.header)} or Accessed to sort">
            <span class="th-stacked-sub-sorts"><span class="th-sub-sort" data-sort="${sortKey}" title="Sort by ${escapeHtml(c.header)}">${escapeHtml(c.header)}<span class="sort-indicator" data-ind="${sortKey}">↕</span></span><span class="th-sub-sort th-sub-sort-accessed" data-sort="accessed" title="Sort by Last Accessed Time">Accessed<span class="sort-indicator" data-ind="accessed">↕</span></span></span>${resizerSpan(sortKey)}
          </th>
        `;
      }
      return `
        <th class="sortable-th" data-sort="${sortKey}" data-col-width-key="${sortKey}" title="Click to sort by ${escapeHtml(c.header)}">
          <span>${escapeHtml(c.header)}</span><span class="sort-indicator" data-ind="${sortKey}">↕</span>${resizerSpan(sortKey)}
        </th>
      `;
    }).join('');

    const customActionsThHtml = hasCustomStatusCol
      ? `<th class="tracker-actions-th" data-col-width-key="actions" title="Row Actions">${resizerSpan('actions')}</th>`
      : `<th class="sortable-th tracker-actions-th" data-sort="accessed" data-col-width-key="actions" title="Click to sort by Last Accessed Time"><span>Accessed</span><span class="sort-indicator" data-ind="accessed">↕</span>${resizerSpan('actions')}</th>`;

    tableHead.innerHTML = `
      <tr>
        <th class="tracker-row-num-th" title="Line / Row Number">#</th>
        <th id="thDeleteSelect" class="${selectColClass}" title="Select rows for batch status update or deletion">
          <input type="checkbox" id="chkSelectAllHeader" title="Select all visible requests">
        </th>
        <th class="sortable-th" data-sort="date" data-col-width-key="date" title="Click to sort by Date Tab">
          <span>Date Tab</span><span class="sort-indicator" data-ind="date">↕</span>${resizerSpan('date')}
        </th>
        ${customThsHtml}
        ${customActionsThHtml}
      </tr>
    `;
    syncSelectAllCheckboxes();
    applySavedTrackerTableColumnWidths();
    return;
  }

  const descriptor = getStandardTableColumnsDescriptor();
  const hasStdStatusSlot = descriptor.slots.some(s => s.kind === 'status');
  const validStdSorts = new Set(['date', 'key', 'accessed']);
  descriptor.slots.forEach(s => {
    if (s.kind === 'master_bidder') {
      if (s.showMaster) validStdSorts.add('master');
      if (s.showBidder) validStdSorts.add('bidder');
    } else if (s.kind === 'yard_lots') {
      if (s.showYard) validStdSorts.add('yard');
      if (s.showLots) validStdSorts.add('lots');
    } else if (s.sortKey) {
      validStdSorts.add(s.sortKey);
    }
  });
  if (sortState.column && !validStdSorts.has(sortState.column)) {
    sortState = { column: null, direction: 'asc' };
  }

  const stdThsHtml = descriptor.slots.map(s => {
    if (s.kind === 'master_bidder') {
      if (s.showMaster && s.showBidder) {
        return `
          <th class="sortable-th tracker-master-bidder-th col-master-bidder" data-sort="master" data-col-width-key="master_bidder" title="Click ${escapeHtml(s.masterLabel)} or ${escapeHtml(s.bidderLabel)} to sort">
            <span class="th-stacked-sub-sorts"><span class="th-sub-sort" data-sort="master" title="Sort by ${escapeHtml(s.masterLabel)}">${escapeHtml(s.masterLabel)}<span class="sort-indicator" data-ind="master">↕</span></span><span class="th-sub-sort" data-sort="bidder" title="Sort by ${escapeHtml(s.bidderLabel)}">${escapeHtml(s.bidderLabel)}<span class="sort-indicator" data-ind="bidder">↕</span></span></span>${resizerSpan('master_bidder')}
          </th>
        `;
      }
      const key = s.showMaster ? 'master' : 'bidder';
      const lbl = s.showMaster ? s.masterLabel : s.bidderLabel;
      return `
        <th class="sortable-th tracker-master-bidder-th col-master-bidder" data-sort="${key}" data-col-width-key="${key}" title="Click to sort by ${escapeHtml(lbl)}">
          <span>${escapeHtml(lbl)}</span><span class="sort-indicator" data-ind="${key}">↕</span>${resizerSpan(key)}
        </th>
      `;
    }
    if (s.kind === 'yard_lots') {
      if (s.showYard && s.showLots) {
        return `
          <th class="sortable-th" data-sort="yard" data-col-width-key="yard_lots" title="Click ${escapeHtml(s.yardLabel)} or ${escapeHtml(s.lotsLabel)} to sort">
            <span class="th-sub-sort" data-sort="yard" title="Sort by ${escapeHtml(s.yardLabel)}">${escapeHtml(s.yardLabel)}<span class="sort-indicator" data-ind="yard">↕</span></span>
            <span class="th-sep">/</span>
            <span class="th-sub-sort" data-sort="lots" title="Sort by ${escapeHtml(s.lotsLabel)}">${escapeHtml(s.lotsLabel)}<span class="sort-indicator" data-ind="lots">↕</span></span>${resizerSpan('yard_lots')}
          </th>
        `;
      }
      const key = s.showYard ? 'yard' : 'lots';
      const lbl = s.showYard ? s.yardLabel : s.lotsLabel;
      return `
        <th class="sortable-th" data-sort="${key}" data-col-width-key="${key}" title="Click to sort by ${escapeHtml(lbl)}">
          <span>${escapeHtml(lbl)}</span><span class="sort-indicator" data-ind="${key}">↕</span>${resizerSpan(key)}
        </th>
      `;
    }
    const widthKey = s.sortKey || s.kind || 'col';
    if (s.kind === 'status') {
      return `
        <th class="sortable-th tracker-status-th" data-sort="${s.sortKey}" data-col-width-key="${widthKey}" title="Click ${escapeHtml(s.label)} or Accessed to sort">
          <span class="th-stacked-sub-sorts"><span class="th-sub-sort" data-sort="${s.sortKey}" title="Sort by ${escapeHtml(s.label)}">${escapeHtml(s.label)}<span class="sort-indicator" data-ind="${s.sortKey}">↕</span></span><span class="th-sub-sort th-sub-sort-accessed" data-sort="accessed" title="Sort by Last Accessed Time">Accessed<span class="sort-indicator" data-ind="accessed">↕</span></span></span>${resizerSpan(widthKey)}
        </th>
      `;
    }
    const extraCls = s.kind === 'subject' ? ' tracker-subject-th' : '';
    return `
      <th class="sortable-th${extraCls}" data-sort="${s.sortKey}" data-col-width-key="${widthKey}" title="Click to sort by ${escapeHtml(s.label)}">
        <span>${escapeHtml(s.label)}</span><span class="sort-indicator" data-ind="${s.sortKey}">↕</span>${resizerSpan(widthKey)}
      </th>
    `;
  }).join('');

  const stdActionsThHtml = hasStdStatusSlot
    ? `<th class="tracker-actions-th" data-col-width-key="actions" title="Row Actions">${resizerSpan('actions')}</th>`
    : `<th class="sortable-th tracker-actions-th" data-sort="accessed" data-col-width-key="actions" title="Click to sort by Last Accessed Time"><span>Accessed</span><span class="sort-indicator" data-ind="accessed">↕</span>${resizerSpan('actions')}</th>`;

  tableHead.innerHTML = `
    <tr>
      <th class="tracker-row-num-th" title="Line / Row Number">#</th>
      <th id="thDeleteSelect" class="${selectColClass}" title="Select rows for batch status update or deletion">
        <input type="checkbox" id="chkSelectAllHeader" title="Select all visible requests">
      </th>
      <th class="sortable-th" data-sort="date" data-col-width-key="date" title="Click to sort by Date Tab">
        <span>Date Tab</span><span class="sort-indicator" data-ind="date">↕</span>${resizerSpan('date')}
      </th>
      ${stdThsHtml}
      ${stdActionsThHtml}
    </tr>
  `;
  syncSelectAllCheckboxes();
  applySavedTrackerTableColumnWidths();
}

function createShadowNewRowElement(colSpan, currentSheet) {
  if (!workbook || isDeleteMode || !currentSheet || !DATE_TAB_REGEX.test(currentSheet)) {
    return null;
  }
  const isCustom = activeSchemaConfig && !activeSchemaConfig.isStandardSchema && activeSchemaConfig.columns.length > 0;
  const ghostPillsHtml = isCustom
    ? getVisibleCustomTableColumns().slice(0, 5).map(c => `<span class="shadow-ghost-pill">${escapeHtml(c.header)}</span>`).join('')
    : getStandardTableColumnsDescriptor().slots
        .filter(s => s.kind !== 'status' && s.kind !== 'notes')
        .slice(0, 5)
        .map(s => {
          if (s.kind === 'master_bidder') {
            const lbl = (s.showMaster && s.showBidder) ? `${s.masterLabel} & ${s.bidderLabel}` : (s.showMaster ? s.masterLabel : s.bidderLabel);
            return `<span class="shadow-ghost-pill">${escapeHtml(lbl)}</span>`;
          }
          if (s.kind === 'yard_lots') {
            const lbl = (s.showYard && s.showLots) ? `${s.yardLabel} / ${s.lotsLabel}` : (s.showYard ? s.yardLabel : s.lotsLabel);
            return `<span class="shadow-ghost-pill">${escapeHtml(lbl)}</span>`;
          }
          return `<span class="shadow-ghost-pill">${escapeHtml(s.label)}</span>`;
        })
        .join('');

  const tr = document.createElement('tr');
  tr.className = 'shadow-new-row';
  tr.title = `Click to add a new request to ${currentSheet}`;
  tr.innerHTML = `
    <td colspan="${colSpan}">
      <div class="shadow-row-inner">
        <div class="shadow-row-left">
          <span class="shadow-plus-btn" aria-hidden="true">＋</span>
          <span class="shadow-row-label">Add new request to <strong>${escapeHtml(currentSheet)}</strong>...</span>
          <div class="shadow-ghost-pills">
            ${ghostPillsHtml}
          </div>
        </div>
        <span class="shadow-action-pill">＋ New Request</span>
      </div>
    </td>
  `;
  return tr;
}

const SORT_COLUMN_LABELS = {
  date: 'Date Tab',
  key: 'Tracking ID',
  subject: 'Subject',
  ticket: 'Ticket #',
  broker: 'Broker',
  master: 'Master #',
  bidder: 'Bidder #',
  action: 'Requested Action',
  yard: 'Yard #',
  lots: 'Lot Numbers',
  notes: 'Action Notes',
  status: 'Status',
  accessed: 'Last Accessed'
};

function getSortColumnLabel(colKey) {
  if (!colKey) return '';
  if (SORT_COLUMN_LABELS[colKey]) return SORT_COLUMN_LABELS[colKey];
  if (colKey.startsWith('col_') && activeSchemaConfig && Array.isArray(activeSchemaConfig.columns)) {
    const colIdx = parseInt(colKey.replace('col_', ''), 10);
    const hit = activeSchemaConfig.columns.find(c => c.colIndex === colIdx);
    if (hit) return hit.header;
  }
  return colKey;
}

function extractLotTokens(lotsStr) {
  if (!lotsStr) return [];
  return String(lotsStr)
    .split(/[,;/|\s]+/)
    .map(s => s.trim())
    .filter(s => s.length > 0 && s !== '—');
}

function isFieldFilterActive(field, val) {
  if (!val) return false;
  if (field === 'key' && activeKeyFilter && activeKeyFilter.toLowerCase() === String(val).trim().toLowerCase()) {
    return true;
  }
  if (!activeFieldFilter) return false;
  return (
    activeFieldFilter.field === field &&
    activeFieldFilter.value.toLowerCase() === String(val).trim().toLowerCase()
  );
}

function renderClickableLotsHtml(lotsStr, includeCopy = true) {
  const trimmed = String(lotsStr || '').trim();
  if (!trimmed || trimmed === '—') return '—';

  const copyBtnFn = (includeCopy && typeof renderInlineCopyBtnHtml === 'function') ? renderInlineCopyBtnHtml : () => '';
  const parts = trimmed.split(/([,;/|\s]+)/);
  return parts
    .map(part => {
      if (!part || /^[,;/|\s]+$/.test(part)) {
        return escapeHtml(part);
      }
      const activeClass = isFieldFilterActive('lot', part) ? ' active-filter-match' : '';
      const copyHtml = copyBtnFn(part, 'Lot #');
      const chipHtml = `<span class="clickable-code${activeClass}" data-filter-field="lot" data-filter-label="Lot #" data-filter-val="${escapeHtml(part)}" title="Click to show all requests with Lot # ${escapeHtml(part)}">${escapeHtml(part)}</span>`;
      return copyHtml ? `<span class="cell-copy-group">${chipHtml}${copyHtml}</span>` : chipHtml;
    })
    .join('');
}

function renderClickableTicketTokensHtml(ticketStr, filterField = 'ticket', filterLabel = 'Ticket #', includeCopy = true) {
  const trimmed = String(ticketStr || '').trim();
  if (!trimmed || trimmed === '—') return '—';
  const copyBtnFn = (includeCopy && typeof renderInlineCopyBtnHtml === 'function') ? renderInlineCopyBtnHtml : () => '';
  const tokens = splitNormalizedTicketTokens(trimmed);
  if (tokens.length <= 1) {
    const activeClass = isFieldFilterActive(filterField, trimmed) ? ' active-filter-match' : '';
    const copyHtml = copyBtnFn(trimmed, filterLabel);
    const chipHtml = `<span class="clickable-code${activeClass}" data-filter-field="${escapeHtml(filterField)}" data-filter-label="${escapeHtml(filterLabel)}" data-filter-val="${escapeHtml(trimmed)}" title="Click to show all requests with ${escapeHtml(filterLabel)}: ${escapeHtml(trimmed)}">${escapeHtml(trimmed)}</span>`;
    return copyHtml ? `<span class="cell-copy-group">${chipHtml}${copyHtml}</span>` : chipHtml;
  }
  return tokens
    .map(tok => {
      const activeClass = isFieldFilterActive(filterField, tok) ? ' active-filter-match' : '';
      const copyHtml = copyBtnFn(tok, filterLabel);
      const chipHtml = `<span class="clickable-code${activeClass}" data-filter-field="${escapeHtml(filterField)}" data-filter-label="${escapeHtml(filterLabel)}" data-filter-val="${escapeHtml(tok)}" title="Click to show all requests with ${escapeHtml(filterLabel)}: ${escapeHtml(tok)}">${escapeHtml(tok)}</span>`;
      return copyHtml ? `<span class="cell-copy-group">${chipHtml}${copyHtml}</span>` : chipHtml;
    })
    .join(', ');
}

function doesItemMatchFieldFilter(item, filter) {
  if (!filter || !filter.value) return true;
  const target = filter.value.trim().toLowerCase();
  if (filter.field && filter.field.startsWith('col_')) {
    const rawVal = item.customValues ? (item.customValues[filter.field] || '') : '';
    if (activeSchemaConfig && `col_${activeSchemaConfig.ticketCol}` === filter.field) {
      const tokens = splitNormalizedTicketTokens(rawVal).map(t => t.toLowerCase());
      return tokens.includes(target) || rawVal.trim().toLowerCase() === target;
    }
    return String(rawVal).trim().toLowerCase() === target;
  }

  switch (filter.field) {
    case 'key':
      return item.key.trim().toLowerCase() === target;
    case 'master':
      return item.master.trim().toLowerCase() === target;
    case 'bidder':
      return item.bidder.trim().toLowerCase() === target;
    case 'yard':
      return item.yard.trim().toLowerCase() === target;
    case 'lot': {
      const tokens = extractLotTokens(item.lots).map(t => t.toLowerCase());
      return tokens.includes(target) || item.lots.trim().toLowerCase() === target;
    }
    case 'ticket': {
      const tTokens = splitNormalizedTicketTokens(item.ticket).map(t => t.toLowerCase());
      return tTokens.includes(target) || item.ticket.trim().toLowerCase() === target;
    }
    case 'broker':
      return item.broker.trim().toLowerCase() === target;
    case 'action':
      return item.action.trim().toLowerCase() === target;
    case 'status':
      return item.status.trim().toLowerCase() === target;
    case 'subject':
      return item.subject.trim().toLowerCase() === target;
    default:
      return true;
  }
}

function toggleFieldFilter(field, label, rawValue) {
  const val = String(rawValue || '').trim();
  if (!val || val === '—') return;

  // Naturally clear search box when clicking a Like-Item or Key filter chip
  if (searchInput && searchInput.value) {
    searchInput.value = '';
  }

  if (
    activeFieldFilter &&
    activeFieldFilter.field === field &&
    activeFieldFilter.value.toLowerCase() === val.toLowerCase()
  ) {
    activeFieldFilter = null;
    activeKeyFilter = null;
  } else {
    activeFieldFilter = { field, label, value: val };
    activeKeyFilter = field === 'key' ? val : null;
  }
  renderList();
}

function getItemSortCellValue(item, col) {
  if (col && col.startsWith('col_')) {
    return String((item.customValues && item.customValues[col]) || '').trim();
  }
  return String(item[col] || '').trim();
}

function sortRowsList(rows) {
  if (!sortState.column) return rows;
  const col = sortState.column;
  const dirMultiplier = sortState.direction === 'asc' ? 1 : -1;

  return [...rows].sort((a, b) => {
    if (col === 'date') {
      const diff = parseDateTab(a.sheetName) - parseDateTab(b.sheetName);
      const primary = diff !== 0 ? diff : a.rowNumber - b.rowNumber;
      return primary * dirMultiplier;
    }

    if (col === 'accessed') {
      const accA = typeof getRowLastAccessedMs === 'function' ? getRowLastAccessedMs(a.key) : 0;
      const accB = typeof getRowLastAccessedMs === 'function' ? getRowLastAccessedMs(b.key) : 0;
      if (!accA && !accB) {
        return parseDateTab(b.sheetName) - parseDateTab(a.sheetName) || (a.rowNumber - b.rowNumber);
      }
      if (!accA) return 1;
      if (!accB) return -1;
      const diff = accA - accB;
      return (diff !== 0 ? diff : a.rowNumber - b.rowNumber) * dirMultiplier;
    }

    const valA = getItemSortCellValue(a, col);
    const valB = getItemSortCellValue(b, col);

    // Always push empty cells to the bottom regardless of sort direction
    if (!valA && !valB) {
      return parseDateTab(b.sheetName) - parseDateTab(a.sheetName) || (a.rowNumber - b.rowNumber);
    }
    if (!valA) return 1;
    if (!valB) return -1;

    const cmp = valA.localeCompare(valB, undefined, { numeric: true, sensitivity: 'base' });
    if (cmp !== 0) return cmp * dirMultiplier;

    return parseDateTab(b.sheetName) - parseDateTab(a.sheetName) || (a.rowNumber - b.rowNumber);
  });
}

function updateSortAndFilterUI(matchCount) {
  const activeFilter = activeFieldFilter || (activeKeyFilter ? { field: 'key', label: 'Tracking ID', value: activeKeyFilter } : null);
  const hasSearch = searchInput.value.trim().length > 0;
  const hasSort = Boolean(sortState.column);

  // 1. Update Sort Header Indicators
  if (tableHead) {
    tableHead.querySelectorAll('th.sortable-th').forEach(th => {
      const thSort = th.dataset.sort;
      const isCompositeMatch =
        (thSort === 'master' && (sortState.column === 'master' || sortState.column === 'bidder')) ||
        (thSort === 'yard' && (sortState.column === 'yard' || sortState.column === 'lots')) ||
        (thSort === 'accessed' && (sortState.column === 'accessed' || sortState.column === 'key')) ||
        (th.classList.contains('tracker-status-th') && sortState.column === 'accessed');
      const isDirectMatch = thSort === sortState.column;
      th.classList.toggle('sorted-active', isDirectMatch || isCompositeMatch);
    });

    tableHead.querySelectorAll('.th-sub-sort').forEach(sub => {
      sub.classList.toggle('active', sub.dataset.sort === sortState.column);
    });

    tableHead.querySelectorAll('.sort-indicator').forEach(ind => {
      const indCol = ind.dataset.ind;
      if (indCol && indCol === sortState.column) {
        ind.textContent = sortState.direction === 'asc' ? '▲' : '▼';
        ind.classList.add('active');
      } else {
        ind.textContent = '↕';
        ind.classList.remove('active');
      }
    });
  }

  // 2. Update Controls Bar "Clear Sort / Filter" Button
  if (btnClearSortFilter) {
    const anyActive = Boolean(activeFilter || hasSearch || hasSort);
    btnClearSortFilter.style.display = anyActive ? 'inline-block' : 'none';
  }

  // 3. Update Filter / Sort Banner above Table
  if (activeFilter || hasSort) {
    filterBanner.style.display = 'flex';
    if (activeFilter) {
      filterBannerTitle.textContent = activeFilter.field === 'key' ? 'History Trace Active:' : 'Like-Item Filter Active:';
      filterBannerDesc.textContent = `Showing all daily entries matching ${activeFilter.label}`;
      filterKeyText.textContent = `${activeFilter.value} (${matchCount} match${matchCount === 1 ? '' : 'es'})`;
      filterKeyText.style.display = 'inline-block';
      if (hasSort) {
        const sortLabel = getSortColumnLabel(sortState.column);
        const arrow = sortState.direction === 'asc' ? '▲ Asc' : '▼ Desc';
        filterBannerDesc.textContent += ` • Sorted by ${sortLabel} (${arrow}) —`;
      }
    } else if (hasSort) {
      const sortLabel = getSortColumnLabel(sortState.column);
      const arrow = sortState.direction === 'asc' ? '▲ Ascending' : '▼ Descending';
      filterBannerTitle.textContent = 'Column Sort Active:';
      filterBannerDesc.textContent = `Sorting visible requests by`;
      filterKeyText.textContent = `${sortLabel} (${arrow})`;
      filterKeyText.style.display = 'inline-block';
    }
  } else {
    filterBanner.style.display = 'none';
  }
}

function clearAllSortAndFilters() {
  activeKeyFilter = null;
  activeFieldFilter = null;
  sortState = { column: null, direction: 'asc' };
  if (searchInput.value) {
    searchInput.value = '';
  }
  renderList();
}

function buildItemSearchHaystack(item) {
  const customParts = item.customValues ? Object.values(item.customValues) : [];
  let accessedText = '';
  if (typeof getRowLastAccessedMs === 'function' && typeof formatRowLastAccessedLabel === 'function') {
    const accMs = getRowLastAccessedMs(item.key);
    if (accMs > 0) {
      const fmt = formatRowLastAccessedLabel(accMs);
      accessedText = `${fmt.short} ${fmt.full}`;
    }
  }
  return [
    item.subject,
    item.ticket,
    item.broker,
    item.master,
    item.bidder,
    item.action,
    item.yard,
    item.lots,
    item.notes,
    item.status,
    ...customParts,
    item.key,
    item.sheetName,
    accessedText
  ].join(' ').toLowerCase();
}

function getEffectiveItemStatus(item) {
  if (!item) return '';
  const isCustomSchema = Boolean(activeSchemaConfig && !activeSchemaConfig.isStandardSchema && activeSchemaConfig.columns.length > 0);
  if (isCustomSchema && activeSchemaConfig.statusCol) {
    const stFieldKey = `col_${activeSchemaConfig.statusCol}`;
    const rawCustomSt = item.customValues && item.customValues[stFieldKey] !== undefined
      ? String(item.customValues[stFieldKey]).trim()
      : '';
    if (rawCustomSt) {
      return rawCustomSt;
    }
  }
  return String(item.status || '');
}

function syncHideCompletedButtonUI(completedInScopeCount = 0) {
  if (!btnToggleHideCompleted) return;
  const isHiding = Boolean(hideCompletedRequests);
  btnToggleHideCompleted.classList.toggle('active', isHiding);
  btnToggleHideCompleted.setAttribute('aria-checked', String(isHiding));
  btnToggleHideCompleted.setAttribute('aria-pressed', String(isHiding));
  const countSuffix = completedInScopeCount > 0 ? ` (${completedInScopeCount})` : '';
  if (hideCompletedLabel) {
    hideCompletedLabel.textContent = isHiding
      ? `Hiding Completed${countSuffix}`
      : `Hide Completed${countSuffix}`;
  }
  btnToggleHideCompleted.title = isHiding
    ? 'Hiding completed requests — click to show completed entries'
    : 'Click to hide completed requests';
}

function syncHideUuidButtonUI() {
  if (typeof btnToggleHideUuid === 'undefined' || !btnToggleHideUuid) return;
  const isHiding = Boolean(typeof hideUuidKey !== 'undefined' && hideUuidKey);
  btnToggleHideUuid.classList.toggle('active', isHiding);
  btnToggleHideUuid.setAttribute('aria-checked', String(isHiding));
  btnToggleHideUuid.setAttribute('aria-pressed', String(isHiding));
  if (typeof document !== 'undefined' && document.body) {
    document.body.classList.toggle('hide-tracker-uuid-keys', isHiding);
  }
  if (typeof listBody !== 'undefined' && listBody && typeof listBody.querySelectorAll === 'function') {
    listBody.querySelectorAll('th.tracker-row-num-cell, td.select-cell, td.delete-cell').forEach(cell => {
      cell.rowSpan = isHiding ? 1 : 2;
    });
  }
  if (typeof hideUuidLabel !== 'undefined' && hideUuidLabel) {
    hideUuidLabel.textContent = isHiding ? 'Hiding ID' : 'Hide ID';
  }
  btnToggleHideUuid.title = isHiding
    ? 'Hiding UUID Tracking ID keys — click to show UUID keys'
    : 'Click to hide the UUID Tracking ID key on each request row';
}

function renderList() {
  const query = searchInput.value.trim().toLowerCase();
  const scope = scopeSelect.value;
  const currentSheet = sheetSelect.value;
  const activeFilter = activeFieldFilter || (activeKeyFilter ? { field: 'key', label: 'Tracking ID', value: activeKeyFilter } : null);
  const isCustomSchema = Boolean(activeSchemaConfig && !activeSchemaConfig.isStandardSchema && activeSchemaConfig.columns.length > 0);

  if (searchScopeHint) {
    searchScopeHint.style.display = (query.length > 0 || activeFilter) ? 'inline-block' : 'none';
    if (activeFilter && query.length === 0) {
      searchScopeHint.textContent = `🔗 Like-Item List (${activeFilter.label})`;
    } else {
      searchScopeHint.textContent = '🌐 Searching All Tabs';
    }
  }

  const baseFiltered = allRows.filter(item => {
    if (activeFilter) {
      if (!doesItemMatchFieldFilter(item, activeFilter)) return false;
      if (query.length > 0) {
        return buildItemSearchHaystack(item).includes(query);
      }
      return true;
    }
    if (query.length > 0) {
      return buildItemSearchHaystack(item).includes(query);
    }
    if (scope === 'current' && item.sheetName !== currentSheet) {
      return false;
    }
    return true;
  });

  const completedInBaseCount = baseFiltered.filter(item => isCompletedStatus(getEffectiveItemStatus(item))).length;
  syncHideCompletedButtonUI(completedInBaseCount);
  syncHideUuidButtonUI();

  const explicitCompletedStatusFilter = Boolean(
    activeFilter && activeFilter.field === 'status' && isCompletedStatus(activeFilter.value)
  );
  const filtered = (hideCompletedRequests && !explicitCompletedStatusFilter)
    ? baseFiltered.filter(item => !isCompletedStatus(getEffectiveItemStatus(item)))
    : baseFiltered;

  const sorted = sortRowsList(filtered);
  currentlyVisibleRows = sorted;
  updateSortAndFilterUI(sorted.length);

  // Prune any stale selection keys that no longer exist in allRows
  if (selectedForDelete.size > 0) {
    const validCompIds = new Set(allRows.map(r => getRowCompositeId(r.sheetName, r.rowNumber)));
    for (const cid of Array.from(selectedForDelete)) {
      if (!validCompIds.has(cid)) {
        selectedForDelete.delete(cid);
      }
    }
  }

  listBody.innerHTML = '';
  const colSpan = getActiveTableColSpan();
  if (sorted.length === 0) {
    let prevSheetForEmpty = null;
    if (
      workbook &&
      scope === 'current' &&
      !activeFilter &&
      query.length === 0 &&
      completedInBaseCount === 0 &&
      currentSheet &&
      DATE_TAB_REGEX.test(currentSheet)
    ) {
      const currentTabTime = parseDateTab(currentSheet);
      const olderTabs = dateSheets
        .filter(name => parseDateTab(name) < currentTabTime)
        .sort((a, b) => parseDateTab(b) - parseDateTab(a));
      if (olderTabs.length > 0) {
        prevSheetForEmpty = olderTabs[0];
      }
    }

    const emptyRolloverBtnHtml = prevSheetForEmpty
      ? `<div>
           <button type="button" id="btnEmptyTabRolloverPrev" class="btn-rollover btn-sm empty-rollover-cta">
             ↪ Rollover Unfinished Requests from ${escapeHtml(prevSheetForEmpty)}
           </button>
         </div>`
      : '';

    const onlyHiddenCompleted = Boolean(hideCompletedRequests && !explicitCompletedStatusFilter && completedInBaseCount > 0);
    const emptyMsgText = onlyHiddenCompleted
      ? `All ${completedInBaseCount} request${completedInBaseCount === 1 ? ' is' : 's are'} completed (Hiding Completed).`
      : 'No matching requests found.';
    const showCompletedInlineHtml = onlyHiddenCompleted
      ? `<div>
           <button type="button" id="btnShowHiddenCompletedInline" class="btn-sm empty-rollover-cta">
             Show Completed (${completedInBaseCount})
           </button>
         </div>`
      : '';

    listBody.innerHTML = `<tr><td colspan="${colSpan}" class="empty-table-cell"><div>${emptyMsgText}</div>${showCompletedInlineHtml}${emptyRolloverBtnHtml}</td></tr>`;
    const btnShowHidden = document.getElementById('btnShowHiddenCompletedInline');
    if (btnShowHidden) {
      btnShowHidden.addEventListener('click', () => {
        if (typeof setHideCompletedPreference === 'function') {
          setHideCompletedPreference(false);
        } else {
          hideCompletedRequests = false;
        }
        renderList();
      });
    }
    const btnEmptyRollover = document.getElementById('btnEmptyTabRolloverPrev');
    if (btnEmptyRollover) {
      btnEmptyRollover.addEventListener('click', () => {
        openRolloverModal({ preferredSourceMode: 'previous', preferredTargetDate: currentSheet });
      });
    }

    const shadowRowEmpty = createShadowNewRowElement(colSpan, currentSheet);
    if (shadowRowEmpty) {
      listBody.appendChild(shadowRowEmpty);
    }
    updateDeleteSelectionUI();
    return;
  }

  const requestStatsByKey = new Map();
  allRows.forEach(r => {
    const normKey = String(r.key || '').trim().toLowerCase();
    if (!normKey) return;
    let stat = requestStatsByKey.get(normKey);
    if (!stat) {
      stat = { firstTime: 0, sheetTimeMap: new Map() };
      requestStatsByKey.set(normKey, stat);
    }
    const t = parseDateTab(r.sheetName);
    if (t > 0 && (!stat.firstTime || t < stat.firstTime)) {
      stat.firstTime = t;
    }
    if (r.sheetName && !stat.sheetTimeMap.has(r.sheetName)) {
      stat.sheetTimeMap.set(r.sheetName, t);
    }
  });

  const copyBtnFn = typeof renderInlineCopyBtnHtml === 'function' ? renderInlineCopyBtnHtml : () => '';

  sorted.forEach((item, idx) => {
    const tr = document.createElement('tr');
    tr.classList.add('tracker-main-row');
    tr.dataset.sheet = item.sheetName;
    tr.dataset.row = String(item.rowNumber);

    const metaTr = document.createElement('tr');
    metaTr.classList.add('tracker-meta-row');
    metaTr.dataset.sheet = item.sheetName;
    metaTr.dataset.row = String(item.rowNumber);

    const compId = getRowCompositeId(item.sheetName, item.rowNumber);
    const isRowChecked = selectedForDelete.has(compId);
    if (isRowChecked) {
      const selCls = isDeleteMode ? 'selected-for-delete' : 'selected-for-batch';
      tr.classList.add(selCls);
      metaTr.classList.add(selCls);
    }
    if (
      dockEditingTarget &&
      dockEditingTarget.sheetName === item.sheetName &&
      dockEditingTarget.rowNumber === item.rowNumber
    ) {
      tr.classList.add('active-editing-row');
      metaTr.classList.add('active-editing-row');
    }

    const normItemKey = String(item.key || '').trim().toLowerCase();
    const itemStat = normItemKey ? requestStatsByKey.get(normItemKey) : null;
    const itemTabTime = parseDateTab(item.sheetName);
    const ageDays = (itemStat && itemStat.firstTime > 0 && itemTabTime >= itemStat.firstTime)
      ? Math.round((itemTabTime - itemStat.firstTime) / (1000 * 60 * 60 * 24))
      : 0;
    let occurrencesUpToSheet = 1;
    if (itemStat) {
      let count = 0;
      itemStat.sheetTimeMap.forEach((tVal) => {
        if (!tVal || !itemTabTime || tVal <= itemTabTime) count++;
      });
      occurrencesUpToSheet = Math.max(1, count);
    }
    const effectiveRowStatus = getEffectiveItemStatus(item);
    const isRowCompleted = isCompletedStatus(effectiveRowStatus);
    const isRowWaiting =
      !isRowCompleted &&
      typeof isWaitingOrHoldStatus === 'function' &&
      isWaitingOrHoldStatus(effectiveRowStatus);
    const isRowOverdue =
      !isRowCompleted &&
      (
        (typeof isOverdueStatus === 'function' && isOverdueStatus(effectiveRowStatus)) ||
        (!isRowWaiting && (ageDays > SMART_OVERDUE_MAX_AGE_DAYS || occurrencesUpToSheet >= SMART_OVERDUE_MIN_CARRIES))
      );
    if (isRowCompleted) {
      tr.classList.add('completed-outline-green');
      metaTr.classList.add('completed-outline-green');
    } else if (isRowOverdue) {
      tr.classList.add('overdue-outline-red');
      metaTr.classList.add('overdue-outline-red');
    } else if (isRowWaiting) {
      tr.classList.add('waiting-outline-purple');
      metaTr.classList.add('waiting-outline-purple');
    }

    const noteText = item.notes || '—';
    const isLongNote = noteText.length > 45 || noteText.includes('\n');
    const shortGuid = item.key.length > 8 ? `${item.key.slice(0, 8)}...` : item.key;

    const isHidingUuid = Boolean(typeof hideUuidKey !== 'undefined' && hideUuidKey);
    const rowSpanAttr = isHidingUuid ? '' : ' rowspan="2"';
    const rowNumCell = `<th scope="row"${rowSpanAttr} class="tracker-row-num-cell" title="Line ${idx + 1} (Excel Row ${item.rowNumber} on ${escapeHtml(item.sheetName)})">${idx + 1}</th>`;
    const rowCheckboxCell = `<td${rowSpanAttr} class="${isDeleteMode ? 'delete-cell' : 'select-cell'}">
      <input type="checkbox" class="chk-delete-row" data-comp="${escapeHtml(compId)}" data-key="${escapeHtml(item.key)}" ${isRowChecked ? 'checked' : ''} aria-label="Select request row ${item.rowNumber} on ${escapeHtml(item.sheetName)}">
    </td>`;
    const accessedInfo = (typeof getRowLastAccessedMs === 'function' && typeof formatRowLastAccessedLabel === 'function')
      ? formatRowLastAccessedLabel(getRowLastAccessedMs(item.key))
      : { ms: 0, short: '—', full: 'Not accessed yet' };
    const statusAccessedCornerHtml = `<span class="status-accessed-time${accessedInfo.ms ? '' : ' is-empty'}" data-accessed-key="${escapeHtml(normItemKey)}" title="${escapeHtml(accessedInfo.full)}">${accessedInfo.ms ? escapeHtml(accessedInfo.short) : ''}</span>`;
    const keyActiveClass = isFieldFilterActive('key', item.key) ? ' active-filter-match' : '';
    const metaBarColSpan = Math.max(1, colSpan - 2);
    const keyCopyBtnHtml = copyBtnFn(item.key, 'Tracking ID');
    metaTr.innerHTML = `
      <td colspan="${metaBarColSpan}" class="tracker-meta-bar-cell">
        <div class="tracker-meta-bar">
          <span class="key-chip-copy-group">
            <span class="key-chip key-chip-compact${keyActiveClass}" title="Click to filter by Tracking ID: ${escapeHtml(item.key)}" data-key="${escapeHtml(item.key)}">
              🔍 ${escapeHtml(item.key)}
            </span>${keyCopyBtnHtml}
          </span>
        </div>
      </td>
    `;

    if (isCustomSchema) {
      const visibleCustomCols = getVisibleCustomTableColumns();
      const hasCustomStatusCol = Boolean(
        activeSchemaConfig.statusCol &&
        visibleCustomCols.some(c => c.colIndex === activeSchemaConfig.statusCol)
      );
      const customCellsHtml = visibleCustomCols.map(c => {
        const fieldKey = `col_${c.colIndex}`;
        const rawVal = (item.customValues && item.customValues[fieldKey] !== undefined)
          ? String(item.customValues[fieldKey])
          : '';
        const trimmedVal = rawVal.trim();

        if (c.colIndex === activeSchemaConfig.statusCol) {
          const stText = trimmedVal || item.status || 'Pending';
          const stBadgeClass = typeof getStatusBadgeClass === 'function'
            ? getStatusBadgeClass(stText)
            : (isCompletedStatus(stText) ? 'badge-complete' : 'badge-pending');
          const stActive = isFieldFilterActive(fieldKey, stText) ? ' active-filter-match' : '';
          return `
            <td class="tracker-status-cell">
              <div class="status-cell-wrap">
                <span class="badge clickable-badge ${stBadgeClass}${stActive}" data-filter-field="${fieldKey}" data-filter-label="${escapeHtml(c.header)}" data-filter-val="${escapeHtml(stText)}" title="Click to filter by ${escapeHtml(c.header)}: ${escapeHtml(stText)}">
                  ${escapeHtml(stText)}
                </span>
                ${statusAccessedCornerHtml}
              </div>
            </td>
          `;
        }

        if (c.colIndex === activeSchemaConfig.notesCol || c.dataType === 'notes') {
          const cNote = trimmedVal || '—';
          const cLong = cNote.length > 45 || cNote.includes('\n');
          return `
            <td class="notes-cell">
              <div class="notes-preview" id="note-${idx}-${c.colIndex}">${escapeHtml(cNote)}</div>
              ${cLong ? `<span class="notes-toggle" data-target="note-${idx}-${c.colIndex}">[+ Expand]</span>` : ''}
            </td>
          `;
        }

        const isCopyableCustomCol = Array.isArray(activeSchemaConfig.copyCols)
          ? activeSchemaConfig.copyCols.includes(c.colIndex)
          : (
              c.colIndex === activeSchemaConfig.ticketCol ||
              c.colIndex === activeSchemaConfig.secondaryCol ||
              /\b(subject|ticket|master\s*#?|account\s*#?|member\s*#?|bidder\s*#?|yard\s*#?|lot\s*#?)\b/i.test(c.excelHeader || c.header)
            );

        if (c.colIndex === activeSchemaConfig.ticketCol) {
          return `<td>${renderClickableTicketTokensHtml(trimmedVal, fieldKey, c.header, isCopyableCustomCol)}</td>`;
        }

        if (!trimmedVal) {
          return '<td>—</td>';
        }
        const cellActive = isFieldFilterActive(fieldKey, trimmedVal) ? ' active-filter-match' : '';
        const copyHtml = isCopyableCustomCol ? copyBtnFn(trimmedVal, c.header) : '';
        const cellSpanHtml = `<span class="clickable-cell${cellActive}" data-filter-field="${fieldKey}" data-filter-label="${escapeHtml(c.header)}" data-filter-val="${escapeHtml(trimmedVal)}" title="Click to filter by ${escapeHtml(c.header)}: ${escapeHtml(trimmedVal)}">${escapeHtml(trimmedVal)}</span>`;
        return `
          <td>
            ${copyHtml ? `<span class="cell-copy-group cell-copy-text">${cellSpanHtml}${copyHtml}</span>` : cellSpanHtml}
          </td>
        `;
      }).join('');

      tr.innerHTML = `
        ${rowNumCell}
        ${rowCheckboxCell}
        <td style="white-space:nowrap;">
          <button type="button" class="clickable-date-tab" data-date-tab="${escapeHtml(item.sheetName)}" title="Click to switch to daily tab ${escapeHtml(item.sheetName)}">
            ${escapeHtml(item.sheetName)}
          </button>
        </td>
        ${customCellsHtml}
        <td>
          <div class="row-actions-group">
            <button type="button" class="btn-details" data-sheet="${escapeHtml(item.sheetName)}" data-row="${item.rowNumber}" title="View Details &amp; Aggregated Timestamped Notes">Details</button>
            <button type="button" class="btn-edit" data-sheet="${escapeHtml(item.sheetName)}" data-row="${item.rowNumber}" title="Load into Bottom Editor (Double-click for Modal)">Edit</button>
          </div>
          ${hasCustomStatusCol ? '' : `<div class="actions-accessed-fallback">${statusAccessedCornerHtml}</div>`}
        </td>
      `;
      listBody.appendChild(tr);
      if (!isHidingUuid) {
        listBody.appendChild(metaTr);
      }
      return;
    }

    const stdDescriptor = getStandardTableColumnsDescriptor();
    const isStdColCopyable = (colIdx) => Array.isArray(activeSchemaConfig && activeSchemaConfig.copyCols)
      ? activeSchemaConfig.copyCols.includes(colIdx)
      : [1, 2, 4, 5, 7, 8].includes(colIdx);

    const subjectActiveClass = isFieldFilterActive('subject', item.subject) ? ' active-filter-match' : '';
    const brokerActiveClass = isFieldFilterActive('broker', item.broker) ? ' active-filter-match' : '';
    const masterActiveClass = isFieldFilterActive('master', item.master) ? ' active-filter-match' : '';
    const bidderActiveClass = isFieldFilterActive('bidder', item.bidder) ? ' active-filter-match' : '';
    const actionActiveClass = isFieldFilterActive('action', item.action) ? ' active-filter-match' : '';
    const yardActiveClass = isFieldFilterActive('yard', item.yard) ? ' active-filter-match' : '';
    const statusActiveClass = isFieldFilterActive('status', item.status) ? ' active-filter-match' : '';

    const subjectCopyHtml = isStdColCopyable(1) ? copyBtnFn(item.subject, 'Email Subject') : '';
    const subjectChipHtml = `<span class="clickable-cell${subjectActiveClass}" data-filter-field="subject" data-filter-label="Subject" data-filter-val="${escapeHtml(item.subject)}" title="Click to show all requests with Subject: ${escapeHtml(item.subject)}">${escapeHtml(item.subject)}</span>`;
    const subjectHtml = item.subject
      ? (subjectCopyHtml ? `<span class="cell-copy-group cell-copy-text">${subjectChipHtml}${subjectCopyHtml}</span>` : subjectChipHtml)
      : '—';

    const ticketHtml = renderClickableTicketTokensHtml(item.ticket, 'ticket', 'Ticket #', isStdColCopyable(2));

    const brokerCopyHtml = isStdColCopyable(3) ? copyBtnFn(item.broker, 'Broker') : '';
    const brokerChipHtml = `<span class="clickable-cell${brokerActiveClass}" data-filter-field="broker" data-filter-label="Broker" data-filter-val="${escapeHtml(item.broker)}" title="Click to show all requests for Broker: ${escapeHtml(item.broker)}">${escapeHtml(item.broker)}</span>`;
    const brokerHtml = item.broker
      ? (brokerCopyHtml ? `<span class="cell-copy-group cell-copy-text">${brokerChipHtml}${brokerCopyHtml}</span>` : brokerChipHtml)
      : '—';

    const masterCopyHtml = isStdColCopyable(4) ? copyBtnFn(item.master, 'Master #') : '';
    const masterChipHtml = `<span class="clickable-code${masterActiveClass}" data-filter-field="master" data-filter-label="Master #" data-filter-val="${escapeHtml(item.master)}" title="Click to show all requests with Master # ${escapeHtml(item.master)}">${escapeHtml(item.master)}</span>`;
    const masterHtml = item.master
      ? (masterCopyHtml ? `<span class="cell-copy-group">${masterChipHtml}${masterCopyHtml}</span>` : masterChipHtml)
      : '—';

    const bidderCopyHtml = isStdColCopyable(5) ? copyBtnFn(item.bidder, 'Bidder #') : '';
    const bidderChipHtml = `<span class="clickable-code${bidderActiveClass}" data-filter-field="bidder" data-filter-label="Bidder #" data-filter-val="${escapeHtml(item.bidder)}" title="Click to show all requests with Bidder # ${escapeHtml(item.bidder)}">${escapeHtml(item.bidder)}</span>`;
    const bidderHtml = item.bidder
      ? (bidderCopyHtml ? `<span class="cell-copy-group">${bidderChipHtml}${bidderCopyHtml}</span>` : bidderChipHtml)
      : '—';

    const actionCopyHtml = isStdColCopyable(6) ? copyBtnFn(item.action, 'Requested Action') : '';
    const actionChipHtml = `<span class="clickable-cell${actionActiveClass}" data-filter-field="action" data-filter-label="Action" data-filter-val="${escapeHtml(item.action)}" title="Click to show all requests with Action: ${escapeHtml(item.action)}">${escapeHtml(item.action)}</span>`;
    const actionHtml = item.action
      ? (actionCopyHtml ? `<span class="cell-copy-group cell-copy-text">${actionChipHtml}${actionCopyHtml}</span>` : actionChipHtml)
      : '—';

    const yardCopyHtml = isStdColCopyable(7) ? copyBtnFn(item.yard, 'Yard #') : '';
    const yardChipHtml = `<span class="clickable-code${yardActiveClass}" data-filter-field="yard" data-filter-label="Yard #" data-filter-val="${escapeHtml(item.yard)}" title="Click to show all requests for Yard # ${escapeHtml(item.yard)}">${escapeHtml(item.yard)}</span>`;
    const yardHtml = item.yard
      ? (yardCopyHtml ? `<span class="cell-copy-group">${yardChipHtml}${yardCopyHtml}</span>` : yardChipHtml)
      : '—';

    const lotsHtml = renderClickableLotsHtml(item.lots, isStdColCopyable(8));
    const stdStatusBadgeClass = typeof getStatusBadgeClass === 'function'
      ? getStatusBadgeClass(item.status)
      : (isCompletedStatus(item.status) ? 'badge-complete' : 'badge-pending');

    const stdCellsHtml = stdDescriptor.slots.map(s => {
      switch (s.kind) {
        case 'subject':
          return `<td class="tracker-subject-cell">${subjectHtml}</td>`;
        case 'ticket':
          return `<td>${ticketHtml}</td>`;
        case 'broker':
          return `<td>${brokerHtml}</td>`;
        case 'master_bidder': {
          if (s.showMaster && s.showBidder) {
            return `<td class="cell-mono-sm col-master-bidder">M: ${masterHtml}<br>B: ${bidderHtml}</td>`;
          }
          return `<td class="cell-mono-sm col-master-bidder">${s.showMaster ? masterHtml : bidderHtml}</td>`;
        }
        case 'action':
          return `<td>${actionHtml}</td>`;
        case 'yard_lots': {
          if (s.showYard && s.showLots) {
            return `<td class="cell-sm">Y: ${yardHtml}<br>L: ${lotsHtml}</td>`;
          }
          return `<td class="cell-sm">${s.showYard ? yardHtml : lotsHtml}</td>`;
        }
        case 'notes':
          return `<td class="notes-cell"><div class="notes-preview" id="note-${idx}">${escapeHtml(noteText)}</div>${isLongNote ? `<span class="notes-toggle" data-target="note-${idx}">[+ Expand]</span>` : ''}</td>`;
        case 'status':
          return `<td class="tracker-status-cell"><div class="status-cell-wrap"><span class="badge clickable-badge ${stdStatusBadgeClass}${statusActiveClass}" data-filter-field="status" data-filter-label="Status" data-filter-val="${escapeHtml(item.status)}" title="Click to show all requests with Status: ${escapeHtml(item.status)}">${escapeHtml(item.status)}</span>${statusAccessedCornerHtml}</div></td>`;
        case 'extra': {
          const fKey = `col_${s.colIndex}`;
          const ev = (item.customValues && item.customValues[fKey] !== undefined) ? String(item.customValues[fKey]).trim() : '';
          if (s.colObj && s.colObj.dataType === 'notes') {
            const eNote = ev || '—';
            const eLong = eNote.length > 45 || eNote.includes('\n');
            return `<td class="notes-cell"><div class="notes-preview" id="note-${idx}-${s.colIndex}">${escapeHtml(eNote)}</div>${eLong ? `<span class="notes-toggle" data-target="note-${idx}-${s.colIndex}">[+ Expand]</span>` : ''}</td>`;
          }
          if (!ev) return '<td>—</td>';
          const eAct = isFieldFilterActive(fKey, ev) ? ' active-filter-match' : '';
          const eCopy = isStdColCopyable(s.colIndex) ? copyBtnFn(ev, s.label) : '';
          const eSpan = `<span class="clickable-cell${eAct}" data-filter-field="${fKey}" data-filter-label="${escapeHtml(s.label)}" data-filter-val="${escapeHtml(ev)}" title="Click to filter by ${escapeHtml(s.label)}: ${escapeHtml(ev)}">${escapeHtml(ev)}</span>`;
          return `<td>${eCopy ? `<span class="cell-copy-group cell-copy-text">${eSpan}${eCopy}</span>` : eSpan}</td>`;
        }
        default:
          return '';
      }
    }).join('');

    tr.innerHTML = `
      ${rowNumCell}
      ${rowCheckboxCell}
      <td style="white-space:nowrap;">
        <button type="button" class="clickable-date-tab" data-date-tab="${escapeHtml(item.sheetName)}" title="Click to switch to daily tab ${escapeHtml(item.sheetName)}">
          ${escapeHtml(item.sheetName)}
        </button>
      </td>
      ${stdCellsHtml}
      <td>
        <div class="row-actions-group">
          <button type="button" class="btn-details" data-sheet="${escapeHtml(item.sheetName)}" data-row="${item.rowNumber}" title="View Ticket Details &amp; Aggregated Timestamped Notes">Details</button>
          <button type="button" class="btn-edit" data-sheet="${escapeHtml(item.sheetName)}" data-row="${item.rowNumber}" title="Load into Bottom Editor (Double-click for Modal)">Edit</button>
        </div>
        ${stdDescriptor.slots.some(s => s.kind === 'status') ? '' : `<div class="actions-accessed-fallback">${statusAccessedCornerHtml}</div>`}
      </td>
    `;
    listBody.appendChild(tr);
    if (!isHidingUuid) {
      listBody.appendChild(metaTr);
    }
  });

  const shadowRow = createShadowNewRowElement(colSpan, currentSheet);
  if (shadowRow) {
    listBody.appendChild(shadowRow);
  }

  updateDeleteSelectionUI();
}

// --- PERMANENT CHECKBOX SELECTION, FLOATING BATCH STATUS DOCK & BATCH DELETE ---
function setDeleteMode(enabled) {
  closeToolsMenu();
  isDeleteMode = Boolean(enabled);
  selectedForDelete.clear();
  if (deleteModeLabel) {
    deleteModeLabel.textContent = isDeleteMode
      ? 'Exit Batch Delete Mode'
      : 'Batch Delete Requests Mode';
  } else {
    btnDeleteMode.textContent = isDeleteMode
      ? 'Exit Batch Delete Mode'
      : 'Batch Delete Requests Mode';
  }
  deleteBanner.style.display = isDeleteMode ? 'flex' : 'none';
  const liveThSelect = document.getElementById('thDeleteSelect');
  if (liveThSelect) {
    liveThSelect.className = isDeleteMode ? 'delete-col' : 'select-col';
  }
  updateDeleteSelectionUI();
  renderList();
}

function updateDeleteSelectionUI() {
  const count = selectedForDelete.size;
  if (deleteSelectedCountEl) {
    deleteSelectedCountEl.textContent = String(count);
  }
  if (btnConfirmDeleteSelected) {
    btnConfirmDeleteSelected.textContent = `Delete Selected (${count})`;
    btnConfirmDeleteSelected.disabled = count === 0;
  }
  if (batchSelectedCountEl) {
    batchSelectedCountEl.textContent = String(count);
  }
  const hasStatusCol = activeSchemaConfig.isStandardSchema || Boolean(activeSchemaConfig.statusCol);
  if (batchStatusControlGroup) {
    batchStatusControlGroup.style.display = hasStatusCol ? 'inline-flex' : 'none';
  }
  if (batchStatusSep) {
    batchStatusSep.style.display = hasStatusCol ? 'inline-block' : 'none';
  }
  if (batchStatusSelect) {
    batchStatusSelect.disabled = !hasStatusCol;
  }
  if (btnApplyBatchStatus) {
    btnApplyBatchStatus.disabled = !hasStatusCol;
    btnApplyBatchStatus.title = hasStatusCol
      ? 'Update status on all selected rows'
      : 'Assign a Status column via Configure Column Mapping to enable batch status updates';
  }
  if (batchActionDock) {
    const showBatchDock = !isDeleteMode && count > 0 && Boolean(workbook);
    batchActionDock.style.display = showBatchDock ? 'block' : 'none';
  }
  updateCopiedRowsToolbarUI();
  syncSelectAllCheckboxes();
  updateStatusBarTelemetry();
}

function updateCopiedRowsToolbarUI() {
  const clipCount = Array.isArray(copiedTrackerRowsClipboard) ? copiedTrackerRowsClipboard.length : 0;
  if (pasteCopiedCountEl) {
    pasteCopiedCountEl.textContent = String(clipCount);
  }
  if (pasteRowsToolbarGroup) {
    const showPaste = clipCount > 0 && Boolean(workbook);
    pasteRowsToolbarGroup.style.display = showPaste ? 'inline-flex' : 'none';
  }
}

function syncSelectAllCheckboxes() {
  const liveChkHeader = document.getElementById('chkSelectAllHeader');
  if (!workbook || currentlyVisibleRows.length === 0) {
    if (liveChkHeader) {
      liveChkHeader.checked = false;
      liveChkHeader.indeterminate = false;
      liveChkHeader.disabled = true;
    }
    if (chkSelectAllBanner) {
      chkSelectAllBanner.checked = false;
      chkSelectAllBanner.indeterminate = false;
    }
    return;
  }

  if (liveChkHeader) {
    liveChkHeader.disabled = false;
  }

  let visibleSelectedCount = 0;
  currentlyVisibleRows.forEach(item => {
    if (selectedForDelete.has(getRowCompositeId(item.sheetName, item.rowNumber))) {
      visibleSelectedCount++;
    }
  });

  const allVisibleSelected = visibleSelectedCount === currentlyVisibleRows.length;
  const someVisibleSelected = visibleSelectedCount > 0 && !allVisibleSelected;

  if (liveChkHeader) {
    liveChkHeader.checked = allVisibleSelected;
    liveChkHeader.indeterminate = someVisibleSelected;
  }
  if (chkSelectAllBanner) {
    chkSelectAllBanner.checked = isDeleteMode && allVisibleSelected;
    chkSelectAllBanner.indeterminate = isDeleteMode && someVisibleSelected;
  }
}

function handleToggleSelectAllVisible(checked) {
  currentlyVisibleRows.forEach(item => {
    const compId = getRowCompositeId(item.sheetName, item.rowNumber);
    if (checked) {
      selectedForDelete.add(compId);
    } else {
      selectedForDelete.delete(compId);
    }
  });
  updateDeleteSelectionUI();
  renderList();
}

btnDeleteMode.addEventListener('click', () => {
  setDeleteMode(!isDeleteMode);
});

btnCancelDeleteMode.addEventListener('click', () => {
  setDeleteMode(false);
});

if (tableHead) {
  tableHead.addEventListener('change', (e) => {
    const hdrChk = e.target.closest('#chkSelectAllHeader');
    if (hdrChk) {
      handleToggleSelectAllVisible(hdrChk.checked);
    }
  });
}

chkSelectAllBanner.addEventListener('change', (e) => {
  handleToggleSelectAllVisible(e.target.checked);
});

if (btnClearBatchSelection) {
  btnClearBatchSelection.addEventListener('click', () => {
    selectedForDelete.clear();
    updateDeleteSelectionUI();
    renderList();
  });
}

listBody.addEventListener('change', (e) => {
  const chk = e.target.closest('.chk-delete-row');
  if (!chk) return;
  const compId = chk.dataset.comp;
  if (!compId) return;

  const tr = chk.closest('tr');
  const pairedMetaTr = (tr && tr.nextElementSibling && tr.nextElementSibling.classList.contains('tracker-meta-row'))
    ? tr.nextElementSibling
    : null;
  if (chk.checked) {
    selectedForDelete.add(compId);
    [tr, pairedMetaTr].forEach(rowEl => {
      if (!rowEl) return;
      rowEl.classList.toggle('selected-for-delete', isDeleteMode);
      rowEl.classList.toggle('selected-for-batch', !isDeleteMode);
    });
  } else {
    selectedForDelete.delete(compId);
    [tr, pairedMetaTr].forEach(rowEl => {
      if (!rowEl) return;
      rowEl.classList.remove('selected-for-delete', 'selected-for-batch');
    });
  }
  updateDeleteSelectionUI();
});

async function handleApplyBatchStatusUpdate() {
  if (selectedForDelete.size === 0 || !workbook || !batchStatusSelect) return;
  if (!activeSchemaConfig.isStandardSchema && !activeSchemaConfig.statusCol) {
    return showModalAlert({
      title: 'No Status Column Mapped',
      message: 'Open Tools & Setup → Configure Column Mapping to assign a Status column for this custom workbook.'
    });
  }
  const targetStatus = batchStatusSelect.value.trim();
  if (!targetStatus) return;

  const selectedCompSet = new Set(selectedForDelete);
  const selectedItems = allRows.filter(item =>
    selectedCompSet.has(getRowCompositeId(item.sheetName, item.rowNumber))
  );
  if (selectedItems.length === 0) return;

  // Check if any selected rows also appear on newer daily tabs with a different status
  const extraNewerItems = [];
  const seenExtraComp = new Set();
  selectedItems.forEach(sel => {
    const selTime = parseDateTab(sel.sheetName);
    const selKey = String(sel.key || '').trim().toLowerCase();
    if (!selKey) return;
    allRows.forEach(candidate => {
      const candComp = getRowCompositeId(candidate.sheetName, candidate.rowNumber);
      if (selectedCompSet.has(candComp) || seenExtraComp.has(candComp)) return;
      if (
        String(candidate.key || '').trim().toLowerCase() === selKey &&
        parseDateTab(candidate.sheetName) > selTime &&
        candidate.status.trim().toLowerCase() !== targetStatus.toLowerCase()
      ) {
        seenExtraComp.add(candComp);
        extraNewerItems.push(candidate);
      }
    });
  });

  let rowsToMutate = [...selectedItems];
  if (extraNewerItems.length > 0) {
    const newerTabsList = Array.from(new Set(extraNewerItems.map(r => r.sheetName))).join(', ');
    const syncNewer = await showModalConfirm({
      title: '🔄 Sync Batch Status to Newer Tab(s)?',
      badge: targetStatus,
      message:
        `${extraNewerItems.length} selected request(s) also exist on newer daily tab(s): ${newerTabsList}.\n\n` +
        `Also update their status to "${targetStatus}" on those newer tab(s) so Rollover stays consistent?`,
      confirmText: '🔄 Update Newer Tab(s) Too',
      cancelText: 'Only Selected Row(s)'
    });
    if (syncNewer) {
      rowsToMutate = [...selectedItems, ...extraNewerItems];
    }
  }

  // 1 & 2. Map via immutable UUIDv4 keys (hidden Tracking ID col) & strictly preserve existing UUIDv4s
  const batchTargets = rowsToMutate.map(item => ({
    itemRef: item,
    sheetName: item.sheetName,
    rowNumber: item.rowNumber,
    immutableKey: item.key // Strictly preserved — never regenerated
  }));

  // 3. Update Request Status in active state array FIRST
  ensureStatusOptionExists(targetStatus);
  batchTargets.forEach(target => {
    target.itemRef.status = targetStatus;
    if (!activeSchemaConfig.isStandardSchema && activeSchemaConfig.statusCol && target.itemRef.customValues) {
      target.itemRef.customValues[`col_${activeSchemaConfig.statusCol}`] = targetStatus;
    }
    if (
      dockEditingTarget &&
      dockEditingTarget.sheetName === target.sheetName &&
      dockEditingTarget.rowNumber === target.rowNumber
    ) {
      if (dockStatus) dockStatus.value = targetStatus;
      if (!activeSchemaConfig.isStandardSchema && activeSchemaConfig.statusCol && dockDynamicFieldsGrid) {
        const dynStEl = dockDynamicFieldsGrid.querySelector(`[data-custom-col="${activeSchemaConfig.statusCol}"]`);
        if (dynStEl) dynStEl.value = targetStatus;
      }
    }
  });

  // 4. Clear selection & re-render the DOM immediately to reflect the new statuses
  selectedForDelete.clear();
  updateDeleteSelectionUI();
  renderSheetTabsBar();
  renderList();

  // 5. Mutate ExcelJS cells (preserving existing Tracking ID UUIDv4s) & perform ONE single batch write-back
  if (btnApplyBatchStatus) btnApplyBatchStatus.disabled = true;
  try {
    batchTargets.forEach(target => {
      const ws = workbook.getWorksheet(target.sheetName);
      if (!ws) return;
      const colMap = getWorksheetColumnMap(ws);
      ws.getColumn(colMap.trackingCol).hidden = true;
      if (colMap.emailSubjCol) {
        ws.getColumn(colMap.emailSubjCol).hidden = false;
      }

      const excelRow = ws.getRow(target.rowNumber);
      if (colMap.statusCol) {
        excelRow.getCell(colMap.statusCol).value = targetStatus;
      }

      // Guarantee existing UUIDv4 in hidden Tracking ID column is never lost or regenerated
      const existingCellKey = cellText(excelRow, colMap.trackingCol).trim();
      if (!existingCellKey && target.immutableKey) {
        excelRow.getCell(colMap.trackingCol).value = target.immutableKey;
      }
      excelRow.commit();
    });

    const saveResult = await saveWorkbookToDisk();
    if (!saveResult) {
      if (!pendingPrimarySaveState) {
        await loadWorkbookFromDisk(sheetSelect.value);
      }
      return;
    }

    showToast(
      `Batch updated ${selectedItems.length} request(s) to "${targetStatus}"${rowsToMutate.length > selectedItems.length ? ` (+${rowsToMutate.length - selectedItems.length} newer-tab row(s))` : ''}.`,
      'success'
    );
  } finally {
    if (btnApplyBatchStatus) {
      btnApplyBatchStatus.disabled = !(activeSchemaConfig.isStandardSchema || Boolean(activeSchemaConfig.statusCol));
    }
  }
}

if (btnApplyBatchStatus) {
  btnApplyBatchStatus.addEventListener('click', handleApplyBatchStatusUpdate);
}

// --- MANUAL ROW COPY & PASTE ACROSS DAILY TABS (Works with or without Status/Validation Lists) ---
function cloneTrackerRowForClipboard(item) {
  return {
    sourceSheetName: item.sheetName || '',
    sourceRowNumber: item.rowNumber || 0,
    subject: item.subject || '',
    ticket: item.ticket || '',
    broker: item.broker || '',
    master: item.master || '',
    bidder: item.bidder || '',
    action: item.action || '',
    yard: item.yard || '',
    lots: item.lots || '',
    notes: item.notes || '',
    status: item.status || '',
    key: item.key || generateTrackingGuid(),
    isCustomSchema: Boolean(item.isCustomSchema || (activeSchemaConfig && !activeSchemaConfig.isStandardSchema)),
    customValues: item.customValues ? { ...item.customValues } : null
  };
}

function getSelectedItemsFromCheckboxSet() {
  if (selectedForDelete.size === 0) return [];
  const selectedCompSet = new Set(selectedForDelete);
  return allRows.filter(item =>
    selectedCompSet.has(getRowCompositeId(item.sheetName, item.rowNumber))
  );
}

function handleCopySelectedTrackerRows() {
  const selectedItems = getSelectedItemsFromCheckboxSet();
  if (selectedItems.length === 0) return;

  copiedTrackerRowsClipboard = selectedItems.map(cloneTrackerRowForClipboard);
  selectedForDelete.clear();
  updateDeleteSelectionUI();
  renderList();
  showToast(
    `📋 Copied ${copiedTrackerRowsClipboard.length} row(s). Switch to the next day tab and click "📥 Paste Row(s)" (or press Ctrl+V).`,
    'info'
  );
}

function clearCopiedTrackerRowsClipboard() {
  copiedTrackerRowsClipboard = [];
  updateCopiedRowsToolbarUI();
}

function ensureTargetDailySheetInWorkbook(targetTabName) {
  let targetSheet = workbook.getWorksheet(targetTabName);
  let createdNewSheet = false;

  if (!targetSheet) {
    createdNewSheet = true;
    let placeholderWs = null;
    workbook.eachSheet(ws => {
      if (isTemplatePlaceholderTab(ws.name)) placeholderWs = ws;
    });

    if (placeholderWs && dateSheets.length === 0) {
      placeholderWs.name = targetTabName;
      targetSheet = placeholderWs;
    } else {
      const templateSheet = findTemplateSheet();
      targetSheet = workbook.addWorksheet(targetTabName);
      placeWorksheetFirst(targetSheet);

      if (templateSheet) {
        cloneTemplateStructure(templateSheet, targetSheet);
      } else if (!activeSchemaConfig.isStandardSchema && activeSchemaConfig.columns.length > 0) {
        applyCustomHeadersSheetFormatting(targetSheet, activeSchemaConfig.columns.map(c => c.excelHeader || c.header));
      } else {
        applyStandardSheetFormatting(targetSheet);
      }
    }
  }

  return { targetSheet, createdNewSheet };
}

async function pasteTrackerRowsToTargetTab(targetTabName, rowsToPaste, { clearClipboardAfter = true } = {}) {
  if (!workbook || !Array.isArray(rowsToPaste) || rowsToPaste.length === 0) return false;
  const cleanTargetTab = String(targetTabName || '').trim();
  if (!DATE_TAB_REGEX.test(cleanTargetTab)) {
    await showModalAlert({
      title: 'Invalid Target Daily Tab',
      message: 'Please select or enter a valid daily tab in MM-DD-YYYY format (e.g., 11-10-2026).'
    });
    return false;
  }

  if (!pendingPrimarySaveState) {
    await loadWorkbookFromDisk(sheetSelect ? sheetSelect.value : cleanTargetTab, { skipStagedRecoveryCheck: true });
  }

  const { targetSheet, createdNewSheet } = ensureTargetDailySheetInWorkbook(cleanTargetTab);
  const targetColMap = getWorksheetColumnMap(targetSheet);
  targetSheet.getColumn(targetColMap.trackingCol).hidden = true;
  if (targetColMap.emailSubjCol) {
    targetSheet.getColumn(targetColMap.emailSubjCol).hidden = false;
  }

  // Track existing Tracking IDs on the target sheet so same-sheet duplicates get a new UUIDv4
  // while cross-tab pastes preserve the original UUIDv4 to keep multi-day history linked
  const usedKeysOnTarget = new Set();
  targetSheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const existingKey = cellText(row, targetColMap.trackingCol).trim().toLowerCase();
    if (existingKey) usedKeysOnTarget.add(existingKey);
  });

  let writeRowIdx = getNextEmptyRow(targetSheet);
  let pastedCount = 0;

  rowsToPaste.forEach(srcItem => {
    const cloned = cloneTrackerRowForClipboard(srcItem);
    const normKey = String(cloned.key || '').trim().toLowerCase();
    const isSameSheetDuplicate = cloned.sourceSheetName === cleanTargetTab || usedKeysOnTarget.has(normKey);
    const finalKey = (!normKey || isSameSheetDuplicate) ? generateTrackingGuid() : cloned.key;
    usedKeysOnTarget.add(finalKey.toLowerCase());

    const rNum = writeRowIdx++;
    const excelRow = targetSheet.getRow(rNum);
    writeItemToWorksheetRow(excelRow, rNum, {
      ...cloned,
      key: finalKey
    }, targetColMap);
    pastedCount++;
  });

  const saveResult = await saveWorkbookToDisk();
  if (saveResult) {
    selectedForDelete.clear();
    if (clearClipboardAfter) {
      copiedTrackerRowsClipboard = [];
    }
    activeWorkspaceTab = DAILY_TRACKER_WS_TAB;
    clearTransientFiltersOnNavigation({ resetScopeToCurrent: true });
    await loadWorkbookFromDisk(cleanTargetTab, {
      skipStagedRecoveryCheck: true,
      preserveStatus: saveResult === 'mirror-locked'
    });
    updateCopiedRowsToolbarUI();
    showToast(
      `📥 Pasted ${pastedCount} row(s) onto "${cleanTargetTab}"${createdNewSheet ? ' (created new tab)' : ''}.`,
      'success'
    );
    return true;
  }
  return false;
}

async function handlePasteCopiedTrackerRows() {
  if (!workbook || !Array.isArray(copiedTrackerRowsClipboard) || copiedTrackerRowsClipboard.length === 0) return;
  const activeTab = sheetSelect ? sheetSelect.value : '';
  if (!activeTab || !DATE_TAB_REGEX.test(activeTab)) {
    return handleCopySelectedRowsToTabPrompt(copiedTrackerRowsClipboard);
  }
  await pasteTrackerRowsToTargetTab(activeTab, copiedTrackerRowsClipboard, { clearClipboardAfter: true });
}

async function handleCopySelectedRowsToTabPrompt(overrideRows = null) {
  if (!workbook) return;
  const rowsToCopy = Array.isArray(overrideRows) && overrideRows.length > 0
    ? overrideRows.map(cloneTrackerRowForClipboard)
    : getSelectedItemsFromCheckboxSet().map(cloneTrackerRowForClipboard);
  if (rowsToCopy.length === 0) return;

  const currentTab = (sheetSelect && DATE_TAB_REGEX.test(sheetSelect.value))
    ? sheetSelect.value
    : (rowsToCopy[0].sourceSheetName && DATE_TAB_REGEX.test(rowsToCopy[0].sourceSheetName)
        ? rowsToCopy[0].sourceSheetName
        : (dateSheets[0] || getTodayFormatted()));

  const nextDayFromCurrent = addDaysToDateTab(currentTab, 1);
  const todayStr = getTodayFormatted();
  const suggestedDate = currentTab !== todayStr && !dateSheets.includes(todayStr)
    ? todayStr
    : nextDayFromCurrent;

  const selectOptions = [
    {
      value: '__CUSTOM_DATE__',
      label: `📅 Use date entered below (creates tab if needed)`
    },
    ...dateSheets.map(tab => ({
      value: tab,
      label: `📄 Existing Tab: ${tab}${tab === currentTab ? ' (Current)' : ''}`
    }))
  ];

  const rawChoice = await showModalPrompt({
    title: `📅 Copy ${rowsToCopy.length} Selected Row(s) to Day Tab`,
    badge: `${rowsToCopy.length} row(s)`,
    message:
      `Choose an existing daily tab or enter the next day's MM-DD-YYYY date below.\n` +
      `If the target tab doesn't exist yet, TrackerUI will create it automatically with your workbook's headers and paste the selected row(s).`,
    selectLabel: 'Destination Tab Option',
    selectOptions,
    defaultSelectValue: '__CUSTOM_DATE__',
    inputLabel: 'Target Daily Tab Date (MM-DD-YYYY)',
    defaultValue: suggestedDate,
    placeholder: 'MM-DD-YYYY',
    confirmText: `📥 Copy & Paste ${rowsToCopy.length} Row(s)`,
    validate: (val) => {
      const selVal = dialogSelect ? dialogSelect.value : '__CUSTOM_DATE__';
      const effective = (selVal && selVal !== '__CUSTOM_DATE__') ? selVal : String(val || '').trim();
      if (!DATE_TAB_REGEX.test(effective)) {
        return 'Please enter a valid date in exact MM-DD-YYYY format (e.g., 11-10-2026).';
      }
      return null;
    }
  });
  if (!rawChoice) return;

  const typedDate = (typeof rawChoice === 'object' ? rawChoice.value : rawChoice).trim();
  const selectedOpt = (typeof rawChoice === 'object' && rawChoice.selectValue) ? rawChoice.selectValue : '__CUSTOM_DATE__';
  const targetTabName = (selectedOpt && selectedOpt !== '__CUSTOM_DATE__') ? selectedOpt : typedDate;

  await pasteTrackerRowsToTargetTab(targetTabName, rowsToCopy, { clearClipboardAfter: true });
}

if (btnBatchCopyRows) {
  btnBatchCopyRows.addEventListener('click', handleCopySelectedTrackerRows);
}
if (btnBatchCopyToTab) {
  btnBatchCopyToTab.addEventListener('click', () => handleCopySelectedRowsToTabPrompt());
}
if (btnPasteTrackerRows) {
  btnPasteTrackerRows.addEventListener('click', handlePasteCopiedTrackerRows);
}
if (btnClearCopiedRows) {
  btnClearCopiedRows.addEventListener('click', () => {
    clearCopiedTrackerRowsClipboard();
    showToast('Cleared copied row(s) from clipboard.', 'info');
  });
}

// Keyboard shortcuts (Ctrl+C to copy checked rows, Ctrl+V to paste copied rows onto active daily tab)
document.addEventListener('keydown', (e) => {
  if (activeWorkspaceTab !== DAILY_TRACKER_WS_TAB || !workbook) return;
  const activeEl = document.activeElement;
  if (
    activeEl &&
    (activeEl.tagName === 'INPUT' ||
      activeEl.tagName === 'TEXTAREA' ||
      activeEl.tagName === 'SELECT' ||
      activeEl.isContentEditable)
  ) {
    return;
  }
  const anyModalOpen = document.querySelector('.modal-backdrop.open');
  if (anyModalOpen) return;

  const isMod = e.ctrlKey || e.metaKey;
  if (!isMod || e.altKey) return;

  const keyLower = String(e.key || '').toLowerCase();
  if (keyLower === 'c' && selectedForDelete.size > 0) {
    const selText = typeof window.getSelection === 'function' ? String(window.getSelection() || '').trim() : '';
    if (!selText) {
      e.preventDefault();
      handleCopySelectedTrackerRows();
    }
  } else if (keyLower === 'v' && Array.isArray(copiedTrackerRowsClipboard) && copiedTrackerRowsClipboard.length > 0) {
    e.preventDefault();
    handlePasteCopiedTrackerRows();
  }
});

async function executeBatchDeleteSelected() {
  if (selectedForDelete.size === 0 || !workbook) return;

  const itemsToDelete = allRows.filter(item =>
    selectedForDelete.has(getRowCompositeId(item.sheetName, item.rowNumber))
  );
  if (itemsToDelete.length === 0) return;

  const previewLines = itemsToDelete
    .slice(0, 12)
    .map(r => `• [${r.sheetName} • Row ${r.rowNumber}] ${r.subject} (${r.broker || 'No Broker'})`)
    .join('\n');
  const moreNotice = itemsToDelete.length > 12 ? `\n...and ${itemsToDelete.length - 12} more.` : '';

  const confirmed = await showModalConfirm({
    title: `🗑️ Delete ${itemsToDelete.length} Selected Request(s)?`,
    badge: `${itemsToDelete.length} selected`,
    message:
      `Are you sure you want to permanently delete ${itemsToDelete.length} selected request(s)?\n\n` +
      `${previewLines}${moreNotice}\n\n` +
      `This will remove the row(s) from the workbook and save immediately.`,
    confirmText: `🗑️ Delete ${itemsToDelete.length} Request(s)`,
    variant: 'danger'
  });
  if (!confirmed) return;

  const currentTab = sheetSelect.value;
  await loadWorkbookFromDisk(currentTab);

  const bySheet = new Map();
  itemsToDelete.forEach(item => {
    if (!bySheet.has(item.sheetName)) {
      bySheet.set(item.sheetName, []);
    }
    bySheet.get(item.sheetName).push(item);
  });

  let deletedCount = 0;
  bySheet.forEach((sheetItems, sheetName) => {
    const ws = workbook.getWorksheet(sheetName);
    if (!ws) return;

    sheetItems.sort((a, b) => b.rowNumber - a.rowNumber);
    sheetItems.forEach(delItem => {
      ws.spliceRows(delItem.rowNumber, 1);
      deletedCount++;
    });
  });

  const saveResult = await saveWorkbookToDisk();
  if (saveResult) {
    selectedForDelete.clear();
    setDeleteMode(false);
    await loadWorkbookFromDisk(currentTab, {
      preserveStatus: saveResult === 'mirror-locked'
    });
    showToast(`Deleted ${deletedCount} request(s) from workbook.`, 'success');
  }
}

btnConfirmDeleteSelected.addEventListener('click', executeBatchDeleteSelected);
if (btnBatchDeleteSelected) {
  btnBatchDeleteSelected.addEventListener('click', executeBatchDeleteSelected);
}

// Column Header Sorting & Drag-to-Resize Handles
if (tableHead) {
  tableHead.addEventListener('mousedown', (e) => {
    if (e.button !== 0) return;
    const resizer = e.target.closest('.tracker-col-resizer[data-resize-col-key]');
    if (!resizer) return;
    const thEl = resizer.closest('th[data-col-width-key]');
    const tableEl = tableHead.closest('table');
    if (!thEl || !tableEl) return;
    e.preventDefault();
    e.stopPropagation();

    // Snapshot all current column widths on first custom resize so adjacent columns don't jump
    if (!tableEl.classList.contains('has-custom-col-widths')) {
      tableHead.querySelectorAll('th[data-col-width-key]').forEach(otherTh => {
        const w = Math.max(64, Math.round(otherTh.getBoundingClientRect().width));
        otherTh.style.width = `${w}px`;
        otherTh.style.minWidth = `${w}px`;
      });
      tableEl.classList.add('has-custom-col-widths');
    }

    const startW = thEl.getBoundingClientRect().width;
    resizer.classList.add('resizing');
    document.body.classList.add('is-resizing-tracker-col');
    trackerColResizeState = {
      colKey: resizer.dataset.resizeColKey,
      thEl,
      resizerEl: resizer,
      tableEl,
      startX: e.clientX,
      startWidthPx: startW,
      currentWidthPx: startW
    };
  });

  tableHead.addEventListener('dblclick', (e) => {
    const resizer = e.target.closest('.tracker-col-resizer[data-resize-col-key]');
    if (!resizer) return;
    const thEl = resizer.closest('th[data-col-width-key]');
    if (!thEl) return;
    e.preventDefault();
    e.stopPropagation();
    lastTrackerColResizeAt = Date.now();
    autoFitTrackerTableColumnWidth(thEl);
  });

  tableHead.addEventListener('click', (e) => {
    if (e.target.closest('.tracker-col-resizer') || Date.now() - lastTrackerColResizeAt < 250) {
      return;
    }
    const subSort = e.target.closest('.th-sub-sort');
    const thSort = e.target.closest('th.sortable-th');
    const col = subSort ? subSort.dataset.sort : (thSort ? thSort.dataset.sort : null);
    if (!col) return;

    if (sortState.column === col) {
      if (sortState.direction === 'asc') {
        sortState.direction = 'desc';
      } else {
        sortState.column = null;
        sortState.direction = 'asc';
      }
    } else {
      sortState.column = col;
      sortState.direction = 'asc';
    }
    renderList();
  });
}

window.addEventListener('mousemove', (e) => {
  if (!trackerColResizeState) return;
  const deltaX = e.clientX - trackerColResizeState.startX;
  const nextW = Math.max(64, Math.min(680, Math.round(trackerColResizeState.startWidthPx + deltaX)));
  trackerColResizeState.currentWidthPx = nextW;
  trackerColResizeState.thEl.style.width = `${nextW}px`;
  trackerColResizeState.thEl.style.minWidth = `${nextW}px`;
});

window.addEventListener('mouseup', () => {
  if (!trackerColResizeState) return;
  const st = trackerColResizeState;
  trackerColResizeState = null;
  lastTrackerColResizeAt = Date.now();
  if (st.resizerEl) st.resizerEl.classList.remove('resizing');
  document.body.classList.remove('is-resizing-tracker-col');
  if (Math.abs(st.currentWidthPx - st.startWidthPx) >= 2) {
    setSavedTrackerTableColWidth(st.colKey, st.currentWidthPx);
  }
});

if (typeof btnResetTableColWidths !== 'undefined' && btnResetTableColWidths) {
  btnResetTableColWidths.addEventListener('click', resetActiveTrackerTableColWidths);
}

if (typeof btnToggleHideCompleted !== 'undefined' && btnToggleHideCompleted) {
  syncHideCompletedButtonUI(0);
  btnToggleHideCompleted.addEventListener('click', () => {
    if (typeof setHideCompletedPreference === 'function') {
      setHideCompletedPreference(!hideCompletedRequests);
    } else {
      hideCompletedRequests = !hideCompletedRequests;
    }
    renderList();
  });
}

if (typeof btnToggleHideUuid !== 'undefined' && btnToggleHideUuid) {
  syncHideUuidButtonUI();
  btnToggleHideUuid.addEventListener('click', () => {
    if (typeof setHideUuidKeyPreference === 'function') {
      setHideUuidKeyPreference(!hideUuidKey);
    } else {
      hideUuidKey = !hideUuidKey;
      syncHideUuidButtonUI();
    }
    renderList();
  });
}

function refreshRowLastAccessedDOM(trackingKey) {
  const normKey = String(trackingKey || '').trim().toLowerCase();
  if (!normKey || !listBody || typeof getRowLastAccessedMs !== 'function' || typeof formatRowLastAccessedLabel !== 'function') {
    return;
  }
  const info = formatRowLastAccessedLabel(getRowLastAccessedMs(normKey));
  listBody.querySelectorAll('[data-accessed-key]').forEach(el => {
    if (el.dataset.accessedKey === normKey) {
      el.title = info.full;
      el.textContent = info.ms ? info.short : '';
      el.classList.toggle('is-empty', !info.ms);
    }
  });
}

// Delegate clicks for Shadow New Request Row, Date Tab jump, Key Chip, Like-Item chips, Collapsible Notes, Details button, Edit button, and Row Selection
listBody.addEventListener('click', async (e) => {
  const copyBtn = e.target.closest('.btn-inline-copy');
  if (copyBtn) {
    e.stopPropagation();
    e.preventDefault();
    if (typeof copyTextToClipboard === 'function') {
      copyTextToClipboard(copyBtn.dataset.copyVal, copyBtn.dataset.copyLabel || 'Value', copyBtn);
    }
    return;
  }

  const shadowRow = e.target.closest('.shadow-new-row');
  if (shadowRow) {
    if (bottomEditorDock && bottomEditorDock.style.display !== 'none') {
      if (!dockEditingTarget && typeof hasUnsavedNewDockDraft === 'function' && hasUnsavedNewDockDraft()) {
        if (isBottomDockCollapsed) setBottomDockCollapsed(false);
        if (typeof focusFirstEditorInput === 'function') focusFirstEditorInput('dock');
      } else {
        setBottomDockNewMode(true, { restoreSavedDraft: true });
      }
    } else {
      openNewRequestModal();
    }
    return;
  }

  const dateTabBtn = e.target.closest('.clickable-date-tab');
  if (dateTabBtn && dateTabBtn.dataset.dateTab) {
    const targetTab = dateTabBtn.dataset.dateTab;
    if (dateSheets.includes(targetTab)) {
      clearTransientFiltersOnNavigation({ resetScopeToCurrent: true });
      sheetSelect.value = targetTab;
      onActiveSheetChanged();
      showToast(`Switched to daily tab "${targetTab}".`, 'info');
    }
    return;
  }

  const keyChip = e.target.closest('.key-chip');
  if (keyChip && keyChip.dataset.key) {
    toggleFieldFilter('key', 'Tracking ID', keyChip.dataset.key);
    return;
  }

  const filterEl = e.target.closest('[data-filter-field]');
  if (filterEl && filterEl.dataset.filterField && filterEl.dataset.filterVal) {
    toggleFieldFilter(
      filterEl.dataset.filterField,
      filterEl.dataset.filterLabel || filterEl.dataset.filterField,
      filterEl.dataset.filterVal
    );
    return;
  }

  const noteToggle = e.target.closest('.notes-toggle');
  if (noteToggle) {
    const targetDiv = document.getElementById(noteToggle.dataset.target);
    if (targetDiv) {
      const expanded = targetDiv.classList.toggle('expanded');
      noteToggle.textContent = expanded ? '[- Collapse]' : '[+ Expand]';
    }
    return;
  }

  const detailsBtn = e.target.closest('.btn-details');
  if (detailsBtn) {
    const detSheet = detailsBtn.dataset.sheet;
    const detRow = Number(detailsBtn.dataset.row);
    const detItem = allRows.find(r => r.sheetName === detSheet && r.rowNumber === detRow);
    if (detItem && detItem.key && typeof recordRowLastAccessed === 'function') {
      recordRowLastAccessed(detItem.key);
      refreshRowLastAccessedDOM(detItem.key);
    }
    openTicketDetailsModal(detSheet, detRow);
    return;
  }

  const editBtn = e.target.closest('.btn-edit');
  if (editBtn) {
    if (bottomEditorDock && bottomEditorDock.style.display !== 'none') {
      if (typeof confirmSwitchAwayFromUnsavedNewDraft === 'function') {
        const proceed = await confirmSwitchAwayFromUnsavedNewDraft('open that row for editing');
        if (!proceed) return;
      }
      loadRowIntoBottomDock(editBtn.dataset.sheet, Number(editBtn.dataset.row), true);
    } else {
      openEditModal(editBtn.dataset.sheet, Number(editBtn.dataset.row));
    }
    return;
  }

  // Handle clicks inside the permanent left-edge checkbox cell without opening the single-row editor
  const selectCellEl = e.target.closest('.select-cell, .delete-cell');
  if (selectCellEl) {
    if (!e.target.closest('.chk-delete-row')) {
      const cellChk = selectCellEl.querySelector('.chk-delete-row');
      if (cellChk) {
        cellChk.checked = !cellChk.checked;
        cellChk.dispatchEvent(new Event('change', { bubbles: true }));
      }
    }
    return;
  }

  // Clicking anywhere else on a data row populates the Bottom Editor Dock
  if (!isDeleteMode && !e.target.closest('.chk-delete-row')) {
    const rowTr = e.target.closest('tr[data-sheet][data-row]');
    if (rowTr && bottomEditorDock && bottomEditorDock.style.display !== 'none') {
      if (typeof confirmSwitchAwayFromUnsavedNewDraft === 'function') {
        const proceed = await confirmSwitchAwayFromUnsavedNewDraft('open that row');
        if (!proceed) return;
      }
      loadRowIntoBottomDock(rowTr.dataset.sheet, Number(rowTr.dataset.row), false);
    }
  }
});

listBody.addEventListener('dblclick', (e) => {
  const editBtn = e.target.closest('.btn-edit');
  if (editBtn) {
    openEditModal(editBtn.dataset.sheet, Number(editBtn.dataset.row));
  }
});

document.getElementById('btnClearKeyFilter').addEventListener('click', clearAllSortAndFilters);
if (btnClearSortFilter) {
  btnClearSortFilter.addEventListener('click', clearAllSortAndFilters);
}

scopeSelect.addEventListener('change', () => {
  clearTransientFiltersOnNavigation({ resetScopeToCurrent: false });
  renderList();
});

searchInput.addEventListener('input', () => {
  if (searchInput.value.trim().length > 0) {
    activeKeyFilter = null;
    activeFieldFilter = null;
  }
  renderList();
});

