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
    const inferredActionColObj = customCols.find(c =>
      c.colIndex !== colMap.primaryCol &&
      c.colIndex !== colMap.statusCol &&
      c.colIndex !== colMap.notesCol &&
      /\b(action|subtype|issue|category|reason|request\s*type|type)\b/i.test(c.header)
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

    const hasData = [subject, ticket, broker, master, bidder, action, yard, lots, notes, status]
      .some(val => val !== '');

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
      key
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
    setBottomDockNewMode();
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

function getActiveTableColSpan() {
  if (activeSchemaConfig && !activeSchemaConfig.isStandardSchema && Array.isArray(activeSchemaConfig.columns) && activeSchemaConfig.columns.length > 0) {
    return 4 + activeSchemaConfig.columns.length;
  }
  return 12;
}

function renderTableHeader() {
  if (!tableHead) return;
  const selectColClass = isDeleteMode ? 'delete-col' : 'select-col';

  if (activeSchemaConfig && !activeSchemaConfig.isStandardSchema && Array.isArray(activeSchemaConfig.columns) && activeSchemaConfig.columns.length > 0) {
    // Reset incompatible sortState if left over from standard schema
    if (sortState.column && !['date', 'key'].includes(sortState.column) && !sortState.column.startsWith('col_')) {
      sortState = { column: null, direction: 'asc' };
    }
    const trkLetter = colNumberToExcelLetters(activeSchemaConfig.trackingCol || (activeSchemaConfig.columns.length + 1));
    const customThsHtml = activeSchemaConfig.columns.map(c => {
      const sortKey = `col_${c.colIndex}`;
      return `
        <th class="sortable-th" data-sort="${sortKey}" title="Click to sort by ${escapeHtml(c.header)}">
          <span>${escapeHtml(c.header)}</span><span class="sort-indicator" data-ind="${sortKey}">↕</span>
        </th>
      `;
    }).join('');

    tableHead.innerHTML = `
      <tr>
        <th id="thDeleteSelect" class="${selectColClass}" title="Select rows for batch status update or deletion">
          <input type="checkbox" id="chkSelectAllHeader" title="Select all visible requests">
        </th>
        <th class="sortable-th" data-sort="date" title="Click to sort by Date Tab">
          <span>Date Tab</span><span class="sort-indicator" data-ind="date">↕</span>
        </th>
        <th class="sortable-th" data-sort="key" title="Click to sort by Tracking ID">
          <span>Tracking ID (Col ${escapeHtml(trkLetter)})</span><span class="sort-indicator" data-ind="key">↕</span>
        </th>
        ${customThsHtml}
        <th>Actions</th>
      </tr>
    `;
    syncSelectAllCheckboxes();
    return;
  }

  if (sortState.column && sortState.column.startsWith('col_')) {
    sortState = { column: null, direction: 'asc' };
  }

  tableHead.innerHTML = `
    <tr>
      <th id="thDeleteSelect" class="${selectColClass}" title="Select rows for batch status update or deletion">
        <input type="checkbox" id="chkSelectAllHeader" title="Select all visible requests">
      </th>
      <th class="sortable-th" data-sort="date" title="Click to sort by Date Tab">
        <span>Date Tab</span><span class="sort-indicator" data-ind="date">↕</span>
      </th>
      <th class="sortable-th" data-sort="key" title="Click to sort by Tracking ID">
        <span>Tracking ID (Col K)</span><span class="sort-indicator" data-ind="key">↕</span>
      </th>
      <th class="sortable-th" data-sort="subject" title="Click to sort by Subject">
        <span>Subject</span><span class="sort-indicator" data-ind="subject">↕</span>
      </th>
      <th class="sortable-th" data-sort="ticket" title="Click to sort by Ticket #">
        <span>Ticket #</span><span class="sort-indicator" data-ind="ticket">↕</span>
      </th>
      <th class="sortable-th" data-sort="broker" title="Click to sort by Broker Name">
        <span>Broker</span><span class="sort-indicator" data-ind="broker">↕</span>
      </th>
      <th class="sortable-th" data-sort="master" title="Click Master or Bidder to sort">
        <span class="th-sub-sort" data-sort="master" title="Sort by Master #">Master<span class="sort-indicator" data-ind="master">↕</span></span>
        <span class="th-sep">/</span>
        <span class="th-sub-sort" data-sort="bidder" title="Sort by Bidder #">Bidder<span class="sort-indicator" data-ind="bidder">↕</span></span>
      </th>
      <th class="sortable-th" data-sort="action" title="Click to sort by Requested Action">
        <span>Action</span><span class="sort-indicator" data-ind="action">↕</span>
      </th>
      <th class="sortable-th" data-sort="yard" title="Click Yard or Lots to sort">
        <span class="th-sub-sort" data-sort="yard" title="Sort by Yard #">Yard<span class="sort-indicator" data-ind="yard">↕</span></span>
        <span class="th-sep">/</span>
        <span class="th-sub-sort" data-sort="lots" title="Sort by Lot Numbers">Lots<span class="sort-indicator" data-ind="lots">↕</span></span>
      </th>
      <th class="sortable-th" data-sort="notes" title="Click to sort by Action Notes">
        <span>Action Notes</span><span class="sort-indicator" data-ind="notes">↕</span>
      </th>
      <th class="sortable-th" data-sort="status" title="Click to sort by Status">
        <span>Status</span><span class="sort-indicator" data-ind="status">↕</span>
      </th>
      <th>Actions</th>
    </tr>
  `;
  syncSelectAllCheckboxes();
}

function createShadowNewRowElement(colSpan, currentSheet) {
  if (!workbook || isDeleteMode || !currentSheet || !DATE_TAB_REGEX.test(currentSheet)) {
    return null;
  }
  const isCustom = activeSchemaConfig && !activeSchemaConfig.isStandardSchema && activeSchemaConfig.columns.length > 0;
  const ghostPillsHtml = isCustom
    ? activeSchemaConfig.columns.slice(0, 5).map(c => `<span class="shadow-ghost-pill">${escapeHtml(c.header)}</span>`).join('')
    : `
        <span class="shadow-ghost-pill">Subject</span>
        <span class="shadow-ghost-pill">Ticket #</span>
        <span class="shadow-ghost-pill">Broker &amp; Master #</span>
        <span class="shadow-ghost-pill">Requested Action</span>
        <span class="shadow-ghost-pill">Yard / Lots</span>
      `;

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
  status: 'Status'
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

function renderClickableLotsHtml(lotsStr) {
  const trimmed = String(lotsStr || '').trim();
  if (!trimmed || trimmed === '—') return '—';

  const parts = trimmed.split(/([,;/|\s]+)/);
  return parts
    .map(part => {
      if (!part || /^[,;/|\s]+$/.test(part)) {
        return escapeHtml(part);
      }
      const activeClass = isFieldFilterActive('lot', part) ? ' active-filter-match' : '';
      return `<span class="clickable-code${activeClass}" data-filter-field="lot" data-filter-label="Lot #" data-filter-val="${escapeHtml(part)}" title="Click to show all requests with Lot # ${escapeHtml(part)}">${escapeHtml(part)}</span>`;
    })
    .join('');
}

function renderClickableTicketTokensHtml(ticketStr, filterField = 'ticket', filterLabel = 'Ticket #') {
  const trimmed = String(ticketStr || '').trim();
  if (!trimmed || trimmed === '—') return '—';
  const tokens = splitNormalizedTicketTokens(trimmed);
  if (tokens.length <= 1) {
    const activeClass = isFieldFilterActive(filterField, trimmed) ? ' active-filter-match' : '';
    return `<span class="clickable-code${activeClass}" data-filter-field="${escapeHtml(filterField)}" data-filter-label="${escapeHtml(filterLabel)}" data-filter-val="${escapeHtml(trimmed)}" title="Click to show all requests with ${escapeHtml(filterLabel)}: ${escapeHtml(trimmed)}">${escapeHtml(trimmed)}</span>`;
  }
  return tokens
    .map(tok => {
      const activeClass = isFieldFilterActive(filterField, tok) ? ' active-filter-match' : '';
      return `<span class="clickable-code${activeClass}" data-filter-field="${escapeHtml(filterField)}" data-filter-label="${escapeHtml(filterLabel)}" data-filter-val="${escapeHtml(tok)}" title="Click to show all requests with ${escapeHtml(filterLabel)}: ${escapeHtml(tok)}">${escapeHtml(tok)}</span>`;
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
        (thSort === 'yard' && (sortState.column === 'yard' || sortState.column === 'lots'));
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
    item.sheetName
  ].join(' ').toLowerCase();
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

  const filtered = allRows.filter(item => {
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

    listBody.innerHTML = `<tr><td colspan="${colSpan}" class="empty-table-cell"><div>No matching requests found.</div>${emptyRolloverBtnHtml}</td></tr>`;
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

  sorted.forEach((item, idx) => {
    const tr = document.createElement('tr');
    tr.dataset.sheet = item.sheetName;
    tr.dataset.row = String(item.rowNumber);

    const compId = getRowCompositeId(item.sheetName, item.rowNumber);
    const isRowChecked = selectedForDelete.has(compId);
    if (isRowChecked) {
      tr.classList.add(isDeleteMode ? 'selected-for-delete' : 'selected-for-batch');
    }
    if (
      dockEditingTarget &&
      dockEditingTarget.sheetName === item.sheetName &&
      dockEditingTarget.rowNumber === item.rowNumber
    ) {
      tr.classList.add('active-editing-row');
    }

    const isComplete = isCompletedStatus(item.status);
    const noteText = item.notes || '—';
    const isLongNote = noteText.length > 45 || noteText.includes('\n');
    const shortGuid = item.key.length > 8 ? `${item.key.slice(0, 8)}...` : item.key;

    const rowCheckboxCell = `<td class="${isDeleteMode ? 'delete-cell' : 'select-cell'}">
      <input type="checkbox" class="chk-delete-row" data-comp="${escapeHtml(compId)}" data-key="${escapeHtml(item.key)}" ${isRowChecked ? 'checked' : ''} aria-label="Select request row ${item.rowNumber} on ${escapeHtml(item.sheetName)}">
    </td>`;

    const keyActiveClass = isFieldFilterActive('key', item.key) ? ' active-filter-match' : '';

    if (isCustomSchema) {
      const customCellsHtml = activeSchemaConfig.columns.map(c => {
        const fieldKey = `col_${c.colIndex}`;
        const rawVal = (item.customValues && item.customValues[fieldKey] !== undefined)
          ? String(item.customValues[fieldKey])
          : '';
        const trimmedVal = rawVal.trim();

        if (c.colIndex === activeSchemaConfig.statusCol) {
          const stText = trimmedVal || item.status || 'Pending';
          const stDone = isCompletedStatus(stText);
          const stActive = isFieldFilterActive(fieldKey, stText) ? ' active-filter-match' : '';
          return `
            <td>
              <span class="badge clickable-badge ${stDone ? 'badge-complete' : 'badge-pending'}${stActive}" data-filter-field="${fieldKey}" data-filter-label="${escapeHtml(c.header)}" data-filter-val="${escapeHtml(stText)}" title="Click to filter by ${escapeHtml(c.header)}: ${escapeHtml(stText)}">
                ${escapeHtml(stText)}
              </span>
            </td>
          `;
        }

        if (c.colIndex === activeSchemaConfig.notesCol) {
          const cNote = trimmedVal || '—';
          const cLong = cNote.length > 45 || cNote.includes('\n');
          return `
            <td class="notes-cell">
              <div class="notes-preview" id="note-${idx}-${c.colIndex}">${escapeHtml(cNote)}</div>
              ${cLong ? `<span class="notes-toggle" data-target="note-${idx}-${c.colIndex}">[+ Expand]</span>` : ''}
            </td>
          `;
        }

        if (c.colIndex === activeSchemaConfig.ticketCol) {
          return `<td>${renderClickableTicketTokensHtml(trimmedVal, fieldKey, c.header)}</td>`;
        }

        if (!trimmedVal) {
          return '<td>—</td>';
        }
        const cellActive = isFieldFilterActive(fieldKey, trimmedVal) ? ' active-filter-match' : '';
        return `
          <td>
            <span class="clickable-cell${cellActive}" data-filter-field="${fieldKey}" data-filter-label="${escapeHtml(c.header)}" data-filter-val="${escapeHtml(trimmedVal)}" title="Click to filter by ${escapeHtml(c.header)}: ${escapeHtml(trimmedVal)}">
              ${escapeHtml(trimmedVal)}
            </span>
          </td>
        `;
      }).join('');

      tr.innerHTML = `
        ${rowCheckboxCell}
        <td style="white-space:nowrap;">
          <button type="button" class="clickable-date-tab" data-date-tab="${escapeHtml(item.sheetName)}" title="Click to switch to daily tab ${escapeHtml(item.sheetName)}">
            ${escapeHtml(item.sheetName)}
          </button>
        </td>
        <td>
          <span class="key-chip${keyActiveClass}" title="Click to filter by Tracking ID: ${escapeHtml(item.key)}" data-key="${escapeHtml(item.key)}">
            🔍 ${escapeHtml(shortGuid)}
          </span>
        </td>
        ${customCellsHtml}
        <td>
          <div class="row-actions-group">
            <button type="button" class="btn-details" data-sheet="${escapeHtml(item.sheetName)}" data-row="${item.rowNumber}" title="View Details &amp; Aggregated Timestamped Notes">Details</button>
            <button type="button" class="btn-edit" data-sheet="${escapeHtml(item.sheetName)}" data-row="${item.rowNumber}" title="Load into Bottom Editor (Double-click for Modal)">Edit</button>
          </div>
        </td>
      `;
      listBody.appendChild(tr);
      return;
    }

    const subjectActiveClass = isFieldFilterActive('subject', item.subject) ? ' active-filter-match' : '';
    const brokerActiveClass = isFieldFilterActive('broker', item.broker) ? ' active-filter-match' : '';
    const masterActiveClass = isFieldFilterActive('master', item.master) ? ' active-filter-match' : '';
    const bidderActiveClass = isFieldFilterActive('bidder', item.bidder) ? ' active-filter-match' : '';
    const actionActiveClass = isFieldFilterActive('action', item.action) ? ' active-filter-match' : '';
    const yardActiveClass = isFieldFilterActive('yard', item.yard) ? ' active-filter-match' : '';
    const statusActiveClass = isFieldFilterActive('status', item.status) ? ' active-filter-match' : '';

    const subjectHtml = item.subject
      ? `<span class="clickable-cell${subjectActiveClass}" data-filter-field="subject" data-filter-label="Subject" data-filter-val="${escapeHtml(item.subject)}" title="Click to show all requests with Subject: ${escapeHtml(item.subject)}">${escapeHtml(item.subject)}</span>`
      : '—';

    const ticketHtml = renderClickableTicketTokensHtml(item.ticket, 'ticket', 'Ticket #');

    const brokerHtml = item.broker
      ? `<span class="clickable-cell${brokerActiveClass}" data-filter-field="broker" data-filter-label="Broker" data-filter-val="${escapeHtml(item.broker)}" title="Click to show all requests for Broker: ${escapeHtml(item.broker)}">${escapeHtml(item.broker)}</span>`
      : '—';

    const masterHtml = item.master
      ? `<span class="clickable-code${masterActiveClass}" data-filter-field="master" data-filter-label="Master #" data-filter-val="${escapeHtml(item.master)}" title="Click to show all requests with Master # ${escapeHtml(item.master)}">${escapeHtml(item.master)}</span>`
      : '—';

    const bidderHtml = item.bidder
      ? `<span class="clickable-code${bidderActiveClass}" data-filter-field="bidder" data-filter-label="Bidder #" data-filter-val="${escapeHtml(item.bidder)}" title="Click to show all requests with Bidder # ${escapeHtml(item.bidder)}">${escapeHtml(item.bidder)}</span>`
      : '—';

    const actionHtml = item.action
      ? `<span class="clickable-cell${actionActiveClass}" data-filter-field="action" data-filter-label="Action" data-filter-val="${escapeHtml(item.action)}" title="Click to show all requests with Action: ${escapeHtml(item.action)}">${escapeHtml(item.action)}</span>`
      : '—';

    const yardHtml = item.yard
      ? `<span class="clickable-code${yardActiveClass}" data-filter-field="yard" data-filter-label="Yard #" data-filter-val="${escapeHtml(item.yard)}" title="Click to show all requests for Yard # ${escapeHtml(item.yard)}">${escapeHtml(item.yard)}</span>`
      : '—';

    const lotsHtml = renderClickableLotsHtml(item.lots);

    tr.innerHTML = `
      ${rowCheckboxCell}
      <td style="white-space:nowrap;">
        <button type="button" class="clickable-date-tab" data-date-tab="${escapeHtml(item.sheetName)}" title="Click to switch to daily tab ${escapeHtml(item.sheetName)}">
          ${escapeHtml(item.sheetName)}
        </button>
      </td>
      <td>
        <span class="key-chip${keyActiveClass}" title="Click to filter by Tracking ID: ${escapeHtml(item.key)}" data-key="${escapeHtml(item.key)}">
          🔍 ${escapeHtml(shortGuid)}
        </span>
      </td>
      <td>${subjectHtml}</td>
      <td>${ticketHtml}</td>
      <td>${brokerHtml}</td>
      <td class="cell-mono-sm">
        M: ${masterHtml}<br>B: ${bidderHtml}
      </td>
      <td>${actionHtml}</td>
      <td class="cell-sm">
        Y: ${yardHtml}<br>L: ${lotsHtml}
      </td>
      <td class="notes-cell">
        <div class="notes-preview" id="note-${idx}">${escapeHtml(noteText)}</div>
        ${isLongNote ? `<span class="notes-toggle" data-target="note-${idx}">[+ Expand]</span>` : ''}
      </td>
      <td>
        <span class="badge clickable-badge ${isComplete ? 'badge-complete' : 'badge-pending'}${statusActiveClass}" data-filter-field="status" data-filter-label="Status" data-filter-val="${escapeHtml(item.status)}" title="Click to show all requests with Status: ${escapeHtml(item.status)}">
          ${escapeHtml(item.status)}
        </span>
      </td>
      <td>
        <div class="row-actions-group">
          <button type="button" class="btn-details" data-sheet="${escapeHtml(item.sheetName)}" data-row="${item.rowNumber}" title="View Ticket Details &amp; Aggregated Timestamped Notes">Details</button>
          <button type="button" class="btn-edit" data-sheet="${escapeHtml(item.sheetName)}" data-row="${item.rowNumber}" title="Load into Bottom Editor (Double-click for Modal)">Edit</button>
        </div>
      </td>
    `;
    listBody.appendChild(tr);
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
  if (chk.checked) {
    selectedForDelete.add(compId);
    if (tr) {
      tr.classList.toggle('selected-for-delete', isDeleteMode);
      tr.classList.toggle('selected-for-batch', !isDeleteMode);
    }
  } else {
    selectedForDelete.delete(compId);
    if (tr) {
      tr.classList.remove('selected-for-delete', 'selected-for-batch');
    }
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
        applyCustomHeadersSheetFormatting(targetSheet, activeSchemaConfig.columns.map(c => c.header));
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

// Column Header Sorting
if (tableHead) {
  tableHead.addEventListener('click', (e) => {
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

// Delegate clicks for Shadow New Request Row, Date Tab jump, Key Chip, Like-Item chips, Collapsible Notes, Details button, Edit button, and Row Selection
listBody.addEventListener('click', (e) => {
  const shadowRow = e.target.closest('.shadow-new-row');
  if (shadowRow) {
    if (bottomEditorDock && bottomEditorDock.style.display !== 'none') {
      setBottomDockNewMode(true);
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
    openTicketDetailsModal(detailsBtn.dataset.sheet, Number(detailsBtn.dataset.row));
    return;
  }

  const editBtn = e.target.closest('.btn-edit');
  if (editBtn) {
    if (bottomEditorDock && bottomEditorDock.style.display !== 'none') {
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

