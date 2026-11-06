// ============================================================================
// tracker-table.js — Worksheet Row Indexing, UUIDv4 Backfill, Sheet Stepper &
//                    Tab Bar, Daily Request Table, Search, Sort, Like-Item
//                    Filters, Permanent Row Selection, Batch Status & Delete
// ============================================================================

// --- INDEX WORKSHEET ROWS & LEGACY GUID BACKFILL ---
function indexWorksheetRows(worksheet, sheetName) {
  const colMap = getWorksheetColumnMap(worksheet);
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
    if (existingKey) {
      usedKeysOnSheet.add(existingKey);
    } else if (colMap.isTemplateLayout) {
      const col11Text = cellText(row, 11);
      if (UUID_REGEX.test(col11Text)) {
        usedKeysOnSheet.add(col11Text.toLowerCase());
      }
    }
  });

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

// --- 5. RENDER LIST VIEW, AUTOMATIC GLOBAL SEARCH, & DELETE SELECTION MODE ---
function getRowCompositeId(sheetName, rowNumber) {
  return `${sheetName}::${rowNumber}`;
}

function createShadowNewRowElement(colSpan, currentSheet) {
  if (!workbook || isDeleteMode || !currentSheet || !DATE_TAB_REGEX.test(currentSheet)) {
    return null;
  }
  const tr = document.createElement('tr');
  tr.className = 'shadow-new-row';
  tr.title = `Click to add a new support request to ${currentSheet}`;
  tr.innerHTML = `
    <td colspan="${colSpan}">
      <div class="shadow-row-inner">
        <div class="shadow-row-left">
          <span class="shadow-plus-btn" aria-hidden="true">＋</span>
          <span class="shadow-row-label">Add new support request to <strong>${escapeHtml(currentSheet)}</strong>...</span>
          <div class="shadow-ghost-pills">
            <span class="shadow-ghost-pill">Subject</span>
            <span class="shadow-ghost-pill">Ticket #</span>
            <span class="shadow-ghost-pill">Broker &amp; Master #</span>
            <span class="shadow-ghost-pill">Requested Action</span>
            <span class="shadow-ghost-pill">Yard / Lots</span>
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

function doesItemMatchFieldFilter(item, filter) {
  if (!filter || !filter.value) return true;
  const target = filter.value.trim().toLowerCase();
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
    case 'ticket':
      return item.ticket.trim().toLowerCase() === target;
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

    const valA = String(a[col] || '').trim();
    const valB = String(b[col] || '').trim();

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
        const sortLabel = SORT_COLUMN_LABELS[sortState.column] || sortState.column;
        const arrow = sortState.direction === 'asc' ? '▲ Asc' : '▼ Desc';
        filterBannerDesc.textContent += ` • Sorted by ${sortLabel} (${arrow}) —`;
      }
    } else if (hasSort) {
      const sortLabel = SORT_COLUMN_LABELS[sortState.column] || sortState.column;
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

function renderList() {
  const query = searchInput.value.trim().toLowerCase();
  const scope = scopeSelect.value;
  const currentSheet = sheetSelect.value;
  const activeFilter = activeFieldFilter || (activeKeyFilter ? { field: 'key', label: 'Tracking ID', value: activeKeyFilter } : null);

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
        const haystack = [
          item.subject, item.ticket, item.broker, item.master,
          item.bidder, item.action, item.yard, item.lots, item.notes, item.key, item.sheetName
        ].join(' ').toLowerCase();
        return haystack.includes(query);
      }
      return true;
    }
    if (query.length > 0) {
      const haystack = [
        item.subject, item.ticket, item.broker, item.master,
        item.bidder, item.action, item.yard, item.lots, item.notes, item.key, item.sheetName
      ].join(' ').toLowerCase();
      return haystack.includes(query);
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
  const colSpan = 12;
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
    const subjectActiveClass = isFieldFilterActive('subject', item.subject) ? ' active-filter-match' : '';
    const ticketActiveClass = isFieldFilterActive('ticket', item.ticket) ? ' active-filter-match' : '';
    const brokerActiveClass = isFieldFilterActive('broker', item.broker) ? ' active-filter-match' : '';
    const masterActiveClass = isFieldFilterActive('master', item.master) ? ' active-filter-match' : '';
    const bidderActiveClass = isFieldFilterActive('bidder', item.bidder) ? ' active-filter-match' : '';
    const actionActiveClass = isFieldFilterActive('action', item.action) ? ' active-filter-match' : '';
    const yardActiveClass = isFieldFilterActive('yard', item.yard) ? ' active-filter-match' : '';
    const statusActiveClass = isFieldFilterActive('status', item.status) ? ' active-filter-match' : '';

    const subjectHtml = item.subject
      ? `<span class="clickable-cell${subjectActiveClass}" data-filter-field="subject" data-filter-label="Subject" data-filter-val="${escapeHtml(item.subject)}" title="Click to show all requests with Subject: ${escapeHtml(item.subject)}">${escapeHtml(item.subject)}</span>`
      : '—';

    const ticketHtml = item.ticket
      ? `<span class="clickable-code${ticketActiveClass}" data-filter-field="ticket" data-filter-label="Ticket #" data-filter-val="${escapeHtml(item.ticket)}" title="Click to show all requests with Ticket # ${escapeHtml(item.ticket)}">${escapeHtml(item.ticket)}</span>`
      : '—';

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
  if (thDeleteSelect) {
    thDeleteSelect.className = isDeleteMode ? 'delete-col' : 'select-col';
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
  if (batchActionDock) {
    const showBatchDock = !isDeleteMode && count > 0 && Boolean(workbook);
    batchActionDock.style.display = showBatchDock ? 'block' : 'none';
  }
  syncSelectAllCheckboxes();
  updateStatusBarTelemetry();
}

function syncSelectAllCheckboxes() {
  if (!workbook || currentlyVisibleRows.length === 0) {
    if (chkSelectAllHeader) {
      chkSelectAllHeader.checked = false;
      chkSelectAllHeader.indeterminate = false;
      chkSelectAllHeader.disabled = true;
    }
    if (chkSelectAllBanner) {
      chkSelectAllBanner.checked = false;
      chkSelectAllBanner.indeterminate = false;
    }
    return;
  }

  if (chkSelectAllHeader) {
    chkSelectAllHeader.disabled = false;
  }

  let visibleSelectedCount = 0;
  currentlyVisibleRows.forEach(item => {
    if (selectedForDelete.has(getRowCompositeId(item.sheetName, item.rowNumber))) {
      visibleSelectedCount++;
    }
  });

  const allVisibleSelected = visibleSelectedCount === currentlyVisibleRows.length;
  const someVisibleSelected = visibleSelectedCount > 0 && !allVisibleSelected;

  if (chkSelectAllHeader) {
    chkSelectAllHeader.checked = allVisibleSelected;
    chkSelectAllHeader.indeterminate = someVisibleSelected;
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

chkSelectAllHeader.addEventListener('change', (e) => {
  handleToggleSelectAllVisible(e.target.checked);
});

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

  // 1 & 2. Map via immutable UUIDv4 keys (hidden Col K/L) & strictly preserve existing UUIDv4s
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
    if (
      dockEditingTarget &&
      dockEditingTarget.sheetName === target.sheetName &&
      dockEditingTarget.rowNumber === target.rowNumber &&
      dockStatus
    ) {
      dockStatus.value = targetStatus;
    }
  });

  // 4. Clear selection & re-render the DOM immediately to reflect the new statuses
  selectedForDelete.clear();
  updateDeleteSelectionUI();
  renderSheetTabsBar();
  renderList();

  // 5. Mutate ExcelJS cells (preserving existing Column K/L UUIDv4s) & perform ONE single batch write-back
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
      excelRow.getCell(colMap.statusCol).value = targetStatus;

      // Guarantee existing UUIDv4 in hidden Tracking ID column is never lost or regenerated
      const existingCellKey = cellText(excelRow, colMap.trackingCol).trim();
      if (!existingCellKey && target.immutableKey) {
        excelRow.getCell(colMap.trackingCol).value = target.immutableKey;
      }
      excelRow.commit();
    });

    const saveResult = await saveWorkbookToDisk();
    if (!saveResult) {
      // Primary working file was locked in Desktop Excel — reload from disk so state stays truthful
      await loadWorkbookFromDisk(sheetSelect.value);
      return;
    }

    showToast(
      `Batch updated ${selectedItems.length} request(s) to "${targetStatus}"${rowsToMutate.length > selectedItems.length ? ` (+${rowsToMutate.length - selectedItems.length} newer-tab row(s))` : ''}.`,
      'success'
    );
  } finally {
    if (btnApplyBatchStatus) btnApplyBatchStatus.disabled = false;
  }
}

if (btnApplyBatchStatus) {
  btnApplyBatchStatus.addEventListener('click', handleApplyBatchStatusUpdate);
}

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

