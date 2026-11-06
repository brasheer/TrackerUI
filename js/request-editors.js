// ============================================================================
// request-editors.js — Docked Bottom Request Editor, Modal Request Editor,
//                      Cross-Tab Status Sync, Ticket Details Timeline Modal &
//                      Supervisor View-Only Mirror Inspector
// ============================================================================

// --- 6. BOTTOM EDITOR DOCK + MODAL EDITOR + CROSS-TAB STATUS PROPAGATION ---
function ensureSelectOptionExists(selectEl, statusValue) {
  if (!selectEl || !statusValue) return;
  const exists = Array.from(selectEl.options).some(opt => opt.value === statusValue);
  if (!exists) {
    const opt = document.createElement('option');
    opt.value = statusValue;
    opt.textContent = statusValue;
    selectEl.appendChild(opt);
  }
}

function ensureStatusOptionExists(statusValue) {
  ensureSelectOptionExists(fStatus, statusValue);
  ensureSelectOptionExists(dockStatus, statusValue);
  ensureSelectOptionExists(batchStatusSelect, statusValue);
}

function formatCurrentTimestampPrefix() {
  const now = new Date();
  const datePart = now.toLocaleDateString('en-US', {
    month: '2-digit',
    day: '2-digit',
    year: 'numeric'
  });
  const timePart = now.toLocaleTimeString('en-US', {
    hour: '2-digit',
    minute: '2-digit'
  });
  return `[${datePart} ${timePart}]`;
}

function insertTimestampIntoTextarea(textareaEl) {
  if (!textareaEl) return;
  const stamp = `${formatCurrentTimestampPrefix()} `;
  const current = textareaEl.value.trim();
  if (!current) {
    textareaEl.value = stamp;
  } else {
    textareaEl.value = `${textareaEl.value.replace(/\s+$/, '')}\n${stamp}`;
  }
  textareaEl.focus();
  textareaEl.setSelectionRange(textareaEl.value.length, textareaEl.value.length);
}

if (btnDockInsertTimestamp) {
  btnDockInsertTimestamp.addEventListener('click', () => {
    insertTimestampIntoTextarea(dockNotes);
  });
}

if (btnModalInsertTimestamp) {
  btnModalInsertTimestamp.addEventListener('click', () => {
    insertTimestampIntoTextarea(fNotes);
  });
}

function setBottomDockCollapsed(collapsed, { persist = true } = {}) {
  isBottomDockCollapsed = Boolean(collapsed);
  if (persist) {
    localStorage.setItem(DOCK_COLLAPSED_STORAGE_KEY, String(isBottomDockCollapsed));
  }
  if (bottomEditorDock) {
    bottomEditorDock.classList.toggle('collapsed', isBottomDockCollapsed);
  }
  if (btnDockToggleCollapse) {
    btnDockToggleCollapse.textContent = isBottomDockCollapsed ? '▴ Expand Editor' : '▾ Collapse';
    btnDockToggleCollapse.setAttribute('aria-expanded', String(!isBottomDockCollapsed));
  }
}

if (btnDockToggleCollapse) {
  btnDockToggleCollapse.addEventListener('click', () => {
    setBottomDockCollapsed(!isBottomDockCollapsed);
  });
}

if (dockHeaderToggleZone) {
  dockHeaderToggleZone.addEventListener('click', (e) => {
    if (e.target.closest('.key-chip')) return;
    setBottomDockCollapsed(!isBottomDockCollapsed);
  });
}

function syncBottomDockStateAfterLoad() {
  if (!bottomEditorDock) return;
  if (!workbook || dateSheets.length === 0) {
    bottomEditorDock.style.display = 'none';
    dockEditingTarget = null;
    return;
  }
  bottomEditorDock.style.display = 'block';
  setBottomDockCollapsed(isBottomDockCollapsed, { persist: false });

  if (dockEditingTarget) {
    const updatedMatch = allRows.find(
      r =>
        r.sheetName === dockEditingTarget.sheetName &&
        (r.rowNumber === dockEditingTarget.rowNumber || r.key === dockEditingTarget.key)
    );
    if (updatedMatch) {
      loadRowIntoBottomDock(updatedMatch.sheetName, updatedMatch.rowNumber, false, { expandIfCollapsed: false });
      return;
    }
  }
  setBottomDockNewMode(false);
}

function setBottomDockNewMode(focusInput = false) {
  dockEditingTarget = null;
  if (!bottomEditorDock) return;

  if (focusInput && isBottomDockCollapsed) {
    setBottomDockCollapsed(false);
  }

  bottomEditorDock.classList.remove('editing-mode');
  const activeTab = sheetSelect.value || 'Daily Tab';

  if (dockModeBadge) {
    dockModeBadge.textContent = '＋ New Request';
    dockModeBadge.classList.add('new-mode');
    dockModeBadge.classList.remove('edit-mode');
  }
  if (dockTargetLabel) {
    dockTargetLabel.textContent = `Creating new support request on ${activeTab} — or click any row above to edit`;
  }
  if (dockKeyBadge) {
    dockKeyBadge.style.display = 'none';
    dockKeyBadge.textContent = '';
  }
  if (dockHintText) {
    dockHintText.textContent = 'Tip: Click any request row in the table above to load it here for instant editing.';
  }
  if (btnDockTicketDetails) {
    btnDockTicketDetails.style.display = 'none';
  }
  if (btnDockCancelEdit) {
    btnDockCancelEdit.style.display = 'none';
  }
  if (btnDockSave) {
    btnDockSave.textContent = `Save New Request to ${activeTab}`;
  }

  if (dockEntryForm) {
    dockEntryForm.reset();
  }
  const defaultSt = getDefaultStatus();
  ensureStatusOptionExists(defaultSt);
  if (dockStatus) {
    dockStatus.value = defaultSt;
  }

  if (listBody) {
    listBody.querySelectorAll('tr.active-editing-row').forEach(tr => tr.classList.remove('active-editing-row'));
  }

  if (focusInput && dockSubject) {
    setTimeout(() => dockSubject.focus(), 30);
  }
}

function loadRowIntoBottomDock(sheetName, rowNumber, focusInput = false, { expandIfCollapsed = true } = {}) {
  const item = allRows.find(r => r.sheetName === sheetName && r.rowNumber === rowNumber);
  if (!item || !bottomEditorDock) return;

  bottomEditorDock.style.display = 'block';
  if (expandIfCollapsed && isBottomDockCollapsed) {
    setBottomDockCollapsed(false);
  }
  bottomEditorDock.classList.add('editing-mode');
  dockEditingTarget = { sheetName, rowNumber, key: item.key };

  if (dockModeBadge) {
    dockModeBadge.textContent = 'Editing Row';
    dockModeBadge.classList.remove('new-mode');
    dockModeBadge.classList.add('edit-mode');
  }
  if (dockTargetLabel) {
    const ticketSuffix = item.ticket ? ` • Ticket #${item.ticket}` : '';
    dockTargetLabel.textContent = `${item.subject || 'Untitled Request'} (${sheetName} • Row ${rowNumber}${ticketSuffix})`;
  }
  if (dockKeyBadge) {
    dockKeyBadge.style.display = 'inline-block';
    dockKeyBadge.textContent = `🔍 ${item.key.slice(0, 8)}...`;
    dockKeyBadge.title = `Immutable Tracking ID (Excluded from editable inputs): ${item.key}`;
  }
  if (dockHintText) {
    dockHintText.textContent = `Editing Row ${rowNumber} on ${sheetName}. Unique Tracking ID is preserved automatically on save.`;
  }
  if (btnDockTicketDetails) {
    btnDockTicketDetails.style.display = 'inline-block';
  }
  if (btnDockCancelEdit) {
    btnDockCancelEdit.style.display = 'inline-block';
  }
  if (btnDockSave) {
    btnDockSave.textContent = `Save Changes (${sheetName} • Row ${rowNumber})`;
  }

  if (dockSubject) dockSubject.value = item.subject;
  if (dockTicket) dockTicket.value = item.ticket;
  if (dockBroker) dockBroker.value = item.broker;
  if (dockMaster) dockMaster.value = item.master;
  if (dockBidder) dockBidder.value = item.bidder;
  if (dockAction) dockAction.value = item.action;
  if (dockYard) dockYard.value = item.yard;
  if (dockLots) dockLots.value = item.lots;
  if (dockNotes) dockNotes.value = item.notes;

  ensureStatusOptionExists(item.status);
  if (dockStatus) dockStatus.value = item.status;

  if (listBody) {
    listBody.querySelectorAll('tr').forEach(tr => {
      const isMatch = tr.dataset.sheet === sheetName && Number(tr.dataset.row) === rowNumber;
      tr.classList.toggle('active-editing-row', isMatch);
    });
  }

  if (focusInput && dockSubject) {
    setTimeout(() => dockSubject.focus(), 30);
  }
}

if (btnDockNewClear) {
  btnDockNewClear.addEventListener('click', () => {
    setBottomDockNewMode(true);
  });
}

if (btnDockCancelEdit) {
  btnDockCancelEdit.addEventListener('click', () => {
    setBottomDockNewMode(false);
  });
}

if (btnDockOpenModal) {
  btnDockOpenModal.addEventListener('click', () => {
    if (dockEditingTarget) {
      openEditModal(dockEditingTarget.sheetName, dockEditingTarget.rowNumber);
    } else {
      openNewRequestModal();
    }
  });
}

if (btnDockTicketDetails) {
  btnDockTicketDetails.addEventListener('click', () => {
    if (dockEditingTarget) {
      openTicketDetailsModal(dockEditingTarget.sheetName, dockEditingTarget.rowNumber);
    }
  });
}

async function openNewRequestModal() {
  if (!sheetSelect.value || !DATE_TAB_REGEX.test(sheetSelect.value)) {
    let hasPlaceholder = false;
    if (workbook) {
      workbook.eachSheet(ws => {
        if (isTemplatePlaceholderTab(ws.name)) hasPlaceholder = true;
      });
    }
    if (hasPlaceholder) {
      await handleRepairCurrentWorkbookSchema();
    }
    if (!sheetSelect.value || !DATE_TAB_REGEX.test(sheetSelect.value)) {
      return showModalAlert({
        title: 'No Daily Tab Active',
        message: 'Create or select a daily tab (MM-DD-YYYY) first using Bootstrap or Rollover / Sync.'
      });
    }
  }
  editingTarget = null;
  document.getElementById('modalTitle').textContent = `New Request (${sheetSelect.value})`;
  const modalKeyEl = document.getElementById('modalKeyDisplay');
  modalKeyEl.textContent = 'New UUIDv4 on Save';
  modalKeyEl.title = 'A unique 36-character GUID will be generated on save';

  document.getElementById('entryForm').reset();
  const defaultSt = getDefaultStatus();
  ensureStatusOptionExists(defaultSt);
  fStatus.value = defaultSt;
  editModal.classList.add('open');
  setTimeout(() => fSubject.focus(), 20);
}

btnNewEntry.addEventListener('click', () => {
  if (bottomEditorDock && bottomEditorDock.style.display !== 'none') {
    setBottomDockNewMode(true);
  } else {
    openNewRequestModal();
  }
});

function openEditModal(sheetName, rowNumber) {
  const item = allRows.find(r => r.sheetName === sheetName && r.rowNumber === rowNumber);
  if (!item) return;

  editingTarget = { sheetName, rowNumber, key: item.key };
  document.getElementById('modalTitle').textContent = `Edit Request (${sheetName} • Row ${rowNumber})`;
  const modalKeyEl = document.getElementById('modalKeyDisplay');
  modalKeyEl.textContent = `🔍 ${item.key.slice(0, 8)}...`;
  modalKeyEl.title = item.key;

  fSubject.value = item.subject;
  fTicket.value = item.ticket;
  fBroker.value = item.broker;
  fMaster.value = item.master;
  fBidder.value = item.bidder;
  fAction.value = item.action;
  fYard.value = item.yard;
  fLots.value = item.lots;
  fNotes.value = item.notes;

  ensureStatusOptionExists(item.status);
  fStatus.value = item.status;

  editModal.classList.add('open');
}

document.getElementById('btnCancelModal').addEventListener('click', () => {
  editModal.classList.remove('open');
});

async function saveRequestToWorkbook(targetInfo, formValues) {
  const currentSheetName = targetInfo ? targetInfo.sheetName : sheetSelect.value;

  await loadWorkbookFromDisk(currentSheetName);
  const worksheet = workbook.getWorksheet(currentSheetName);
  if (!worksheet) {
    await showModalAlert({
      title: 'Worksheet Not Found',
      message: `Target worksheet "${currentSheetName}" was not found in the workbook.`
    });
    return false;
  }

  const subjectVal = formValues.subject.trim();
  const ticketVal = formValues.ticket.trim().toUpperCase();
  const brokerVal = formValues.broker.trim();
  const masterVal = formValues.master.trim().toUpperCase();
  const bidderVal = formValues.bidder.trim().toUpperCase();
  const actionVal = formValues.action.trim();
  const yardVal = formValues.yard.trim().toUpperCase();
  const lotsVal = formValues.lots.trim();
  const notesVal = formValues.notes.trim();
  const statusVal = formValues.status.trim();

  let targetRowNum;
  let trackingKey;

  if (targetInfo) {
    targetRowNum = targetInfo.rowNumber;
    trackingKey = targetInfo.key;
  } else {
    targetRowNum = getNextEmptyRow(worksheet);
    trackingKey = generateTrackingGuid();
  }

  const colMap = getWorksheetColumnMap(worksheet);
  worksheet.getColumn(colMap.trackingCol).hidden = true;
  if (colMap.emailSubjCol) {
    worksheet.getColumn(colMap.emailSubjCol).hidden = false;
  }

  const row = worksheet.getRow(targetRowNum);
  row.getCell(1).value = subjectVal;
  row.getCell(2).value = ticketVal;
  setBrokerCellWithFormula(row, targetRowNum, brokerVal, masterVal, colMap);
  row.getCell(4).value = masterVal;
  row.getCell(5).value = bidderVal;
  row.getCell(6).value = actionVal;
  row.getCell(7).value = yardVal;
  row.getCell(8).value = lotsVal;
  row.getCell(colMap.notesCol).value = notesVal;
  row.getCell(colMap.statusCol).value = statusVal;
  row.getCell(colMap.trackingCol).value = trackingKey;
  row.commit();

  if (targetInfo) {
    const editedTime = parseDateTab(currentSheetName);
    const newerInstances = allRows.filter(
      r =>
        r.key.toLowerCase() === trackingKey.toLowerCase() &&
        parseDateTab(r.sheetName) > editedTime &&
        r.status.trim().toLowerCase() !== statusVal.toLowerCase()
    );
    if (newerInstances.length > 0) {
      const newerSummary = newerInstances.map(r => `${r.sheetName} (${r.status})`).join(', ');
      const syncNewer = await showModalConfirm({
        title: '🔄 Sync Status Across Newer Tab(s)?',
        badge: statusVal,
        message:
          `This request also exists on newer tab(s): ${newerSummary}.\n\n` +
          `Update its status to "${statusVal}" on those newer tab(s) as well so Rollover uses your updated status?`,
        confirmText: '🔄 Update Newer Tab(s)',
        cancelText: 'Only This Tab'
      });
      if (syncNewer) {
        newerInstances.forEach(inst => {
          const newerWs = workbook.getWorksheet(inst.sheetName);
          if (newerWs) {
            const newerColMap = getWorksheetColumnMap(newerWs);
            const nr = newerWs.getRow(inst.rowNumber);
            nr.getCell(newerColMap.statusCol).value = statusVal;
            nr.commit();
          }
        });
      }
    }
  }

  const saveResult = await saveWorkbookToDisk();
  if (saveResult) {
    await loadWorkbookFromDisk(currentSheetName, {
      preserveStatus: saveResult === 'mirror-locked'
    });
    return true;
  }
  return false;
}

document.getElementById('entryForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const ok = await saveRequestToWorkbook(editingTarget, {
    subject: fSubject.value,
    ticket: fTicket.value,
    broker: fBroker.value,
    master: fMaster.value,
    bidder: fBidder.value,
    action: fAction.value,
    yard: fYard.value,
    lots: fLots.value,
    notes: fNotes.value,
    status: fStatus.value
  });
  if (ok) {
    editModal.classList.remove('open');
  }
});

if (dockEntryForm) {
  dockEntryForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const wasEditing = Boolean(dockEditingTarget);
    const ok = await saveRequestToWorkbook(dockEditingTarget, {
      subject: dockSubject.value,
      ticket: dockTicket.value,
      broker: dockBroker.value,
      master: dockMaster.value,
      bidder: dockBidder.value,
      action: dockAction.value,
      yard: dockYard.value,
      lots: dockLots.value,
      notes: dockNotes.value,
      status: dockStatus.value
    });
    if (ok) {
      if (!wasEditing) {
        setBottomDockNewMode(false);
        showToast('New support request added and saved.', 'success');
      } else {
        showToast('Request changes saved to workbook.', 'success');
      }
    }
  });
}

// --- 6B. TICKET DETAILS & AGGREGATED TIMESTAMPED NOTES MODAL ---
function findTicketHistoryEntries(seedItem) {
  if (!seedItem) return [];
  const normKey = seedItem.key ? seedItem.key.trim().toLowerCase() : '';

  const exactKeyMatches = normKey
    ? allRows.filter(r => r.key && r.key.trim().toLowerCase() === normKey)
    : [];

  const matchedSheets = new Set(exactKeyMatches.map(r => r.sheetName));
  const combined = [...exactKeyMatches];

  const candidateRows = [...allRows].sort((a, b) => {
    const diff = parseDateTab(a.sheetName) - parseDateTab(b.sheetName);
    return diff !== 0 ? diff : a.rowNumber - b.rowNumber;
  });

  candidateRows.forEach(r => {
    if (normKey && r.key && r.key.trim().toLowerCase() === normKey) return;
    if (matchedSheets.has(r.sheetName)) return;
    if (areRowsSameUniqueRequest(seedItem, r)) {
      matchedSheets.add(r.sheetName);
      combined.push(r);
    }
  });

  // Sort chronologically oldest to newest
  return combined.sort((a, b) => {
    const diff = parseDateTab(a.sheetName) - parseDateTab(b.sheetName);
    return diff !== 0 ? diff : a.rowNumber - b.rowNumber;
  });
}

function formatNoteHtmlWithTimestamps(rawNote) {
  if (!rawNote || !rawNote.trim()) {
    return `<span style="color:var(--text-muted); font-style:italic;">No action notes recorded on this tab.</span>`;
  }
  const escaped = escapeHtml(rawNote);
  return escaped.replace(
    /\[(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}[^\]]*)\]/g,
    '<span class="timeline-timestamp-tag">$1</span>'
  );
}

function openTicketDetailsModal(sheetName, rowNumber) {
  const seedItem = allRows.find(r => r.sheetName === sheetName && r.rowNumber === rowNumber);
  if (!seedItem || !ticketDetailsModal) return;

  activeTicketDetailsTarget = { sheetName, rowNumber, key: seedItem.key };
  renderTicketDetailsModal(seedItem);
  if (ticketQuickNoteInput) {
    ticketQuickNoteInput.value = '';
  }
  ticketDetailsModal.classList.add('open');
}

function renderTicketDetailsModal(seedItem) {
  const history = findTicketHistoryEntries(seedItem);
  if (history.length === 0) return;

  const oldest = history[0];
  const latest = history[history.length - 1];

  const oldestTime = parseDateTab(oldest.sheetName);
  const latestTime = parseDateTab(latest.sheetName);
  const spanDays = oldestTime && latestTime
    ? Math.max(1, Math.round((latestTime - oldestTime) / (1000 * 60 * 60 * 24)) + 1)
    : history.length;

  if (ticketDetailsTitle) {
    ticketDetailsTitle.textContent = latest.subject || 'Ticket Details';
  }
  if (ticketDetailsSubtitle) {
    ticketDetailsSubtitle.textContent = `Tracked across ${history.length} daily tab${history.length === 1 ? '' : 's'} (${spanDays} day span • First logged ${oldest.sheetName})`;
  }
  if (ticketDetailsKeyBadge) {
    ticketDetailsKeyBadge.textContent = latest.ticket ? `Ticket #${latest.ticket}` : `🔍 ${latest.key.slice(0, 8)}...`;
    ticketDetailsKeyBadge.title = `Tracking ID: ${latest.key}`;
  }
  if (ticketTimelineHeader) {
    ticketTimelineHeader.textContent = `Aggregated Notes & Daily Timeline (${history.length} Tab${history.length === 1 ? '' : 's'} • ${spanDays} Day Span)`;
  }
  if (ticketQuickNoteTargetBadge) {
    ticketQuickNoteTargetBadge.textContent = `Target: ${latest.sheetName} (Row ${latest.rowNumber})`;
  }

  if (ticketDetailsMetaGrid) {
    const isDone = isCompletedStatus(latest.status);
    ticketDetailsMetaGrid.innerHTML = `
      <div class="ticket-meta-card">
        <div class="ticket-meta-card-label">Ticket Number</div>
        <div class="ticket-meta-card-val">${escapeHtml(latest.ticket || '—')}</div>
      </div>
      <div class="ticket-meta-card">
        <div class="ticket-meta-card-label">Current Status (${escapeHtml(latest.sheetName)})</div>
        <div class="ticket-meta-card-val">
          <span class="badge ${isDone ? 'badge-complete' : 'badge-pending'}">${escapeHtml(latest.status)}</span>
        </div>
      </div>
      <div class="ticket-meta-card">
        <div class="ticket-meta-card-label">Original Creation Tab</div>
        <div class="ticket-meta-card-val">${escapeHtml(oldest.sheetName)} (${history.length} tab${history.length === 1 ? '' : 's'})</div>
      </div>
      <div class="ticket-meta-card">
        <div class="ticket-meta-card-label">Broker &amp; Accounts</div>
        <div class="ticket-meta-card-val">${escapeHtml(latest.broker || '—')} (M: ${escapeHtml(latest.master || '—')} / B: ${escapeHtml(latest.bidder || '—')})</div>
      </div>
      <div class="ticket-meta-card">
        <div class="ticket-meta-card-label">Requested Action</div>
        <div class="ticket-meta-card-val">${escapeHtml(latest.action || '—')}</div>
      </div>
      <div class="ticket-meta-card">
        <div class="ticket-meta-card-label">Yard # &amp; Lot Numbers</div>
        <div class="ticket-meta-card-val">Y: ${escapeHtml(latest.yard || '—')} • Lots: ${escapeHtml(latest.lots || '—')}</div>
      </div>
    `;
  }

  if (ticketTimelineList) {
    ticketTimelineList.innerHTML = '';
    history.forEach((entry, idx) => {
      const isOrigin = idx === 0;
      const isLatest = idx === history.length - 1;
      const isDone = isCompletedStatus(entry.status);
      const stepLabel = isOrigin
        ? (isLatest ? 'Created & Current Tab' : 'Originally Created')
        : (isLatest ? 'Latest Daily Tab' : `Day ${idx + 1} Rollover`);

      const div = document.createElement('div');
      div.className = 'timeline-entry' + (isDone ? ' completed-entry' : '');
      div.innerHTML = `
        <div class="timeline-entry-header">
          <div>
            <strong>${escapeHtml(entry.sheetName)}</strong>
            <span style="font-size:0.6875rem; color:var(--text-muted); font-weight:500; margin-left:6px;">(Row ${entry.rowNumber} • ${stepLabel})</span>
          </div>
          <div style="display:flex; align-items:center; gap:8px;">
            <span class="badge ${isDone ? 'badge-complete' : 'badge-pending'}">${escapeHtml(entry.status)}</span>
            <button type="button" class="btn-edit btn-sm btn-timeline-load-dock" data-sheet="${escapeHtml(entry.sheetName)}" data-row="${entry.rowNumber}">Edit Row</button>
          </div>
        </div>
        <div class="timeline-entry-body">${formatNoteHtmlWithTimestamps(entry.notes)}</div>
      `;
      ticketTimelineList.appendChild(div);
    });
  }
}

if (ticketTimelineList) {
  ticketTimelineList.addEventListener('click', (e) => {
    const loadBtn = e.target.closest('.btn-timeline-load-dock');
    if (loadBtn) {
      const sName = loadBtn.dataset.sheet;
      const rNum = Number(loadBtn.dataset.row);
      ticketDetailsModal.classList.remove('open');
      loadRowIntoBottomDock(sName, rNum, true);
    }
  });
}

if (btnCloseTicketDetails) {
  btnCloseTicketDetails.addEventListener('click', () => {
    ticketDetailsModal.classList.remove('open');
  });
}

if (btnTicketEditInDock) {
  btnTicketEditInDock.addEventListener('click', () => {
    if (!activeTicketDetailsTarget) return;
    ticketDetailsModal.classList.remove('open');
    loadRowIntoBottomDock(activeTicketDetailsTarget.sheetName, activeTicketDetailsTarget.rowNumber, true);
  });
}

if (ticketQuickNoteForm) {
  ticketQuickNoteForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!activeTicketDetailsTarget || !ticketQuickNoteInput) return;
    const newNoteText = ticketQuickNoteInput.value.trim();
    if (!newNoteText) return;

    const seedItem = allRows.find(
      r =>
        r.sheetName === activeTicketDetailsTarget.sheetName &&
        r.rowNumber === activeTicketDetailsTarget.rowNumber
    ) || allRows.find(r => r.key === activeTicketDetailsTarget.key);

    if (!seedItem) return;
    const history = findTicketHistoryEntries(seedItem);
    const latestEntry = history[history.length - 1] || seedItem;

    const stampedLine = `${formatCurrentTimestampPrefix()} ${newNoteText}`;
    const combinedNotes = latestEntry.notes
      ? `${latestEntry.notes.replace(/\s+$/, '')}\n${stampedLine}`
      : stampedLine;

    const ok = await saveRequestToWorkbook(
      { sheetName: latestEntry.sheetName, rowNumber: latestEntry.rowNumber, key: latestEntry.key },
      {
        subject: latestEntry.subject,
        ticket: latestEntry.ticket,
        broker: latestEntry.broker,
        master: latestEntry.master,
        bidder: latestEntry.bidder,
        action: latestEntry.action,
        yard: latestEntry.yard,
        lots: latestEntry.lots,
        notes: combinedNotes,
        status: latestEntry.status
      }
    );

    if (ok) {
      ticketQuickNoteInput.value = '';
      const refreshedSeed = allRows.find(
        r => r.sheetName === latestEntry.sheetName && r.rowNumber === latestEntry.rowNumber
      );
      if (refreshedSeed) {
        activeTicketDetailsTarget = {
          sheetName: refreshedSeed.sheetName,
          rowNumber: refreshedSeed.rowNumber,
          key: refreshedSeed.key
        };
        renderTicketDetailsModal(refreshedSeed);
      }
      showToast(`Timestamped note appended to ${latestEntry.sheetName}.`, 'success');
    }
  });
}

// --- 6C. SUPERVISOR VIEW-ONLY TRACKER INSPECTOR (NON-LOCKING READ-ONLY SNAPSHOT) ---
async function openViewOnlyTrackerModal() {
  closeToolsMenu();
  const targetMirrorDir = supervisorDirHandle || parentDirHandle;
  if (!targetMirrorDir || !workingFileHandle) {
    return showModalAlert({
      title: 'No Active Workbook',
      message: 'Please connect a Tracker folder and load a workbook first.'
    });
  }

  const mirrorFileName = workingFileHandle.name.replace(/\.xlsx$/i, '') + ' (VIEW ONLY).xlsx';
  let mirrorHandle = null;

  try {
    mirrorHandle = await targetMirrorDir.getFileHandle(mirrorFileName, { create: false });
  } catch (notFoundErr) {
    const createNow = await showModalConfirm({
      title: '👁️ View-Only Mirror File Not Created Yet',
      badge: mirrorFileName,
      message:
        `The supervisor file "${mirrorFileName}" does not exist yet in "${targetMirrorDir.name}".\n\n` +
        `Would you like to generate and sync "${mirrorFileName}" right now from your active workbook?`,
      confirmText: '✨ Generate & Open View-Only Copy'
    });
    if (!createNow) return;

    const buf = await serializeWorkbookBuffer();
    const ok = await attemptMirrorWrite(buf, false);
    if (!ok) return;
    mirrorHandle = await targetMirrorDir.getFileHandle(mirrorFileName, { create: false });
  }

  try {
    const file = await mirrorHandle.getFile();
    const arrayBuffer = await file.arrayBuffer();
    const sanitizedBuffer = sanitizeXlsxBufferForExcelJS(arrayBuffer);

    viewOnlyState.buffer = arrayBuffer;
    viewOnlyState.fileName = mirrorFileName;
    viewOnlyState.fullPath = `${targetMirrorDir.name}\\${mirrorFileName}`;
    viewOnlyState.lastModified = file.lastModified || Date.now();
    viewOnlyState.rows = [];
    viewOnlyState.sheets = [];

    const mirrorWb = new ExcelJS.Workbook();
    await mirrorWb.xlsx.load(sanitizedBuffer);

    mirrorWb.eachSheet((ws) => {
      const sName = ws.name.trim();
      if (DATE_TAB_REGEX.test(sName)) {
        viewOnlyState.sheets.push(sName);
      }
    });

    viewOnlyState.sheets.sort((a, b) => parseDateTab(b) - parseDateTab(a));

    viewOnlyState.sheets.forEach(sName => {
      const ws = mirrorWb.getWorksheet(sName);
      if (!ws) return;
      const colMap = getWorksheetColumnMap(ws);
      ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber === 1) return;
        const subject = cellText(row, 1);
        const ticket = cellText(row, 2);
        const master = cellText(row, 4);
        const broker = resolveBrokerFromMaster(master, cellText(row, 3));
        const bidder = cellText(row, 5);
        const action = cellText(row, 6);
        const yard = cellText(row, 7);
        const lots = cellText(row, 8);
        const status = cellText(row, colMap.statusCol);
        const notes = cellText(row, colMap.notesCol);
        const hasData = [subject, ticket, broker, master, bidder, action, yard, lots, notes, status]
          .some(v => v !== '');
        if (!hasData) return;

        viewOnlyState.rows.push({
          sheetName: sName,
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
          status: status || 'Pending'
        });
      });
    });

    if (viewOnlyTitle) {
      viewOnlyTitle.textContent = `View-Only Inspector: ${mirrorFileName}`;
    }
    if (viewOnlyPathSubtitle) {
      const modTimeStr = file.lastModified
        ? ` • Last synced: ${new Date(file.lastModified).toLocaleString()}`
        : '';
      viewOnlyPathSubtitle.textContent = `Path: ${viewOnlyState.fullPath}${modTimeStr}`;
    }
    if (viewOnlySyncBadge) {
      viewOnlySyncBadge.textContent = pendingMirrorBuffer
        ? 'Pending Mirror Retry'
        : 'Read-Only Snapshot';
      viewOnlySyncBadge.className = pendingMirrorBuffer
        ? 'badge badge-pending'
        : 'badge badge-complete';
    }

    if (viewOnlySheetSelect) {
      viewOnlySheetSelect.innerHTML = '';
      const allOpt = document.createElement('option');
      allOpt.value = '__ALL__';
      allOpt.textContent = `All Daily Tabs (${viewOnlyState.rows.length} total)`;
      viewOnlySheetSelect.appendChild(allOpt);

      viewOnlyState.sheets.forEach(tab => {
        const count = viewOnlyState.rows.filter(r => r.sheetName === tab).length;
        const opt = document.createElement('option');
        opt.value = tab;
        opt.textContent = `${tab} (${count})`;
        viewOnlySheetSelect.appendChild(opt);
      });

      if (viewOnlyState.sheets.length > 0) {
        viewOnlySheetSelect.value = viewOnlyState.sheets[0];
      }
    }

    if (viewOnlySearchInput) {
      viewOnlySearchInput.value = '';
    }

    renderViewOnlyTable();
    if (viewOnlyModal) {
      viewOnlyModal.classList.add('open');
    }
  } catch (err) {
    console.error('Failed to inspect View-Only file:', err);
    await showModalAlert({
      title: 'View-Only Read Error',
      message: `Could not read "${mirrorFileName}": ${err.message}`,
      variant: 'danger'
    });
  }
}

function renderViewOnlyTable() {
  if (!viewOnlyBody) return;
  const selectedTab = viewOnlySheetSelect ? viewOnlySheetSelect.value : '__ALL__';
  const query = viewOnlySearchInput ? viewOnlySearchInput.value.trim().toLowerCase() : '';

  const filtered = viewOnlyState.rows.filter(item => {
    if (query.length > 0) {
      const hay = [
        item.sheetName, item.subject, item.ticket, item.broker,
        item.master, item.bidder, item.action, item.yard, item.lots, item.notes, item.status
      ].join(' ').toLowerCase();
      return hay.includes(query);
    }
    if (selectedTab !== '__ALL__' && item.sheetName !== selectedTab) {
      return false;
    }
    return true;
  });

  if (viewOnlyStatsText) {
    viewOnlyStatsText.textContent = `Showing ${filtered.length} of ${viewOnlyState.rows.length} request(s)`;
  }

  viewOnlyBody.innerHTML = '';
  if (filtered.length === 0) {
    viewOnlyBody.innerHTML = `<tr><td colspan="9" class="empty-table-cell">No matching requests in View-Only snapshot.</td></tr>`;
    return;
  }

  filtered.forEach(item => {
    const tr = document.createElement('tr');
    const isDone = isCompletedStatus(item.status);
    tr.innerHTML = `
      <td class="cell-mono-sm" style="white-space:nowrap;">${escapeHtml(item.sheetName)}</td>
      <td>${escapeHtml(item.subject || '—')}</td>
      <td class="cell-mono-sm">${escapeHtml(item.ticket || '—')}</td>
      <td>${escapeHtml(item.broker || '—')}</td>
      <td class="cell-mono-sm">M: ${escapeHtml(item.master || '—')}<br>B: ${escapeHtml(item.bidder || '—')}</td>
      <td>${escapeHtml(item.action || '—')}</td>
      <td class="cell-sm">Y: ${escapeHtml(item.yard || '—')}<br>L: ${escapeHtml(item.lots || '—')}</td>
      <td class="cell-sm" style="max-width:260px; white-space:pre-wrap;">${escapeHtml(item.notes || '—')}</td>
      <td><span class="badge ${isDone ? 'badge-complete' : 'badge-pending'}">${escapeHtml(item.status)}</span></td>
    `;
    viewOnlyBody.appendChild(tr);
  });
}

if (btnOpenViewOnly) {
  btnOpenViewOnly.addEventListener('click', openViewOnlyTrackerModal);
}
if (menuOpenViewOnly) {
  menuOpenViewOnly.addEventListener('click', openViewOnlyTrackerModal);
}
if (btnRefreshViewOnly) {
  btnRefreshViewOnly.addEventListener('click', openViewOnlyTrackerModal);
}
if (viewOnlySheetSelect) {
  viewOnlySheetSelect.addEventListener('change', renderViewOnlyTable);
}
if (viewOnlySearchInput) {
  viewOnlySearchInput.addEventListener('input', renderViewOnlyTable);
}
if (btnCloseViewOnly) {
  btnCloseViewOnly.addEventListener('click', () => {
    viewOnlyModal.classList.remove('open');
  });
}
if (btnDownloadViewOnly) {
  btnDownloadViewOnly.addEventListener('click', () => {
    if (!viewOnlyState.buffer) return;
    const blob = new Blob([viewOnlyState.buffer], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = viewOnlyState.fileName || 'Tracker (VIEW ONLY).xlsx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
    showToast(`Downloaded copy of "${viewOnlyState.fileName}" (safe to open in Excel).`, 'info');
  });
}

