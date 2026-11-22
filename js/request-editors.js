// ============================================================================
// request-editors.js — Docked Bottom Request Editor, Modal Request Editor,
//                      Grab Ticket ([XXXX-XX-XXXX]), Pull from Prior Tab,
//                      Forward Batch Sync, Dynamic Custom Schema Forms,
//                      Ticket Details Timeline & View-Only Mirror Inspector
// ============================================================================

// --- 6. BOTTOM EDITOR DOCK + MODAL EDITOR + CROSS-TAB SYNC & DYNAMIC FORMS ---
function ensureSelectOptionExists(selectEl, statusValue) {
  if (!selectEl || !statusValue) return '';
  const cleanVal = String(statusValue).replace(/\s+/g, ' ').trim();
  if (!cleanVal || UUID_REGEX.test(cleanVal)) return '';
  const cleanLower = cleanVal.toLowerCase();
  const existingOpt = Array.from(selectEl.options).find(
    opt => String(opt.value || '').replace(/\s+/g, ' ').trim().toLowerCase() === cleanLower
  );
  if (existingOpt) {
    return existingOpt.value;
  }
  const opt = document.createElement('option');
  opt.value = cleanVal;
  opt.textContent = cleanVal;
  selectEl.appendChild(opt);
  return cleanVal;
}

function ensureStatusOptionExists(statusValue) {
  let canonical = ensureSelectOptionExists(fStatus, statusValue);
  canonical = ensureSelectOptionExists(dockStatus, statusValue) || canonical;
  canonical = ensureSelectOptionExists(batchStatusSelect, statusValue) || canonical;
  if (!activeSchemaConfig.isStandardSchema && activeSchemaConfig.statusCol) {
    if (dockDynamicFieldsGrid) {
      const dockSel = dockDynamicFieldsGrid.querySelector(`select[data-custom-col="${activeSchemaConfig.statusCol}"]`);
      canonical = ensureSelectOptionExists(dockSel, statusValue) || canonical;
    }
    if (modalDynamicFieldsGrid) {
      const modSel = modalDynamicFieldsGrid.querySelector(`select[data-custom-col="${activeSchemaConfig.statusCol}"]`);
      canonical = ensureSelectOptionExists(modSel, statusValue) || canonical;
    }
  }
  return canonical || String(statusValue || '').trim();
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

// --- DYNAMIC CUSTOM SCHEMA FORM RENDERING & LOCKABLE 12-COL GRID LAYOUT ---
let isFormLayoutEditMode = false;
let activeFieldResizeDrag = null;

function getActiveFormLayoutSchemaKey() {
  if (
    activeSchemaConfig &&
    !activeSchemaConfig.isStandardSchema &&
    Array.isArray(activeSchemaConfig.columns) &&
    activeSchemaConfig.columns.length > 0
  ) {
    const sig = activeSchemaConfig.columns
      .map(c => `${c.colIndex}:${String(c.header || '').trim().toLowerCase()}`)
      .join('|');
    return `custom::${sig}`;
  }
  return '__standard__';
}

function getVisibleCustomEditorColumns() {
  if (!activeSchemaConfig || !Array.isArray(activeSchemaConfig.columns)) return [];
  const visible = activeSchemaConfig.columns.filter(c => c.showInEditor !== false);
  return visible.length > 0 ? visible : activeSchemaConfig.columns;
}

function buildDefaultFormLayoutForActiveSchema() {
  if (
    !activeSchemaConfig ||
    activeSchemaConfig.isStandardSchema ||
    !Array.isArray(activeSchemaConfig.columns) ||
    activeSchemaConfig.columns.length === 0
  ) {
    const nIdx = (activeSchemaConfig && activeSchemaConfig.notesCol) || 9;
    const stIdx = (activeSchemaConfig && activeSchemaConfig.statusCol) || 10;
    const trkIdx = (activeSchemaConfig && activeSchemaConfig.trackingCol) || 11;
    const coreStdSet = new Set([1, 2, 3, 4, 5, 6, 7, 8, nIdx, stIdx, trkIdx]);
    const extraStdCols = (activeSchemaConfig && Array.isArray(activeSchemaConfig.columns))
      ? activeSchemaConfig.columns.filter(c => !coreStdSet.has(c.colIndex) && c.showInEditor !== false)
      : [];

    const order = [
      'subject', 'ticket', 'broker', 'master', 'bidder', 'action', 'yard', 'lots', 'status',
      ...extraStdCols.filter(c => c.dataType !== 'notes').map(c => `col_${c.colIndex}`),
      'notes',
      ...extraStdCols.filter(c => c.dataType === 'notes').map(c => `col_${c.colIndex}`)
    ];
    const spans = {
      subject: 3,
      ticket: 3,
      broker: 2,
      master: 2,
      bidder: 2,
      action: 3,
      yard: 2,
      lots: 4,
      status: 3,
      notes: 12
    };
    const rows = { notes: 2 };
    extraStdCols.forEach(c => {
      const fid = `col_${c.colIndex}`;
      spans[fid] = c.dataType === 'notes' ? 12 : 4;
      if (c.dataType === 'notes') rows[fid] = 2;
    });

    return { order, spans, rows };
  }

  const cols = getVisibleCustomEditorColumns();
  const isNotesLikeCol = (c) => c.colIndex === activeSchemaConfig.notesCol || c.dataType === 'notes';
  const nonNotesCols = cols.filter(c => !isNotesLikeCol(c));
  const notesCols = cols.filter(c => isNotesLikeCol(c));

  // Smart default order: place multi-line Notes at the bottom so mid-sheet Notes columns (e.g. Col E of A..H)
  // never split single-line inputs into awkward half-empty rows, while still saving to their exact Excel column.
  const orderedCols = [...nonNotesCols, ...notesCols];
  const order = orderedCols.map(c => `col_${c.colIndex}`);
  const spans = {};
  const rows = {};

  // Assign base spans on a 12-column grid and auto-balance each row to eliminate dead trailing whitespace
  const baseItems = nonNotesCols.map(col => {
    const idx = col.colIndex;
    const hdr = String(col.excelHeader || col.header || '');
    const isSubjectLike =
      idx === activeSchemaConfig.primaryCol ||
      (/\b(subject|summary|description|title|detail|reason|message)\b/i.test(hdr) &&
        !/\btitle\s*(status|state|type|#|no|number|received|sent|date)\b/i.test(hdr));
    const isTicket = idx === activeSchemaConfig.ticketCol;
    const isMedText = /\b(broker|dealer|customer|company|vendor|account\s*name|action|type|category|status|disposition)\b/i.test(hdr);

    let base = 2;
    if (isSubjectLike) {
      base = nonNotesCols.length >= 8 ? 3 : 4;
    } else if (isTicket || isMedText) {
      base = 3;
    } else if (nonNotesCols.length <= 6) {
      base = 3;
    }
    return { id: `col_${idx}`, span: base, priority: isSubjectLike ? 2 : (isTicket || isMedText ? 1 : 0) };
  });

  const packedRows = [];
  let currentRow = [];
  let currentSum = 0;
  baseItems.forEach(item => {
    if (currentSum + item.span > 12 && currentRow.length > 0) {
      packedRows.push(currentRow);
      currentRow = [];
      currentSum = 0;
    }
    currentRow.push(item);
    currentSum += item.span;
  });
  if (currentRow.length > 0) packedRows.push(currentRow);

  packedRows.forEach(rowItems => {
    let sum = rowItems.reduce((acc, it) => acc + it.span, 0);
    const targetSum = rowItems.length === 1 ? Math.max(rowItems[0].span, 6) : 12;
    if (sum < targetSum) {
      const sortedByPriority = [...rowItems].sort((a, b) => b.priority - a.priority);
      let cursor = 0;
      while (sum < targetSum && cursor < 24) {
        const target = sortedByPriority[cursor % sortedByPriority.length];
        if (target.span < 12) {
          target.span += 1;
          sum += 1;
        }
        cursor++;
      }
    }
    rowItems.forEach(it => {
      spans[it.id] = it.span;
    });
  });

  notesCols.forEach(nc => {
    const nId = `col_${nc.colIndex}`;
    spans[nId] = 12;
    rows[nId] = 2;
  });

  return { order, spans, rows };
}

function getResolvedFormLayoutForActiveSchema() {
  const key = getActiveFormLayoutSchemaKey();
  const defaults = buildDefaultFormLayoutForActiveSchema();
  const saved = (formLayoutsStore && formLayoutsStore[key] && typeof formLayoutsStore[key] === 'object')
    ? formLayoutsStore[key]
    : null;
  if (!saved) return defaults;

  const validIds = new Set(defaults.order);
  const mergedOrder = [];
  if (Array.isArray(saved.order)) {
    saved.order.forEach(id => {
      if (validIds.has(id) && !mergedOrder.includes(id)) {
        mergedOrder.push(id);
      }
    });
  }
  defaults.order.forEach(id => {
    if (!mergedOrder.includes(id)) mergedOrder.push(id);
  });

  const mergedSpans = { ...defaults.spans };
  if (saved.spans && typeof saved.spans === 'object') {
    Object.entries(saved.spans).forEach(([id, val]) => {
      const n = Number(val);
      if (validIds.has(id) && Number.isFinite(n)) {
        mergedSpans[id] = Math.max(2, Math.min(12, Math.round(n)));
      }
    });
  }

  const mergedRows = { ...defaults.rows };
  if (saved.rows && typeof saved.rows === 'object') {
    Object.entries(saved.rows).forEach(([id, val]) => {
      const n = Number(val);
      if (validIds.has(id) && Number.isFinite(n)) {
        mergedRows[id] = Math.max(1, Math.min(6, Math.round(n)));
      }
    });
  }

  return { order: mergedOrder, spans: mergedSpans, rows: mergedRows };
}

function persistActiveFormLayoutMutation(mutatorFn) {
  const key = getActiveFormLayoutSchemaKey();
  const current = getResolvedFormLayoutForActiveSchema();
  const next = {
    order: [...current.order],
    spans: { ...current.spans },
    rows: { ...current.rows }
  };
  mutatorFn(next);
  if (!formLayoutsStore || typeof formLayoutsStore !== 'object') {
    formLayoutsStore = {};
  }
  formLayoutsStore[key] = next;
  if (typeof saveFormLayoutsStore === 'function') {
    saveFormLayoutsStore();
  }
  applyActiveFormLayoutToGrids();
}

function ensureDockFieldLayoutChrome(groupEl, fieldId, indexInOrder, totalCount, span, rowsVal) {
  let toolbar = groupEl.querySelector(':scope > .field-layout-toolbar');
  const hasTextarea = Boolean(groupEl.querySelector('textarea'));
  if (!toolbar) {
    toolbar = document.createElement('div');
    toolbar.className = 'field-layout-toolbar';
    groupEl.insertBefore(toolbar, groupEl.firstChild);
  }
  const rowsControlsHtml = hasTextarea
    ? `<span class="field-layout-btn-group" title="Adjust textarea height (rows)">
         <button type="button" class="field-layout-step-btn" data-layout-act="rows-dec" data-field-id="${escapeHtml(fieldId)}" ${rowsVal <= 1 ? 'disabled' : ''} title="Fewer rows">↕−</button>
         <span class="field-layout-span-pill">${rowsVal}r</span>
         <button type="button" class="field-layout-step-btn" data-layout-act="rows-inc" data-field-id="${escapeHtml(fieldId)}" ${rowsVal >= 6 ? 'disabled' : ''} title="More rows">↕+</button>
       </span>`
    : '';

  toolbar.innerHTML = `
    <span class="field-layout-btn-group" title="Move field earlier or later in the form (Excel column mapping stays unchanged)">
      <button type="button" class="field-layout-step-btn" data-layout-act="left" data-field-id="${escapeHtml(fieldId)}" ${indexInOrder === 0 ? 'disabled' : ''} title="Move field left / up">◀</button>
      <button type="button" class="field-layout-step-btn" data-layout-act="right" data-field-id="${escapeHtml(fieldId)}" ${indexInOrder >= totalCount - 1 ? 'disabled' : ''} title="Move field right / down">▶</button>
    </span>
    <span class="field-layout-btn-group" title="Adjust field width on 12-column grid (or drag right border)">
      <button type="button" class="field-layout-step-btn" data-layout-act="narrow" data-field-id="${escapeHtml(fieldId)}" ${span <= 2 ? 'disabled' : ''} title="Narrower (−1 col)">−</button>
      <span class="field-layout-span-pill">${span}/12</span>
      <button type="button" class="field-layout-step-btn" data-layout-act="widen" data-field-id="${escapeHtml(fieldId)}" ${span >= 12 ? 'disabled' : ''} title="Wider (+1 col)">+</button>
    </span>
    ${rowsControlsHtml}
  `;

  let dragHandle = groupEl.querySelector(':scope > .field-width-drag-handle');
  if (!dragHandle) {
    dragHandle = document.createElement('div');
    dragHandle.className = 'field-width-drag-handle';
    dragHandle.title = 'Drag horizontally to resize field width';
    groupEl.appendChild(dragHandle);
  }
  dragHandle.dataset.fieldId = fieldId;
}

function syncFormLayoutButtonsUI() {
  if (bottomEditorDock) {
    bottomEditorDock.classList.toggle('layout-edit-active', isFormLayoutEditMode);
  }
  if (btnDockToggleLayout) {
    btnDockToggleLayout.classList.toggle('unlocked', isFormLayoutEditMode);
    btnDockToggleLayout.setAttribute('aria-pressed', String(isFormLayoutEditMode));
    btnDockToggleLayout.textContent = isFormLayoutEditMode ? '🔓 Lock Layout' : '🔒 Layout';
    btnDockToggleLayout.title = isFormLayoutEditMode
      ? 'Lock form layout to hide adjustment controls'
      : 'Unlock form grid to adjust field widths, order, or textarea rows (locked by default)';
  }
  if (btnDockResetLayout) {
    const key = getActiveFormLayoutSchemaKey();
    const hasCustomSaved = Boolean(formLayoutsStore && formLayoutsStore[key]);
    btnDockResetLayout.style.display = isFormLayoutEditMode ? 'inline-flex' : 'none';
    btnDockResetLayout.disabled = !hasCustomSaved;
  }
}

function applyActiveFormLayoutToGrids() {
  const isCustom = Boolean(
    activeSchemaConfig &&
    !activeSchemaConfig.isStandardSchema &&
    Array.isArray(activeSchemaConfig.columns) &&
    activeSchemaConfig.columns.length > 0
  );
  const layout = getResolvedFormLayoutForActiveSchema();
  const activeDockGrid = isCustom ? dockDynamicFieldsGrid : dockStandardFieldsGrid;

  if (activeDockGrid) {
    const groupMap = new Map();
    Array.from(activeDockGrid.children).forEach(el => {
      if (el && el.dataset && el.dataset.fieldId) {
        groupMap.set(el.dataset.fieldId, el);
      }
    });

    const presentOrder = layout.order.filter(id => groupMap.has(id));
    presentOrder.forEach((fieldId, idx) => {
      const groupEl = groupMap.get(fieldId);
      activeDockGrid.appendChild(groupEl);
      const span = Math.max(2, Math.min(12, Number(layout.spans[fieldId]) || 3));
      groupEl.style.gridColumn = span >= 12 ? '1 / -1' : `span ${span}`;
      const rowsVal = Math.max(1, Math.min(6, Number(layout.rows[fieldId]) || 2));
      const textarea = groupEl.querySelector('textarea');
      if (textarea) {
        textarea.rows = rowsVal;
      }
      ensureDockFieldLayoutChrome(groupEl, fieldId, idx, presentOrder.length, span, rowsVal);
    });
  }

  // Mirror field order & proportional spans into Popup Modal grids
  const mirrorGridEl = isCustom ? modalDynamicFieldsGrid : modalStandardFieldsGrid;
  if (mirrorGridEl) {
    const modalMap = new Map();
    Array.from(mirrorGridEl.children).forEach(el => {
      if (el && el.dataset && el.dataset.fieldId) {
        modalMap.set(el.dataset.fieldId, el);
      }
    });
    if (isCustom) {
      layout.order.forEach(fieldId => {
        const mEl = modalMap.get(fieldId);
        if (!mEl) return;
        mirrorGridEl.appendChild(mEl);
        const dockSpan = Math.max(2, Math.min(12, Number(layout.spans[fieldId]) || 3));
        const modalSpan = dockSpan >= 8 ? 12 : (dockSpan >= 4 ? 6 : 4);
        mEl.style.gridColumn = modalSpan >= 12 ? '1 / -1' : `span ${modalSpan}`;
      });
    } else {
      layout.order.forEach(fieldId => {
        const mEl = modalMap.get(fieldId);
        if (!mEl || !mEl.classList.contains('std-extra-col-group')) return;
        const notesEl = modalMap.get('notes');
        if (notesEl && mEl.dataset.isNotesCol !== 'true') {
          mirrorGridEl.insertBefore(mEl, notesEl);
        } else {
          mirrorGridEl.appendChild(mEl);
        }
      });
    }
  }

  syncFormLayoutButtonsUI();
}

function handleDockLayoutControlAction(e) {
  const stepBtn = e.target.closest('[data-layout-act]');
  if (!stepBtn || !isFormLayoutEditMode) return;
  e.preventDefault();
  e.stopPropagation();

  const act = stepBtn.dataset.layoutAct;
  const fieldId = stepBtn.dataset.fieldId;
  if (!fieldId) return;

  persistActiveFormLayoutMutation(next => {
    const idx = next.order.indexOf(fieldId);
    if (act === 'left' && idx > 0) {
      [next.order[idx - 1], next.order[idx]] = [next.order[idx], next.order[idx - 1]];
    } else if (act === 'right' && idx !== -1 && idx < next.order.length - 1) {
      [next.order[idx + 1], next.order[idx]] = [next.order[idx], next.order[idx + 1]];
    } else if (act === 'narrow') {
      next.spans[fieldId] = Math.max(2, (Number(next.spans[fieldId]) || 3) - 1);
    } else if (act === 'widen') {
      next.spans[fieldId] = Math.min(12, (Number(next.spans[fieldId]) || 3) + 1);
    } else if (act === 'rows-dec') {
      next.rows[fieldId] = Math.max(1, (Number(next.rows[fieldId]) || 2) - 1);
    } else if (act === 'rows-inc') {
      next.rows[fieldId] = Math.min(6, (Number(next.rows[fieldId]) || 2) + 1);
    }
  });
}

function startFieldWidthDragResize(e) {
  if (!isFormLayoutEditMode) return;
  const handle = e.target.closest('.field-width-drag-handle');
  if (!handle) return;
  const groupEl = handle.closest('.form-group[data-field-id]');
  const gridEl = groupEl ? groupEl.parentElement : null;
  if (!groupEl || !gridEl) return;

  e.preventDefault();
  const fieldId = groupEl.dataset.fieldId;
  const layout = getResolvedFormLayoutForActiveSchema();
  const startSpan = Math.max(2, Math.min(12, Number(layout.spans[fieldId]) || 3));
  const gridWidth = Math.max(240, gridEl.getBoundingClientRect().width || 960);

  handle.classList.add('dragging');
  activeFieldResizeDrag = {
    fieldId,
    startX: e.clientX,
    startSpan,
    colPx: gridWidth / 12,
    lastSpan: startSpan,
    handle
  };
}

window.addEventListener('mousemove', (e) => {
  if (!activeFieldResizeDrag) return;
  const deltaCols = Math.round((e.clientX - activeFieldResizeDrag.startX) / activeFieldResizeDrag.colPx);
  const nextSpan = Math.max(2, Math.min(12, activeFieldResizeDrag.startSpan + deltaCols));
  if (nextSpan !== activeFieldResizeDrag.lastSpan) {
    activeFieldResizeDrag.lastSpan = nextSpan;
    persistActiveFormLayoutMutation(next => {
      next.spans[activeFieldResizeDrag.fieldId] = nextSpan;
    });
  }
});

window.addEventListener('mouseup', () => {
  if (!activeFieldResizeDrag) return;
  if (activeFieldResizeDrag.handle) {
    activeFieldResizeDrag.handle.classList.remove('dragging');
  }
  activeFieldResizeDrag = null;
});

if (btnDockToggleLayout) {
  btnDockToggleLayout.addEventListener('click', () => {
    if (!isFormLayoutEditMode && isBottomDockCollapsed) {
      setBottomDockCollapsed(false);
    }
    isFormLayoutEditMode = !isFormLayoutEditMode;
    applyActiveFormLayoutToGrids();
    if (isFormLayoutEditMode) {
      showToast('Layout unlocked — adjust field order (◀ ▶), width (−/+ or drag right border), or rows, then click 🔓 Lock Layout.', 'info');
    }
  });
}

if (btnDockResetLayout) {
  btnDockResetLayout.addEventListener('click', () => {
    const key = getActiveFormLayoutSchemaKey();
    if (formLayoutsStore && formLayoutsStore[key]) {
      delete formLayoutsStore[key];
      if (typeof saveFormLayoutsStore === 'function') {
        saveFormLayoutsStore();
      }
    }
    applyActiveFormLayoutToGrids();
    showToast('Form field layout reset to smart defaults.', 'success');
  });
}

function handleStandardGridDelegatedClick(e) {
  if (e.target.closest('[data-layout-act]')) {
    handleDockLayoutControlAction(e);
    return;
  }
  const tsBtn = e.target.closest('[data-dyn-timestamp]');
  if (tsBtn) {
    const targetArea = document.getElementById(tsBtn.dataset.dynTimestamp);
    if (targetArea) insertTimestampIntoTextarea(targetArea);
  }
}

if (dockStandardFieldsGrid) {
  dockStandardFieldsGrid.addEventListener('click', handleStandardGridDelegatedClick);
  dockStandardFieldsGrid.addEventListener('mousedown', startFieldWidthDragResize);
}
if (modalStandardFieldsGrid) {
  modalStandardFieldsGrid.addEventListener('click', handleStandardGridDelegatedClick);
}

function buildSingleDynamicFieldHtml(col, scopePrefix, extraClass = '') {
  const isDock = scopePrefix === 'dock';
  const idx = col.colIndex;
  const fieldId = `col_${idx}`;
  const colLetter = colNumberToExcelLetters(idx);
  const inputId = `${scopePrefix}CustomCol_${idx}`;
  const isPrimary = idx === activeSchemaConfig.primaryCol;
  const isStatus = idx === activeSchemaConfig.statusCol;
  const isNotes = idx === activeSchemaConfig.notesCol || col.dataType === 'notes';
  const isTicket = idx === activeSchemaConfig.ticketCol;
  const colDataType = col.dataType || 'auto';
  const colValSource = col.validationSource || 'auto';
  const colControlMode = col.controlMode || (isStatus ? 'select' : 'datalist');
  const excelHdr = String(col.excelHeader || col.header || '').trim();
  const fullTitleAttr = col.header !== excelHdr
    ? `${escapeHtml(col.header)} (Excel Column ${escapeHtml(colLetter)}: ${escapeHtml(excelHdr)})`
    : `${escapeHtml(col.header)} (Column ${escapeHtml(colLetter)})`;
  const labelInnerHtml = `<span class="field-lbl-main">${escapeHtml(col.header)}</span> <span class="field-col-tag">Col ${escapeHtml(colLetter)}</span>`;
  const extraClsStr = extraClass ? ` ${extraClass}` : '';

  if (isNotes) {
    const groupClass = (isDock ? 'form-group dock-col-notes' : 'form-group full') + extraClsStr;
    return `
      <div class="${groupClass}" data-field-id="${fieldId}" data-is-notes-col="true">
        <div class="dock-notes-label-row">
          <label for="${inputId}" title="${fullTitleAttr}">${labelInnerHtml}</label>
          <button type="button" class="btn-timestamp-chip" data-dyn-timestamp="${inputId}" title="Insert current date &amp; time stamp">
            + Insert Timestamp
          </button>
        </div>
        <textarea id="${inputId}" data-custom-col="${idx}" rows="${isDock ? 2 : 3}" placeholder="Enter ${escapeHtml(col.header)}..."></textarea>
      </div>
    `;
  }

  const defaultStatusesList = (validationStatuses.length > 0 ? validationStatuses : DEFAULT_STATUSES)
    .filter(s => s && !UUID_REGEX.test(String(s).trim()));
  const cleanColOpts = Array.isArray(col.dropdownOptions)
    ? col.dropdownOptions.filter(s => s && !UUID_REGEX.test(String(s).trim()))
    : [];

  if (isStatus || (colControlMode === 'select' && cleanColOpts.length > 0)) {
    const baseStatusOpts = cleanColOpts.length > 0
      ? cleanColOpts
      : defaultStatusesList;
    const statusOpts = isStatus && typeof mergeRequiredOperationalStatuses === 'function'
      ? mergeRequiredOperationalStatuses(baseStatusOpts)
      : baseStatusOpts;
    const emptyOptHtml = !isStatus ? '<option value="">— Select —</option>' : '';
    const isYardCol = colValSource === 'yard_list' ||
      (colValSource === 'auto' && /\b(yard|yards|yard\s*#|yard\s*name)\b/i.test(excelHdr));
    const optionsHtml = emptyOptHtml + statusOpts
      .map(st => {
        let lbl = st;
        if (isYardCol && Array.isArray(validationYards)) {
          const yHit = validationYards.find(y => String(y.value || '').toLowerCase() === String(st).toLowerCase());
          if (yHit && yHit.label) lbl = yHit.label;
        }
        return `<option value="${escapeHtml(st)}">${escapeHtml(lbl)}</option>`;
      })
      .join('');
    return `
      <div class="form-group${extraClsStr}" data-field-id="${fieldId}">
        <label for="${inputId}" title="${fullTitleAttr}">${labelInnerHtml}</label>
        <select id="${inputId}" data-custom-col="${idx}">
          ${optionsHtml}
        </select>
      </div>
    `;
  }

  const isYardSource =
    colValSource === 'yard_list' ||
    (colValSource === 'auto' && /\b(yard|yards|yard\s*#|yard\s*name)\b/i.test(excelHdr) && Array.isArray(validationYards) && validationYards.length > 0);
  const hasDatalist = isYardSource ? (Array.isArray(validationYards) && validationYards.length > 0) : cleanColOpts.length > 0;
  const datalistId = isYardSource ? 'yardList' : (hasDatalist ? `${scopePrefix}CustomList_${idx}` : '');
  const datalistHtml = (!isYardSource && hasDatalist)
    ? `<datalist id="${datalistId}">${cleanColOpts.map(v => `<option value="${escapeHtml(v)}"></option>`).join('')}</datalist>`
    : '';
  const groupClass = (isPrimary
    ? (isDock ? 'form-group dock-col-subject' : 'form-group full')
    : 'form-group') + extraClsStr;

  const headerLabelBlock = isTicket
    ? `<div class="dock-notes-label-row">
         <label for="${inputId}" title="${fullTitleAttr}">${labelInnerHtml}</label>
         <button type="button" class="btn-timestamp-chip" data-dyn-grab-ticket="${scopePrefix}" title="Extract bracketed [XXXX-XX-XXXX] ticket number(s) from primary subject column">
           🎯 Grab Ticket
         </button>
       </div>`
    : `<label for="${inputId}" title="${fullTitleAttr}">${labelInnerHtml}</label>`;

  const inputModeAttr = (colDataType === 'number' || colDataType === 'currency') ? ' inputmode="decimal"' : '';
  const placeholderSuffix = colDataType === 'date'
    ? ' (MM/DD/YYYY)'
    : (colDataType === 'currency' ? ' ($0.00)' : '...');

  return `
    <div class="${groupClass}" data-field-id="${fieldId}">
      ${headerLabelBlock}
      <input type="text" id="${inputId}" data-custom-col="${idx}" data-col-type="${escapeHtml(colDataType)}"${inputModeAttr} ${hasDatalist ? `list="${datalistId}"` : ''} placeholder="${escapeHtml(col.header)}${placeholderSuffix}">
      ${datalistHtml}
    </div>
  `;
}

function buildDynamicFieldsHtmlForScope(scopePrefix) {
  if (!activeSchemaConfig || activeSchemaConfig.isStandardSchema || !Array.isArray(activeSchemaConfig.columns)) {
    return '';
  }
  return getVisibleCustomEditorColumns()
    .map(col => buildSingleDynamicFieldHtml(col, scopePrefix))
    .join('');
}

function getStandardFieldDomDefs(scopePrefix) {
  const isDock = scopePrefix === 'dock';
  const stIdx = (activeSchemaConfig && activeSchemaConfig.statusCol) || 10;
  const nIdx = (activeSchemaConfig && activeSchemaConfig.notesCol) || 9;
  return [
    { colIndex: 1, fieldId: 'subject', defaultLabel: 'Subject', inputEl: isDock ? dockSubject : fSubject, defaultUpper: false, datalistId: '', requiredDefault: true },
    { colIndex: 2, fieldId: 'ticket', defaultLabel: 'Ticket Number', inputEl: isDock ? dockTicket : fTicket, defaultUpper: true, datalistId: '', requiredDefault: false },
    { colIndex: 3, fieldId: 'broker', defaultLabel: 'Broker Name', inputEl: isDock ? dockBroker : fBroker, defaultUpper: false, datalistId: 'brokerList', requiredDefault: true },
    { colIndex: 4, fieldId: 'master', defaultLabel: 'Master #', inputEl: isDock ? dockMaster : fMaster, defaultUpper: true, datalistId: '', requiredDefault: false },
    { colIndex: 5, fieldId: 'bidder', defaultLabel: 'Bidder #', inputEl: isDock ? dockBidder : fBidder, defaultUpper: true, datalistId: '', requiredDefault: false },
    { colIndex: 6, fieldId: 'action', defaultLabel: 'Requested Action', inputEl: isDock ? dockAction : fAction, defaultUpper: false, datalistId: 'actionList', requiredDefault: false },
    { colIndex: 7, fieldId: 'yard', defaultLabel: 'Yard #', inputEl: isDock ? dockYard : fYard, defaultUpper: true, datalistId: 'yardList', requiredDefault: false },
    { colIndex: 8, fieldId: 'lots', defaultLabel: 'Lot Numbers', inputEl: isDock ? dockLots : fLots, defaultUpper: false, datalistId: '', requiredDefault: false },
    { colIndex: nIdx, fieldId: 'notes', defaultLabel: 'Action Notes', inputEl: isDock ? dockNotes : fNotes, defaultUpper: false, datalistId: '', requiredDefault: false },
    { colIndex: stIdx, fieldId: 'status', defaultLabel: 'Request Status', inputEl: isDock ? dockStatus : fStatus, defaultUpper: false, datalistId: '', requiredDefault: false }
  ];
}

function syncStandardMirrorSelects(gridEl) {
  if (!gridEl) return;
  gridEl.querySelectorAll('.std-override-select[data-mirror-for]').forEach(sel => {
    const inp = document.getElementById(sel.dataset.mirrorFor);
    if (!inp) return;
    const val = String(inp.value || '').trim();
    if (!val) {
      sel.value = '';
      return;
    }
    const matched = ensureSelectOptionExists(sel, val);
    sel.value = matched || val;
  });
}

function syncStandardEditorFieldsFromScope(scopePrefix) {
  const isDock = scopePrefix === 'dock';
  const gridEl = isDock ? dockStandardFieldsGrid : modalStandardFieldsGrid;
  if (!gridEl) return;

  gridEl.querySelectorAll('.std-extra-col-group, .std-override-datalist, .std-override-select').forEach(el => el.remove());

  const isStd = Boolean(!activeSchemaConfig || activeSchemaConfig.isStandardSchema);
  const cols = (isStd && activeSchemaConfig && Array.isArray(activeSchemaConfig.columns)) ? activeSchemaConfig.columns : [];
  const byIdx = new Map();
  cols.forEach(c => byIdx.set(c.colIndex, c));

  const defs = getStandardFieldDomDefs(scopePrefix);
  defs.forEach(def => {
    const groupEl = gridEl.querySelector(`.form-group[data-field-id="${def.fieldId}"]`);
    if (!groupEl || !def.inputEl) return;

    const colObj = byIdx.get(def.colIndex) || null;
    const isVisible = colObj ? (colObj.showInEditor !== false) : true;
    groupEl.style.display = isVisible ? '' : 'none';
    if (def.requiredDefault) {
      def.inputEl.required = Boolean(isStd && isVisible);
    }

    const colLetter = colNumberToExcelLetters(def.colIndex);
    const labelMainEl = groupEl.querySelector('.field-lbl-main');
    const colTagEl = groupEl.querySelector('.field-col-tag');
    if (labelMainEl) {
      labelMainEl.textContent = (colObj && colObj.header)
        ? `${colObj.header}${def.requiredDefault ? ' *' : ''}`
        : `${def.defaultLabel}${def.requiredDefault ? ' *' : ''}`;
    }
    if (colTagEl) {
      colTagEl.textContent = `Col ${colLetter}`;
    }

    if (def.fieldId === 'status' && def.inputEl.tagName === 'SELECT') {
      if (colObj && colObj.validationSource && !['auto', 'status_list', 'none'].includes(colObj.validationSource) && Array.isArray(colObj.dropdownOptions) && colObj.dropdownOptions.length > 0) {
        const prev = def.inputEl.value;
        def.inputEl.innerHTML = colObj.dropdownOptions
          .map(o => `<option value="${escapeHtml(o)}">${escapeHtml(o)}</option>`)
          .join('');
        if (prev) ensureSelectOptionExists(def.inputEl, prev);
      } else if (typeof populateSelectWithStatuses === 'function') {
        populateSelectWithStatuses(def.inputEl);
      }
      return;
    }

    if (def.inputEl.tagName !== 'INPUT') return;

    def.inputEl.style.display = '';
    const dt = (colObj && colObj.dataType) || 'auto';
    if (dt === 'number' || dt === 'currency') {
      def.inputEl.setAttribute('inputmode', 'decimal');
    } else {
      def.inputEl.removeAttribute('inputmode');
    }

    const valSrc = (colObj && colObj.validationSource) || 'auto';
    const ctrlMode = (colObj && colObj.controlMode) || 'datalist';
    let targetListId = def.datalistId || '';
    let optObjects = [];

    if (valSrc === 'none') {
      targetListId = '';
      def.inputEl.removeAttribute('list');
    } else if (valSrc === 'auto') {
      if (targetListId) {
        def.inputEl.setAttribute('list', targetListId);
      } else if (colObj && Array.isArray(colObj.dropdownOptions) && colObj.dropdownOptions.length > 0) {
        optObjects = colObj.dropdownOptions.map(v => ({ value: v, label: v }));
        targetListId = `${scopePrefix}StdList_${def.colIndex}`;
      } else {
        def.inputEl.removeAttribute('list');
      }
    } else if (valSrc === 'yard_list') {
      targetListId = 'yardList';
      def.inputEl.setAttribute('list', 'yardList');
      optObjects = (Array.isArray(validationYards) ? validationYards : []).map(y => ({ value: y.value, label: y.label || y.value }));
    } else if (valSrc === 'broker_list') {
      targetListId = 'brokerList';
      def.inputEl.setAttribute('list', 'brokerList');
      const seenB = new Set();
      (Array.isArray(masterAccountsMap) ? masterAccountsMap : []).forEach(m => {
        const b = String((m && m.broker) || '').trim();
        if (b && !seenB.has(b.toLowerCase())) {
          seenB.add(b.toLowerCase());
          optObjects.push({ value: b, label: b });
        }
      });
    } else if (valSrc === 'action_list') {
      targetListId = 'actionList';
      def.inputEl.setAttribute('list', 'actionList');
      const acts = (Array.isArray(validationActions) && validationActions.length > 0) ? validationActions : DEFAULT_ACTIONS;
      optObjects = acts.map(a => ({ value: a, label: a }));
    } else {
      const opts = (colObj && Array.isArray(colObj.dropdownOptions)) ? colObj.dropdownOptions : [];
      optObjects = opts.map(v => ({ value: v, label: v }));
      targetListId = opts.length > 0 ? `${scopePrefix}StdList_${def.colIndex}` : '';
    }

    if (targetListId && targetListId.startsWith(`${scopePrefix}StdList_`) && optObjects.length > 0) {
      const dlEl = document.createElement('datalist');
      dlEl.id = targetListId;
      dlEl.className = 'std-override-datalist';
      dlEl.innerHTML = optObjects.map(o => `<option value="${escapeHtml(o.value)}"></option>`).join('');
      groupEl.appendChild(dlEl);
      def.inputEl.setAttribute('list', targetListId);
    }

    if (optObjects.length === 0 && targetListId) {
      if (targetListId === 'yardList') {
        optObjects = (Array.isArray(validationYards) ? validationYards : []).map(y => ({ value: y.value, label: y.label || y.value }));
      } else if (targetListId === 'brokerList') {
        const seenB = new Set();
        (Array.isArray(masterAccountsMap) ? masterAccountsMap : []).forEach(m => {
          const b = String((m && m.broker) || '').trim();
          if (b && !seenB.has(b.toLowerCase())) {
            seenB.add(b.toLowerCase());
            optObjects.push({ value: b, label: b });
          }
        });
      } else if (targetListId === 'actionList') {
        const acts = (Array.isArray(validationActions) && validationActions.length > 0) ? validationActions : DEFAULT_ACTIONS;
        optObjects = acts.map(a => ({ value: a, label: a }));
      }
    }

    if (ctrlMode === 'select' && valSrc !== 'none' && optObjects.length > 0) {
      def.inputEl.style.display = 'none';
      const selEl = document.createElement('select');
      selEl.className = 'std-override-select';
      selEl.dataset.mirrorFor = def.inputEl.id;
      selEl.innerHTML = '<option value="">— Select —</option>' + optObjects
        .map(o => `<option value="${escapeHtml(o.value)}">${escapeHtml(o.label)}</option>`)
        .join('');
      selEl.addEventListener('change', () => {
        def.inputEl.value = selEl.value;
        def.inputEl.dispatchEvent(new Event('input', { bubbles: true }));
        def.inputEl.dispatchEvent(new Event('change', { bubbles: true }));
      });
      groupEl.appendChild(selEl);
    }
  });

  if (isStd && cols.length > 0) {
    const stIdx = (activeSchemaConfig && activeSchemaConfig.statusCol) || 10;
    const nIdx = (activeSchemaConfig && activeSchemaConfig.notesCol) || 9;
    const trkIdx = (activeSchemaConfig && activeSchemaConfig.trackingCol) || 11;
    const coreStdSet = new Set([1, 2, 3, 4, 5, 6, 7, 8, nIdx, stIdx, trkIdx]);
    const extraCols = cols.filter(c => !coreStdSet.has(c.colIndex) && c.showInEditor !== false);
    if (extraCols.length > 0) {
      const notesGroup = gridEl.querySelector('.form-group[data-field-id="notes"]');
      extraCols.forEach(c => {
        const tempWrap = document.createElement('div');
        tempWrap.innerHTML = buildSingleDynamicFieldHtml(c, scopePrefix, 'std-extra-col-group').trim();
        const extraGroupEl = tempWrap.firstElementChild;
        if (!extraGroupEl) return;
        if (notesGroup && c.dataType !== 'notes') {
          gridEl.insertBefore(extraGroupEl, notesGroup);
        } else {
          gridEl.appendChild(extraGroupEl);
        }
      });
    }
  }

  syncStandardMirrorSelects(gridEl);
}

function syncStandardEditorFieldsFromSchema() {
  syncStandardEditorFieldsFromScope('dock');
  syncStandardEditorFieldsFromScope('modal');
}

function populateStandardExtraFields(gridEl, itemOrNull) {
  if (!gridEl) return;
  gridEl.querySelectorAll('.std-extra-col-group [data-custom-col]').forEach(el => {
    const colIdx = Number(el.dataset.customCol);
    const key = `col_${colIdx}`;
    if (!itemOrNull || !itemOrNull.customValues) {
      el.value = '';
      return;
    }
    const rawVal = itemOrNull.customValues[key] !== undefined ? String(itemOrNull.customValues[key] || '') : '';
    el.value = UUID_REGEX.test(rawVal.trim()) ? '' : rawVal;
  });
}

function syncDynamicEditorFormsUI() {
  const isCustom = Boolean(
    activeSchemaConfig &&
    !activeSchemaConfig.isStandardSchema &&
    Array.isArray(activeSchemaConfig.columns) &&
    activeSchemaConfig.columns.length > 0
  );

  if (dockStandardFieldsGrid && dockDynamicFieldsGrid) {
    dockStandardFieldsGrid.style.display = isCustom ? 'none' : 'grid';
    dockDynamicFieldsGrid.style.display = isCustom ? 'grid' : 'none';
    if (dockSubject) dockSubject.required = !isCustom;
    if (dockBroker) dockBroker.required = !isCustom;
    dockDynamicFieldsGrid.innerHTML = isCustom ? buildDynamicFieldsHtmlForScope('dock') : '';
  }

  if (modalStandardFieldsGrid && modalDynamicFieldsGrid) {
    modalStandardFieldsGrid.style.display = isCustom ? 'none' : 'grid';
    modalDynamicFieldsGrid.style.display = isCustom ? 'grid' : 'none';
    if (fSubject) fSubject.required = !isCustom;
    if (fBroker) fBroker.required = !isCustom;
    modalDynamicFieldsGrid.innerHTML = isCustom ? buildDynamicFieldsHtmlForScope('modal') : '';
  }

  if (!isCustom) {
    syncStandardEditorFieldsFromSchema();
  }

  // Toggle Grab Ticket buttons visibility for custom schemas when no ticket column is mapped
  const showCustomTicketBtn = isCustom ? Boolean(activeSchemaConfig.ticketCol) : true;
  if (btnDockGrabTicket) btnDockGrabTicket.style.display = !isCustom ? 'inline-flex' : 'none';
  if (btnModalGrabTicket) btnModalGrabTicket.style.display = !isCustom ? 'inline-flex' : 'none';
  if (dockDynamicFieldsGrid) {
    dockDynamicFieldsGrid.querySelectorAll('[data-dyn-grab-ticket]').forEach(b => {
      b.style.display = showCustomTicketBtn ? 'inline-flex' : 'none';
    });
  }

  applyActiveFormLayoutToGrids();
}

function handleDynamicGridDelegatedClick(e) {
  if (e.target.closest('[data-layout-act]')) {
    handleDockLayoutControlAction(e);
    return;
  }
  const tsBtn = e.target.closest('[data-dyn-timestamp]');
  if (tsBtn) {
    const targetArea = document.getElementById(tsBtn.dataset.dynTimestamp);
    if (targetArea) insertTimestampIntoTextarea(targetArea);
    return;
  }
  const grabBtn = e.target.closest('[data-dyn-grab-ticket]');
  if (grabBtn) {
    handleGrabTicketFromSubject(grabBtn.dataset.dynGrabTicket || 'dock');
  }
}

if (dockDynamicFieldsGrid) {
  dockDynamicFieldsGrid.addEventListener('click', handleDynamicGridDelegatedClick);
  dockDynamicFieldsGrid.addEventListener('mousedown', startFieldWidthDragResize);
}
if (modalDynamicFieldsGrid) {
  modalDynamicFieldsGrid.addEventListener('click', handleDynamicGridDelegatedClick);
}

applyActiveFormLayoutToGrids();

function populateDynamicGridFromItem(gridEl, itemOrNull) {
  if (!gridEl || !activeSchemaConfig || activeSchemaConfig.isStandardSchema) return;
  const defaultSt = getDefaultStatus();
  gridEl.querySelectorAll('[data-custom-col]').forEach(inputEl => {
    const colIdx = Number(inputEl.dataset.customCol);
    const key = `col_${colIdx}`;
    if (!itemOrNull) {
      if (colIdx === activeSchemaConfig.statusCol) {
        const canonicalDefault = ensureSelectOptionExists(inputEl, defaultSt) || defaultSt;
        inputEl.value = canonicalDefault;
      } else {
        inputEl.value = '';
      }
      return;
    }
    const rawVal = (itemOrNull.customValues && itemOrNull.customValues[key] !== undefined)
      ? String(itemOrNull.customValues[key] || '')
      : '';
    const val = UUID_REGEX.test(rawVal.trim()) ? '' : rawVal;
    if (colIdx === activeSchemaConfig.statusCol) {
      const candSt = val || (itemOrNull.status && !UUID_REGEX.test(String(itemOrNull.status).trim()) ? itemOrNull.status : '');
      const st = candSt || defaultSt;
      const canonicalSt = ensureSelectOptionExists(inputEl, st) || st;
      inputEl.value = canonicalSt;
    } else {
      inputEl.value = val;
    }
  });
}

function readFormPayloadFromScope(scopeMode) {
  const isDock = scopeMode === 'dock';
  const targetObj = isDock ? dockEditingTarget : editingTarget;
  const existingRow = targetObj
    ? allRows.find(r => r.sheetName === targetObj.sheetName && r.rowNumber === targetObj.rowNumber)
    : null;

  if (activeSchemaConfig && !activeSchemaConfig.isStandardSchema && activeSchemaConfig.columns.length > 0) {
    const gridEl = isDock ? dockDynamicFieldsGrid : modalDynamicFieldsGrid;
    const customValues = (existingRow && existingRow.customValues)
      ? { ...existingRow.customValues }
      : {};
    if (gridEl) {
      gridEl.querySelectorAll('[data-custom-col]').forEach(el => {
        const idx = Number(el.dataset.customCol);
        const colObj = activeSchemaConfig.columns.find(c => c.colIndex === idx);
        let val = String(el.value || '').trim();
        if (colObj && (colObj.validationSource === 'yard_list' || /\b(yard|yards)\b/i.test(String(colObj.excelHeader || colObj.header || ''))) && typeof resolveYardValueFromTypedInput === 'function') {
          val = resolveYardValueFromTypedInput(val);
        }
        if (idx === activeSchemaConfig.ticketCol && val) {
          val = val.toUpperCase();
        } else if (colObj && colObj.dataType && colObj.dataType !== 'auto' && typeof formatValueForColumnDataType === 'function') {
          val = formatValueForColumnDataType(val, colObj.dataType);
        }
        customValues[`col_${idx}`] = val;
      });
    }
    const pCol = activeSchemaConfig.primaryCol;
    const tCol = activeSchemaConfig.ticketCol;
    const secCol = activeSchemaConfig.secondaryCol;
    const actCol = activeSchemaConfig.actionCol;
    const nCol = activeSchemaConfig.notesCol;
    const stCol = activeSchemaConfig.statusCol;

    return {
      isCustomSchema: true,
      customValues,
      subject: pCol ? (customValues[`col_${pCol}`] || '') : '',
      ticket: tCol ? (customValues[`col_${tCol}`] || '').toUpperCase() : '',
      broker: secCol ? (customValues[`col_${secCol}`] || '') : '',
      master: '',
      bidder: '',
      action: actCol ? (customValues[`col_${actCol}`] || '') : '',
      yard: '',
      lots: '',
      notes: nCol ? (customValues[`col_${nCol}`] || '') : '',
      status: stCol ? (customValues[`col_${stCol}`] || getDefaultStatus()) : 'Open'
    };
  }

  const stdSettings = (activeSchemaConfig && activeSchemaConfig.colSettings) || {};
  const fmtStd = (colIdx, rawVal, defaultUpper = false) => {
    let v = String(rawVal || '').trim();
    if (colIdx === 7 && typeof resolveYardValueFromTypedInput === 'function') {
      v = resolveYardValueFromTypedInput(v);
    }
    const cfg = stdSettings[colIdx];
    if (cfg && cfg.dataType && cfg.dataType !== 'auto' && typeof formatValueForColumnDataType === 'function') {
      return formatValueForColumnDataType(v, cfg.dataType);
    }
    if (cfg && cfg.dataType === 'auto') {
      return v;
    }
    return defaultUpper ? v.toUpperCase() : v;
  };

  const stdGridEl = isDock ? dockStandardFieldsGrid : modalStandardFieldsGrid;
  const extraCustomValues = (existingRow && existingRow.customValues)
    ? { ...existingRow.customValues }
    : {};
  let hasExtraStdValues = Object.keys(extraCustomValues).length > 0;
  if (stdGridEl) {
    stdGridEl.querySelectorAll('.std-extra-col-group [data-custom-col]').forEach(el => {
      const idx = Number(el.dataset.customCol);
      const colObj = activeSchemaConfig && activeSchemaConfig.columns
        ? activeSchemaConfig.columns.find(c => c.colIndex === idx)
        : null;
      let val = String(el.value || '').trim();
      if (colObj && (colObj.validationSource === 'yard_list' || /\b(yard|yards)\b/i.test(String(colObj.excelHeader || colObj.header || ''))) && typeof resolveYardValueFromTypedInput === 'function') {
        val = resolveYardValueFromTypedInput(val);
      }
      if (colObj && colObj.dataType && colObj.dataType !== 'auto' && typeof formatValueForColumnDataType === 'function') {
        val = formatValueForColumnDataType(val, colObj.dataType);
      }
      extraCustomValues[`col_${idx}`] = val;
      hasExtraStdValues = true;
    });
  }

  return {
    isCustomSchema: false,
    subject: fmtStd(1, isDock ? dockSubject.value : fSubject.value, false),
    ticket: fmtStd(2, isDock ? dockTicket.value : fTicket.value, true),
    broker: fmtStd(3, isDock ? dockBroker.value : fBroker.value, false),
    master: fmtStd(4, isDock ? dockMaster.value : fMaster.value, true),
    bidder: fmtStd(5, isDock ? dockBidder.value : fBidder.value, true),
    action: fmtStd(6, isDock ? dockAction.value : fAction.value, false),
    yard: fmtStd(7, isDock ? dockYard.value : fYard.value, true),
    lots: fmtStd(8, isDock ? dockLots.value : fLots.value, false),
    notes: (isDock ? dockNotes.value : fNotes.value) || '',
    status: (isDock ? dockStatus.value : fStatus.value) || getDefaultStatus(),
    ...(hasExtraStdValues ? { customValues: extraCustomValues } : {})
  };
}

// ============================================================================
// FEATURE 1: GRAB TICKET NUMBER FROM EMAIL SUBJECT ([XXXX-XX-XXXX])
// ============================================================================
function handleGrabTicketFromSubject(scopeMode = 'dock') {
  const isDock = scopeMode === 'dock';
  const isCustom = Boolean(activeSchemaConfig && !activeSchemaConfig.isStandardSchema);

  let subjectStr = '';
  let ticketInputEl = null;

  if (isCustom) {
    const gridEl = isDock ? dockDynamicFieldsGrid : modalDynamicFieldsGrid;
    if (!gridEl) return;
    if (!activeSchemaConfig.ticketCol) {
      showToast('Assign a Ticket # column in Configure Column Mapping first.', 'warning');
      return;
    }
    ticketInputEl = gridEl.querySelector(`[data-custom-col="${activeSchemaConfig.ticketCol}"]`);
    const primaryInputEl = activeSchemaConfig.primaryCol
      ? gridEl.querySelector(`[data-custom-col="${activeSchemaConfig.primaryCol}"]`)
      : null;
    subjectStr = primaryInputEl ? primaryInputEl.value : '';
    // Also check other text fields if primary column didn't contain the bracketed subject
    if (!extractBracketedTicketsFromSubject(subjectStr).length) {
      gridEl.querySelectorAll('input[type="text"], textarea').forEach(inp => {
        if (inp !== ticketInputEl && !extractBracketedTicketsFromSubject(subjectStr).length) {
          if (extractBracketedTicketsFromSubject(inp.value).length > 0) {
            subjectStr = inp.value;
          }
        }
      });
    }
  } else {
    subjectStr = isDock ? (dockSubject ? dockSubject.value : '') : (fSubject ? fSubject.value : '');
    ticketInputEl = isDock ? dockTicket : fTicket;
  }

  if (!ticketInputEl) return;

  const extracted = extractBracketedTicketsFromSubject(subjectStr);
  if (extracted.length === 0) {
    showToast('No bracketed ticket pattern (e.g. [ABCD-12-3456]) found in Subject.', 'warning');
    return;
  }

  const { merged, added } = mergeCommaDelimitedTickets(ticketInputEl.value, extracted);
  ticketInputEl.value = merged;

  if (added.length > 0) {
    showToast(`🎯 Grabbed ticket: ${added.join(', ')}`, 'success');
  } else {
    showToast(`Ticket ${extracted.join(', ')} is already listed in Ticket Number.`, 'info');
  }
}

if (btnDockGrabTicket) {
  btnDockGrabTicket.addEventListener('click', () => handleGrabTicketFromSubject('dock'));
}
if (btnModalGrabTicket) {
  btnModalGrabTicket.addEventListener('click', () => handleGrabTicketFromSubject('modal'));
}

// ============================================================================
// FEATURE 2: SYNC FROM PREVIOUS ENTRY ("PULL FROM PRIOR TAB")
// ============================================================================
function stripEmailReplyPrefixes(str) {
  return String(str || '')
    .replace(/^\s*((re|fw|fwd)\s*:\s*)+/gi, '')
    .trim()
    .toLowerCase();
}

function findMatchingPriorTabEntries(scopeMode = 'dock') {
  const isDock = scopeMode === 'dock';
  const targetObj = isDock ? dockEditingTarget : editingTarget;
  const refTabName = (targetObj && targetObj.sheetName) ? targetObj.sheetName : (sheetSelect ? sheetSelect.value : '');
  const refTime = DATE_TAB_REGEX.test(refTabName) ? parseDateTab(refTabName) : Infinity;

  const olderRows = allRows
    .filter(r => parseDateTab(r.sheetName) < refTime)
    .sort((a, b) => parseDateTab(b.sheetName) - parseDateTab(a.sheetName) || b.rowNumber - a.rowNumber);

  if (olderRows.length === 0) {
    return { error: 'No earlier daily tabs exist before this tab.' };
  }

  const formPayload = readFormPayloadFromScope(scopeMode);
  const targetKey = targetObj && targetObj.key ? targetObj.key.trim().toLowerCase() : '';
  const formTicketTokens = splitNormalizedTicketTokens(formPayload.ticket);
  const cleanFormSubj = stripEmailReplyPrefixes(formPayload.subject);

  // Also check if Subject has a bracketed ticket not yet grabbed!
  const subjectTickets = extractBracketedTicketsFromSubject(formPayload.subject);
  const combinedTickets = Array.from(new Set([...formTicketTokens, ...subjectTickets]));

  // Build chronological match chain (newest prior first)
  const matches = olderRows.filter(prior => {
    if (targetKey && prior.key && prior.key.trim().toLowerCase() === targetKey) {
      return true;
    }
    if (combinedTickets.length > 0 && prior.ticket) {
      const priorTokens = splitNormalizedTicketTokens(prior.ticket);
      if (priorTokens.some(t => combinedTickets.includes(t))) {
        return true;
      }
    }
    if (typeof areRowsSameUniqueRequest === 'function') {
      const candidate = { ...formPayload, key: targetKey };
      if (areRowsSameUniqueRequest(prior, candidate)) {
        return true;
      }
    }
    if (cleanFormSubj.length >= 5) {
      const cleanPriorSubj = stripEmailReplyPrefixes(prior.subject);
      if (cleanPriorSubj && cleanPriorSubj === cleanFormSubj) {
        return true;
      }
    }
    return false;
  });

  if (matches.length === 0) {
    return {
      error: 'No matching entry found on earlier daily tabs (enter a Subject, Ticket #, or Broker/Lot first).'
    };
  }
  return { matches };
}

function handlePullFromPriorTab(scopeMode = 'dock') {
  if (!workbook || allRows.length === 0) return;
  const isDock = scopeMode === 'dock';
  const result = findMatchingPriorTabEntries(scopeMode);
  if (result.error) {
    showToast(result.error, 'warning');
    return;
  }

  const matches = result.matches; // sorted newest-prior -> oldest-prior
  const sourceTabName = matches[0].sheetName;
  const updatedLabels = [];

  if (activeSchemaConfig && !activeSchemaConfig.isStandardSchema && activeSchemaConfig.columns.length > 0) {
    const gridEl = isDock ? dockDynamicFieldsGrid : modalDynamicFieldsGrid;
    if (!gridEl) return;

    activeSchemaConfig.columns.forEach(col => {
      const idx = col.colIndex;
      if (idx === activeSchemaConfig.statusCol) return; // Keep current tab's status untouched
      const el = gridEl.querySelector(`[data-custom-col="${idx}"]`);
      if (!el) return;

      if (idx === activeSchemaConfig.ticketCol) {
        const priorTickets = matches.map(m => (m.customValues && m.customValues[`col_${idx}`]) || m.ticket || '').filter(Boolean);
        const { merged, added } = mergeCommaDelimitedTickets(el.value, priorTickets);
        if (added.length > 0 && merged !== el.value.trim()) {
          el.value = merged;
          updatedLabels.push(col.header);
        }
        return;
      }

      if (el.value.trim() !== '') return; // Strictly non-destructive: never overwrite filled fields
      const donor = matches.find(m => {
        const cand = m.customValues ? String(m.customValues[`col_${idx}`] || '').trim() : '';
        return cand !== '' && !UUID_REGEX.test(cand);
      });
      if (donor) {
        el.value = String(donor.customValues[`col_${idx}`]).trim();
        updatedLabels.push(col.header);
      }
    });
  } else {
    const fields = isDock
      ? [
          { key: 'subject', label: 'Subject', el: dockSubject },
          { key: 'broker', label: 'Broker', el: dockBroker },
          { key: 'master', label: 'Master #', el: dockMaster },
          { key: 'bidder', label: 'Bidder #', el: dockBidder },
          { key: 'action', label: 'Action', el: dockAction },
          { key: 'yard', label: 'Yard #', el: dockYard },
          { key: 'lots', label: 'Lots', el: dockLots },
          { key: 'notes', label: 'Notes', el: dockNotes }
        ]
      : [
          { key: 'subject', label: 'Subject', el: fSubject },
          { key: 'broker', label: 'Broker', el: fBroker },
          { key: 'master', label: 'Master #', el: fMaster },
          { key: 'bidder', label: 'Bidder #', el: fBidder },
          { key: 'action', label: 'Action', el: fAction },
          { key: 'yard', label: 'Yard #', el: fYard },
          { key: 'lots', label: 'Lots', el: fLots },
          { key: 'notes', label: 'Notes', el: fNotes }
        ];

    const ticketEl = isDock ? dockTicket : fTicket;
    if (ticketEl) {
      const priorTickets = matches.map(m => m.ticket).filter(Boolean);
      const { merged, added } = mergeCommaDelimitedTickets(ticketEl.value, priorTickets);
      if (added.length > 0 && merged !== ticketEl.value.trim()) {
        ticketEl.value = merged;
        updatedLabels.push('Ticket #');
      }
    }

    fields.forEach(({ key, label, el }) => {
      if (!el || el.value.trim() !== '') return;
      const donor = matches.find(m => String(m[key] || '').trim() !== '');
      if (donor) {
        el.value = String(donor[key]).trim();
        updatedLabels.push(label);
      }
    });

    const stdGridEl = isDock ? dockStandardFieldsGrid : modalStandardFieldsGrid;
    if (stdGridEl) {
      stdGridEl.querySelectorAll('.std-extra-col-group [data-custom-col]').forEach(el => {
        if (el.value.trim() !== '') return;
        const cIdx = Number(el.dataset.customCol);
        const donor = matches.find(m => m.customValues && String(m.customValues[`col_${cIdx}`] || '').trim() !== '');
        if (donor) {
          el.value = String(donor.customValues[`col_${cIdx}`]).trim();
          const cObj = activeSchemaConfig?.columns?.find(c => c.colIndex === cIdx);
          updatedLabels.push(cObj ? cObj.header : `Col ${cIdx}`);
        }
      });
      syncStandardMirrorSelects(stdGridEl);
    }
  }

  if (updatedLabels.length === 0) {
    showToast(`Matched prior entry on ${sourceTabName}, and all non-empty fields are already filled here.`, 'info');
  } else {
    showToast(`📥 Pulled ${updatedLabels.join(', ')} from prior tab (${sourceTabName}).`, 'success');
  }
}

if (btnDockPullPrior) {
  btnDockPullPrior.addEventListener('click', () => handlePullFromPriorTab('dock'));
}
if (btnModalPullPrior) {
  btnModalPullPrior.addEventListener('click', () => handlePullFromPriorTab('modal'));
}

// ============================================================================
// FEATURE 3: FORWARD BATCH SYNC ACROSS DAILY TABS (NON-DESTRUCTIVE + TICKET MERGE)
// ============================================================================
async function handleForwardBatchSync(triggerSource = 'dock') {
  if (!workbook) return;

  let sourceSheetName = '';
  let sourceRowNumber = 0;
  let sourceKey = '';
  let sourceSnapshot = null;
  let saveDockSourceToo = false;

  if (triggerSource === 'dock') {
    if (!dockEditingTarget) {
      showToast('Select an existing request row first to sync forward to newer tabs.', 'warning');
      return;
    }
    sourceSheetName = dockEditingTarget.sheetName;
    sourceRowNumber = dockEditingTarget.rowNumber;
    sourceKey = dockEditingTarget.key;
    sourceSnapshot = {
      ...readFormPayloadFromScope('dock'),
      sheetName: sourceSheetName,
      rowNumber: sourceRowNumber,
      key: sourceKey
    };
    saveDockSourceToo = true;
  } else if (triggerSource === 'details') {
    if (!activeTicketDetailsTarget) return;
    const seed = allRows.find(
      r => r.sheetName === activeTicketDetailsTarget.sheetName && r.rowNumber === activeTicketDetailsTarget.rowNumber
    ) || allRows.find(r => r.key === activeTicketDetailsTarget.key);
    if (!seed) return;

    // Consolidate known values across the ticket's history up to each newer tab!
    const history = findTicketHistoryEntries(seed);
    sourceSheetName = history[0].sheetName;
    sourceRowNumber = history[0].rowNumber;
    sourceKey = seed.key;
    sourceSnapshot = seed;
  }

  if (!sourceSnapshot) return;

  const sourceTime = parseDateTab(sourceSheetName);
  const normKey = String(sourceKey || '').trim().toLowerCase();

  // Find all matching rolled-over rows on newer tabs (or across full forward chain when triggered from Details)
  const forwardTargets = allRows
    .filter(r => {
      const rTime = parseDateTab(r.sheetName);
      if (triggerSource === 'dock' && rTime <= sourceTime) return false;
      if (triggerSource === 'details' && rTime <= sourceTime && r.rowNumber === sourceRowNumber) return false;
      if (normKey && r.key && r.key.trim().toLowerCase() === normKey) return true;
      return typeof areRowsSameUniqueRequest === 'function' && areRowsSameUniqueRequest(sourceSnapshot, r);
    })
    .sort((a, b) => parseDateTab(a.sheetName) - parseDateTab(b.sheetName) || a.rowNumber - b.rowNumber);

  if (forwardTargets.length === 0) {
    showToast(`No newer daily tabs contain a rolled-over copy of this request after ${sourceSheetName}.`, 'info');
    return;
  }

  // Build cumulative donor values chronologically so data added on any earlier tab flows forward cleanly
  const isCustom = Boolean(activeSchemaConfig && !activeSchemaConfig.isStandardSchema && activeSchemaConfig.columns.length > 0);
  const cumulativeDonor = JSON.parse(JSON.stringify(sourceSnapshot));
  if (triggerSource === 'details') {
    const earliest = allRows.find(r => r.sheetName === sourceSheetName && r.rowNumber === sourceRowNumber);
    if (earliest) {
      Object.assign(cumulativeDonor, JSON.parse(JSON.stringify(earliest)));
    }
  }

  const plannedMutations = [];

  forwardTargets.forEach(targetRow => {
    const diffs = [];
    const updatedCopy = JSON.parse(JSON.stringify(targetRow));

    if (isCustom) {
      if (!updatedCopy.customValues) updatedCopy.customValues = {};
      if (!cumulativeDonor.customValues) cumulativeDonor.customValues = {};

      activeSchemaConfig.columns.forEach(col => {
        const idx = col.colIndex;
        if (idx === activeSchemaConfig.statusCol) return; // Do not clobber forward tab's status
        const cKey = `col_${idx}`;
        const rawDonorVal = String(cumulativeDonor.customValues[cKey] || '').trim();
        const rawCurrVal = String(updatedCopy.customValues[cKey] || '').trim();
        const donorVal = UUID_REGEX.test(rawDonorVal) ? '' : rawDonorVal;
        const currVal = UUID_REGEX.test(rawCurrVal) ? '' : rawCurrVal;

        if (idx === activeSchemaConfig.ticketCol) {
          const { merged, added } = mergeCommaDelimitedTickets(currVal, donorVal);
          if (added.length > 0 && merged !== currVal) {
            diffs.push(`${col.header}: "${currVal || '(blank)'}" → "${merged}"`);
            updatedCopy.customValues[cKey] = merged;
            updatedCopy.ticket = merged;
          }
          // Accumulate target's tickets into donor for subsequent even-newer tabs
          cumulativeDonor.customValues[cKey] = mergeCommaDelimitedTickets(donorVal, updatedCopy.customValues[cKey]).merged;
          return;
        }

        if (!currVal && donorVal) {
          diffs.push(`${col.header}: (blank) → "${donorVal.slice(0, 45)}"`);
          updatedCopy.customValues[cKey] = donorVal;
          if (idx === activeSchemaConfig.primaryCol) updatedCopy.subject = donorVal;
          if (idx === activeSchemaConfig.secondaryCol) updatedCopy.broker = donorVal;
          if (idx === activeSchemaConfig.actionCol) updatedCopy.action = donorVal;
          if (idx === activeSchemaConfig.notesCol) updatedCopy.notes = donorVal;
        } else if (currVal && !donorVal) {
          cumulativeDonor.customValues[cKey] = currVal;
        }
      });
    } else {
      // 1. Ticket Number: comma-delimited deduplicated merge
      const { merged: mergedTicket, added: addedTickets } = mergeCommaDelimitedTickets(updatedCopy.ticket, cumulativeDonor.ticket);
      if (addedTickets.length > 0 && mergedTicket !== String(updatedCopy.ticket || '').trim()) {
        diffs.push(`Ticket #: "${updatedCopy.ticket || '(blank)'}" → "${mergedTicket}"`);
        updatedCopy.ticket = mergedTicket;
      }
      cumulativeDonor.ticket = mergeCommaDelimitedTickets(cumulativeDonor.ticket, updatedCopy.ticket).merged;

      // 2. Non-destructive fill of blank fields only (never wipes existing values)
      const syncableFields = [
        { key: 'subject', label: 'Subject', upper: false },
        { key: 'broker', label: 'Broker', upper: false },
        { key: 'master', label: 'Master #', upper: true },
        { key: 'bidder', label: 'Bidder #', upper: true },
        { key: 'action', label: 'Action', upper: false },
        { key: 'yard', label: 'Yard #', upper: true },
        { key: 'lots', label: 'Lots', upper: false },
        { key: 'notes', label: 'Notes', upper: false }
      ];

      syncableFields.forEach(({ key, label, upper }) => {
        const rawDonor = String(cumulativeDonor[key] || '').trim();
        const donorVal = upper ? rawDonor.toUpperCase() : rawDonor;
        const currVal = String(updatedCopy[key] || '').trim();

        if (!currVal && donorVal) {
          diffs.push(`${label}: (blank) → "${donorVal.slice(0, 45)}"`);
          updatedCopy[key] = donorVal;
        } else if (currVal && !donorVal) {
          cumulativeDonor[key] = upper ? currVal.toUpperCase() : currVal;
        }
      });
    }

    if (normKey && (!updatedCopy.key || updatedCopy.key.toLowerCase() !== normKey)) {
      updatedCopy.key = normKey;
    }

    if (diffs.length > 0) {
      plannedMutations.push({
        sheetName: targetRow.sheetName,
        rowNumber: targetRow.rowNumber,
        updatedItem: updatedCopy,
        diffs
      });
    }
  });

  if (plannedMutations.length === 0) {
    showToast('✅ All newer daily tabs already contain these values — nothing needed forward-syncing.', 'info');
    return;
  }

  const previewText = plannedMutations
    .slice(0, 8)
    .map(m => `• [${m.sheetName} • Row ${m.rowNumber}] ${m.diffs.join('; ')}`)
    .join('\n');
  const moreNote = plannedMutations.length > 8 ? `\n...and ${plannedMutations.length - 8} more tab(s).` : '';

  const confirmed = await showModalConfirm({
    title: `🔄 Sync Data Forward to ${plannedMutations.length} Newer Tab(s)?`,
    badge: 'Non-Destructive Merge',
    message:
      `Propagate newly added details forward without overwriting any existing values (Ticket Numbers merge comma-delimited):\n\n` +
      `${previewText}${moreNote}`,
    confirmText: `🔄 Sync ${plannedMutations.length} Tab(s) Now`
  });
  if (!confirmed) return;

  const activeTab = sheetSelect ? sheetSelect.value : sourceSheetName;
  await loadWorkbookFromDisk(activeTab, { skipStagedRecoveryCheck: true });

  // Also write source row if triggered from dock with live edits
  if (saveDockSourceToo) {
    const srcWs = workbook.getWorksheet(sourceSheetName);
    if (srcWs) {
      const srcMap = getWorksheetColumnMap(srcWs);
      srcWs.getColumn(srcMap.trackingCol).hidden = true;
      const srcRow = srcWs.getRow(sourceRowNumber);
      writeItemToWorksheetRow(srcRow, sourceRowNumber, {
        ...sourceSnapshot,
        ticket: sourceSnapshot.ticket ? sourceSnapshot.ticket.toUpperCase() : '',
        master: sourceSnapshot.master ? sourceSnapshot.master.toUpperCase() : '',
        bidder: sourceSnapshot.bidder ? sourceSnapshot.bidder.toUpperCase() : '',
        yard: sourceSnapshot.yard ? sourceSnapshot.yard.toUpperCase() : '',
        key: sourceKey
      }, srcMap);
    }
  }

  plannedMutations.forEach(mut => {
    const ws = workbook.getWorksheet(mut.sheetName);
    if (!ws) return;
    const colMap = getWorksheetColumnMap(ws);
    ws.getColumn(colMap.trackingCol).hidden = true;
    const excelRow = ws.getRow(mut.rowNumber);
    writeItemToWorksheetRow(excelRow, mut.rowNumber, mut.updatedItem, colMap);
  });

  const saveResult = await saveWorkbookToDisk();
  if (saveResult) {
    await loadWorkbookFromDisk(activeTab, {
      skipStagedRecoveryCheck: true,
      preserveStatus: saveResult === 'mirror-locked'
    });
    if (ticketDetailsModal && ticketDetailsModal.classList.contains('open') && activeTicketDetailsTarget) {
      const refreshed = allRows.find(
        r => r.sheetName === activeTicketDetailsTarget.sheetName && r.rowNumber === activeTicketDetailsTarget.rowNumber
      ) || allRows.find(r => r.key === activeTicketDetailsTarget.key);
      if (refreshed) renderTicketDetailsModal(refreshed);
    }
    showToast(`🔄 Forward-synced updates across ${plannedMutations.length} newer daily tab(s)!`, 'success');
  }
}

if (btnDockSyncForward) {
  btnDockSyncForward.addEventListener('click', () => handleForwardBatchSync('dock'));
}
if (btnTicketSyncForward) {
  btnTicketSyncForward.addEventListener('click', () => handleForwardBatchSync('details'));
}

// ============================================================================
// BOTTOM DOCK COLLAPSE, NEW MODE, EDIT MODE & MODAL EDITORS
// ============================================================================
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

// --- UNSAVED NEW REQUEST DRAFT AUTOSAVE & OVERWRITE PROTECTION ---
const DOCK_NEW_DRAFT_STORAGE_KEY = 'trackerui_dock_new_draft_v1';
let dockDraftSaveTimer = null;

function getActiveDockDraftScopeKey() {
  const prof = (typeof activeProfileId === 'string' && activeProfileId) ? activeProfileId : 'default';
  const wbName = (typeof workingFileHandle !== 'undefined' && workingFileHandle && workingFileHandle.name)
    ? workingFileHandle.name
    : '__no_workbook__';
  const schemaTag = (activeSchemaConfig && !activeSchemaConfig.isStandardSchema) ? 'custom' : 'standard';
  return `${prof}::${wbName}::${schemaTag}`;
}

function readAllSavedDockDraftsMap() {
  try {
    const raw = localStorage.getItem(DOCK_NEW_DRAFT_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) return parsed;
    }
  } catch (_) { /* ignore */ }
  return {};
}

function writeAllSavedDockDraftsMap(mapObj) {
  try {
    localStorage.setItem(DOCK_NEW_DRAFT_STORAGE_KEY, JSON.stringify(mapObj || {}));
  } catch (_) { /* ignore */ }
}

function readNewDockDraftSnapshot() {
  const isCustom = Boolean(activeSchemaConfig && !activeSchemaConfig.isStandardSchema);
  if (isCustom) {
    const customValues = {};
    if (dockDynamicFieldsGrid) {
      dockDynamicFieldsGrid.querySelectorAll('[data-custom-col]').forEach(el => {
        const colIdx = Number(el.dataset.customCol);
        if (colIdx >= 1) {
          customValues[`col_${colIdx}`] = String(el.value || '');
        }
      });
    }
    return {
      savedAt: Date.now(),
      isStandard: false,
      customValues
    };
  }
  const extraCustomValues = {};
  if (dockStandardFieldsGrid) {
    dockStandardFieldsGrid.querySelectorAll('.std-extra-col-group [data-custom-col]').forEach(el => {
      const colIdx = Number(el.dataset.customCol);
      if (colIdx >= 1) extraCustomValues[`col_${colIdx}`] = String(el.value || '');
    });
  }
  return {
    savedAt: Date.now(),
    isStandard: true,
    fields: {
      subject: dockSubject ? dockSubject.value : '',
      ticket: dockTicket ? dockTicket.value : '',
      broker: dockBroker ? dockBroker.value : '',
      master: dockMaster ? dockMaster.value : '',
      bidder: dockBidder ? dockBidder.value : '',
      action: dockAction ? dockAction.value : '',
      yard: dockYard ? dockYard.value : '',
      lots: dockLots ? dockLots.value : '',
      notes: dockNotes ? dockNotes.value : '',
      status: dockStatus ? dockStatus.value : '',
      extraCustomValues
    }
  };
}

function isNewDockDraftSnapshotNonEmpty(snap) {
  if (!snap || typeof snap !== 'object') return false;
  if (snap.isStandard && snap.fields) {
    const f = snap.fields;
    const extraNonEmpty = f.extraCustomValues
      ? Object.values(f.extraCustomValues).some(v => String(v || '').trim() !== '')
      : false;
    return [f.subject, f.ticket, f.broker, f.master, f.bidder, f.action, f.yard, f.lots, f.notes]
      .some(v => String(v || '').trim() !== '') || extraNonEmpty;
  }
  if (!snap.isStandard && snap.customValues) {
    const stColKey = (activeSchemaConfig && activeSchemaConfig.statusCol) ? `col_${activeSchemaConfig.statusCol}` : '';
    return Object.entries(snap.customValues).some(([k, v]) => {
      if (k === stColKey) return false;
      return String(v || '').trim() !== '';
    });
  }
  return false;
}

function hasUnsavedNewDockDraft() {
  if (dockEditingTarget || !bottomEditorDock || bottomEditorDock.style.display === 'none') {
    return false;
  }
  return isNewDockDraftSnapshotNonEmpty(readNewDockDraftSnapshot());
}

function getNewDockDraftSummaryLabel(snap) {
  if (!snap) return 'Untitled Draft';
  let preview = '';
  if (snap.isStandard && snap.fields) {
    preview = String(snap.fields.subject || snap.fields.ticket || snap.fields.broker || snap.fields.notes || '').trim();
  } else if (snap.customValues) {
    const stColKey = (activeSchemaConfig && activeSchemaConfig.statusCol) ? `col_${activeSchemaConfig.statusCol}` : '';
    for (const [k, v] of Object.entries(snap.customValues)) {
      if (k === stColKey) continue;
      const t = String(v || '').trim();
      if (t) {
        preview = t;
        break;
      }
    }
  }
  if (!preview) return 'Untitled Draft';
  return preview.length > 48 ? `${preview.slice(0, 48)}…` : preview;
}

function loadSavedNewDockDraft() {
  const map = readAllSavedDockDraftsMap();
  const snap = map[getActiveDockDraftScopeKey()];
  return isNewDockDraftSnapshotNonEmpty(snap) ? snap : null;
}

function saveNewDockDraftToStorage() {
  if (dockEditingTarget) return false;
  const snap = readNewDockDraftSnapshot();
  const map = readAllSavedDockDraftsMap();
  const scopeKey = getActiveDockDraftScopeKey();
  if (isNewDockDraftSnapshotNonEmpty(snap)) {
    map[scopeKey] = snap;
    writeAllSavedDockDraftsMap(map);
    syncDockDraftIndicatorsUI();
    return true;
  } else {
    if (map[scopeKey]) {
      delete map[scopeKey];
      writeAllSavedDockDraftsMap(map);
    }
    syncDockDraftIndicatorsUI();
    return false;
  }
}

function clearSavedNewDockDraft() {
  if (dockDraftSaveTimer) {
    clearTimeout(dockDraftSaveTimer);
    dockDraftSaveTimer = null;
  }
  const map = readAllSavedDockDraftsMap();
  const scopeKey = getActiveDockDraftScopeKey();
  if (map[scopeKey]) {
    delete map[scopeKey];
    writeAllSavedDockDraftsMap(map);
  }
  syncDockDraftIndicatorsUI();
}

function applySavedDraftToBottomDock(snap) {
  if (!snap || typeof snap !== 'object') return false;
  if (snap.isStandard && snap.fields && (!activeSchemaConfig || activeSchemaConfig.isStandardSchema)) {
    const f = snap.fields;
    if (dockSubject) dockSubject.value = f.subject || '';
    if (dockTicket) dockTicket.value = f.ticket || '';
    if (dockBroker) dockBroker.value = f.broker || '';
    if (dockMaster) dockMaster.value = f.master || '';
    if (dockBidder) dockBidder.value = f.bidder || '';
    if (dockAction) dockAction.value = f.action || '';
    if (dockYard) dockYard.value = f.yard || '';
    if (dockLots) dockLots.value = f.lots || '';
    if (dockNotes) dockNotes.value = f.notes || '';
    if (dockStatus && f.status) {
      const canonicalDraftSt = ensureStatusOptionExists(f.status);
      dockStatus.value = canonicalDraftSt;
    }
    if (f.extraCustomValues) {
      populateStandardExtraFields(dockStandardFieldsGrid, { customValues: f.extraCustomValues });
    }
    syncStandardMirrorSelects(dockStandardFieldsGrid);
    return true;
  }
  if (!snap.isStandard && snap.customValues && activeSchemaConfig && !activeSchemaConfig.isStandardSchema) {
    populateDynamicGridFromItem(dockDynamicFieldsGrid, { customValues: snap.customValues });
    return true;
  }
  return false;
}

function syncDockDraftIndicatorsUI() {
  const savedDraft = loadSavedNewDockDraft();
  const isNewModeWithDraft = !dockEditingTarget && (hasUnsavedNewDockDraft() || Boolean(savedDraft));
  if (dockDraftBadge) {
    if (isNewModeWithDraft) {
      dockDraftBadge.style.display = 'inline-flex';
      const tsStr = savedDraft && savedDraft.savedAt
        ? new Date(savedDraft.savedAt).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })
        : '';
      dockDraftBadge.textContent = tsStr ? `● Draft Autosaved (${tsStr})` : '● Draft Autosaved';
    } else {
      dockDraftBadge.style.display = 'none';
    }
  }
  if (btnDockResumeDraft) {
    if (dockEditingTarget && savedDraft) {
      btnDockResumeDraft.style.display = 'inline-block';
      btnDockResumeDraft.title = `Resume unsaved new request draft: "${getNewDockDraftSummaryLabel(savedDraft)}"`;
    } else {
      btnDockResumeDraft.style.display = 'none';
    }
  }
}

async function confirmSwitchAwayFromUnsavedNewDraft(actionDesc = 'switch to another row') {
  if (!hasUnsavedNewDockDraft()) return true;
  saveNewDockDraftToStorage();
  const snap = readNewDockDraftSnapshot();
  const summary = getNewDockDraftSummaryLabel(snap);
  return await showModalConfirm({
    title: 'Unsaved New Request in Progress',
    badge: 'Draft Auto-Saved',
    message:
      `You have an unsaved new request in the bottom editor dock ("${summary}").\n\n` +
      `Your draft has been auto-saved in browser storage so you won't lose your work. ` +
      `Do you want to ${actionDesc} now? (You can return to your draft anytime via "📝 Resume Draft".)`,
    confirmText: 'Switch & Keep Draft',
    cancelText: 'Stay on New Request',
    variant: 'primary'
  });
}

if (dockEntryForm) {
  const scheduleDraftAutosave = () => {
    if (dockEditingTarget) return;
    if (dockDraftSaveTimer) clearTimeout(dockDraftSaveTimer);
    dockDraftSaveTimer = setTimeout(() => {
      dockDraftSaveTimer = null;
      saveNewDockDraftToStorage();
    }, 200);
  };
  dockEntryForm.addEventListener('input', scheduleDraftAutosave);
  dockEntryForm.addEventListener('change', scheduleDraftAutosave);
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
      loadRowIntoBottomDock(updatedMatch.sheetName, updatedMatch.rowNumber, false, { expandIfCollapsed: false, recordAccess: false });
      return;
    }
  }
  setBottomDockNewMode(false, { restoreSavedDraft: true });
}

function focusFirstEditorInput(scopeMode = 'dock') {
  const isDock = scopeMode === 'dock';
  if (activeSchemaConfig && !activeSchemaConfig.isStandardSchema) {
    const gridEl = isDock ? dockDynamicFieldsGrid : modalDynamicFieldsGrid;
    const firstEl = gridEl ? gridEl.querySelector('input, textarea, select') : null;
    if (firstEl) firstEl.focus();
    return;
  }
  const stdFirst = isDock ? dockSubject : fSubject;
  if (stdFirst) stdFirst.focus();
}

function setBottomDockNewMode(focusInput = false, { restoreSavedDraft = false, clearDraft = false } = {}) {
  dockEditingTarget = null;
  if (!bottomEditorDock) return;

  if (clearDraft) {
    clearSavedNewDockDraft();
  }

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
    dockTargetLabel.textContent = `Creating new request on ${activeTab} — or click any row above to edit`;
  }
  if (dockKeyBadge) {
    dockKeyBadge.style.display = 'none';
    dockKeyBadge.textContent = '';
  }
  if (dockHintText) {
    dockHintText.textContent = 'Tip: Click any row above to edit, or use 🎯 Grab Ticket / 📥 Pull Prior to auto-fill.';
  }
  if (btnDockSyncForward) {
    btnDockSyncForward.style.display = 'none';
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
  populateStandardExtraFields(dockStandardFieldsGrid, null);
  syncStandardMirrorSelects(dockStandardFieldsGrid);
  populateDynamicGridFromItem(dockDynamicFieldsGrid, null);

  if (restoreSavedDraft) {
    const savedDraft = loadSavedNewDockDraft();
    if (savedDraft) {
      applySavedDraftToBottomDock(savedDraft);
    }
  }

  syncDockDraftIndicatorsUI();

  if (listBody) {
    listBody.querySelectorAll('tr.active-editing-row').forEach(tr => tr.classList.remove('active-editing-row'));
  }

  if (focusInput) {
    setTimeout(() => focusFirstEditorInput('dock'), 30);
  }
}

function loadRowIntoBottomDock(sheetName, rowNumber, focusInput = false, { expandIfCollapsed = true, recordAccess = true } = {}) {
  const item = allRows.find(r => r.sheetName === sheetName && r.rowNumber === rowNumber);
  if (!item || !bottomEditorDock) return;

  if (recordAccess && item.key && typeof recordRowLastAccessed === 'function') {
    recordRowLastAccessed(item.key);
    if (typeof refreshRowLastAccessedDOM === 'function') {
      refreshRowLastAccessedDOM(item.key);
    }
  }

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
  if (btnDockSyncForward) {
    btnDockSyncForward.style.display = 'inline-block';
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

  if (activeSchemaConfig && !activeSchemaConfig.isStandardSchema) {
    populateDynamicGridFromItem(dockDynamicFieldsGrid, item);
  } else {
    if (dockSubject) dockSubject.value = item.subject;
    if (dockTicket) dockTicket.value = item.ticket;
    if (dockBroker) dockBroker.value = item.broker;
    if (dockMaster) dockMaster.value = item.master;
    if (dockBidder) dockBidder.value = item.bidder;
    if (dockAction) dockAction.value = item.action;
    if (dockYard) dockYard.value = item.yard;
    if (dockLots) dockLots.value = item.lots;
    if (dockNotes) dockNotes.value = item.notes;

    const canonicalDockSt = ensureStatusOptionExists(item.status);
    if (dockStatus) dockStatus.value = canonicalDockSt;
    populateStandardExtraFields(dockStandardFieldsGrid, item);
    syncStandardMirrorSelects(dockStandardFieldsGrid);
  }

  syncDockDraftIndicatorsUI();

  if (listBody) {
    listBody.querySelectorAll('tr').forEach(tr => {
      const isMatch = tr.dataset.sheet === sheetName && Number(tr.dataset.row) === rowNumber;
      tr.classList.toggle('active-editing-row', isMatch);
    });
  }

  if (focusInput) {
    setTimeout(() => focusFirstEditorInput('dock'), 30);
  }
}

if (btnDockResumeDraft) {
  btnDockResumeDraft.addEventListener('click', () => {
    setBottomDockNewMode(true, { restoreSavedDraft: true });
    showToast('Restored your auto-saved new request draft.', 'info');
  });
}

if (btnDockNewClear) {
  btnDockNewClear.addEventListener('click', async () => {
    if (!dockEditingTarget && hasUnsavedNewDockDraft()) {
      const clearConfirmed = await showModalConfirm({
        title: 'Discard Unsaved New Request Draft?',
        badge: 'Clear Form',
        message: 'Clearing the New Request form will discard your current auto-saved draft. Continue?',
        confirmText: 'Clear & Discard Draft',
        cancelText: 'Keep Draft',
        variant: 'danger'
      });
      if (!clearConfirmed) return;
      setBottomDockNewMode(true, { clearDraft: true });
      return;
    }
    const hasDraft = Boolean(loadSavedNewDockDraft());
    setBottomDockNewMode(true, { restoreSavedDraft: hasDraft });
    if (hasDraft) {
      showToast('Switched to New Request and restored your auto-saved draft.', 'info');
    }
  });
}

if (btnDockCancelEdit) {
  btnDockCancelEdit.addEventListener('click', () => {
    setBottomDockNewMode(false, { restoreSavedDraft: true });
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
  populateStandardExtraFields(modalStandardFieldsGrid, null);
  syncStandardMirrorSelects(modalStandardFieldsGrid);
  populateDynamicGridFromItem(modalDynamicFieldsGrid, null);

  editModal.classList.add('open');
  setTimeout(() => focusFirstEditorInput('modal'), 20);
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

  if (activeSchemaConfig && !activeSchemaConfig.isStandardSchema) {
    populateDynamicGridFromItem(modalDynamicFieldsGrid, item);
  } else {
    fSubject.value = item.subject;
    fTicket.value = item.ticket;
    fBroker.value = item.broker;
    fMaster.value = item.master;
    fBidder.value = item.bidder;
    fAction.value = item.action;
    fYard.value = item.yard;
    fLots.value = item.lots;
    fNotes.value = item.notes;

    const canonicalModalSt = ensureStatusOptionExists(item.status);
    fStatus.value = canonicalModalSt;
    populateStandardExtraFields(modalStandardFieldsGrid, item);
    syncStandardMirrorSelects(modalStandardFieldsGrid);
  }

  editModal.classList.add('open');
}

document.getElementById('btnCancelModal').addEventListener('click', () => {
  editModal.classList.remove('open');
});

async function saveRequestToWorkbook(targetInfo, formValues) {
  const currentSheetName = targetInfo ? targetInfo.sheetName : sheetSelect.value;

  if (formValues && formValues.isCustomSchema && formValues.customValues) {
    const stColKey = activeSchemaConfig && activeSchemaConfig.statusCol ? `col_${activeSchemaConfig.statusCol}` : '';
    const hasAnyField = Object.entries(formValues.customValues).some(
      ([k, v]) => k !== stColKey && String(v || '').trim() !== ''
    );
    if (!hasAnyField) {
      showToast('Please enter at least one field before saving.', 'warning');
      return false;
    }
  }

  if (!pendingPrimarySaveState) {
    await loadWorkbookFromDisk(currentSheetName, { skipStagedRecoveryCheck: true });
  }
  const worksheet = workbook.getWorksheet(currentSheetName);
  if (!worksheet) {
    await showModalAlert({
      title: 'Worksheet Not Found',
      message: `Target worksheet "${currentSheetName}" was not found in the workbook.`
    });
    return false;
  }

  const stdSettings = (activeSchemaConfig && activeSchemaConfig.colSettings) || {};
  const shouldDefaultUpper = (colIdx) => !(stdSettings[colIdx] && stdSettings[colIdx].dataType === 'auto') &&
    !(stdSettings[colIdx] && ['number', 'currency', 'date', 'notes'].includes(stdSettings[colIdx].dataType));
  const subjectVal = String(formValues.subject || '').trim();
  const ticketVal = shouldDefaultUpper(2) ? String(formValues.ticket || '').trim().toUpperCase() : String(formValues.ticket || '').trim();
  const brokerVal = String(formValues.broker || '').trim();
  const masterVal = shouldDefaultUpper(4) ? String(formValues.master || '').trim().toUpperCase() : String(formValues.master || '').trim();
  const bidderVal = shouldDefaultUpper(5) ? String(formValues.bidder || '').trim().toUpperCase() : String(formValues.bidder || '').trim();
  const actionVal = String(formValues.action || '').trim();
  const yardVal = shouldDefaultUpper(7) ? String(formValues.yard || '').trim().toUpperCase() : String(formValues.yard || '').trim();
  const lotsVal = String(formValues.lots || '').trim();
  const notesVal = String(formValues.notes || '').trim();
  const statusVal = String(formValues.status || '').trim();

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
  writeItemToWorksheetRow(row, targetRowNum, {
    subject: subjectVal,
    ticket: ticketVal,
    broker: brokerVal,
    master: masterVal,
    bidder: bidderVal,
    action: actionVal,
    yard: yardVal,
    lots: lotsVal,
    notes: notesVal,
    status: statusVal,
    key: trackingKey,
    customValues: formValues.customValues || null
  }, colMap);

  if (targetInfo && colMap.statusCol && statusVal) {
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
            if (newerColMap.statusCol) {
              const nr = newerWs.getRow(inst.rowNumber);
              nr.getCell(newerColMap.statusCol).value = statusVal;
              nr.commit();
            }
          }
        });
      }
    }
  }

  const saveResult = await saveWorkbookToDisk();
  if (saveResult) {
    await loadWorkbookFromDisk(currentSheetName, {
      skipStagedRecoveryCheck: true,
      preserveStatus: saveResult === 'mirror-locked'
    });
    return true;
  }

  // If primary file was locked and staged in fail-safe cache, re-index in-memory workbook so UI shows the user's staged edit!
  if (pendingPrimarySaveState && workbook) {
    allRows = [];
    const dailyObjs = [];
    workbook.eachSheet(ws => {
      const raw = ws.name.trim();
      if (DATE_TAB_REGEX.test(raw)) dailyObjs.push({ ws, raw });
    });
    dailyObjs.sort((a, b) => parseDateTab(a.raw) - parseDateTab(b.raw));
    dailyObjs.forEach(({ ws, raw }) => indexWorksheetRows(ws, raw));
    renderSheetTabsBar();
    renderList();
  }
  return false;
}

document.getElementById('entryForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const payload = readFormPayloadFromScope('modal');
  const ok = await saveRequestToWorkbook(editingTarget, payload);
  if (ok || pendingPrimarySaveState) {
    editModal.classList.remove('open');
  }
});

if (dockEntryForm) {
  dockEntryForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    const wasEditing = Boolean(dockEditingTarget);
    const payload = readFormPayloadFromScope('dock');
    const ok = await saveRequestToWorkbook(dockEditingTarget, payload);
    if (ok) {
      if (!wasEditing) {
        clearSavedNewDockDraft();
        setBottomDockNewMode(false);
        showToast('New request added and saved.', 'success');
      } else {
        showToast('Request changes saved to workbook.', 'success');
      }
    } else if (pendingPrimarySaveState && !wasEditing) {
      clearSavedNewDockDraft();
      setBottomDockNewMode(false);
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
  if (ticketQuickNoteForm) {
    const hasNotesCol = activeSchemaConfig.isStandardSchema || Boolean(activeSchemaConfig.notesCol);
    ticketQuickNoteForm.style.display = hasNotesCol ? 'block' : 'none';
  }

  if (ticketDetailsMetaGrid) {
    if (!activeSchemaConfig.isStandardSchema && activeSchemaConfig.columns.length > 0) {
      const customCards = activeSchemaConfig.columns
        .filter(c => {
          if (c.colIndex === activeSchemaConfig.notesCol) return false;
          const v = (latest.customValues && latest.customValues[`col_${c.colIndex}`]) || '';
          return c.showInTable !== false || c.showInEditor !== false || Boolean(String(v).trim());
        })
        .map(c => {
          const val = (latest.customValues && latest.customValues[`col_${c.colIndex}`]) || '';
          const isCopyableField = Array.isArray(activeSchemaConfig.copyCols)
            ? activeSchemaConfig.copyCols.includes(c.colIndex)
            : (
                c.colIndex === activeSchemaConfig.ticketCol ||
                c.colIndex === activeSchemaConfig.secondaryCol ||
                /\b(subject|ticket|master\s*#?|account\s*#?|member\s*#?|bidder\s*#?|yard\s*#?|lot\s*#?)\b/i.test(c.excelHeader || c.header)
              );
          const inner = c.colIndex === activeSchemaConfig.statusCol
            ? `<span class="badge ${getStatusBadgeClass(val || latest.status)}">${escapeHtml(val || latest.status)}</span>`
            : (isCopyableField && val
                ? `<span class="cell-copy-group cell-copy-text"><span>${escapeHtml(val)}</span>${renderInlineCopyBtnHtml(val, c.header)}</span>`
                : escapeHtml(val || '—'));
          return `
            <div class="ticket-meta-card">
              <div class="ticket-meta-card-label">${escapeHtml(c.header)}</div>
              <div class="ticket-meta-card-val">${inner}</div>
            </div>
          `;
        })
        .join('');
      ticketDetailsMetaGrid.innerHTML = `
        <div class="ticket-meta-card">
          <div class="ticket-meta-card-label">Original Creation Tab</div>
          <div class="ticket-meta-card-val">${escapeHtml(oldest.sheetName)} (${history.length} tab${history.length === 1 ? '' : 's'})</div>
        </div>
        ${customCards}
      `;
    } else {
      const parentCategory = resolveActionParentCategory(latest.action);
      const isStdCopy = (colIdx) => Array.isArray(activeSchemaConfig && activeSchemaConfig.copyCols)
        ? activeSchemaConfig.copyCols.includes(colIdx)
        : [1, 2, 4, 5, 7, 8].includes(colIdx);
      const stdColHdr = (colIdx, fallback) => {
        const hit = activeSchemaConfig && Array.isArray(activeSchemaConfig.columns)
          ? activeSchemaConfig.columns.find(c => c.colIndex === colIdx)
          : null;
        return (hit && hit.header) ? hit.header : fallback;
      };
      const wrapStdCopy = (colIdx, val, lbl, mono = true) => {
        if (!val) return '—';
        if (!isStdCopy(colIdx)) return escapeHtml(val);
        const cls = mono ? 'cell-copy-group' : 'cell-copy-group cell-copy-text';
        return `<span class="${cls}"><span>${escapeHtml(val)}</span>${renderInlineCopyBtnHtml(val, lbl)}</span>`;
      };

      const subjectCardVal = wrapStdCopy(1, latest.subject, stdColHdr(1, 'Email Subject'), false);
      const ticketCardVal = wrapStdCopy(2, latest.ticket, stdColHdr(2, 'Ticket #'), true);
      const brokerCardVal = wrapStdCopy(3, latest.broker, stdColHdr(3, 'Broker Name'), false);
      const masterCardVal = wrapStdCopy(4, latest.master, stdColHdr(4, 'Master #'), true);
      const bidderCardVal = wrapStdCopy(5, latest.bidder, stdColHdr(5, 'Bidder #'), true);
      const actionCardVal = wrapStdCopy(6, latest.action, stdColHdr(6, 'Requested Action'), false);
      const yardCardVal = wrapStdCopy(7, latest.yard, stdColHdr(7, 'Yard #'), true);
      const lotsCardVal = wrapStdCopy(8, latest.lots, stdColHdr(8, 'Lot Numbers'), true);

      const coreStdSet = new Set([1, 2, 3, 4, 5, 6, 7, 8, activeSchemaConfig.notesCol || 9, activeSchemaConfig.statusCol || 10, activeSchemaConfig.trackingCol || 11]);
      const extraStdCardsHtml = (activeSchemaConfig && Array.isArray(activeSchemaConfig.columns))
        ? activeSchemaConfig.columns
            .filter(c => !coreStdSet.has(c.colIndex))
            .map(c => {
              const ev = (latest.customValues && latest.customValues[`col_${c.colIndex}`]) || '';
              if (c.showInTable === false && c.showInEditor === false && !String(ev).trim()) return '';
              const inner = ev ? wrapStdCopy(c.colIndex, ev, c.header, false) : '—';
              return `
                <div class="ticket-meta-card">
                  <div class="ticket-meta-card-label">${escapeHtml(c.header)}</div>
                  <div class="ticket-meta-card-val">${inner}</div>
                </div>
              `;
            }).join('')
        : '';

      ticketDetailsMetaGrid.innerHTML = `
        <div class="ticket-meta-card" style="grid-column: 1 / -1;">
          <div class="ticket-meta-card-label">${escapeHtml(stdColHdr(1, 'Email Subject'))}</div>
          <div class="ticket-meta-card-val">${subjectCardVal}</div>
        </div>
        <div class="ticket-meta-card">
          <div class="ticket-meta-card-label">${escapeHtml(stdColHdr(2, 'Ticket Number'))}</div>
          <div class="ticket-meta-card-val">${ticketCardVal}</div>
        </div>
        <div class="ticket-meta-card">
          <div class="ticket-meta-card-label">Current ${escapeHtml(stdColHdr(10, 'Status'))} (${escapeHtml(latest.sheetName)})</div>
          <div class="ticket-meta-card-val">
            <span class="badge ${getStatusBadgeClass(latest.status)}">${escapeHtml(latest.status)}</span>
          </div>
        </div>
        <div class="ticket-meta-card">
          <div class="ticket-meta-card-label">Original Creation Tab</div>
          <div class="ticket-meta-card-val">${escapeHtml(oldest.sheetName)} (${history.length} tab${history.length === 1 ? '' : 's'})</div>
        </div>
        <div class="ticket-meta-card">
          <div class="ticket-meta-card-label">${escapeHtml(stdColHdr(3, 'Broker'))} &amp; Accounts</div>
          <div class="ticket-meta-card-val">${brokerCardVal} (M: ${masterCardVal} / B: ${bidderCardVal})</div>
        </div>
        <div class="ticket-meta-card">
          <div class="ticket-meta-card-label">${escapeHtml(stdColHdr(6, 'Requested Action'))} &amp; Category</div>
          <div class="ticket-meta-card-val">${actionCardVal} <span class="badge badge-progress" style="margin-left:4px;">${escapeHtml(parentCategory)}</span></div>
        </div>
        <div class="ticket-meta-card">
          <div class="ticket-meta-card-label">${escapeHtml(stdColHdr(7, 'Yard #'))} &amp; ${escapeHtml(stdColHdr(8, 'Lot Numbers'))}</div>
          <div class="ticket-meta-card-val">Y: ${yardCardVal} • Lots: ${lotsCardVal}</div>
        </div>
        ${extraStdCardsHtml}
      `;
    }
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
            <span class="badge ${getStatusBadgeClass(entry.status)}">${escapeHtml(entry.status)}</span>
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

    const updatedCustomValues = latestEntry.customValues
      ? { ...latestEntry.customValues }
      : null;
    if (updatedCustomValues && activeSchemaConfig && activeSchemaConfig.notesCol) {
      updatedCustomValues[`col_${activeSchemaConfig.notesCol}`] = combinedNotes;
    }

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
        status: latestEntry.status,
        customValues: updatedCustomValues
      }
    );

    if (ok || pendingPrimarySaveState) {
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
      if (ok) {
        showToast(`Timestamped note appended to ${latestEntry.sheetName}.`, 'success');
      }
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
    viewOnlyState.isCustomSchema = Boolean(activeSchemaConfig && !activeSchemaConfig.isStandardSchema && activeSchemaConfig.columns.length > 0);
    viewOnlyState.customColumns = viewOnlyState.isCustomSchema ? [...activeSchemaConfig.columns] : [];

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
      if (viewOnlyState.isCustomSchema) {
        const defaultMirrorSt = getDefaultStatus();
        ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
          if (rowNumber === 1) return;
          const customValues = {};
          let hasAny = false;
          viewOnlyState.customColumns.forEach(c => {
            const rawV = cellText(row, c.colIndex);
            const v = (rawV && UUID_REGEX.test(rawV.trim())) ? '' : rawV;
            customValues[`col_${c.colIndex}`] = v;
            if (v !== '') hasAny = true;
          });
          if (!hasAny) return;
          const rawSt = activeSchemaConfig.statusCol
            ? (customValues[`col_${activeSchemaConfig.statusCol}`] || '')
            : '';
          const status = activeSchemaConfig.statusCol
            ? ((rawSt && !UUID_REGEX.test(rawSt.trim())) ? rawSt : defaultMirrorSt)
            : 'Open';
          if (activeSchemaConfig.statusCol) {
            customValues[`col_${activeSchemaConfig.statusCol}`] = status;
          }
          viewOnlyState.rows.push({
            sheetName: sName,
            rowNumber,
            customValues,
            status
          });
        });
        return;
      }

      const colMap = getStandardWorksheetColumnMapOnly(ws);
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

    if (viewOnlyHead) {
      if (viewOnlyState.isCustomSchema) {
        viewOnlyHead.innerHTML = `
          <tr>
            <th>Date Tab</th>
            ${viewOnlyState.customColumns.map(c => `<th>${escapeHtml(c.header)}</th>`).join('')}
          </tr>
        `;
      } else {
        viewOnlyHead.innerHTML = `
          <tr>
            <th>Date Tab</th>
            <th>Subject</th>
            <th>Ticket #</th>
            <th>Broker</th>
            <th>Master / Bidder</th>
            <th>Action</th>
            <th>Yard / Lots</th>
            <th>Action Notes</th>
            <th>Status</th>
          </tr>
        `;
      }
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
  const isCustom = Boolean(viewOnlyState.isCustomSchema && Array.isArray(viewOnlyState.customColumns));
  const colSpan = isCustom ? (1 + viewOnlyState.customColumns.length) : 9;

  const filtered = viewOnlyState.rows.filter(item => {
    if (query.length > 0) {
      const hay = isCustom
        ? [item.sheetName, ...Object.values(item.customValues || {})].join(' ').toLowerCase()
        : [
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
    viewOnlyBody.innerHTML = `<tr><td colspan="${colSpan}" class="empty-table-cell">No matching requests in View-Only snapshot.</td></tr>`;
    return;
  }

  filtered.forEach(item => {
    const tr = document.createElement('tr');
    if (isCustom) {
      const cells = viewOnlyState.customColumns.map(c => {
        const val = (item.customValues && item.customValues[`col_${c.colIndex}`]) || '';
        if (c.colIndex === activeSchemaConfig.statusCol) {
          return `<td><span class="badge ${getStatusBadgeClass(val || item.status)}">${escapeHtml(val || item.status)}</span></td>`;
        }
        return `<td class="cell-sm" style="max-width:260px; white-space:pre-wrap;">${escapeHtml(val || '—')}</td>`;
      }).join('');
      tr.innerHTML = `<td class="cell-mono-sm" style="white-space:nowrap;">${escapeHtml(item.sheetName)}</td>${cells}`;
      viewOnlyBody.appendChild(tr);
      return;
    }

    tr.innerHTML = `
      <td class="cell-mono-sm" style="white-space:nowrap;">${escapeHtml(item.sheetName)}</td>
      <td>${escapeHtml(item.subject || '—')}</td>
      <td class="cell-mono-sm">${escapeHtml(item.ticket || '—')}</td>
      <td>${escapeHtml(item.broker || '—')}</td>
      <td class="cell-mono-sm">M: ${escapeHtml(item.master || '—')}<br>B: ${escapeHtml(item.bidder || '—')}</td>
      <td>${escapeHtml(item.action || '—')}</td>
      <td class="cell-sm">Y: ${escapeHtml(item.yard || '—')}<br>L: ${escapeHtml(item.lots || '—')}</td>
      <td class="cell-sm" style="max-width:260px; white-space:pre-wrap;">${escapeHtml(item.notes || '—')}</td>
      <td><span class="badge ${getStatusBadgeClass(item.status)}">${escapeHtml(item.status)}</span></td>
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
