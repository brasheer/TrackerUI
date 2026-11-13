// ============================================================================
// rollover.js — Unique Request Identity Matching, Chronological Notes Merging,
//               Smart Multi-Sheet / Previous-Sheet Rollover & Startup Init
// ============================================================================

// --- UNIQUE REQUEST IDENTITY & CHRONOLOGICAL ACTION NOTES MERGING ---
const IGNORED_TICKET_PLACEHOLDERS = new Set([
  'n/a', 'na', 'none', 'null', 'tbd', 'pending', '-', '—', '--', '0', 'unknown'
]);

function normalizeTicketId(ticket) {
  const clean = String(ticket || '').trim().toLowerCase().replace(/^#+\s*/, '');
  if (!clean || IGNORED_TICKET_PLACEHOLDERS.has(clean)) return '';
  return clean;
}

function normalizeMatchText(str) {
  return String(str || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function normalizeLotsSignature(lots) {
  if (!lots) return '';
  const tokens = String(lots)
    .split(/[,;/|\s]+/)
    .map(s => s.trim().toLowerCase())
    .filter(s => s.length > 0 && s !== '—')
    .sort();
  return tokens.join(',');
}

function normalizeDateTabString(tabName) {
  const parts = String(tabName || '').trim().split('-').map(s => s.trim());
  if (parts.length !== 3) return String(tabName || '').trim();
  let [m, d, y] = parts;
  if (y.length === 2) y = '20' + y;
  return `${m.padStart(2, '0')}-${d.padStart(2, '0')}-${y}`;
}

/**
 * Determines whether two rows across daily sheets represent the same unique
 * support request—even if an older sheet had no Column K Tracking ID or received
 * a separate backfilled UUID before rollover.
 */
function areRowsSameUniqueRequest(rowA, rowB) {
  if (!rowA || !rowB) return false;

  const keyA = String(rowA.key || '').trim().toLowerCase();
  const keyB = String(rowB.key || '').trim().toLowerCase();
  if (keyA && keyB && keyA === keyB) return true;

  const tokensA = typeof splitNormalizedTicketTokens === 'function'
    ? splitNormalizedTicketTokens(rowA.ticket)
    : (normalizeTicketId(rowA.ticket) ? [normalizeTicketId(rowA.ticket).toUpperCase()] : []);
  const tokensB = typeof splitNormalizedTicketTokens === 'function'
    ? splitNormalizedTicketTokens(rowB.ticket)
    : (normalizeTicketId(rowB.ticket) ? [normalizeTicketId(rowB.ticket).toUpperCase()] : []);
  const hasSharedTicket = tokensA.length > 0 && tokensB.length > 0 && tokensA.some(t => tokensB.includes(t));

  // Explicit, non-overlapping Ticket Numbers indicate distinct requests
  if (tokensA.length > 0 && tokensB.length > 0 && !hasSharedTicket) {
    return false;
  }

  // Custom Department Schema identity matching
  if (rowA.isCustomSchema || rowB.isCustomSchema || (activeSchemaConfig && !activeSchemaConfig.isStandardSchema)) {
    if (hasSharedTicket) return true;
    const cols = (activeSchemaConfig && Array.isArray(activeSchemaConfig.columns)) ? activeSchemaConfig.columns : [];
    const cvA = rowA.customValues || {};
    const cvB = rowB.customValues || {};
    const pCol = activeSchemaConfig ? activeSchemaConfig.primaryCol : 1;
    const pA = normalizeMatchText(cvA[`col_${pCol}`] || rowA.subject);
    const pB = normalizeMatchText(cvB[`col_${pCol}`] || rowB.subject);
    if ((pA || pB) && pA !== pB) return false;

    let populatedMatches = 0;
    for (const c of cols) {
      if (c.colIndex === activeSchemaConfig.statusCol || c.colIndex === activeSchemaConfig.notesCol || c.colIndex === activeSchemaConfig.ticketCol) {
        continue;
      }
      const rawA = String(cvA[`col_${c.colIndex}`] || '').trim();
      const rawB = String(cvB[`col_${c.colIndex}`] || '').trim();
      const vA = UUID_REGEX.test(rawA) ? '' : normalizeMatchText(rawA);
      const vB = UUID_REGEX.test(rawB) ? '' : normalizeMatchText(rawB);
      if (vA && vB && vA !== vB) return false;
      if (vA && vB && vA === vB) populatedMatches++;
    }
    return populatedMatches >= 2 || (populatedMatches >= 1 && cols.length <= 3);
  }

  const subjA = normalizeMatchText(rowA.subject);
  const subjB = normalizeMatchText(rowB.subject);
  const masterA = normalizeMatchText(rowA.master);
  const masterB = normalizeMatchText(rowB.master);
  const brokerA = normalizeMatchText(rowA.broker);
  const brokerB = normalizeMatchText(rowB.broker);
  const acctA = masterA || brokerA;
  const acctB = masterB || brokerB;
  const bidderA = normalizeMatchText(rowA.bidder);
  const bidderB = normalizeMatchText(rowB.bidder);
  const actionA = normalizeMatchText(rowA.action);
  const actionB = normalizeMatchText(rowB.action);
  const yardA = normalizeMatchText(rowA.yard);
  const yardB = normalizeMatchText(rowB.yard);
  const lotsA = normalizeLotsSignature(rowA.lots);
  const lotsB = normalizeLotsSignature(rowB.lots);

  const accountsMatch =
    (masterA && masterB && masterA === masterB) ||
    (brokerA && brokerB && brokerA === brokerB) ||
    (!acctA && !acctB);

  // Case A: Shared valid Ticket Number + compatible Lot Numbers & Account
  if (hasSharedTicket) {
    const lotsCompatible = (lotsA === lotsB) || !lotsA || !lotsB;
    if (lotsCompatible && (accountsMatch || !acctA || !acctB)) {
      return true;
    }
  }

  // Case B: Exact core field match (Subject + Account + Bidder + Action + Yard + Lots)
  const populatedCount = [subjA, acctA, bidderA, actionA, lotsA].filter(Boolean).length;
  if (
    populatedCount >= 2 &&
    subjA === subjB &&
    accountsMatch &&
    bidderA === bidderB &&
    actionA === actionB &&
    yardA === yardB &&
    lotsA === lotsB
  ) {
    return true;
  }

  // Case C: Same Subject + same Account + same non-empty Lot Numbers
  if (
    subjA && subjA === subjB &&
    acctA && acctB && accountsMatch &&
    lotsA && lotsA === lotsB &&
    (!bidderA || !bidderB || bidderA === bidderB)
  ) {
    return true;
  }

  return false;
}

function parseTimestampTaggedNoteLine(rawLine) {
  const trimmed = String(rawLine || '').trim();
  if (!trimmed || trimmed === '—') return null;

  const match = trimmed.match(/^\[(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}(?:\s+[^\]]*)?)\]\s*(.*)$/);
  if (match) {
    const rawStamp = match[1].trim();
    const coreText = match[2].trim();
    if (!coreText) return null;
    const datePartMatch = rawStamp.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})(.*)$/);
    let normDate = rawStamp.toLowerCase();
    let hasTime = false;
    if (datePartMatch) {
      let [, m, d, y, rest] = datePartMatch;
      if (y.length === 2) y = '20' + y;
      normDate = `${m.padStart(2, '0')}-${d.padStart(2, '0')}-${y}`;
      hasTime = Boolean(rest && rest.trim().length > 0);
    }
    return {
      hasExplicitStamp: true,
      rawStamp,
      normDate,
      hasTime,
      coreText,
      normCore: normalizeMatchText(coreText)
    };
  }

  return {
    hasExplicitStamp: false,
    rawStamp: null,
    normDate: null,
    hasTime: false,
    coreText: trimmed,
    normCore: normalizeMatchText(trimmed)
  };
}

/**
 * Chronologically appends and deduplicates Action Notes across all daily entries
 * for a unique request so rolling over carries forward every day's notes without
 * duplicating lines already carried over on intermediate days.
 */
function mergeChronologicalActionNotes(entries) {
  if (!Array.isArray(entries) || entries.length === 0) return '';

  const sortedEntries = [...entries].sort((a, b) => {
    const diff = parseDateTab(a.sheetName) - parseDateTab(b.sheetName);
    return diff !== 0 ? diff : (a.rowNumber || 0) - (b.rowNumber || 0);
  });

  const collected = []; // [{ normDate, hasTime, normCore, formatted }]

  sortedEntries.forEach(entry => {
    const rawNotes = String(entry.notes || '').trim();
    if (!rawNotes || rawNotes === '—') return;

    const entryNormDate = normalizeDateTabString(entry.sheetName);
    const lines = rawNotes.split(/\r?\n/);

    lines.forEach(line => {
      const parsed = parseTimestampTaggedNoteLine(line);
      if (!parsed) return;

      const effectiveDate = parsed.normDate || entryNormDate;
      const formattedLine = parsed.hasExplicitStamp
        ? `[${parsed.rawStamp}] ${parsed.coreText}`
        : (effectiveDate ? `[${effectiveDate}] ${parsed.coreText}` : parsed.coreText);

      const existingIdx = collected.findIndex(item => {
        if (item.normCore !== parsed.normCore) return false;
        if (!parsed.hasExplicitStamp) return true;
        if (item.normDate === effectiveDate) return true;
        return false;
      });

      if (existingIdx !== -1) {
        if (parsed.hasTime && !collected[existingIdx].hasTime) {
          collected[existingIdx].hasTime = true;
          collected[existingIdx].formatted = formattedLine;
        }
        return;
      }

      collected.push({
        normDate: effectiveDate,
        hasTime: parsed.hasTime,
        normCore: parsed.normCore,
        formatted: formattedLine
      });
    });
  });

  return collected.map(c => c.formatted).join('\n');
}

function extractNoteCoresSet(rawNotes) {
  const set = new Set();
  const trimmed = String(rawNotes || '').trim();
  if (!trimmed || trimmed === '—') return set;
  trimmed.split(/\r?\n/).forEach(line => {
    const parsed = parseTimestampTaggedNoteLine(line);
    if (parsed && parsed.normCore) {
      set.add(parsed.normCore);
    }
  });
  return set;
}

// --- 7. SMART MULTI-SHEET & IDEMPOTENT ROLLOVER ENGINE WITH FUTURE DATE HANDLING ---
btnRollover.addEventListener('click', async () => {
  if (!workbook) return;
  try {
    await loadWorkbookFromDisk(sheetSelect.value);
    openRolloverModal();
  } catch (err) {
    console.error('Error preparing rollover modal:', err);
    await showModalAlert({
      title: 'Rollover Read Error',
      message: 'Could not read workbook for rollover: ' + err.message,
      variant: 'danger'
    });
  }
});

btnCancelRollover.addEventListener('click', () => {
  rolloverModal.classList.remove('open');
});

function openRolloverModal(options = {}) {
  const { preferredSourceMode = null, preferredTargetDate = null } = options || {};
  rolloverSourceMode = preferredSourceMode === 'all' ? 'all' : (preferredSourceMode === 'previous' ? 'previous' : 'previous');

  const todayStr = getTodayFormatted();
  const todayTime = parseDateTab(todayStr);
  rolloverTodayBadge.textContent = `Today: ${todayStr}`;

  const latestSheet = dateSheets.length > 0 ? dateSheets[0] : null;
  const latestTime = latestSheet ? parseDateTab(latestSheet) : 0;
  const hasFutureSheet = latestSheet && latestTime > todayTime;

  rolloverPresetBtns.innerHTML = '';
  const presets = [];

  if (hasFutureSheet) {
    const latestCount = allRows.filter(r => r.sheetName === latestSheet).length;
    const nextAfterLatest = addDaysToDateTab(latestSheet, 1);

    rolloverFutureBanner.style.display = 'block';
    rolloverFutureBanner.innerHTML = `
      <strong>Future-Dated Tab Detected:</strong> Your workbook's latest sheet is
      <code>${escapeHtml(latestSheet)}</code> (${latestCount} entry${latestCount === 1 ? '' : 'ies'}),
      which is ahead of today's date (<code>${escapeHtml(todayStr)}</code>).
      Select whether to sync into <strong>${escapeHtml(latestSheet)}</strong>, roll forward to
      <strong>${escapeHtml(nextAfterLatest)}</strong>, or sync into <strong>${escapeHtml(todayStr)}</strong>.
    `;

    presets.push({ label: `Latest Tab (${latestSheet})`, date: latestSheet });
    presets.push({ label: `Next Day (${nextAfterLatest})`, date: nextAfterLatest });
    presets.push({ label: `Today (${todayStr})`, date: todayStr });

    rolloverTargetDate.value = preferredTargetDate || latestSheet;
  } else {
    rolloverFutureBanner.style.display = 'none';
    presets.push({ label: `Today (${todayStr})`, date: todayStr });
    const nextDay = addDaysToDateTab(latestSheet || todayStr, 1);
    if (nextDay !== todayStr) {
      presets.push({ label: `Tomorrow (${nextDay})`, date: nextDay });
    }
    if (latestSheet && latestSheet !== todayStr) {
      presets.push({ label: `Latest Tab (${latestSheet})`, date: latestSheet });
    }
    rolloverTargetDate.value = preferredTargetDate || todayStr;
  }

  if (preferredTargetDate && !presets.some(p => p.date === preferredTargetDate)) {
    presets.unshift({ label: `Selected Tab (${preferredTargetDate})`, date: preferredTargetDate });
  }

  presets.forEach(p => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'btn-preset';
    btn.textContent = p.label;
    btn.dataset.date = p.date;
    btn.addEventListener('click', () => {
      rolloverTargetDate.value = p.date;
      updateRolloverModalPreview();
    });
    rolloverPresetBtns.appendChild(btn);
  });

  updateRolloverModalPreview();
  rolloverModal.classList.add('open');
}

rolloverTargetDate.addEventListener('input', updateRolloverModalPreview);
chkIncludeFutureTabs.addEventListener('change', updateRolloverModalPreview);
if (rolloverSourceGroup) {
  rolloverSourceGroup.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-rollover-source]');
    if (!btn || btn.disabled) return;
    const nextMode = btn.dataset.rolloverSource;
    if (nextMode === 'previous' || nextMode === 'all') {
      rolloverSourceMode = nextMode;
      updateRolloverModalPreview();
    }
  });
}

function buildConsolidatedRolloverItem(entries, baseRow) {
  const mergedNotes = mergeChronologicalActionNotes(entries);
  const canonicalRow = entries[entries.length - 1] || baseRow;
  const pickLatestIfUnkeyedMerge = (primaryVal, fieldName) => {
    if (primaryVal && String(primaryVal).trim() !== '') return primaryVal;
    for (let i = entries.length - 1; i >= 0; i--) {
      const candidate = entries[i][fieldName];
      if (candidate && String(candidate).trim() !== '') return candidate;
    }
    return primaryVal;
  };

  // Merge ticket tokens comma-delimited across daily entries without duplicates
  const mergedTicket = typeof mergeCommaDelimitedTickets === 'function'
    ? mergeCommaDelimitedTickets(
        baseRow.ticket || '',
        entries.map(e => e.ticket || '')
      ).merged
    : pickLatestIfUnkeyedMerge(baseRow.ticket, 'ticket');

  let consolidatedCustomValues = undefined;
  let resolvedCustomStatus = baseRow.status && !UUID_REGEX.test(String(baseRow.status).trim())
    ? baseRow.status
    : '';
  if (!resolvedCustomStatus) {
    for (let i = entries.length - 1; i >= 0; i--) {
      const candSt = entries[i].status;
      if (candSt && String(candSt).trim() !== '' && !UUID_REGEX.test(String(candSt).trim())) {
        resolvedCustomStatus = candSt;
        break;
      }
    }
  }
  if (!resolvedCustomStatus) {
    resolvedCustomStatus = getDefaultStatus();
  }

  if (baseRow.customValues || (activeSchemaConfig && !activeSchemaConfig.isStandardSchema)) {
    consolidatedCustomValues = { ...(baseRow.customValues || {}) };
    const cols = (activeSchemaConfig && Array.isArray(activeSchemaConfig.columns)) ? activeSchemaConfig.columns : [];
    cols.forEach(c => {
      const k = `col_${c.colIndex}`;
      if (consolidatedCustomValues[k] && UUID_REGEX.test(String(consolidatedCustomValues[k]).trim())) {
        consolidatedCustomValues[k] = '';
      }
      if (c.colIndex === activeSchemaConfig.notesCol) {
        consolidatedCustomValues[k] = mergedNotes;
        return;
      }
      if (c.colIndex === activeSchemaConfig.ticketCol) {
        consolidatedCustomValues[k] = mergedTicket;
        return;
      }
      if (c.colIndex === activeSchemaConfig.statusCol) {
        consolidatedCustomValues[k] = resolvedCustomStatus;
        return;
      }
      if (!consolidatedCustomValues[k] || String(consolidatedCustomValues[k]).trim() === '') {
        for (let i = entries.length - 1; i >= 0; i--) {
          const cand = entries[i].customValues && entries[i].customValues[k];
          if (cand && String(cand).trim() !== '' && !UUID_REGEX.test(String(cand).trim())) {
            consolidatedCustomValues[k] = cand;
            break;
          }
        }
      }
    });
  }

  return {
    ...baseRow,
    key: canonicalRow.key || baseRow.key,
    subject: pickLatestIfUnkeyedMerge(baseRow.subject, 'subject'),
    ticket: mergedTicket,
    broker: pickLatestIfUnkeyedMerge(baseRow.broker, 'broker'),
    master: pickLatestIfUnkeyedMerge(baseRow.master, 'master'),
    bidder: pickLatestIfUnkeyedMerge(baseRow.bidder, 'bidder'),
    action: pickLatestIfUnkeyedMerge(baseRow.action, 'action'),
    yard: pickLatestIfUnkeyedMerge(baseRow.yard, 'yard'),
    lots: pickLatestIfUnkeyedMerge(baseRow.lots, 'lots'),
    notes: mergedNotes,
    status: resolvedCustomStatus,
    ...(consolidatedCustomValues ? { customValues: consolidatedCustomValues } : {}),
    _entries: entries,
    _sourceTabCount: entries.length,
    _firstSheet: entries[0] ? entries[0].sheetName : baseRow.sheetName
  };
}

function getRolloverItemDisplayTitle(item) {
  if (!item) return '(Untitled Request)';
  if (item.subject && String(item.subject).trim()) return String(item.subject).trim();
  if (item.ticket && String(item.ticket).trim()) return `Ticket #${String(item.ticket).trim()}`;
  if (item.customValues && activeSchemaConfig && Array.isArray(activeSchemaConfig.columns)) {
    for (const c of activeSchemaConfig.columns) {
      if (c.colIndex === activeSchemaConfig.statusCol || c.colIndex === activeSchemaConfig.notesCol) continue;
      const v = String(item.customValues[`col_${c.colIndex}`] || '').trim();
      if (v && !UUID_REGEX.test(v)) return `${c.header}: ${v}`;
    }
  }
  return `Row ${item.rowNumber || '?'} (${item.sheetName || ''})`;
}

function unifyHistoricalGroupTrackingIds(entries, canonicalKey) {
  if (!Array.isArray(entries) || !canonicalKey || !workbook) return;
  const normCanonical = canonicalKey.trim().toLowerCase();

  entries.forEach(histRow => {
    if (!histRow || !histRow.sheetName || !histRow.rowNumber) return;

    const ws = workbook.getWorksheet(histRow.sheetName);
    if (!ws) return;
    const colMap = getWorksheetColumnMap(ws);
    const excelRow = ws.getRow(histRow.rowNumber);
    let mutated = false;

    // Auto-heal any earlier sheet where Status column held the retroactive UUIDv4
    if (colMap.isCustomSchema && colMap.statusCol) {
      const curStCell = cellText(excelRow, colMap.statusCol);
      if (curStCell && UUID_REGEX.test(curStCell)) {
        const cleanHistSt = (histRow.status && !UUID_REGEX.test(String(histRow.status).trim()))
          ? histRow.status
          : getDefaultStatus();
        excelRow.getCell(colMap.statusCol).value = cleanHistSt;
        mutated = true;
      }
    }

    if (!histRow.key || histRow.key.trim().toLowerCase() !== normCanonical || cellText(excelRow, colMap.trackingCol).trim().toLowerCase() !== normCanonical) {
      // Ensure no other distinct row on the same sheet already holds canonicalKey
      let collision = false;
      ws.eachRow({ includeEmpty: false }, (r, rNum) => {
        if (rNum === 1 || rNum === histRow.rowNumber) return;
        if (cellText(r, colMap.trackingCol).trim().toLowerCase() === normCanonical) {
          collision = true;
        }
      });
      if (!collision) {
        excelRow.getCell(colMap.trackingCol).value = canonicalKey;
        mutated = true;
      }
    }

    if (mutated) {
      excelRow.commit();
    }
  });
}

function computeRolloverAnalysis(targetTabName, includeFutureTabs, sourceMode = rolloverSourceMode) {
  if (!DATE_TAB_REGEX.test(targetTabName)) {
    return { valid: false };
  }

  const effectiveSourceMode = sourceMode === 'all' ? 'all' : 'previous';
  const targetTime = parseDateTab(targetTabName);
  const targetExists = Boolean(workbook.getWorksheet(targetTabName));

  const priorSheets = dateSheets
    .filter(name => parseDateTab(name) < targetTime)
    .sort((a, b) => parseDateTab(a) - parseDateTab(b));

  const immediatePrevSheet = priorSheets.length > 0 ? priorSheets[priorSheets.length - 1] : null;

  const futureSheets = dateSheets
    .filter(name => parseDateTab(name) > targetTime)
    .sort((a, b) => parseDateTab(a) - parseDateTab(b));

  const sourceSheets = includeFutureTabs
    ? [...priorSheets, ...futureSheets]
    : priorSheets;

  const allowedActiveSourceSheets = effectiveSourceMode === 'previous'
    ? new Set([
        ...(immediatePrevSheet ? [immediatePrevSheet] : []),
        ...(includeFutureTabs ? futureSheets : [])
      ])
    : null;

  // Consolidate each unique request across all source daily tabs chronologically
  // so only ONE consolidated item per unique request rolls forward with all prior notes appended.
  const requestGroups = [];
  let nextGroupId = 1;

  sourceSheets.forEach(sheetName => {
    const sheetRows = allRows
      .filter(r => r.sheetName === sheetName)
      .sort((a, b) => a.rowNumber - b.rowNumber);

    const matchedGroupIdsOnSheet = new Set();

    sheetRows.forEach(row => {
      const normKey = row.key.trim().toLowerCase();

      // 1. Exact Tracking ID match across sheets
      let matchedGroup = requestGroups.find(g =>
        g.entries.some(e => e.key.trim().toLowerCase() === normKey)
      );

      // 2. Fallback unique-request identity match across different days (1-to-1 per sheet)
      if (!matchedGroup) {
        for (let i = requestGroups.length - 1; i >= 0; i--) {
          const candidate = requestGroups[i];
          if (matchedGroupIdsOnSheet.has(candidate.id)) continue;
          if (
            areRowsSameUniqueRequest(candidate.latestRow, row) ||
            areRowsSameUniqueRequest(candidate.entries[0], row)
          ) {
            matchedGroup = candidate;
            break;
          }
        }
      }

      if (matchedGroup) {
        matchedGroupIdsOnSheet.add(matchedGroup.id);
        matchedGroup.entries.push(row);
        matchedGroup.latestRow = row;
        if (!isCompletedStatus(row.status)) {
          matchedGroup.lastOpenRow = row;
        }
      } else {
        const newGroup = {
          id: nextGroupId++,
          entries: [row],
          latestRow: row,
          lastOpenRow: isCompletedStatus(row.status) ? null : row
        };
        matchedGroupIdsOnSheet.add(newGroup.id);
        requestGroups.push(newGroup);
      }
    });
  });

  const existingTargetRows = allRows
    .filter(r => r.sheetName === targetTabName)
    .sort((a, b) => a.rowNumber - b.rowNumber);

  const matchedTargetRowNums = new Set();

  const toRollOver = [];
  const alreadyOnTarget = [];
  const supersededByComplete = [];

  requestGroups.forEach(group => {
    const latestRow = group.latestRow;
    const isLatestComplete = isCompletedStatus(latestRow.status);
    const consolidatedLatest = buildConsolidatedRolloverItem(group.entries, latestRow);
    const inSelectedSourceScope = !allowedActiveSourceSheets ||
      group.entries.some(e => allowedActiveSourceSheets.has(e.sheetName));

    // Check whether this unique request is already present on targetTabName
    let targetRow = existingTargetRows.find(t =>
      group.entries.some(e => e.key.trim().toLowerCase() === t.key.trim().toLowerCase())
    );
    if (!targetRow) {
      targetRow = existingTargetRows.find(t =>
        !matchedTargetRowNums.has(t.rowNumber) &&
        (areRowsSameUniqueRequest(latestRow, t) || areRowsSameUniqueRequest(group.entries[0], t))
      );
    }

    if (targetRow) {
      matchedTargetRowNums.add(targetRow.rowNumber);
      const isTargetComplete = isCompletedStatus(targetRow.status);
      const priorCores = extractNoteCoresSet(consolidatedLatest.notes);
      const targetCores = extractNoteCoresSet(targetRow.notes);
      let hasMissingPriorNotes = false;
      for (const core of priorCores) {
        if (!targetCores.has(core)) {
          hasMissingPriorNotes = true;
          break;
        }
      }
      const mergedTargetNotes = hasMissingPriorNotes
        ? mergeChronologicalActionNotes([...group.entries, targetRow])
        : targetRow.notes;

      alreadyOnTarget.push({
        latestRow: consolidatedLatest,
        targetRow,
        reopenedOnSource: inSelectedSourceScope && isTargetComplete && !isLatestComplete,
        hasMissingPriorNotes: inSelectedSourceScope && hasMissingPriorNotes,
        mergedTargetNotes: inSelectedSourceScope ? mergedTargetNotes : targetRow.notes,
        groupEntries: group.entries
      });
    } else if (!inSelectedSourceScope) {
      // Excluded when Source Sheet Scope is "Previous Sheet Only" and request is not on immediatePrevSheet
      return;
    } else if (!isLatestComplete) {
      toRollOver.push(consolidatedLatest);
    } else if (group.lastOpenRow) {
      const canOfferReopen = effectiveSourceMode === 'all' || (
        immediatePrevSheet &&
        group.entries.some(e => e.sheetName === immediatePrevSheet && !isCompletedStatus(e.status))
      );
      if (canOfferReopen) {
        const consolidatedOpen = buildConsolidatedRolloverItem(group.entries, group.lastOpenRow);
        supersededByComplete.push({
          openRow: consolidatedOpen,
          completedRow: latestRow,
          groupEntries: group.entries
        });
      }
    }
  });

  existingTargetRows.forEach(targetRow => {
    if (!matchedTargetRowNums.has(targetRow.rowNumber)) {
      alreadyOnTarget.push({
        latestRow: targetRow,
        targetRow,
        reopenedOnSource: false,
        hasMissingPriorNotes: false,
        mergedTargetNotes: targetRow.notes,
        groupEntries: []
      });
    }
  });

  return {
    valid: true,
    targetTabName,
    targetExists,
    priorSheets,
    immediatePrevSheet,
    sourceMode: effectiveSourceMode,
    futureSheets,
    toRollOver,
    alreadyOnTarget,
    supersededByComplete
  };
}

function updateRolloverModalPreview() {
  const targetVal = rolloverTargetDate.value.trim();

  Array.from(rolloverPresetBtns.children).forEach(btn => {
    btn.classList.toggle('active', btn.dataset.date === targetVal);
  });

  if (!DATE_TAB_REGEX.test(targetVal)) {
    rolloverFutureOptionWrap.style.display = 'none';
    if (rolloverSourceWrap) rolloverSourceWrap.style.display = 'none';
    rolloverPreviewBox.innerHTML = `<div style="color:var(--warning-text);">Please enter a valid target date in exact <code>MM-DD-YYYY</code> format.</div>`;
    return;
  }

  const targetTime = parseDateTab(targetVal);
  const priorTabs = dateSheets
    .filter(name => parseDateTab(name) < targetTime)
    .sort((a, b) => parseDateTab(a) - parseDateTab(b));
  const immediatePrev = priorTabs.length > 0 ? priorTabs[priorTabs.length - 1] : null;

  if (rolloverSourceWrap) {
    rolloverSourceWrap.style.display = 'flex';
  }
  if (btnRolloverSourcePrev) {
    btnRolloverSourcePrev.textContent = immediatePrev
      ? `Previous Sheet Only (${immediatePrev})`
      : 'Previous Sheet Only (None)';
    btnRolloverSourcePrev.classList.toggle('active', rolloverSourceMode === 'previous');
    btnRolloverSourcePrev.setAttribute('aria-pressed', String(rolloverSourceMode === 'previous'));
  }
  if (btnRolloverSourceAll) {
    btnRolloverSourceAll.textContent = `All Prior Sheets (${priorTabs.length})`;
    btnRolloverSourceAll.classList.toggle('active', rolloverSourceMode === 'all');
    btnRolloverSourceAll.setAttribute('aria-pressed', String(rolloverSourceMode === 'all'));
  }
  if (rolloverSourceDesc) {
    if (rolloverSourceMode === 'previous') {
      rolloverSourceDesc.textContent = immediatePrev
        ? `Carries forward unfinished requests from ${immediatePrev} only (Complete / Completed / Closed items are excluded) while preserving historical notes & Tracking IDs.`
        : `No daily sheet exists prior to ${targetVal}. Create or select a later date, or include future-dated tabs below.`;
    } else {
      rolloverSourceDesc.textContent = `Scans all ${priorTabs.length} daily tab(s) prior to ${targetVal} chronologically so older unfinished requests are also carried forward (Complete / Completed / Closed items are excluded).`;
    }
  }

  const futureTabs = dateSheets.filter(name => parseDateTab(name) > targetTime);
  if (futureTabs.length > 0) {
    rolloverFutureOptionWrap.style.display = 'block';
    chkIncludeFutureText.textContent = `Also include unfinished requests from future-dated tab(s): ${futureTabs.join(', ')}`;
  } else {
    rolloverFutureOptionWrap.style.display = 'none';
  }

  const includeFuture = futureTabs.length > 0 && chkIncludeFutureTabs.checked;
  const plan = computeRolloverAnalysis(targetVal, includeFuture, rolloverSourceMode);

  const tabStateBadge = plan.targetExists
    ? `<span class="badge badge-complete">Existing Tab (${plan.alreadyOnTarget.length} request${plan.alreadyOnTarget.length === 1 ? '' : 's'} currently on sheet)</span>`
    : `<span class="badge badge-pending">New Tab (Will be created &amp; cloned from template)</span>`;

  const sourceSummaryLabel = plan.sourceMode === 'previous'
    ? (plan.immediatePrevSheet ? `Source: <code>${escapeHtml(plan.immediatePrevSheet)}</code>` : 'Source: None prior')
    : `Source: All ${plan.priorSheets.length} prior sheet(s)`;

  let html = `
    <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px; padding-bottom:8px; border-bottom:1px solid var(--border);">
      <div><strong>Target:</strong> <code>${escapeHtml(targetVal)}</code> <span style="color:var(--text-muted); font-size:0.75rem; margin-left:6px;">• ${sourceSummaryLabel}</span></div>
      ${tabStateBadge}
    </div>
  `;

  html += `<div>
    <div class="preview-section-title">Unfinished Unique Requests to Roll Over (${plan.toRollOver.length})</div>`;
  if (plan.toRollOver.length === 0) {
    const emptyMsg = plan.sourceMode === 'previous' && plan.immediatePrevSheet
      ? `No unfinished requests found on ${escapeHtml(plan.immediatePrevSheet)} that aren't already on ${escapeHtml(targetVal)}.`
      : 'No new unfinished requests found across source sheets.';
    html += `<div style="color:var(--text-muted); font-size:0.78125rem;">${emptyMsg}</div>`;
  } else {
    html += `<div class="preview-list">`;
    plan.toRollOver.forEach(item => {
      const spanInfo = item._sourceTabCount > 1
        ? `From ${escapeHtml(item._firstSheet)} → ${escapeHtml(item.sheetName)} (${item._sourceTabCount} days consolidated)`
        : `From ${escapeHtml(item.sheetName)}`;
      const noteLineCount = item.notes ? item.notes.split(/\r?\n/).filter(Boolean).length : 0;
      const notesMeta = noteLineCount > 0 ? ` • ${noteLineCount} note line${noteLineCount === 1 ? '' : 's'} appended` : '';
      html += `
        <div class="preview-item">
          <div class="preview-item-main">
            <strong>${escapeHtml(getRolloverItemDisplayTitle(item))}</strong>
            <span class="preview-item-sub">${escapeHtml(item.broker || 'No Broker')} • ${spanInfo}${notesMeta} • ID: ${escapeHtml(item.key.slice(0, 8))}...</span>
          </div>
          <span class="badge badge-pending">${escapeHtml(item.status)}</span>
        </div>
      `;
    });
    html += `</div>`;
  }
  html += `</div>`;

  if (plan.alreadyOnTarget.length > 0) {
    html += `<div>
      <div class="preview-section-title">Already Present on ${escapeHtml(targetVal)} (${plan.alreadyOnTarget.length})</div>
      <div class="preview-list">`;
    plan.alreadyOnTarget.forEach(({ targetRow, latestRow, reopenedOnSource, hasMissingPriorNotes }) => {
      const syncNoteLabel = hasMissingPriorNotes ? ' • +Append missing prior notes' : '';
      html += `
        <div class="preview-item">
          <div class="preview-item-main">
            <strong>${escapeHtml(getRolloverItemDisplayTitle(targetRow))}</strong>
            <span class="preview-item-sub">${escapeHtml(targetRow.broker || 'No Broker')} • Row ${targetRow.rowNumber} on ${escapeHtml(targetVal)}${syncNoteLabel}</span>
          </div>
          <span class="badge ${isCompletedStatus(targetRow.status) ? 'badge-complete' : 'badge-pending'}">
            ${escapeHtml(targetRow.status)}${reopenedOnSource ? ` → Reopen (${escapeHtml(latestRow.status)})` : ''}
          </span>
        </div>
      `;
    });
    html += `</div></div>`;
  }

  if (plan.supersededByComplete.length > 0) {
    html += `<div>
      <div class="preview-section-title warning-title">
        Open on Older Tab, Marked Closed / Complete on Later Tab (${plan.supersededByComplete.length})
      </div>
      <div style="font-size:0.75rem; color:var(--text-muted); margin-bottom:6px;">
        These requests had an open status on an earlier sheet, but were marked <strong>${escapeHtml(plan.supersededByComplete[0].completedRow.status)}</strong> on a newer sheet. Check any item below to reopen and carry it over into <code>${escapeHtml(targetVal)}</code>:
      </div>
      <div class="preview-list">`;
    plan.supersededByComplete.forEach(({ openRow, completedRow }) => {
      html += `
        <label class="preview-item" style="cursor:pointer;">
          <div style="display:flex; align-items:center; gap:10px;">
            <input type="checkbox" class="chk-force-rollover" data-key="${escapeHtml(openRow.key)}" style="width:16px; height:16px;">
            <div class="preview-item-main">
              <strong>${escapeHtml(getRolloverItemDisplayTitle(openRow))}</strong>
              <span class="preview-item-sub">
                ${escapeHtml(openRow.broker || 'No Broker')} • Was <strong>${escapeHtml(openRow.status)}</strong> on ${escapeHtml(openRow.sheetName)}, marked <strong>${escapeHtml(completedRow.status)}</strong> on ${escapeHtml(completedRow.sheetName)}
              </span>
            </div>
          </div>
          <span class="badge badge-pending">Reopen as ${escapeHtml(openRow.status)}</span>
        </label>
      `;
    });
    html += `</div></div>`;
  }

  rolloverPreviewBox.innerHTML = html;
}

rolloverForm.addEventListener('submit', async (e) => {
  e.preventDefault();
  const targetTabName = rolloverTargetDate.value.trim();
  if (!DATE_TAB_REGEX.test(targetTabName)) {
    return showModalAlert({
      title: 'Invalid Date Format',
      message: 'Please use exact MM-DD-YYYY format (e.g., 10-25-2026).'
    });
  }

  const forcedKeys = new Set(
    Array.from(rolloverPreviewBox.querySelectorAll('.chk-force-rollover:checked'))
      .map(cb => cb.dataset.key.toLowerCase())
  );

  const targetTime = parseDateTab(targetTabName);
  const hasFuture = dateSheets.some(name => parseDateTab(name) > targetTime);
  const includeFuture = hasFuture && chkIncludeFutureTabs.checked;
  const selectedSourceMode = rolloverSourceMode;

  try {
    await loadWorkbookFromDisk(sheetSelect.value);

    const plan = computeRolloverAnalysis(targetTabName, includeFuture, selectedSourceMode);
    if (!plan.valid) return;

    const itemsToCarry = [...plan.toRollOver];
    plan.supersededByComplete.forEach(({ openRow, completedRow }) => {
      if (forcedKeys.has(openRow.key.toLowerCase())) {
        itemsToCarry.push(openRow);
        const completedWs = workbook.getWorksheet(completedRow.sheetName);
        if (completedWs) {
          const compColMap = getWorksheetColumnMap(completedWs);
          if (compColMap.statusCol) {
            const cr = completedWs.getRow(completedRow.rowNumber);
            cr.getCell(compColMap.statusCol).value = openRow.status;
            cr.commit();
          }
        }
      }
    });

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

    const targetColMap = getWorksheetColumnMap(targetSheet);
    targetSheet.getColumn(targetColMap.trackingCol).hidden = true;
    if (targetColMap.emailSubjCol) {
      targetSheet.getColumn(targetColMap.emailSubjCol).hidden = false;
    }

    const existingTargetList = [...allRows.filter(r => r.sheetName === targetTabName)];

    let writeRowIdx = getNextEmptyRow(targetSheet);
    let addedCount = 0;
    let alreadyPresentCount = existingTargetList.length;
    let reopenedCount = 0;
    let notesSyncedCount = 0;

    for (const item of itemsToCarry) {
      const normKey = item.key.toLowerCase();
      const duplicateOnTarget = existingTargetList.some(
        t => t.key.toLowerCase() === normKey || areRowsSameUniqueRequest(t, item)
      );
      if (!duplicateOnTarget) {
        const rNum = writeRowIdx++;
        const r = targetSheet.getRow(rNum);
        writeItemToWorksheetRow(r, rNum, item, targetColMap);
        existingTargetList.push(item);
        unifyHistoricalGroupTrackingIds(item._entries, item.key);
        addedCount++;
      }
    }

    for (const entry of plan.alreadyOnTarget) {
      if (entry.groupEntries && entry.groupEntries.length > 0) {
        unifyHistoricalGroupTrackingIds(entry.groupEntries, entry.targetRow.key);
      }

      if (entry.reopenedOnSource || entry.hasMissingPriorNotes) {
        const excelRow = targetSheet.getRow(entry.targetRow.rowNumber);
        if (entry.reopenedOnSource && targetColMap.statusCol) {
          excelRow.getCell(targetColMap.statusCol).value = entry.latestRow.status;
          reopenedCount++;
        }
        if (entry.hasMissingPriorNotes && entry.mergedTargetNotes && targetColMap.notesCol) {
          excelRow.getCell(targetColMap.notesCol).value = entry.mergedTargetNotes;
          notesSyncedCount++;
        }
        excelRow.commit();
      }
    }

    const saveResult = await saveWorkbookToDisk();
    if (saveResult) {
      rolloverModal.classList.remove('open');
      activeWorkspaceTab = DAILY_TRACKER_WS_TAB;
      clearTransientFiltersOnNavigation({ resetScopeToCurrent: true });
      await loadWorkbookFromDisk(targetTabName, {
        preserveStatus: saveResult === 'mirror-locked'
      });
      const detailsParts = [`${alreadyPresentCount} already present`];
      if (notesSyncedCount > 0) detailsParts.push(`${notesSyncedCount} enriched with prior notes`);
      if (reopenedCount > 0) detailsParts.push(`${reopenedCount} reopened`);
      const actionLabel = createdNewSheet ? `Created & synced ${targetTabName}` : `Synced ${targetTabName}`;
      const msg = `${actionLabel}: Rolled over ${addedCount} unique request(s) (${detailsParts.join(', ')}).`;
      showToast(msg, 'success');
    }
  } catch (err) {
    console.error('Error running rollover:', err);
    await showModalAlert({
      title: 'Rollover Error',
      message: 'Failed to execute rollover: ' + err.message,
      variant: 'danger'
    });
  }
});


// --- STARTUP INITIALIZATION (Runs once all TrackerUI script modules are loaded) ---
initCachedDirectoryHandles();
