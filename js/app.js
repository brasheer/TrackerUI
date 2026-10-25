// --- STATE ---
let parentDirHandle = null;
let cachedDirHandle = null;
let supervisorDirHandle = null;
let cachedSupervisorDirHandle = null;
let availableWorkbooks = []; // [{ name, handle, lastModified, sortWeight }]
let workingFileHandle = null;
let workbook = null;
let allRows = []; // Flattened index across all MM-DD-YYYY tabs
let dateSheets = []; // Ordered list of MM-DD-YYYY sheet names (newest first)
let activeKeyFilter = null;
let editingTarget = null; // { sheetName, rowNumber, key } or null for new

// Delete Mode State
let isDeleteMode = false;
let currentlyVisibleRows = [];
const selectedForDelete = new Set(); // Stores "${sheetName}::${rowNumber}"

// Dynamic Reference Data
let masterAccountsMap = []; // [{ broker, master, bidder }]
let validationActions = [];
let validationStatuses = [];

// Stable in-memory backfill cache so legacy rows without Col K keep the same UUID before & during save
const legacyBackfillMap = new Map();

const DEFAULT_STATUSES = [
  'Pending',
  'In Progress',
  'Waiting on Buyer',
  'Escalated',
  'Complete'
];

const DATE_TAB_REGEX = /^\d{2}-\d{2}-\d{4}$/;

const IDB_NAME = 'TrackerUI_DirectoryDB';
const IDB_STORE = 'handles';
const IDB_AGENT_KEY = 'agentFolderHandle';
const IDB_SUPERVISOR_KEY = 'supervisorFolderHandle';

// --- DOM ELEMENTS ---
const statusPill = document.getElementById('statusPill');
const workbookPickerWrap = document.getElementById('workbookPickerWrap');
const workbookSelect = document.getElementById('workbookSelect');
const btnReconnect = document.getElementById('btnReconnect');
const btnLinkFile = document.getElementById('btnLinkFile');
const btnSupervisorDir = document.getElementById('btnSupervisorDir');
const btnResetSupervisorDir = document.getElementById('btnResetSupervisorDir');
const btnRollover = document.getElementById('btnRollover');
const btnDeleteMode = document.getElementById('btnDeleteMode');
const btnNewEntry = document.getElementById('btnNewEntry');
const sheetSelect = document.getElementById('sheetSelect');
const scopeSelect = document.getElementById('scopeSelect');
const searchInput = document.getElementById('searchInput');
const searchScopeHint = document.getElementById('searchScopeHint');
const listBody = document.getElementById('listBody');
const filterBanner = document.getElementById('filterBanner');
const filterKeyText = document.getElementById('filterKeyText');
const editModal = document.getElementById('editModal');
const rolloverModal = document.getElementById('rolloverModal');
const toastContainer = document.getElementById('toastContainer');

// Delete Mode DOM Elements
const deleteBanner = document.getElementById('deleteBanner');
const thDeleteSelect = document.getElementById('thDeleteSelect');
const chkSelectAllHeader = document.getElementById('chkSelectAllHeader');
const chkSelectAllBanner = document.getElementById('chkSelectAllBanner');
const deleteSelectedCountEl = document.getElementById('deleteSelectedCount');
const btnConfirmDeleteSelected = document.getElementById('btnConfirmDeleteSelected');
const btnCancelDeleteMode = document.getElementById('btnCancelDeleteMode');

const fSubject = document.getElementById('fSubject');
const fTicket = document.getElementById('fTicket');
const fBroker = document.getElementById('fBroker');
const fMaster = document.getElementById('fMaster');
const fBidder = document.getElementById('fBidder');
const fAction = document.getElementById('fAction');
const fStatus = document.getElementById('fStatus');
const fYard = document.getElementById('fYard');
const fLots = document.getElementById('fLots');
const fNotes = document.getElementById('fNotes');
const brokerListEl = document.getElementById('brokerList');
const actionListEl = document.getElementById('actionList');

// Rollover Modal DOM Elements
const rolloverForm = document.getElementById('rolloverForm');
const rolloverTodayBadge = document.getElementById('rolloverTodayBadge');
const rolloverFutureBanner = document.getElementById('rolloverFutureBanner');
const rolloverTargetDate = document.getElementById('rolloverTargetDate');
const rolloverPresetBtns = document.getElementById('rolloverPresetBtns');
const rolloverFutureOptionWrap = document.getElementById('rolloverFutureOptionWrap');
const chkIncludeFutureTabs = document.getElementById('chkIncludeFutureTabs');
const chkIncludeFutureText = document.getElementById('chkIncludeFutureText');
const rolloverPreviewBox = document.getElementById('rolloverPreviewBox');
const btnCancelRollover = document.getElementById('btnCancelRollover');

// --- UUID GENERATOR (PURE IMMUTABLE GUID FOR COLUMN K) ---
function generateTrackingGuid() {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID().toLowerCase();
  }
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    const v = c === 'x' ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}

// --- INDEXEDDB DIRECTORY HANDLE PERSISTENCE ---
function openHandleDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) {
        db.createObjectStore(IDB_STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function saveHandleToIDB(key, dirHandle) {
  try {
    const db = await openHandleDB();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readwrite');
      tx.objectStore(IDB_STORE).put(dirHandle, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn(`Could not persist handle (${key}) in IndexedDB:`, err);
  }
}

async function getHandleFromIDB(key) {
  try {
    const db = await openHandleDB();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readonly');
      const req = tx.objectStore(IDB_STORE).get(key);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  } catch (err) {
    console.warn(`Could not read handle (${key}) from IndexedDB:`, err);
    return null;
  }
}

async function deleteHandleFromIDB(key) {
  try {
    const db = await openHandleDB();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readwrite');
      tx.objectStore(IDB_STORE).delete(key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn(`Could not delete handle (${key}) from IndexedDB:`, err);
  }
}

function updateSupervisorDirUI() {
  if (supervisorDirHandle) {
    btnSupervisorDir.textContent = `👁️ Supervisor Folder: ${supervisorDirHandle.name}`;
    btnSupervisorDir.title = `Saving (VIEW ONLY).xlsx copies to "${supervisorDirHandle.name}". Click to change folder.`;
    btnSupervisorDir.classList.add('custom-active');
    btnResetSupervisorDir.style.display = 'inline-block';
  } else if (cachedSupervisorDirHandle) {
    btnSupervisorDir.textContent = `👁️ Reconnect Supervisor: ${cachedSupervisorDirHandle.name}`;
    btnSupervisorDir.title = `Click to reconnect saved Supervisor folder "${cachedSupervisorDirHandle.name}"`;
    btnSupervisorDir.classList.add('custom-active');
    btnResetSupervisorDir.style.display = 'inline-block';
  } else {
    btnSupervisorDir.textContent = '👁️ Supervisor Folder: Same as Agent';
    btnSupervisorDir.title = 'Select a custom folder where (VIEW ONLY).xlsx copies are saved';
    btnSupervisorDir.classList.remove('custom-active');
    btnResetSupervisorDir.style.display = 'none';
  }
}

async function initCachedDirectoryHandles() {
  if (!('indexedDB' in window)) return;
  const savedAgentHandle = await getHandleFromIDB(IDB_AGENT_KEY);
  if (savedAgentHandle) {
    cachedDirHandle = savedAgentHandle;
    btnReconnect.textContent = `🔄 Reconnect Folder (${savedAgentHandle.name})`;
    btnReconnect.style.display = 'inline-block';
  }

  const savedSupHandle = await getHandleFromIDB(IDB_SUPERVISOR_KEY);
  if (savedSupHandle) {
    cachedSupervisorDirHandle = savedSupHandle;
  }
  updateSupervisorDirUI();
}

initCachedDirectoryHandles();

// --- 1. DIRECTORY SELECTION, SUPERVISOR EXPORT FOLDER & WORKBOOK DISCOVERY ---
btnLinkFile.addEventListener('click', async () => {
  try {
    if (typeof window.showDirectoryPicker !== 'function') {
      return alert('Your browser does not support showDirectoryPicker. Please use Chrome or Edge over localhost/HTTPS.');
    }
    const dirHandle = await window.showDirectoryPicker({ mode: 'readwrite' });
    parentDirHandle = dirHandle;
    cachedDirHandle = dirHandle;
    await saveHandleToIDB(IDB_AGENT_KEY, dirHandle);
    btnReconnect.style.display = 'none';
    await scanDirectoryAndLoadWorkbooks();
  } catch (err) {
    if (err.name !== 'AbortError') {
      alert('Could not open agent folder: ' + err.message);
    }
  }
});

btnReconnect.addEventListener('click', async () => {
  if (!cachedDirHandle) return;
  try {
    let perm = await cachedDirHandle.queryPermission({ mode: 'readwrite' });
    if (perm !== 'granted') {
      perm = await cachedDirHandle.requestPermission({ mode: 'readwrite' });
    }
    if (perm === 'granted') {
      parentDirHandle = cachedDirHandle;
      btnReconnect.style.display = 'none';

      // Also attempt to reconnect cached Supervisor folder if one was configured
      if (cachedSupervisorDirHandle && !supervisorDirHandle) {
        try {
          let supPerm = await cachedSupervisorDirHandle.queryPermission({ mode: 'readwrite' });
          if (supPerm !== 'granted') {
            supPerm = await cachedSupervisorDirHandle.requestPermission({ mode: 'readwrite' });
          }
          if (supPerm === 'granted') {
            supervisorDirHandle = cachedSupervisorDirHandle;
          }
        } catch (supErr) {
          console.warn('Could not auto-reconnect supervisor folder:', supErr);
        }
        updateSupervisorDirUI();
      }

      await scanDirectoryAndLoadWorkbooks();
    } else {
      alert('Read/Write permission was not granted for the folder.');
    }
  } catch (err) {
    console.error('Reconnect folder error:', err);
    alert('Could not reconnect folder: ' + err.message);
  }
});

btnSupervisorDir.addEventListener('click', async () => {
  try {
    if (typeof window.showDirectoryPicker !== 'function') {
      return alert('Your browser does not support showDirectoryPicker. Please use Chrome or Edge over localhost/HTTPS.');
    }

    // If a cached supervisor folder exists but isn't connected yet, try reconnecting it first
    if (cachedSupervisorDirHandle && !supervisorDirHandle) {
      let perm = await cachedSupervisorDirHandle.queryPermission({ mode: 'readwrite' });
      if (perm !== 'granted') {
        perm = await cachedSupervisorDirHandle.requestPermission({ mode: 'readwrite' });
      }
      if (perm === 'granted') {
        supervisorDirHandle = cachedSupervisorDirHandle;
        updateSupervisorDirUI();
        showToast(`Reconnected Supervisor export folder: "${supervisorDirHandle.name}"`, 'success');
        if (workbook && workingFileHandle) {
          await writeMirrorOnly();
        }
        return;
      }
    }

    const supDir = await window.showDirectoryPicker({ mode: 'readwrite' });
    supervisorDirHandle = supDir;
    cachedSupervisorDirHandle = supDir;
    await saveHandleToIDB(IDB_SUPERVISOR_KEY, supDir);
    updateSupervisorDirUI();
    showToast(`Supervisor (VIEW ONLY) export folder set to "${supDir.name}"`, 'success');

    // If a workbook is already loaded, immediately write a fresh (VIEW ONLY) copy to the selected folder
    if (workbook && workingFileHandle) {
      await writeMirrorOnly();
    }
  } catch (err) {
    if (err.name !== 'AbortError') {
      alert('Could not select Supervisor export folder: ' + err.message);
    }
  }
});

btnResetSupervisorDir.addEventListener('click', async () => {
  supervisorDirHandle = null;
  cachedSupervisorDirHandle = null;
  await deleteHandleFromIDB(IDB_SUPERVISOR_KEY);
  updateSupervisorDirUI();
  showToast('Supervisor (VIEW ONLY) export reset to Agent workbook folder.', 'info');
});

async function writeMirrorOnly() {
  const targetMirrorDir = supervisorDirHandle || parentDirHandle;
  if (!workbook || !workingFileHandle || !targetMirrorDir) return;

  const mirrorFileName = workingFileHandle.name.replace(/\.xlsx$/i, '') + ' (VIEW ONLY).xlsx';
  try {
    const buffer = await workbook.xlsx.writeBuffer();
    const mirrorHandle = await targetMirrorDir.getFileHandle(mirrorFileName, { create: true });
    const mirrorWritable = await mirrorHandle.createWritable();
    await mirrorWritable.write(buffer);
    await mirrorWritable.close();
    showToast(`Synced "${mirrorFileName}" to "${targetMirrorDir.name}".`, 'success');
  } catch (err) {
    console.warn('Immediate mirror export failed:', err);
    showToast(`⚠️ Could not write "${mirrorFileName}" to "${targetMirrorDir.name}" (file may be open in Excel).`, 'warning');
  }
}

function extractMonthYearWeight(fileName, fallbackLastModified) {
  const months = {
    january: 0, jan: 0,
    february: 1, feb: 1,
    march: 2, mar: 2,
    april: 3, apr: 3,
    may: 4,
    june: 5, jun: 5,
    july: 6, jul: 6,
    august: 7, aug: 7,
    september: 8, sep: 8, sept: 8,
    october: 9, oct: 9,
    november: 10, nov: 10,
    december: 11, dec: 11
  };
  const regex = /\b(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)[\s\-_]+(20\d{2})\b/i;
  const match = fileName.match(regex);
  if (match) {
    const mIdx = months[match[1].toLowerCase()];
    const year = parseInt(match[2], 10);
    if (mIdx !== undefined && !isNaN(year)) {
      return new Date(year, mIdx, 28).getTime();
    }
  }
  return fallbackLastModified;
}

async function scanDirectoryAndLoadWorkbooks(preserveFileName = null) {
  if (!parentDirHandle) return;
  setStatus(`Scanning ${parentDirHandle.name}...`, '');

  availableWorkbooks = [];
  for await (const entry of parentDirHandle.values()) {
    if (entry.kind !== 'file') continue;
    const name = entry.name;
    const lower = name.toLowerCase();

    // Must be .xlsx, NOT end with (VIEW ONLY).xlsx, and NOT start with ~$
    if (
      lower.endsWith('.xlsx') &&
      !lower.endsWith('(view only).xlsx') &&
      !name.startsWith('~$')
    ) {
      try {
        const fileObj = await entry.getFile();
        const sortWeight = extractMonthYearWeight(name, fileObj.lastModified);
        availableWorkbooks.push({
          name,
          handle: entry,
          lastModified: fileObj.lastModified,
          sortWeight
        });
      } catch (e) {
        console.warn('Skipping unreadable file:', name, e);
      }
    }
  }

  if (availableWorkbooks.length === 0) {
    workbookPickerWrap.style.display = 'none';
    workingFileHandle = null;
    setStatus(`No working .xlsx in ${parentDirHandle.name}`, 'locked');
    listBody.innerHTML = `<tr><td colspan="12" style="text-align:center; color:var(--text-muted); padding: 30px;">No working <code>.xlsx</code> files found in folder <strong>${escapeHtml(parentDirHandle.name)}</strong> (ignoring <code>~$</code> lock files and <code>(VIEW ONLY).xlsx</code>).</td></tr>`;
    return;
  }

  // Sort most recent monthly workbook first
  availableWorkbooks.sort((a, b) => {
    if (b.sortWeight !== a.sortWeight) return b.sortWeight - a.sortWeight;
    return b.lastModified - a.lastModified;
  });

  // Populate top-bar workbook selector dropdown
  workbookSelect.innerHTML = '';
  availableWorkbooks.forEach((wb, idx) => {
    const opt = document.createElement('option');
    opt.value = String(idx);
    opt.textContent = wb.name;
    workbookSelect.appendChild(opt);
  });

  workbookPickerWrap.style.display = availableWorkbooks.length >= 1 ? 'flex' : 'none';

  let selectedIdx = 0;
  if (preserveFileName) {
    const foundIdx = availableWorkbooks.findIndex(w => w.name === preserveFileName);
    if (foundIdx !== -1) selectedIdx = foundIdx;
  }

  workbookSelect.value = String(selectedIdx);
  workingFileHandle = availableWorkbooks[selectedIdx].handle;

  await loadWorkbookFromDisk();
}

workbookSelect.addEventListener('change', async () => {
  const idx = Number(workbookSelect.value);
  if (availableWorkbooks[idx]) {
    workingFileHandle = availableWorkbooks[idx].handle;
    activeKeyFilter = null;
    filterBanner.style.display = 'none';
    setDeleteMode(false);
    await loadWorkbookFromDisk();
  }
});

// --- 2. READ WORKBOOK, PARSE REFERENCE SHEETS & INDEX ALL DAILY TABS ---
async function loadWorkbookFromDisk(preserveSelectedSheet = null, options = {}) {
  if (!workingFileHandle) return;
  if (!options.preserveStatus) {
    setStatus('Reading...', '');
  }

  const file = await workingFileHandle.getFile();
  const arrayBuffer = await file.arrayBuffer();

  workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);

  allRows = [];
  dateSheets = [];
  masterAccountsMap = [];
  validationActions = [];
  validationStatuses = [];

  workbook.eachSheet((worksheet) => {
    const rawName = worksheet.name.trim();
    const lowerName = rawName.toLowerCase();

    if (DATE_TAB_REGEX.test(rawName)) {
      dateSheets.push(rawName);
      indexWorksheetRows(worksheet, rawName);
    } else if (lowerName === 'validation lists' || lowerName === 'validation list') {
      parseValidationSheet(worksheet);
    } else if (lowerName === 'master accounts' || lowerName === 'master account') {
      parseMasterAccountsSheet(worksheet);
    }
  });

  // Populate dynamic datalists and status dropdown
  populateReferenceControls();

  // Sort date tabs newest first for UI dropdown
  dateSheets.sort((a, b) => parseDateTab(b) - parseDateTab(a));

  populateSheetDropdown(preserveSelectedSheet);
  sheetSelect.disabled = false;
  scopeSelect.disabled = false;
  searchInput.disabled = false;
  btnNewEntry.style.display = 'inline-block';
  btnRollover.style.display = 'inline-block';
  btnDeleteMode.style.display = 'inline-block';

  if (!options.preserveStatus) {
    setStatus(`Synced: ${file.name}`, 'ready');
  }
  renderList();
}

function parseDateTab(tabName) {
  const [mm, dd, yyyy] = tabName.split('-').map(Number);
  return new Date(yyyy, mm - 1, dd).getTime();
}

function addDaysToDateTab(tabName, daysToAdd) {
  const [mm, dd, yyyy] = tabName.split('-').map(Number);
  const d = new Date(yyyy, mm - 1, dd);
  d.setDate(d.getDate() + daysToAdd);
  const outMm = String(d.getMonth() + 1).padStart(2, '0');
  const outDd = String(d.getDate()).padStart(2, '0');
  const outYyyy = d.getFullYear();
  return `${outMm}-${outDd}-${outYyyy}`;
}

function cellText(row, colIndex) {
  const val = row.getCell(colIndex).value;
  if (val === null || val === undefined) return '';
  if (typeof val === 'object') {
    if (val.result !== undefined && val.result !== null) return String(val.result).trim();
    if (val.text !== undefined && val.text !== null) return String(val.text).trim();
    if (Array.isArray(val.richText)) return val.richText.map(t => t.text).join('').trim();
    if (val instanceof Date) return val.toISOString().split('T')[0];
  }
  return String(val).trim();
}

// --- 3. DYNAMIC REFERENCE SHEET PARSING (validation lists & Master Accounts) ---
function parseValidationSheet(worksheet) {
  const actionsSet = new Set();
  const statusesSet = new Set();

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return; // Skip header row
    const actionVal = cellText(row, 1);
    const statusVal = cellText(row, 2);

    if (actionVal) actionsSet.add(actionVal);
    if (statusVal) statusesSet.add(statusVal);
  });

  validationActions = Array.from(actionsSet);
  validationStatuses = Array.from(statusesSet);
}

function parseMasterAccountsSheet(worksheet) {
  const seen = new Set();
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return; // Skip header row
    const broker = cellText(row, 1);
    const master = cellText(row, 2).toUpperCase();
    const bidder = cellText(row, 3).toUpperCase();

    if (!broker && !master) return;
    const dedupeKey = `${broker.toLowerCase()}::${master.toLowerCase()}`;
    if (seen.has(dedupeKey)) return;
    seen.add(dedupeKey);

    masterAccountsMap.push({ broker, master, bidder });
  });
}

function populateReferenceControls() {
  // 1. Populate Action Datalist (#actionList)
  actionListEl.innerHTML = '';
  validationActions.forEach(action => {
    const opt = document.createElement('option');
    opt.value = action;
    actionListEl.appendChild(opt);
  });

  // 2. Populate Status Select (#fStatus) with fallback
  const statusesToUse = validationStatuses.length > 0 ? validationStatuses : DEFAULT_STATUSES;
  const currentSelected = fStatus.value;
  fStatus.innerHTML = '';
  statusesToUse.forEach(st => {
    const opt = document.createElement('option');
    opt.value = st;
    opt.textContent = st;
    fStatus.appendChild(opt);
  });
  if (currentSelected && statusesToUse.includes(currentSelected)) {
    fStatus.value = currentSelected;
  }

  // 3. Populate Broker Datalist (#brokerList)
  brokerListEl.innerHTML = '';
  const uniqueBrokers = Array.from(new Set(masterAccountsMap.map(m => m.broker).filter(Boolean)));
  uniqueBrokers.forEach(broker => {
    const opt = document.createElement('option');
    opt.value = broker;
    brokerListEl.appendChild(opt);
  });
}

// Auto-Fill Behavior in Modal: Broker <-> Master# (& Default Bidder#)
function handleBrokerAutoFill() {
  const typedBroker = fBroker.value.trim().toLowerCase();
  if (!typedBroker) return;

  const match = masterAccountsMap.find(m => m.broker.toLowerCase() === typedBroker);
  if (match) {
    if (!fMaster.value.trim() && match.master) {
      fMaster.value = match.master;
    }
    if (!fBidder.value.trim() && match.bidder) {
      fBidder.value = match.bidder;
    }
  }
}

function handleMasterReverseAutoFill() {
  const typedMaster = fMaster.value.trim().toLowerCase();
  if (!typedMaster) return;

  const match = masterAccountsMap.find(m => m.master.toLowerCase() === typedMaster);
  if (match) {
    if (!fBroker.value.trim() && match.broker) {
      fBroker.value = match.broker;
    }
    if (!fBidder.value.trim() && match.bidder) {
      fBidder.value = match.bidder;
    }
  }
}

fBroker.addEventListener('input', handleBrokerAutoFill);
fBroker.addEventListener('change', handleBrokerAutoFill);
fMaster.addEventListener('input', handleMasterReverseAutoFill);
fMaster.addEventListener('change', handleMasterReverseAutoFill);

// --- INDEX WORKSHEET ROWS & LEGACY GUID BACKFILL ---
function indexWorksheetRows(worksheet, sheetName) {
  // Keep Column K hidden
  worksheet.getColumn(11).hidden = true;

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return; // Skip header row

    const subject = cellText(row, 1);
    const ticket = cellText(row, 2);
    const broker = cellText(row, 3);
    const master = cellText(row, 4);
    const bidder = cellText(row, 5);
    const action = cellText(row, 6);
    const yard = cellText(row, 7);
    const lots = cellText(row, 8);
    const notes = cellText(row, 9);
    const status = cellText(row, 10);

    // Check if row has any data across Columns A-J (1..10)
    const hasData = [subject, ticket, broker, master, bidder, action, yard, lots, notes, status]
      .some(val => val !== '');
    if (!hasData) return;

    // Pure Immutable GUID in Hidden Column K (Cell 11)
    let key = cellText(row, 11).toLowerCase();
    if (!key) {
      // Legacy Backfill: generate a UUIDv4 in memory and stage it on Cell 11 for the next save
      const cacheKey = `${workingFileHandle ? workingFileHandle.name : 'wb'}::${sheetName}::${rowNumber}`;
      if (legacyBackfillMap.has(cacheKey)) {
        key = legacyBackfillMap.get(cacheKey);
      } else {
        key = generateTrackingGuid();
        legacyBackfillMap.set(cacheKey, key);
      }
      row.getCell(11).value = key;
      row.commit();
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
      status: status || 'Pending',
      key
    });
  });
}

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

// --- 4. RENDER LIST VIEW, AUTOMATIC GLOBAL SEARCH, & DELETE SELECTION MODE ---
function getRowCompositeId(sheetName, rowNumber) {
  return `${sheetName}::${rowNumber}`;
}

function renderList() {
  const query = searchInput.value.trim().toLowerCase();
  const scope = scopeSelect.value;
  const currentSheet = sheetSelect.value;

  // Show visual indicator when search automatically pulls across all tabs
  if (searchScopeHint) {
    searchScopeHint.style.display = query.length > 0 ? 'inline-block' : 'none';
  }

  const filtered = allRows.filter(item => {
    // 1. If a unique Tracking ID filter is active, show ALL occurrences across all MM-DD-YYYY tabs
    if (activeKeyFilter) {
      return item.key.toLowerCase() === activeKeyFilter.toLowerCase();
    }
    // 2. When typing a search query, automatically pull matching results across ALL sheets
    if (query.length > 0) {
      const haystack = [
        item.subject, item.ticket, item.broker, item.master,
        item.bidder, item.action, item.yard, item.lots, item.notes, item.key, item.sheetName
      ].join(' ').toLowerCase();
      return haystack.includes(query);
    }
    // 3. When search is empty, respect the View Scope dropdown
    if (scope === 'current' && item.sheetName !== currentSheet) {
      return false;
    }
    return true;
  });

  currentlyVisibleRows = filtered;

  listBody.innerHTML = '';
  const colSpan = isDeleteMode ? 12 : 11;
  if (filtered.length === 0) {
    listBody.innerHTML = `<tr><td colspan="${colSpan}" style="text-align:center; color:var(--text-muted); padding: 24px;">No matching requests found.</td></tr>`;
    syncSelectAllCheckboxes();
    return;
  }

  filtered.forEach((item, idx) => {
    const tr = document.createElement('tr');
    const compId = getRowCompositeId(item.sheetName, item.rowNumber);
    const isCheckedForDelete = isDeleteMode && selectedForDelete.has(compId);
    if (isCheckedForDelete) {
      tr.classList.add('selected-for-delete');
    }

    const isComplete = item.status.toLowerCase() === 'complete';
    const noteText = item.notes || '—';
    const isLongNote = noteText.length > 45 || noteText.includes('\n');
    const shortGuid = item.key.length > 8 ? `${item.key.slice(0, 8)}...` : item.key;

    const deleteCheckboxCell = isDeleteMode
      ? `<td class="delete-cell">
           <input type="checkbox" class="chk-delete-row" data-comp="${escapeHtml(compId)}" ${isCheckedForDelete ? 'checked' : ''}>
         </td>`
      : '';

    tr.innerHTML = `
      ${deleteCheckboxCell}
      <td style="white-space:nowrap; font-weight:500;">${escapeHtml(item.sheetName)}</td>
      <td>
        <span class="key-chip" title="${escapeHtml(item.key)}" data-key="${escapeHtml(item.key)}">
          🔍 ${escapeHtml(shortGuid)}
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

  syncSelectAllCheckboxes();
}

// --- DELETE SELECTION MODE & CONFIRMATION ---
function setDeleteMode(enabled) {
  isDeleteMode = Boolean(enabled);
  selectedForDelete.clear();
  btnDeleteMode.classList.toggle('active', isDeleteMode);
  btnDeleteMode.textContent = isDeleteMode ? '✕ Exit Delete Mode' : '🗑️ Delete Requests';
  deleteBanner.style.display = isDeleteMode ? 'flex' : 'none';
  thDeleteSelect.style.display = isDeleteMode ? 'table-cell' : 'none';
  updateDeleteSelectionUI();
  renderList();
}

function updateDeleteSelectionUI() {
  const count = selectedForDelete.size;
  deleteSelectedCountEl.textContent = String(count);
  btnConfirmDeleteSelected.textContent = `🗑️ Delete Selected (${count})`;
  btnConfirmDeleteSelected.disabled = count === 0;
  syncSelectAllCheckboxes();
}

function syncSelectAllCheckboxes() {
  if (!isDeleteMode || currentlyVisibleRows.length === 0) {
    chkSelectAllHeader.checked = false;
    chkSelectAllBanner.checked = false;
    return;
  }
  const allVisibleSelected = currentlyVisibleRows.every(item =>
    selectedForDelete.has(getRowCompositeId(item.sheetName, item.rowNumber))
  );
  chkSelectAllHeader.checked = allVisibleSelected;
  chkSelectAllBanner.checked = allVisibleSelected;
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

listBody.addEventListener('change', (e) => {
  const chk = e.target.closest('.chk-delete-row');
  if (!chk) return;
  const compId = chk.dataset.comp;
  if (!compId) return;

  const tr = chk.closest('tr');
  if (chk.checked) {
    selectedForDelete.add(compId);
    if (tr) tr.classList.add('selected-for-delete');
  } else {
    selectedForDelete.delete(compId);
    if (tr) tr.classList.remove('selected-for-delete');
  }
  updateDeleteSelectionUI();
});

btnConfirmDeleteSelected.addEventListener('click', async () => {
  if (selectedForDelete.size === 0 || !workbook) return;

  // Gather selected items
  const itemsToDelete = allRows.filter(item =>
    selectedForDelete.has(getRowCompositeId(item.sheetName, item.rowNumber))
  );
  if (itemsToDelete.length === 0) return;

  const previewLines = itemsToDelete
    .slice(0, 12)
    .map(r => `• [${r.sheetName} • Row ${r.rowNumber}] ${r.subject} (${r.broker || 'No Broker'})`)
    .join('\n');
  const moreNotice = itemsToDelete.length > 12 ? `\n...and ${itemsToDelete.length - 12} more.` : '';

  const confirmed = confirm(
    `Are you sure you want to permanently delete ${itemsToDelete.length} selected request(s)?\n\n` +
    `${previewLines}${moreNotice}\n\n` +
    `This will remove the row(s) from the workbook and save immediately.`
  );
  if (!confirmed) return;

  const currentTab = sheetSelect.value;
  await loadWorkbookFromDisk(currentTab);

  // Group deletions by sheetName and sort rowNumber DESCENDING so spliceRows never shifts earlier row indices
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
    setDeleteMode(false);
    await loadWorkbookFromDisk(currentTab, {
      preserveStatus: saveResult === 'mirror-locked'
    });
    showToast(`Deleted ${deletedCount} request(s) from workbook.`, 'success');
  }
});

// Delegate clicks for Key Chip filter, Collapsible Notes, and Edit button
listBody.addEventListener('click', (e) => {
  const keyChip = e.target.closest('.key-chip');
  if (keyChip && keyChip.dataset.key) {
    activeKeyFilter = keyChip.dataset.key;
    filterKeyText.textContent = activeKeyFilter;
    filterBanner.style.display = 'flex';
    renderList();
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
searchInput.addEventListener('input', renderList);

// --- 5. EDIT & CREATE MODAL + INPUT HYGIENE + CROSS-TAB STATUS PROPAGATION ---
btnNewEntry.addEventListener('click', () => {
  if (!sheetSelect.value) {
    return alert('Create or select a daily tab (MM-DD-YYYY) first using the Rollover / Sync button.');
  }
  editingTarget = null;
  document.getElementById('modalTitle').textContent = `New Request (${sheetSelect.value})`;
  const modalKeyEl = document.getElementById('modalKeyDisplay');
  modalKeyEl.textContent = 'New UUIDv4 on Save';
  modalKeyEl.title = 'A unique 36-character GUID will be generated in hidden Column K';

  document.getElementById('entryForm').reset();
  ensureStatusOptionExists('Pending');
  fStatus.value = 'Pending';
  editModal.classList.add('open');
});

function ensureStatusOptionExists(statusValue) {
  if (!statusValue) return;
  const exists = Array.from(fStatus.options).some(opt => opt.value === statusValue);
  if (!exists) {
    const opt = document.createElement('option');
    opt.value = statusValue;
    opt.textContent = statusValue;
    fStatus.appendChild(opt);
  }
}

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

document.getElementById('entryForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const currentSheetName = editingTarget ? editingTarget.sheetName : sheetSelect.value;

  // Reload fresh copy from disk right before mutating to prevent stale overwrites
  await loadWorkbookFromDisk(currentSheetName);
  const worksheet = workbook.getWorksheet(currentSheetName);
  if (!worksheet) return alert('Target worksheet not found!');

  // Input Hygiene: trim all text inputs & normalize Master#, Bidder#, Yard#, and Ticket# to uppercase
  const subjectVal = fSubject.value.trim();
  const ticketVal = fTicket.value.trim().toUpperCase();
  const brokerVal = fBroker.value.trim();
  const masterVal = fMaster.value.trim().toUpperCase();
  const bidderVal = fBidder.value.trim().toUpperCase();
  const actionVal = fAction.value.trim();
  const yardVal = fYard.value.trim().toUpperCase();
  const lotsVal = fLots.value.trim();
  const notesVal = fNotes.value.trim();
  const statusVal = fStatus.value.trim();

  let targetRowNum;
  let trackingKey;

  if (editingTarget) {
    // Strictly preserve existing Column K UUID even if Ticket number changed or was deleted
    targetRowNum = editingTarget.rowNumber;
    trackingKey = editingTarget.key;
  } else {
    // Always generate a fresh lowercase UUIDv4 for new entries
    targetRowNum = getNextEmptyRow(worksheet);
    trackingKey = generateTrackingGuid();
  }

  // Ensure Column K stays hidden
  worksheet.getColumn(11).hidden = true;

  const row = worksheet.getRow(targetRowNum);
  row.getCell(1).value = subjectVal;
  row.getCell(2).value = ticketVal;
  row.getCell(3).value = brokerVal;
  row.getCell(4).value = masterVal;
  row.getCell(5).value = bidderVal;
  row.getCell(6).value = actionVal;
  row.getCell(7).value = yardVal;
  row.getCell(8).value = lotsVal;
  row.getCell(9).value = notesVal;
  row.getCell(10).value = statusVal;
  row.getCell(11).value = trackingKey; // Pure immutable UUIDv4 in hidden Column K
  row.commit();

  // If editing an older sheet, check if any NEWER sheets have the same Tracking ID with a conflicting status
  if (editingTarget) {
    const editedTime = parseDateTab(currentSheetName);
    const newerInstances = allRows.filter(
      r =>
        r.key.toLowerCase() === trackingKey.toLowerCase() &&
        parseDateTab(r.sheetName) > editedTime &&
        r.status.trim().toLowerCase() !== statusVal.toLowerCase()
    );
    if (newerInstances.length > 0) {
      const newerSummary = newerInstances.map(r => `${r.sheetName} (${r.status})`).join(', ');
      const syncNewer = confirm(
        `This request also exists on newer tab(s): ${newerSummary}.\n\n` +
        `Update its status to "${statusVal}" on those newer tab(s) as well so Rollover uses your updated status?`
      );
      if (syncNewer) {
        newerInstances.forEach(inst => {
          const newerWs = workbook.getWorksheet(inst.sheetName);
          if (newerWs) {
            const nr = newerWs.getRow(inst.rowNumber);
            nr.getCell(10).value = statusVal;
            nr.commit();
          }
        });
      }
    }
  }

  const saveResult = await saveWorkbookToDisk();
  if (saveResult) {
    editModal.classList.remove('open');
    await loadWorkbookFromDisk(currentSheetName, {
      preserveStatus: saveResult === 'mirror-locked'
    });
  }
});

function getNextEmptyRow(worksheet) {
  let lastRow = 1;
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const hasData = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10].some(col => cellText(row, col) !== '');
    if (hasData && rowNumber > lastRow) lastRow = rowNumber;
  });
  return lastRow + 1;
}

// --- 6. SMART MULTI-SHEET & IDEMPOTENT ROLLOVER ENGINE WITH FUTURE DATE HANDLING ---
btnRollover.addEventListener('click', async () => {
  if (!workbook) return;
  await loadWorkbookFromDisk(sheetSelect.value);
  openRolloverModal();
});

btnCancelRollover.addEventListener('click', () => {
  rolloverModal.classList.remove('open');
});

function openRolloverModal() {
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
      <strong>📅 Future-Dated Tab Detected:</strong> Your workbook's latest sheet is
      <code>${escapeHtml(latestSheet)}</code> (${latestCount} entry${latestCount === 1 ? '' : 'ies'}),
      which is ahead of today's date (<code>${escapeHtml(todayStr)}</code>).
      Select whether to sync into <strong>${escapeHtml(latestSheet)}</strong>, roll forward to
      <strong>${escapeHtml(nextAfterLatest)}</strong>, or sync into <strong>${escapeHtml(todayStr)}</strong>.
    `;

    presets.push({ label: `Latest Tab (${latestSheet})`, date: latestSheet });
    presets.push({ label: `Next Day (${nextAfterLatest})`, date: nextAfterLatest });
    presets.push({ label: `Today (${todayStr})`, date: todayStr });

    rolloverTargetDate.value = latestSheet;
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
    rolloverTargetDate.value = todayStr;
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

function computeRolloverAnalysis(targetTabName, includeFutureTabs) {
  if (!DATE_TAB_REGEX.test(targetTabName)) {
    return { valid: false };
  }

  const targetTime = parseDateTab(targetTabName);
  const targetExists = Boolean(workbook.getWorksheet(targetTabName));

  const priorSheets = dateSheets
    .filter(name => parseDateTab(name) < targetTime)
    .sort((a, b) => parseDateTab(a) - parseDateTab(b));

  const futureSheets = dateSheets
    .filter(name => parseDateTab(name) > targetTime)
    .sort((a, b) => parseDateTab(a) - parseDateTab(b));

  const sourceSheets = includeFutureTabs
    ? [...priorSheets, ...futureSheets]
    : priorSheets;

  const latestByKey = new Map();
  const lastOpenByKey = new Map();

  sourceSheets.forEach(sheetName => {
    const sheetRows = allRows
      .filter(r => r.sheetName === sheetName)
      .sort((a, b) => a.rowNumber - b.rowNumber);

    sheetRows.forEach(row => {
      const normKey = row.key.toLowerCase();
      latestByKey.set(normKey, row);
      if (row.status.trim().toLowerCase() !== 'complete') {
        lastOpenByKey.set(normKey, row);
      }
    });
  });

  const existingTargetRows = allRows.filter(r => r.sheetName === targetTabName);
  const existingOnTargetByKey = new Map();
  existingTargetRows.forEach(r => {
    existingOnTargetByKey.set(r.key.toLowerCase(), r);
  });

  const toRollOver = [];
  const alreadyOnTarget = [];
  const supersededByComplete = [];

  latestByKey.forEach((latestRow, normKey) => {
    const isLatestComplete = latestRow.status.trim().toLowerCase() === 'complete';

    if (existingOnTargetByKey.has(normKey)) {
      const targetRow = existingOnTargetByKey.get(normKey);
      const isTargetComplete = targetRow.status.trim().toLowerCase() === 'complete';
      alreadyOnTarget.push({
        latestRow,
        targetRow,
        reopenedOnSource: isTargetComplete && !isLatestComplete
      });
    } else if (!isLatestComplete) {
      toRollOver.push(latestRow);
    } else if (lastOpenByKey.has(normKey)) {
      supersededByComplete.push({
        openRow: lastOpenByKey.get(normKey),
        completedRow: latestRow
      });
    }
  });

  existingTargetRows.forEach(targetRow => {
    const normKey = targetRow.key.toLowerCase();
    if (!latestByKey.has(normKey)) {
      alreadyOnTarget.push({
        latestRow: targetRow,
        targetRow,
        reopenedOnSource: false
      });
    }
  });

  return {
    valid: true,
    targetTabName,
    targetExists,
    priorSheets,
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
    rolloverPreviewBox.innerHTML = `<div style="color:var(--warning);">Please enter a valid target date in exact <code>MM-DD-YYYY</code> format.</div>`;
    return;
  }

  const targetTime = parseDateTab(targetVal);
  const futureTabs = dateSheets.filter(name => parseDateTab(name) > targetTime);
  if (futureTabs.length > 0) {
    rolloverFutureOptionWrap.style.display = 'block';
    chkIncludeFutureText.textContent = `Also include unfinished requests from future-dated tab(s): ${futureTabs.join(', ')}`;
  } else {
    rolloverFutureOptionWrap.style.display = 'none';
  }

  const includeFuture = futureTabs.length > 0 && chkIncludeFutureTabs.checked;
  const plan = computeRolloverAnalysis(targetVal, includeFuture);

  const tabStateBadge = plan.targetExists
    ? `<span class="badge badge-complete">Existing Tab (${plan.alreadyOnTarget.length} request${plan.alreadyOnTarget.length === 1 ? '' : 's'} currently on sheet)</span>`
    : `<span class="badge badge-pending">New Tab (Will be created & cloned from template)</span>`;

  let html = `
    <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px; padding-bottom:8px; border-bottom:1px solid var(--border);">
      <div><strong>Target:</strong> <code>${escapeHtml(targetVal)}</code></div>
      ${tabStateBadge}
    </div>
  `;

  // 1. Unfinished Requests Ready to Roll Over
  html += `<div>
    <div class="preview-section-title">✅ Unfinished Requests to Roll Over (${plan.toRollOver.length})</div>`;
  if (plan.toRollOver.length === 0) {
    html += `<div style="color:var(--text-muted); font-size:12.5px;">No new unfinished requests found across source sheets.</div>`;
  } else {
    html += `<div class="preview-list">`;
    plan.toRollOver.forEach(item => {
      html += `
        <div class="preview-item">
          <div class="preview-item-main">
            <strong>${escapeHtml(item.subject)}</strong>
            <span class="preview-item-sub">${escapeHtml(item.broker)} • From ${escapeHtml(item.sheetName)} • ID: ${escapeHtml(item.key.slice(0, 8))}...</span>
          </div>
          <span class="badge badge-pending">${escapeHtml(item.status)}</span>
        </div>
      `;
    });
    html += `</div>`;
  }
  html += `</div>`;

  // 2. Already Present on Target Sheet
  if (plan.alreadyOnTarget.length > 0) {
    html += `<div>
      <div class="preview-section-title">📌 Already Present on ${escapeHtml(targetVal)} (${plan.alreadyOnTarget.length})</div>
      <div class="preview-list">`;
    plan.alreadyOnTarget.forEach(({ targetRow, latestRow, reopenedOnSource }) => {
      html += `
        <div class="preview-item">
          <div class="preview-item-main">
            <strong>${escapeHtml(targetRow.subject)}</strong>
            <span class="preview-item-sub">${escapeHtml(targetRow.broker)} • Row ${targetRow.rowNumber} on ${escapeHtml(targetVal)}</span>
          </div>
          <span class="badge ${targetRow.status.toLowerCase() === 'complete' ? 'badge-complete' : 'badge-pending'}">
            ${escapeHtml(targetRow.status)}${reopenedOnSource ? ` → Reopen (${escapeHtml(latestRow.status)})` : ''}
          </span>
        </div>
      `;
    });
    html += `</div></div>`;
  }

  // 3. Older Open Requests Superseded by a Later "Complete" Sheet
  if (plan.supersededByComplete.length > 0) {
    html += `<div>
      <div class="preview-section-title" style="color:#fbbf24;">
        ⚠️ Open on Older Tab, but Marked "Complete" on Later Tab (${plan.supersededByComplete.length})
      </div>
      <div style="font-size:12px; color:var(--text-muted); margin-bottom:6px;">
        These requests had an open status on an earlier sheet, but were marked <strong>Complete</strong> on a newer sheet. Check any item below to reopen and carry it over into <code>${escapeHtml(targetVal)}</code>:
      </div>
      <div class="preview-list">`;
    plan.supersededByComplete.forEach(({ openRow, completedRow }) => {
      html += `
        <label class="preview-item" style="cursor:pointer;">
          <div style="display:flex; align-items:center; gap:10px;">
            <input type="checkbox" class="chk-force-rollover" data-key="${escapeHtml(openRow.key)}" style="width:16px; height:16px;">
            <div class="preview-item-main">
              <strong>${escapeHtml(openRow.subject)}</strong>
              <span class="preview-item-sub">
                ${escapeHtml(openRow.broker)} • Was <strong>${escapeHtml(openRow.status)}</strong> on ${escapeHtml(openRow.sheetName)}, marked <strong>Complete</strong> on ${escapeHtml(completedRow.sheetName)}
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
    return alert('Please use exact MM-DD-YYYY format (e.g., 10-25-2026).');
  }

  const forcedKeys = new Set(
    Array.from(rolloverPreviewBox.querySelectorAll('.chk-force-rollover:checked'))
      .map(cb => cb.dataset.key.toLowerCase())
  );

  const targetTime = parseDateTab(targetTabName);
  const hasFuture = dateSheets.some(name => parseDateTab(name) > targetTime);
  const includeFuture = hasFuture && chkIncludeFutureTabs.checked;

  await loadWorkbookFromDisk(sheetSelect.value);

  const plan = computeRolloverAnalysis(targetTabName, includeFuture);
  if (!plan.valid) return;

  const itemsToCarry = [...plan.toRollOver];
  plan.supersededByComplete.forEach(({ openRow, completedRow }) => {
    if (forcedKeys.has(openRow.key.toLowerCase())) {
      itemsToCarry.push(openRow);
      const completedWs = workbook.getWorksheet(completedRow.sheetName);
      if (completedWs) {
        const cr = completedWs.getRow(completedRow.rowNumber);
        cr.getCell(10).value = openRow.status;
        cr.commit();
      }
    }
  });

  let targetSheet = workbook.getWorksheet(targetTabName);
  let createdNewSheet = false;

  if (!targetSheet) {
    createdNewSheet = true;
    const templateSheet = findTemplateSheet();
    targetSheet = workbook.addWorksheet(targetTabName);

    placeWorksheetFirst(targetSheet);

    if (templateSheet) {
      cloneTemplateStructure(templateSheet, targetSheet);
    } else {
      initializeDefaultSheetHeaders(targetSheet);
    }
  }

  targetSheet.getColumn(11).hidden = true;

  const existingTargetRows = allRows.filter(r => r.sheetName === targetTabName);
  const existingOnTargetByKey = new Map();
  existingTargetRows.forEach(r => {
    existingOnTargetByKey.set(r.key.toLowerCase(), r);
  });

  let writeRowIdx = getNextEmptyRow(targetSheet);
  let addedCount = 0;
  let alreadyPresentCount = existingTargetRows.length;
  let reopenedCount = 0;

  for (const item of itemsToCarry) {
    const normKey = item.key.toLowerCase();
    if (!existingOnTargetByKey.has(normKey)) {
      const r = targetSheet.getRow(writeRowIdx++);
      writeItemToWorksheetRow(r, item);
      existingOnTargetByKey.set(normKey, item);
      addedCount++;
    }
  }

  for (const entry of plan.alreadyOnTarget) {
    if (entry.reopenedOnSource) {
      const excelRow = targetSheet.getRow(entry.targetRow.rowNumber);
      excelRow.getCell(10).value = entry.latestRow.status;
      excelRow.commit();
      reopenedCount++;
    }
  }

  const saveResult = await saveWorkbookToDisk();
  if (saveResult) {
    rolloverModal.classList.remove('open');
    await loadWorkbookFromDisk(targetTabName, {
      preserveStatus: saveResult === 'mirror-locked'
    });
    const reopenSuffix = reopenedCount > 0 ? `, ${reopenedCount} reopened` : '';
    const actionLabel = createdNewSheet ? `Created & synced ${targetTabName}` : `Synced ${targetTabName}`;
    const msg = `${actionLabel}: Added ${addedCount} newly rolled-over request(s) (${alreadyPresentCount} already present on sheet${reopenSuffix}).`;
    showToast(msg, 'success');
    alert(msg);
  }
});

function findTemplateSheet() {
  let found = null;
  workbook.eachSheet((ws) => {
    if (ws.name.trim().toLowerCase() === 'template') {
      found = ws;
    }
  });
  if (!found && dateSheets.length > 0) {
    found = workbook.getWorksheet(dateSheets[0]);
  }
  return found;
}

function placeWorksheetFirst(targetSheet) {
  try {
    let order = 2;
    workbook.eachSheet((ws) => {
      if (ws.id === targetSheet.id) {
        ws.orderNo = 1;
      } else {
        ws.orderNo = order++;
      }
    });
  } catch (e) {
    console.warn('Could not reorder worksheet tabs:', e);
  }
}

function cloneTemplateStructure(templateSheet, newSheet) {
  templateSheet.columns.forEach((col, idx) => {
    const targetCol = newSheet.getColumn(idx + 1);
    targetCol.width = col.width || 15;
    targetCol.hidden = col.hidden || false;
  });

  const headerRow = templateSheet.getRow(1);
  const newHeader = newSheet.getRow(1);
  if (headerRow.height) newHeader.height = headerRow.height;

  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const targetCell = newHeader.getCell(colNumber);
    targetCell.value = cell.value;
    if (cell.style) {
      targetCell.style = JSON.parse(JSON.stringify(cell.style));
    }
  });
  newHeader.commit();
  newSheet.getColumn(11).hidden = true;
}

function initializeDefaultSheetHeaders(newSheet) {
  const headers = [
    'Subject', 'Ticket #', 'Broker Name', 'Master #', 'Bidder #',
    'Requested Action', 'Yard #', 'Lot Numbers', 'Action Notes',
    'Request Status', 'Tracking ID'
  ];
  const headerRow = newSheet.getRow(1);
  headers.forEach((h, idx) => {
    headerRow.getCell(idx + 1).value = h;
    newSheet.getColumn(idx + 1).width = idx === 0 || idx === 8 ? 28 : 16;
  });
  headerRow.commit();
  newSheet.getColumn(11).hidden = true;
}

function writeItemToWorksheetRow(row, item) {
  row.getCell(1).value = item.subject;
  row.getCell(2).value = item.ticket ? item.ticket.toUpperCase() : '';
  row.getCell(3).value = item.broker;
  row.getCell(4).value = item.master ? item.master.toUpperCase() : '';
  row.getCell(5).value = item.bidder ? item.bidder.toUpperCase() : '';
  row.getCell(6).value = item.action;
  row.getCell(7).value = item.yard ? item.yard.toUpperCase() : '';
  row.getCell(8).value = item.lots;
  row.getCell(9).value = item.notes;
  row.getCell(10).value = item.status;
  row.getCell(11).value = item.key; // Preserves immutable UUIDv4 across days
  row.commit();
}

// --- 7. DUAL-WRITE MIRRORING + CUSTOM SUPERVISOR DIRECTORY + LOCK DETECTION ---
async function saveWorkbookToDisk() {
  if (!workingFileHandle) return false;

  setStatus('Saving to Excel...', '');
  let buffer;
  try {
    buffer = await workbook.xlsx.writeBuffer();
  } catch (err) {
    console.error('Failed to serialize workbook buffer:', err);
    setStatus('Serialization Error', 'locked');
    alert('Failed to generate Excel buffer: ' + err.message);
    return false;
  }

  // 1. Primary Working File Write
  try {
    const writable = await workingFileHandle.createWritable();
    await writable.write(buffer);
    await writable.close();
  } catch (err) {
    console.error('Primary file write error:', err);
    setStatus('File Locked! Close Excel & Retry', 'locked');
    alert(
      'SAVE BLOCKED BY WINDOWS:\n\n' +
      `The working workbook ("${workingFileHandle.name}") is currently open in Desktop Excel.\n\n` +
      'Close the file in Excel and click Save again.'
    );
    return false;
  }

  // 2. Dual-Write Mirroring to (VIEW ONLY).xlsx in Supervisor Folder (or Agent Folder fallback)
  const targetMirrorDir = supervisorDirHandle || parentDirHandle;
  if (targetMirrorDir) {
    const mirrorFileName = workingFileHandle.name.replace(/\.xlsx$/i, '') + ' (VIEW ONLY).xlsx';
    try {
      const mirrorHandle = await targetMirrorDir.getFileHandle(mirrorFileName, { create: true });
      const mirrorWritable = await mirrorHandle.createWritable();
      await mirrorWritable.write(buffer);
      await mirrorWritable.close();
    } catch (mirrorErr) {
      console.warn(`Supervisor mirror file "${mirrorFileName}" is locked or busy:`, mirrorErr);
      setStatus('⚠️ Saved Working File (Mirror in use by Supervisor)', 'warning');
      showToast('⚠️ Saved Working File (Mirror in use by Supervisor)', 'warning');
      return 'mirror-locked';
    }
  }

  const mirrorLocationNote = supervisorDirHandle ? ` (Mirrored to ${supervisorDirHandle.name})` : '';
  setStatus(`Saved & Mirrored: ${workingFileHandle.name}${mirrorLocationNote}`, 'ready');
  return true;
}

// --- UTILITIES & TOAST ---
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

function showToast(message, type = 'info') {
  if (!toastContainer) return;
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.textContent = message;
  toastContainer.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.25s ease';
    setTimeout(() => toast.remove(), 260);
  }, 5000);
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// --- 8. PWA SERVICE WORKER REGISTRATION ---
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js')
      .then((reg) => console.log('Service Worker registered with scope:', reg.scope))
      .catch((err) => console.warn('Service Worker registration failed:', err));
  });
}
