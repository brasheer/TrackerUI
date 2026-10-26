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
let activeFieldFilter = null; // { field, label, value } for Like-Item filtering across tabs
let sortState = { column: null, direction: 'asc' }; // Column sort state
let editingTarget = null; // { sheetName, rowNumber, key } or null for new

// Mirror Retry State
let pendingMirrorBuffer = null;
let mirrorCountdownInterval = null;
let mirrorCountdownSec = 0;
let mirrorRetryCount = 0;
const MIRROR_RETRY_INTERVAL_SEC = 15;
const MIRROR_MAX_AUTO_RETRIES = 20;

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

// Original OpenXML archive state preserved on load so ExcelJS does not strip dynamic array metadata, external links, or data validations
let originalArchiveState = null;

const UUID_REGEX = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

const DEFAULT_ACTIONS = [
  'Title Inquiry',
  'Gate Pass Reissue',
  'Relist Dispute',
  'Wire Transfer Trace',
  'Bidder Account Unlock',
  'Storage Fee Waiver',
  'Lot Cancellation'
];

const DEFAULT_STATUSES = [
  'Pending',
  'In Progress',
  'Waiting on Buyer',
  'Escalated',
  'Complete'
];

const KNOWN_STATUS_VALUES = new Set([
  'pending',
  'in progress',
  'waiting on buyer',
  'escalated',
  'complete',
  'new',
  'waiting',
  'overdue',
  'closed'
]);

const STANDARD_HEADERS = [
  'Subject',
  'Ticket Number',
  'Broker Name',
  'Master #',
  'Bidder #',
  'Requested Action',
  'Yard #',
  'Lot Numbers',
  'Action Notes',
  'Request Status',
  'Tracking ID'
];

const DATE_TAB_REGEX = /^\d{1,2}-\d{1,2}-\d{2,4}$/;

function isTemplatePlaceholderTab(sheetName) {
  const norm = String(sheetName || '').trim().toLowerCase();
  return (
    norm === 'rename to current date' ||
    norm === 'rename to today' ||
    norm === 'current date'
  );
}

function isCompletedStatus(status) {
  const norm = String(status || '').trim().toLowerCase();
  return norm === 'complete' || norm === 'closed';
}

function getDefaultStatus() {
  if (validationStatuses.length > 0) {
    const preferred = validationStatuses.find(s => !isCompletedStatus(s));
    if (preferred) return preferred;
    return validationStatuses[0];
  }
  return 'Pending';
}

const IDB_NAME = 'TrackerUI_DirectoryDB';
const IDB_STORE = 'handles';
const IDB_AGENT_KEY = 'agentFolderHandle';
const IDB_SUPERVISOR_KEY = 'supervisorFolderHandle';

// --- DOM ELEMENTS ---
const statusPill = document.getElementById('statusPill');
const btnRetryMirror = document.getElementById('btnRetryMirror');
const workbookPickerWrap = document.getElementById('workbookPickerWrap');
const workbookSelect = document.getElementById('workbookSelect');
const btnReconnect = document.getElementById('btnReconnect');
const btnLinkFile = document.getElementById('btnLinkFile');
const btnRollover = document.getElementById('btnRollover');
const btnNewEntry = document.getElementById('btnNewEntry');

// Popover Tools Menu Elements
const toolsMenuWrap = document.getElementById('toolsMenuWrap');
const btnToolsMenu = document.getElementById('btnToolsMenu');
const toolsMenuDropdown = document.getElementById('toolsMenuDropdown');
const menuSelectAgentFolder = document.getElementById('menuSelectAgentFolder');
const btnSupervisorDir = document.getElementById('btnSupervisorDir');
const btnResetSupervisorDir = document.getElementById('btnResetSupervisorDir');
const btnNewWorkbook = document.getElementById('btnNewWorkbook');
const btnRepairSchema = document.getElementById('btnRepairSchema');
const btnAddBlankTab = document.getElementById('btnAddBlankTab');
const btnRenameTab = document.getElementById('btnRenameTab');
const btnDeleteTab = document.getElementById('btnDeleteTab');
const btnDeleteMode = document.getElementById('btnDeleteMode');

// Sheet Navigation & Controls
const btnPrevSheet = document.getElementById('btnPrevSheet');
const btnNextSheet = document.getElementById('btnNextSheet');
const sheetSelect = document.getElementById('sheetSelect');
const sheetTabsBar = document.getElementById('sheetTabsBar');
const scopeSelect = document.getElementById('scopeSelect');
const searchInput = document.getElementById('searchInput');
const searchScopeHint = document.getElementById('searchScopeHint');
const btnClearSortFilter = document.getElementById('btnClearSortFilter');
const tableHead = document.getElementById('tableHead');
const listBody = document.getElementById('listBody');

// Banners & Modals
const schemaBanner = document.getElementById('schemaBanner');
const schemaBannerText = document.getElementById('schemaBannerText');
const btnBannerBootstrap = document.getElementById('btnBannerBootstrap');
const filterBanner = document.getElementById('filterBanner');
const filterBannerTitle = document.getElementById('filterBannerTitle');
const filterBannerDesc = document.getElementById('filterBannerDesc');
const filterKeyText = document.getElementById('filterKeyText');
const editModal = document.getElementById('editModal');
const rolloverModal = document.getElementById('rolloverModal');
const toastContainer = document.getElementById('toastContainer');
const statusFolderInfo = document.getElementById('statusFolderInfo');
const statusRowStats = document.getElementById('statusRowStats');
const statusTimestamp = document.getElementById('statusTimestamp');

// Delete Mode DOM Elements
const deleteBanner = document.getElementById('deleteBanner');
const thDeleteSelect = document.getElementById('thDeleteSelect');
const chkSelectAllHeader = document.getElementById('chkSelectAllHeader');
const chkSelectAllBanner = document.getElementById('chkSelectAllBanner');
const deleteSelectedCountEl = document.getElementById('deleteSelectedCount');
const btnConfirmDeleteSelected = document.getElementById('btnConfirmDeleteSelected');
const btnCancelDeleteMode = document.getElementById('btnCancelDeleteMode');

// Form Fields
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

// In-App Dialog Modal DOM Elements
const dialogModal = document.getElementById('dialogModal');
const dialogForm = document.getElementById('dialogForm');
const dialogTitle = document.getElementById('dialogTitle');
const dialogBadge = document.getElementById('dialogBadge');
const dialogMessage = document.getElementById('dialogMessage');
const dialogInputWrap = document.getElementById('dialogInputWrap');
const dialogInputLabel = document.getElementById('dialogInputLabel');
const dialogInput = document.getElementById('dialogInput');
const dialogInputError = document.getElementById('dialogInputError');
const btnDialogCancel = document.getElementById('btnDialogCancel');
const btnDialogConfirm = document.getElementById('btnDialogConfirm');

// --- IN-APP POPUP DIALOG ENGINE (Replaces native alert / confirm / prompt) ---
let activeDialogResolver = null;
let activeDialogMode = 'alert'; // 'alert' | 'confirm' | 'prompt'
let activeDialogValidator = null;

function closeDialogModal(resultValue) {
  if (!dialogModal) return;
  dialogModal.classList.remove('open');
  dialogInputError.style.display = 'none';
  dialogInputError.textContent = '';
  const resolver = activeDialogResolver;
  activeDialogResolver = null;
  activeDialogValidator = null;
  if (resolver) {
    resolver(resultValue);
  }
}

function openDialogBase({
  mode,
  title = 'Notification',
  message = '',
  badge = '',
  confirmText = 'OK',
  cancelText = 'Cancel',
  variant = 'primary',
  inputLabel = 'Value',
  defaultValue = '',
  placeholder = '',
  validate = null
}) {
  if (activeDialogResolver) {
    closeDialogModal(mode === 'prompt' ? null : false);
  }

  return new Promise((resolve) => {
    activeDialogResolver = resolve;
    activeDialogMode = mode;
    activeDialogValidator = typeof validate === 'function' ? validate : null;

    dialogTitle.textContent = title;
    if (badge) {
      dialogBadge.textContent = badge;
      dialogBadge.style.display = 'inline-block';
    } else {
      dialogBadge.style.display = 'none';
    }

    dialogMessage.textContent = message;
    dialogInputError.style.display = 'none';
    dialogInputError.textContent = '';

    if (mode === 'prompt') {
      dialogInputWrap.style.display = 'flex';
      dialogInputLabel.textContent = inputLabel;
      dialogInput.value = defaultValue;
      dialogInput.placeholder = placeholder;
    } else {
      dialogInputWrap.style.display = 'none';
      dialogInput.value = '';
    }

    btnDialogCancel.style.display = mode === 'alert' ? 'none' : 'inline-block';
    btnDialogCancel.textContent = cancelText;

    btnDialogConfirm.textContent = confirmText;
    btnDialogConfirm.className = variant === 'danger' ? 'btn-danger' : 'btn-primary';

    dialogModal.classList.add('open');

    setTimeout(() => {
      if (mode === 'prompt') {
        dialogInput.focus();
        dialogInput.select();
      } else {
        btnDialogConfirm.focus();
      }
    }, 20);
  });
}

function showModalAlert(optionsOrMessage, title = 'TrackerUI Notice') {
  if (typeof optionsOrMessage === 'string') {
    return openDialogBase({ mode: 'alert', title, message: optionsOrMessage, confirmText: 'OK' });
  }
  return openDialogBase({ mode: 'alert', ...optionsOrMessage });
}

function showModalConfirm(optionsOrMessage, title = 'Confirm Action') {
  if (typeof optionsOrMessage === 'string') {
    return openDialogBase({ mode: 'confirm', title, message: optionsOrMessage });
  }
  return openDialogBase({ mode: 'confirm', ...optionsOrMessage });
}

function showModalPrompt(options) {
  return openDialogBase({ mode: 'prompt', ...options });
}

btnDialogCancel.addEventListener('click', () => {
  closeDialogModal(activeDialogMode === 'prompt' ? null : false);
});

dialogInput.addEventListener('input', () => {
  if (dialogInputError.style.display !== 'none') {
    dialogInputError.style.display = 'none';
  }
});

dialogForm.addEventListener('submit', (e) => {
  e.preventDefault();
  if (activeDialogMode === 'prompt') {
    const val = dialogInput.value.trim();
    if (!val) {
      dialogInputError.textContent = 'Please enter a value.';
      dialogInputError.style.display = 'block';
      dialogInput.focus();
      return;
    }
    if (activeDialogValidator) {
      const errMsg = activeDialogValidator(val);
      if (errMsg) {
        dialogInputError.textContent = errMsg;
        dialogInputError.style.display = 'block';
        dialogInput.focus();
        return;
      }
    }
    closeDialogModal(val);
  } else {
    closeDialogModal(true);
  }
});

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

// --- POPOVER TOOLS MENU LOGIC ---
function closeToolsMenu() {
  toolsMenuDropdown.classList.remove('open');
  btnToolsMenu.classList.remove('open');
}

btnToolsMenu.addEventListener('click', (e) => {
  e.stopPropagation();
  const isOpen = toolsMenuDropdown.classList.toggle('open');
  btnToolsMenu.classList.toggle('open', isOpen);
});

document.addEventListener('click', (e) => {
  if (toolsMenuWrap && !toolsMenuWrap.contains(e.target)) {
    closeToolsMenu();
  }
});

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    if (dialogModal && dialogModal.classList.contains('open')) {
      closeDialogModal(activeDialogMode === 'prompt' ? null : false);
      return;
    }
    closeToolsMenu();
  }
});

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
  updateStatusBarTelemetry();
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
async function promptSelectAgentFolder() {
  closeToolsMenu();
  try {
    if (typeof window.showDirectoryPicker !== 'function') {
      return showModalAlert({
        title: 'Browser Not Supported',
        message: 'Your browser does not support showDirectoryPicker. Please use Chrome or Edge over localhost or HTTPS.'
      });
    }
    const dirHandle = await window.showDirectoryPicker({ mode: 'readwrite' });
    parentDirHandle = dirHandle;
    cachedDirHandle = dirHandle;
    await saveHandleToIDB(IDB_AGENT_KEY, dirHandle);
    btnReconnect.style.display = 'none';
    btnLinkFile.style.display = 'none';
    btnNewWorkbook.disabled = false;
    await scanDirectoryAndLoadWorkbooks();
  } catch (err) {
    if (err.name !== 'AbortError') {
      await showModalAlert({
        title: 'Could Not Open Agent Folder',
        message: err.message
      });
    }
  }
}

btnLinkFile.addEventListener('click', promptSelectAgentFolder);
menuSelectAgentFolder.addEventListener('click', promptSelectAgentFolder);

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
      btnLinkFile.style.display = 'none';
      btnNewWorkbook.disabled = false;

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
      await showModalAlert({
        title: 'Permission Required',
        message: 'Read/Write permission was not granted for the folder.'
      });
    }
  } catch (err) {
    console.error('Reconnect folder error:', err);
    await showModalAlert({
      title: 'Could Not Reconnect Folder',
      message: err.message
    });
  }
});

btnSupervisorDir.addEventListener('click', async () => {
  closeToolsMenu();
  try {
    if (typeof window.showDirectoryPicker !== 'function') {
      return showModalAlert({
        title: 'Browser Not Supported',
        message: 'Your browser does not support showDirectoryPicker. Please use Chrome or Edge over localhost or HTTPS.'
      });
    }

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
          const buf = await serializeWorkbookBuffer();
          await attemptMirrorWrite(buf, false);
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

    if (workbook && workingFileHandle) {
      const buf = await serializeWorkbookBuffer();
      await attemptMirrorWrite(buf, false);
    }
  } catch (err) {
    if (err.name !== 'AbortError') {
      await showModalAlert({
        title: 'Supervisor Folder Error',
        message: 'Could not select Supervisor export folder: ' + err.message
      });
    }
  }
});

btnResetSupervisorDir.addEventListener('click', async () => {
  closeToolsMenu();
  supervisorDirHandle = null;
  cachedSupervisorDirHandle = null;
  await deleteHandleFromIDB(IDB_SUPERVISOR_KEY);
  updateSupervisorDirUI();
  showToast('Supervisor (VIEW ONLY) export reset to Agent workbook folder.', 'info');
});

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
  btnNewWorkbook.disabled = false;

  availableWorkbooks = [];
  for await (const entry of parentDirHandle.values()) {
    if (entry.kind !== 'file') continue;
    const name = entry.name;
    const lower = name.toLowerCase();

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
    workbook = null;
    sheetTabsBar.style.display = 'none';
    schemaBanner.style.display = 'none';
    setStatus(`No working .xlsx in ${parentDirHandle.name}`, 'warning');
    listBody.innerHTML = `
      <tr>
        <td colspan="12" style="text-align:center; color:var(--text-muted); padding: 36px;">
          <div style="margin-bottom:12px;">
            No working <code>.xlsx</code> files found in folder <strong>${escapeHtml(parentDirHandle.name)}</strong>.
          </div>
          <button type="button" id="btnEmptyStateBootstrap" class="btn-primary">
            ✨ Bootstrap New Tracker Workbook (.xlsx)
          </button>
        </td>
      </tr>
    `;
    const emptyBtn = document.getElementById('btnEmptyStateBootstrap');
    if (emptyBtn) {
      emptyBtn.addEventListener('click', promptCreateNewWorkbook);
    }
    return;
  }

  availableWorkbooks.sort((a, b) => {
    if (b.sortWeight !== a.sortWeight) return b.sortWeight - a.sortWeight;
    return b.lastModified - a.lastModified;
  });

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
    activeFieldFilter = null;
    filterBanner.style.display = 'none';
    clearMirrorRetryState();
    setDeleteMode(false);
    try {
      await loadWorkbookFromDisk();
    } catch (err) {
      console.error('Error loading selected workbook:', err);
      await showModalAlert({
        title: 'Could Not Load Workbook',
        message: err.message
      });
    }
  }
});

// --- 2. READ WORKBOOK, PARSE REFERENCE SHEETS, CHECK SCHEMA & INDEX ALL DAILY TABS ---
const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32Bytes(u8) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < u8.length; i++) {
    crc = CRC32_TABLE[(crc ^ u8[i]) & 0xFF] ^ (crc >>> 8);
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

async function inflateRawBytes(compBytes) {
  const ds = new DecompressionStream('deflate-raw');
  const writer = ds.writable.getWriter();
  writer.write(compBytes);
  writer.close();
  const ab = await new Response(ds.readable).arrayBuffer();
  return new Uint8Array(ab);
}

async function deflateRawBytes(rawBytes) {
  const cs = new CompressionStream('deflate-raw');
  const writer = cs.writable.getWriter();
  writer.write(rawBytes);
  writer.close();
  const ab = await new Response(cs.readable).arrayBuffer();
  return new Uint8Array(ab);
}

async function unpackZipArchive(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const len = bytes.length;

  let eocdOffset = -1;
  for (let i = len - 22; i >= Math.max(0, len - 65557); i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocdOffset = i;
      break;
    }
  }
  if (eocdOffset === -1) {
    throw new Error('Invalid XLSX archive: End of Central Directory not found.');
  }

  const totalEntries = view.getUint16(eocdOffset + 10, true);
  const cdOffset = view.getUint32(eocdOffset + 16, true);
  const entries = new Map();
  const dec = new TextDecoder('utf-8');

  let ptr = cdOffset;
  for (let e = 0; e < totalEntries; e++) {
    if (ptr + 46 > len || view.getUint32(ptr, true) !== 0x02014b50) break;
    const compMethod = view.getUint16(ptr + 10, true);
    const compSize = view.getUint32(ptr + 20, true);
    const nameLen = view.getUint16(ptr + 28, true);
    const extraLen = view.getUint16(ptr + 30, true);
    const commentLen = view.getUint16(ptr + 32, true);
    const localHeaderOffset = view.getUint32(ptr + 42, true);

    const name = dec.decode(bytes.subarray(ptr + 46, ptr + 46 + nameLen)).replace(/\\/g, '/');
    const localNameLen = view.getUint16(localHeaderOffset + 26, true);
    const localExtraLen = view.getUint16(localHeaderOffset + 28, true);
    const dataStart = localHeaderOffset + 30 + localNameLen + localExtraLen;
    const compData = bytes.subarray(dataStart, dataStart + compSize);

    let rawData;
    if (compMethod === 0) {
      rawData = compData.slice();
    } else if (compMethod === 8) {
      rawData = await inflateRawBytes(compData);
    } else {
      rawData = compData.slice();
    }
    entries.set(name, rawData);
    ptr += 46 + nameLen + extraLen + commentLen;
  }

  return entries;
}

async function packZipArchive(entriesMap) {
  const enc = new TextEncoder();
  const localParts = [];
  const cdParts = [];
  let localOffset = 0;
  let totalCdBytes = 0;

  for (const [name, rawData] of entriesMap.entries()) {
    const nameBytes = enc.encode(name);
    const u8Raw = rawData instanceof Uint8Array ? rawData : enc.encode(String(rawData));
    const crc = crc32Bytes(u8Raw);
    const compData = await deflateRawBytes(u8Raw);
    const compSize = compData.length;
    const uncompSize = u8Raw.length;

    const lfh = new Uint8Array(30 + nameBytes.length);
    const lfhView = new DataView(lfh.buffer);
    lfhView.setUint32(0, 0x04034b50, true);
    lfhView.setUint16(4, 20, true);
    lfhView.setUint16(6, 0, true);
    lfhView.setUint16(8, 8, true); // deflate
    lfhView.setUint16(10, 0, true);
    lfhView.setUint16(12, 0x21, true);
    lfhView.setUint32(14, crc, true);
    lfhView.setUint32(18, compSize, true);
    lfhView.setUint32(22, uncompSize, true);
    lfhView.setUint16(26, nameBytes.length, true);
    lfhView.setUint16(28, 0, true);
    lfh.set(nameBytes, 30);

    localParts.push(lfh, compData);

    const cdfh = new Uint8Array(46 + nameBytes.length);
    const cdView = new DataView(cdfh.buffer);
    cdView.setUint32(0, 0x02014b50, true);
    cdView.setUint16(4, 20, true);
    cdView.setUint16(6, 20, true);
    cdView.setUint16(8, 0, true);
    cdView.setUint16(10, 8, true);
    cdView.setUint16(12, 0, true);
    cdView.setUint16(14, 0x21, true);
    cdView.setUint32(16, crc, true);
    cdView.setUint32(20, compSize, true);
    cdView.setUint32(24, uncompSize, true);
    cdView.setUint16(28, nameBytes.length, true);
    cdView.setUint16(30, 0, true);
    cdView.setUint16(32, 0, true);
    cdView.setUint16(34, 0, true);
    cdView.setUint16(36, 0, true);
    cdView.setUint32(38, 0, true);
    cdView.setUint32(42, localOffset, true);
    cdfh.set(nameBytes, 46);

    cdParts.push(cdfh);
    totalCdBytes += cdfh.length;
    localOffset += lfh.length + compData.length;
  }

  const eocd = new Uint8Array(22);
  const eocdView = new DataView(eocd.buffer);
  const count = cdParts.length;
  eocdView.setUint32(0, 0x06054b50, true);
  eocdView.setUint16(4, 0, true);
  eocdView.setUint16(6, 0, true);
  eocdView.setUint16(8, count, true);
  eocdView.setUint16(10, count, true);
  eocdView.setUint32(12, totalCdBytes, true);
  eocdView.setUint32(16, localOffset, true);
  eocdView.setUint16(20, 0, true);

  const out = new Uint8Array(localOffset + totalCdBytes + 22);
  let pos = 0;
  for (const part of localParts) {
    out.set(part, pos);
    pos += part.length;
  }
  for (const part of cdParts) {
    out.set(part, pos);
    pos += part.length;
  }
  out.set(eocd, pos);
  return out.buffer;
}

const STANDARD_DYNAMIC_ARRAY_METADATA_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">' +
  '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>' +
  '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
  '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';

const WORKBOOK_CALC_FEATURES_EXT =
  '<ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures">' +
  '<xcalcf:calcFeatures>' +
  '<xcalcf:feature name="microsoft.com:RD"/>' +
  '<xcalcf:feature name="microsoft.com:Single"/>' +
  '<xcalcf:feature name="microsoft.com:FV"/>' +
  '<xcalcf:feature name="microsoft.com:CNMTM"/>' +
  '<xcalcf:feature name="microsoft.com:LET_WF"/>' +
  '<xcalcf:feature name="microsoft.com:LAMBDA_WF"/>' +
  '<xcalcf:feature name="microsoft.com:ARRAYTEXT_WF"/>' +
  '</xcalcf:calcFeatures></ext>';

const DAILY_SHEET_DATA_VALIDATIONS_EXTLST =
  '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
  '<x14:dataValidations count="2" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
  '<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1">' +
  '<x14:formula1><xm:f>\'Validation List\'!$E$2:$E$100</xm:f></x14:formula1><xm:sqref>I2:I101</xm:sqref>' +
  '</x14:dataValidation>' +
  '<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1">' +
  '<x14:formula1><xm:f>\'Validation List\'!$A$2:$A$50</xm:f></x14:formula1><xm:sqref>F2:F108</xm:sqref>' +
  '</x14:dataValidation>' +
  '</x14:dataValidations></ext></extLst>';

const DEFAULT_EXTERNAL_LINK_RELS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="LKQ-Bidder%20Support%20shared%20lists.xlsx" TargetMode="External"/>' +
  '</Relationships>';

function parseWorksheetPathsFromZipEntries(entries) {
  const dec = new TextDecoder('utf-8');
  const wbBytes = entries.get('xl/workbook.xml');
  const relsBytes = entries.get('xl/_rels/workbook.xml.rels');
  const map = new Map(); // sheetName -> zipPath
  if (!wbBytes || !relsBytes) return map;

  const wbXml = dec.decode(wbBytes);
  const relsXml = dec.decode(relsBytes);

  const rIdToTarget = new Map();
  const relRegex = /<Relationship\b[^>]*>/g;
  let m;
  while ((m = relRegex.exec(relsXml)) !== null) {
    const tag = m[0];
    const idMatch = tag.match(/\bId="([^"]+)"/);
    const targetMatch = tag.match(/\bTarget="([^"]+)"/);
    if (idMatch && targetMatch) {
      let target = targetMatch[1].replace(/^\/+/, '');
      if (!target.startsWith('xl/')) target = 'xl/' + target;
      rIdToTarget.set(idMatch[1], target);
    }
  }

  const sheetRegex = /<sheet\b[^>]*>/g;
  while ((m = sheetRegex.exec(wbXml)) !== null) {
    const tag = m[0];
    const nameMatch = tag.match(/\bname="([^"]+)"/);
    const rIdMatch = tag.match(/\br:id="([^"]+)"/);
    if (nameMatch && rIdMatch) {
      const sName = nameMatch[1]
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'");
      const targetPath = rIdToTarget.get(rIdMatch[1]);
      if (targetPath && entries.has(targetPath)) {
        map.set(sName, targetPath);
      }
    }
  }
  return map;
}

async function captureOriginalArchiveState(arrayBuffer) {
  try {
    const entries = await unpackZipArchive(arrayBuffer);
    const dec = new TextDecoder('utf-8');
    const externalLinkFiles = new Map();
    for (const [k, v] of entries.entries()) {
      if (k.startsWith('xl/externalLinks/')) {
        externalLinkFiles.set(k, v);
      }
    }
    const metadataXmlBytes = entries.get('xl/metadata.xml') || null;
    const sheetPaths = parseWorksheetPathsFromZipEntries(entries);
    const arrayFormulasBySheet = new Map(); // lowerSheetName -> Map(cellRef -> { fTag, cm, t })
    let templateExtLstXml = null;

    for (const [sheetName, zipPath] of sheetPaths.entries()) {
      const lower = sheetName.trim().toLowerCase();
      const xml = dec.decode(entries.get(zipPath));

      if (lower === 'template' || isTemplatePlaceholderTab(lower) || DATE_TAB_REGEX.test(sheetName.trim())) {
        const extMatch = xml.match(/<extLst>[\s\S]*?<\/extLst>/);
        if (extMatch && extMatch[0].includes('dataValidations') && !templateExtLstXml) {
          templateExtLstXml = extMatch[0].replace(/\s+xr:uid="[^"]*"/g, '');
        }
      }

      const cellMap = new Map();
      const cellRegex = /<c\b([^>]*\br="([A-Z]+\d+)"[^>]*)>(<f\b[^>]*>[\s\S]*?<\/f>)(<v[^>]*>[\s\S]*?<\/v>|<v\/>)?<\/c>/g;
      let cm;
      while ((cm = cellRegex.exec(xml)) !== null) {
        const cAttrs = cm[1];
        const cellRef = cm[2];
        const fTag = cm[3];
        if (/\bt="array"/.test(fTag) || /\bcm="/.test(cAttrs) || /_xlfn\._xlws\.FILTER|_xlfn\.LET/.test(fTag)) {
          const tAttr = (cAttrs.match(/\bt="([^"]+)"/) || [])[1] || null;
          cellMap.set(cellRef, { fTag, tAttr });
        }
      }
      if (cellMap.size > 0) {
        arrayFormulasBySheet.set(lower, cellMap);
      }
    }

    return {
      externalLinkFiles,
      metadataXmlBytes,
      arrayFormulasBySheet,
      templateExtLstXml: templateExtLstXml || DAILY_SHEET_DATA_VALIDATIONS_EXTLST
    };
  } catch (err) {
    console.warn('Could not capture original archive metadata:', err);
    return null;
  }
}

function parseSharedStringsFromXml(ssXml) {
  if (!ssXml) return [];
  const list = [];
  const siRegex = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
  let m;
  while ((m = siRegex.exec(ssXml)) !== null) {
    const inner = m[1];
    const tParts = [];
    const tRegex = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
    let tm;
    while ((tm = tRegex.exec(inner)) !== null) {
      tParts.push(tm[1]);
    }
    list.push(tParts.join(''));
  }
  return list;
}

function synthesizeExternalLinkXmlFromSheets(validationXml, masterXml) {
  const extractRowsForExternalSheet = (sheetXml) => {
    if (!sheetXml) return '';
    const rowsOut = [];
    const rowRegex = /<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
    let rm;
    while ((rm = rowRegex.exec(sheetXml)) !== null) {
      const rNum = rm[1];
      const rInner = rm[2];
      const cellsOut = [];
      const cRegex = /<c\b([^>]*\br="([A-Z]+\d+)"[^>]*)>(?:<f\b[^>]*>[\s\S]*?<\/f>)?(?:<v([^>]*)>([\s\S]*?)<\/v>|<v\/>)?<\/c>/g;
      let cm;
      while ((cm = cRegex.exec(rInner)) !== null) {
        const cAttrs = cm[1];
        const cRef = cm[2];
        const vAttrs = cm[3] || '';
        const vVal = cm[4];
        if (vVal === undefined || vVal === '') continue;
        const isStr = /\bt="str"/.test(cAttrs);
        cellsOut.push(`<cell r="${cRef}"${isStr ? ' t="str"' : ''}><v${vAttrs}>${vVal}</v></cell>`);
      }
      if (cellsOut.length > 0) {
        rowsOut.push(`<row r="${rNum}">${cellsOut.join('')}</row>`);
      }
    }
    return rowsOut.join('');
  };

  const valRows = extractRowsForExternalSheet(validationXml);
  const masterRows = extractRowsForExternalSheet(masterXml);

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1">' +
    '<sheetNames><sheetName val="Validation List"/><sheetName val="Master Account List"/></sheetNames>' +
    '<sheetDataSet>' +
    `<sheetData sheetId="0">${valRows}</sheetData>` +
    `<sheetData sheetId="1">${masterRows}</sheetData>` +
    '</sheetDataSet>' +
    '</externalBook></externalLink>'
  );
}

/**
 * Post-processes the XLSX archive produced by ExcelJS so Desktop Excel never flags
 * the workbook as corrupted or strips formulas:
 * 1. Restores dynamic array metadata (`xl/metadata.xml` + `cm="1"` on array cells) and
 *    converts ghost spill cells on `Master Accounts` & `Validation List` from `t="s"` back to `t="str"`.
 * 2. Restores `xl/externalLinks/*` and `<externalReferences>` when `[1]` external links exist.
 * 3. Ensures `Column C` XLOOKUP formula cells on `Template` and daily sheets have `t="str"` and `<v/>`.
 * 4. Restores `<extLst>` cross-sheet Data Validation dropdowns (`Requested Action` & `Status`) on daily sheets.
 */
async function repairSavedXlsxArchive(excelJsBuffer, archiveState) {
  try {
    const entries = await unpackZipArchive(excelJsBuffer);
    const dec = new TextDecoder('utf-8');
    const enc = new TextEncoder();

    const ssBytes = entries.get('xl/sharedStrings.xml');
    const sharedStrings = ssBytes ? parseSharedStringsFromXml(dec.decode(ssBytes)) : [];
    const sheetPaths = parseWorksheetPathsFromZipEntries(entries);

    let hasAnyDynamicArrayFormula = false;
    let hasAnyExternalRefFormula = false;
    let hasValidationSheetInWb = false;
    let repairedValidationXml = '';
    let repairedMasterXml = '';

    for (const [sName] of sheetPaths.entries()) {
      const low = sName.trim().toLowerCase();
      if (low === 'validation list' || low === 'validation lists') {
        hasValidationSheetInWb = true;
      }
    }

    for (const [sheetName, zipPath] of sheetPaths.entries()) {
      const lower = sheetName.trim().toLowerCase();
      let xml = dec.decode(entries.get(zipPath));

      const isDynamicArrayRefSheet =
        lower === 'master accounts' ||
        lower === 'master account' ||
        lower === 'validation list' ||
        lower === 'validation lists' ||
        /\bt="array"|_xlfn\._xlws\.FILTER|_xlfn\.LET/.test(xml);

      if (isDynamicArrayRefSheet) {
        const origCellMap = archiveState && archiveState.arrayFormulasBySheet
          ? archiveState.arrayFormulasBySheet.get(lower)
          : null;

        // 1. Restore cm="1" and exact <f t="array" ref="..."> on dynamic array anchor cells
        xml = xml.replace(
          /<c\b([^>]*\br="([A-Z]+\d+)"[^>]*)>(<f\b[^>]*>[\s\S]*?<\/f>)(<v[^>]*>([\s\S]*?)<\/v>|<v\/>)?<\/c>/g,
          (full, cAttrs, cellRef, fTag, vFull, vInner) => {
            const origInfo = origCellMap ? origCellMap.get(cellRef) : null;
            const isArrayF = Boolean(
              origInfo ||
              /\bt="array"/.test(fTag) ||
              /_xlfn\._xlws\.FILTER|_xlfn\.LET/.test(fTag)
            );
            if (!isArrayF) return full;

            hasAnyDynamicArrayFormula = true;
            const finalFTag = origInfo ? origInfo.fTag : fTag;
            if (/\[1\]/.test(finalFTag)) {
              hasAnyExternalRefFormula = true;
            }

            let attrs = cAttrs;
            if (!/\bcm="/.test(attrs)) {
              attrs += ' cm="1"';
            }
            const isNumVal = vInner !== undefined && vInner !== '' && !isNaN(Number(vInner));
            if (origInfo && origInfo.tAttr && !/\bt="/.test(attrs)) {
              attrs += ` t="${origInfo.tAttr}"`;
            } else if (!isNumVal && !/\bt="/.test(attrs)) {
              attrs += ' t="str"';
            }
            return `<c${attrs}>${finalFTag}${vFull || '<v/>'}</c>`;
          }
        );

        // 2. Convert spilled string cells in rows 2+ from t="s" (shared string) back to t="str" (formula string result)
        xml = xml.replace(
          /<row\b([^>]*\br="(\d+)"[^>]*)>([\s\S]*?)<\/row>/g,
          (rowFull, rowAttrs, rNumStr, rowInner) => {
            const rNum = parseInt(rNumStr, 10);
            if (rNum < 2) return rowFull;
            const fixedInner = rowInner.replace(
              /<c\b([^>]*?)\bt="s"([^>]*)><v>(\d+)<\/v><\/c>/g,
              (cFull, a1, a2, idxStr) => {
                const idx = parseInt(idxStr, 10);
                const resolved = sharedStrings[idx] !== undefined ? sharedStrings[idx] : '';
                const needsSpace = resolved.startsWith(' ') || resolved.endsWith(' ');
                return `<c${a1}t="str"${a2}><v${needsSpace ? ' xml:space="preserve"' : ''}>${resolved}</v></c>`;
              }
            );
            return `<row${rowAttrs}>${fixedInner}</row>`;
          }
        );

        if (lower === 'validation list' || lower === 'validation lists') {
          repairedValidationXml = xml;
        } else if (lower === 'master accounts' || lower === 'master account') {
          repairedMasterXml = xml;
        }
      }

      // 3. For Template & Daily Sheets: ensure Column C XLOOKUP formulas have t="str" and <v/>, and restore Data Validations <extLst>
      const isDailyOrTemplate =
        lower === 'template' ||
        isTemplatePlaceholderTab(lower) ||
        DATE_TAB_REGEX.test(sheetName.trim());

      if (isDailyOrTemplate) {
        xml = xml.replace(
          /<c\b([^>]*\br="C(\d+)"[^>]*)>(<f\b[^>]*>[\s\S]*?<\/f>)(<v[^>]*>[\s\S]*?<\/v>|<v\/>)?<\/c>/g,
          (full, cAttrs, rNumStr, fTag, vFull) => {
            let cleanAttrs = cAttrs.replace(/\s+cm="[^"]*"/g, '');
            if (/\bt="[^"]*"/.test(cleanAttrs)) {
              cleanAttrs = cleanAttrs.replace(/\bt="[^"]*"/, 't="str"');
            } else {
              cleanAttrs += ' t="str"';
            }
            const cleanFTag = fTag.replace(/^<f\b[^>]*>/, '<f>');
            let cleanV = vFull || '<v/>';
            if (cleanV === '<v></v>') cleanV = '<v/>';
            return `<c${cleanAttrs}>${cleanFTag}${cleanV}</c>`;
          }
        );

        if (hasValidationSheetInWb && !xml.includes('dataValidations')) {
          const extLstToInject = (archiveState && archiveState.templateExtLstXml)
            ? archiveState.templateExtLstXml
            : DAILY_SHEET_DATA_VALIDATIONS_EXTLST;
          if (!xml.includes('<extLst>')) {
            xml = xml.replace(/<\/worksheet>\s*$/, `${extLstToInject}</worksheet>`);
          }
        }
      }

      entries.set(zipPath, enc.encode(xml));
    }

    // 4. Remove stale calcChain.xml if present so Excel cleanly rebuilds the calculation chain
    entries.delete('xl/calcChain.xml');

    // 5. Restore or synthesize xl/externalLinks/* and xl/metadata.xml when needed
    let ctXml = dec.decode(entries.get('[Content_Types].xml'));
    let relsXml = dec.decode(entries.get('xl/_rels/workbook.xml.rels'));
    let wbXml = dec.decode(entries.get('xl/workbook.xml'));

    ctXml = ctXml.replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g, '');
    relsXml = relsXml.replace(/<Relationship[^>]*Target="calcChain\.xml"[^>]*\/>/g, '');

    let maxRId = 0;
    const rIdMatches = relsXml.match(/\bId="rId(\d+)"/g) || [];
    for (const rm of rIdMatches) {
      const n = parseInt(rm.replace(/\D+/g, ''), 10);
      if (n > maxRId) maxRId = n;
    }

    if (hasAnyExternalRefFormula) {
      if (archiveState && archiveState.externalLinkFiles && archiveState.externalLinkFiles.size > 0) {
        for (const [extPath, extBytes] of archiveState.externalLinkFiles.entries()) {
          entries.set(extPath, extBytes);
        }
      } else {
        entries.set(
          'xl/externalLinks/externalLink1.xml',
          enc.encode(synthesizeExternalLinkXmlFromSheets(repairedValidationXml, repairedMasterXml))
        );
        entries.set(
          'xl/externalLinks/_rels/externalLink1.xml.rels',
          enc.encode(DEFAULT_EXTERNAL_LINK_RELS_XML)
        );
      }

      if (!ctXml.includes('/xl/externalLinks/externalLink1.xml')) {
        ctXml = ctXml.replace(
          /<\/Types>\s*$/,
          '<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>'
        );
      }

      let extRId = null;
      const existingExtRel = relsXml.match(/<Relationship\b[^>]*\bId="(rId\d+)"[^>]*externalLinks\/externalLink1\.xml[^>]*\/>/);
      if (existingExtRel) {
        extRId = existingExtRel[1];
      } else {
        extRId = `rId${++maxRId}`;
        relsXml = relsXml.replace(
          /<\/Relationships>\s*$/,
          `<Relationship Id="${extRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>`
        );
      }

      if (!wbXml.includes('<externalReferences')) {
        wbXml = wbXml.replace(
          /<\/sheets>/,
          `</sheets><externalReferences><externalReference r:id="${extRId}"/></externalReferences>`
        );
      }
    }

    if (hasAnyDynamicArrayFormula) {
      const metaBytes = (archiveState && archiveState.metadataXmlBytes)
        ? archiveState.metadataXmlBytes
        : enc.encode(STANDARD_DYNAMIC_ARRAY_METADATA_XML);
      entries.set('xl/metadata.xml', metaBytes);

      if (!ctXml.includes('/xl/metadata.xml')) {
        ctXml = ctXml.replace(
          /<\/Types>\s*$/,
          '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/></Types>'
        );
      }

      if (!relsXml.includes('Target="metadata.xml"')) {
        const metaRId = `rId${++maxRId}`;
        relsXml = relsXml.replace(
          /<\/Relationships>\s*$/,
          `<Relationship Id="${metaRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>`
        );
      }
    }

    if (!wbXml.includes('xcalcf:calcFeatures')) {
      if (wbXml.includes('</extLst>')) {
        wbXml = wbXml.replace(/<\/extLst>/, `${WORKBOOK_CALC_FEATURES_EXT}</extLst>`);
      } else {
        wbXml = wbXml.replace(/<\/workbook>\s*$/, `<extLst>${WORKBOOK_CALC_FEATURES_EXT}</extLst></workbook>`);
      }
    }

    entries.set('[Content_Types].xml', enc.encode(ctXml));
    entries.set('xl/_rels/workbook.xml.rels', enc.encode(relsXml));
    entries.set('xl/workbook.xml', enc.encode(wbXml));

    return await packZipArchive(entries);
  } catch (err) {
    console.warn('OpenXML post-processor warning (falling back to ExcelJS buffer):', err);
    return excelJsBuffer;
  }
}

/**
 * Neutralizes docProps/core.xml and docProps/app.xml in the in-memory ZIP headers
 * before passing to ExcelJS. ExcelJS's CoreXform / AppXform parsers throw
 * `Unexpected xml node in parseOpen: {"name":"lastModifiedBy",...}` when reading
 * workbooks whose docProps/core.xml uses a default XML namespace without `cp:` prefixes.
 */
function sanitizeXlsxBufferForExcelJS(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer.slice(0));
  const len = bytes.length;

  for (let i = 0; i <= len - 30; i++) {
    if (bytes[i] !== 0x50 || bytes[i + 1] !== 0x4b) continue;

    const b2 = bytes[i + 2];
    const b3 = bytes[i + 3];
    let nameLen = 0;
    let nameOffset = 0;

    if (b2 === 0x03 && b3 === 0x04) {
      nameLen = bytes[i + 26] | (bytes[i + 27] << 8);
      nameOffset = i + 30;
    } else if (b2 === 0x01 && b3 === 0x02 && i + 46 <= len) {
      nameLen = bytes[i + 28] | (bytes[i + 29] << 8);
      nameOffset = i + 46;
    } else {
      continue;
    }

    if (nameLen <= 0 || nameLen > 256 || nameOffset + nameLen > len) continue;

    let rawName = '';
    for (let k = 0; k < nameLen; k++) {
      rawName += String.fromCharCode(bytes[nameOffset + k]);
    }
    const normName = rawName.replace(/\\/g, '/').replace(/^\/+/, '').toLowerCase();

    if (normName === 'docprops/core.xml') {
      const replacement = 'skip.xml';
      const targetStart = nameOffset + nameLen - replacement.length;
      for (let k = 0; k < replacement.length; k++) {
        bytes[targetStart + k] = replacement.charCodeAt(k);
      }
    } else if (normName === 'docprops/app.xml') {
      const replacement = 'skp.xml';
      const targetStart = nameOffset + nameLen - replacement.length;
      for (let k = 0; k < replacement.length; k++) {
        bytes[targetStart + k] = replacement.charCodeAt(k);
      }
    }
  }

  return bytes.buffer;
}

async function loadWorkbookFromDisk(preserveSelectedSheet = null, options = {}) {
  if (!workingFileHandle) return;
  if (!options.preserveStatus && !pendingMirrorBuffer) {
    setStatus('Reading...', '');
  }

  const file = await workingFileHandle.getFile();
  const arrayBuffer = await file.arrayBuffer();
  originalArchiveState = await captureOriginalArchiveState(arrayBuffer);
  const sanitizedBuffer = sanitizeXlsxBufferForExcelJS(arrayBuffer);

  const loadedWorkbook = new ExcelJS.Workbook();
  await loadedWorkbook.xlsx.load(sanitizedBuffer);
  workbook = loadedWorkbook;

  allRows = [];
  dateSheets = [];
  masterAccountsMap = [];
  validationActions = [];
  validationStatuses = [];

  let hasTemplate = false;
  let hasValidation = false;
  let hasMasterAccounts = false;
  let placeholderDateSheet = null;
  const dailySheetObjs = [];

  // Pass 1: Parse Reference Sheets (Validation List & Master Accounts) first so XLOOKUP Broker names can resolve
  workbook.eachSheet((worksheet) => {
    const rawName = worksheet.name.trim();
    const lowerName = rawName.toLowerCase();

    if (DATE_TAB_REGEX.test(rawName)) {
      dateSheets.push(rawName);
      dailySheetObjs.push({ worksheet, rawName });
    } else if (isTemplatePlaceholderTab(rawName)) {
      placeholderDateSheet = worksheet;
    } else if (lowerName === 'template') {
      hasTemplate = true;
    } else if (lowerName === 'validation lists' || lowerName === 'validation list') {
      hasValidation = true;
      parseValidationSheet(worksheet);
    } else if (lowerName === 'master accounts' || lowerName === 'master account') {
      hasMasterAccounts = true;
      parseMasterAccountsSheet(worksheet);
    }
  });

  // Pass 2: Index Daily MM-DD-YYYY Tabs (and if none exist yet but "Rename to Current Date" has entries, index it too)
  dailySheetObjs.forEach(({ worksheet, rawName }) => {
    indexWorksheetRows(worksheet, rawName);
  });

  populateReferenceControls();

  // Sort date tabs newest first for UI dropdown
  dateSheets.sort((a, b) => parseDateTab(b) - parseDateTab(a));

  populateSheetDropdown(preserveSelectedSheet);
  renderSheetTabsBar();
  updateSheetStepperState();

  sheetSelect.disabled = dateSheets.length === 0;
  scopeSelect.disabled = dateSheets.length === 0;
  searchInput.disabled = dateSheets.length === 0;

  btnNewEntry.style.display = 'inline-flex';
  btnRollover.style.display = 'inline-block';
  btnRepairSchema.disabled = false;
  btnAddBlankTab.disabled = false;
  btnRenameTab.disabled = dateSheets.length === 0;
  btnDeleteTab.disabled = dateSheets.length === 0;
  btnDeleteMode.disabled = allRows.length === 0;

  // Check if workbook is a fresh template with "Rename to Current Date" or missing reference tabs
  const missingParts = [];
  if (dateSheets.length === 0 && !placeholderDateSheet) missingParts.push('Daily MM-DD-YYYY Tab');
  if (!hasTemplate) missingParts.push('template');
  if (!hasValidation) missingParts.push('validation lists');
  if (!hasMasterAccounts) missingParts.push('Master Accounts');

  if (dateSheets.length === 0 && placeholderDateSheet && missingParts.length === 0) {
    const todayStr = getTodayFormatted();
    schemaBanner.style.display = 'flex';
    schemaBannerText.innerHTML = `
      <strong>📅 New Tracker Template Ready:</strong> Found initial tab <code>${escapeHtml(placeholderDateSheet.name)}</code>.
      Click <strong>Set Today's Date (${escapeHtml(todayStr)})</strong> to rename it to today's tab with all formulas &amp; dropdowns intact.
    `;
    btnBannerBootstrap.textContent = `📅 Set Today's Date (${todayStr})`;
  } else if (missingParts.length > 0) {
    schemaBanner.style.display = 'flex';
    schemaBannerText.innerHTML = `
      <strong>✨ Workbook Schema Incomplete:</strong> Missing <code>${escapeHtml(missingParts.join(', '))}</code>.
      Click Bootstrap to initialize all missing standard tabs and headers non-destructively.
    `;
    btnBannerBootstrap.textContent = '🛠️ Bootstrap Schema';
  } else {
    schemaBanner.style.display = 'none';
    btnBannerBootstrap.textContent = '🛠️ Bootstrap Schema';
  }

  if (!options.preserveStatus && !pendingMirrorBuffer) {
    setStatus(`Synced: ${file.name}`, 'ready');
  }
  renderList();
}

function parseDateTab(tabName) {
  const parts = String(tabName || '').trim().split('-').map(Number);
  if (parts.length !== 3) return 0;
  let [mm, dd, yyyy] = parts;
  if (yyyy < 100) yyyy += 2000;
  return new Date(yyyy, mm - 1, dd).getTime();
}

function addDaysToDateTab(tabName, daysToAdd) {
  const parts = String(tabName || '').trim().split('-').map(Number);
  let [mm, dd, yyyy] = parts;
  if (yyyy < 100) yyyy += 2000;
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
    if (val instanceof Date) return val.toISOString().split('T')[0];
    if (Array.isArray(val.richText)) return val.richText.map(t => t.text || '').join('').trim();
    if (val.text !== undefined && val.text !== null && typeof val.text !== 'object') {
      const s = String(val.text).trim();
      return s === '[object Object]' ? '' : s;
    }
    if (val.formula !== undefined || val.sharedFormula !== undefined) {
      if (val.result === undefined || val.result === null || typeof val.result === 'object') {
        return '';
      }
      const s = String(val.result).trim();
      return s === '[object Object]' ? '' : s;
    }
    if (val.result !== undefined && val.result !== null && typeof val.result !== 'object') {
      const s = String(val.result).trim();
      return s === '[object Object]' ? '' : s;
    }
    return '';
  }
  const str = String(val).trim();
  return str === '[object Object]' ? '' : str;
}

function getWorksheetColumnMap(worksheet) {
  const row1 = worksheet.getRow(1);
  const h9 = cellText(row1, 9).toLowerCase();
  const h10 = cellText(row1, 10).toLowerCase();
  const h11 = cellText(row1, 11).toLowerCase();

  const isTemplateLayout =
    h9.includes('status') ||
    h10.includes('action taken') ||
    h11.includes('sent email');

  if (isTemplateLayout) {
    return {
      isTemplateLayout: true,
      statusCol: 9,     // Col I: Status
      notesCol: 10,     // Col J: Action Taken/Notes
      emailSubjCol: 11, // Col K: Sent Email Subject (remains visible)
      trackingCol: 12   // Col L: Hidden Tracking ID (UUIDv4)
    };
  }

  return {
    isTemplateLayout: false,
    notesCol: 9,        // Col I: Action Notes
    statusCol: 10,      // Col J: Request Status
    emailSubjCol: null,
    trackingCol: 11     // Col K: Hidden Tracking ID (UUIDv4)
  };
}

function buildTemplateBrokerFormula(rowNumber, brokerVal, masterVal) {
  let fallbackLiteral = '""';
  const cleanBroker = String(brokerVal || '').trim();
  const cleanMaster = String(masterVal || '').trim();

  if (cleanBroker) {
    const matchedInMaster = cleanMaster && masterAccountsMap.some(
      m => m.master.toLowerCase() === cleanMaster.toLowerCase()
    );
    if (!matchedInMaster) {
      fallbackLiteral = `"${cleanBroker.replace(/"/g, '""')}"`;
    }
  }
  return `IF($D${rowNumber}="", ${fallbackLiteral}, IF(_xlfn.XLOOKUP($D${rowNumber}, 'Master Accounts'!A:A, 'Master Accounts'!B:B, "")="", ${fallbackLiteral}, _xlfn.XLOOKUP($D${rowNumber}, 'Master Accounts'!A:A, 'Master Accounts'!B:B, "")))`;
}

function resolveBrokerFromMaster(masterVal, currentBrokerVal) {
  const cleanBroker = String(currentBrokerVal || '').trim();
  if (cleanBroker) return cleanBroker;
  const cleanMaster = String(masterVal || '').trim().toLowerCase();
  if (!cleanMaster) return '';
  const match = masterAccountsMap.find(m => m.master.toLowerCase() === cleanMaster);
  return match ? match.broker : '';
}

function setBrokerCellWithFormula(row, rowNumber, brokerVal, masterVal, colMap) {
  const resolvedBroker = resolveBrokerFromMaster(masterVal, brokerVal);
  if (colMap && colMap.isTemplateLayout) {
    row.getCell(3).value = {
      formula: buildTemplateBrokerFormula(rowNumber, resolvedBroker, masterVal),
      result: resolvedBroker || ''
    };
  } else {
    row.getCell(3).value = resolvedBroker;
  }
}

// --- 3. WORKBOOK BOOTSTRAPPING, SCHEMA REPAIR & TAB MANAGEMENT ---
function applyStandardSheetFormatting(ws) {
  const widths = [30, 16, 22, 14, 14, 22, 10, 20, 36, 16, 38];
  const headerRow = ws.getRow(1);
  headerRow.height = 22;

  STANDARD_HEADERS.forEach((title, idx) => {
    const colIdx = idx + 1;
    ws.getColumn(colIdx).width = widths[idx] || 16;
    const cell = headerRow.getCell(colIdx);
    cell.value = title;
    cell.font = { bold: true, color: { argb: 'FFF8FAFC' }, size: 11 };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1E293B' }
    };
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
  });

  headerRow.commit();
  ws.getColumn(11).hidden = true;
}

function bootstrapWorkbookSchemaInMemory(wb) {
  const todayStr = getTodayFormatted();
  let templateWs = null;
  let validationWs = null;
  let masterWs = null;
  let placeholderDateWs = null;
  let hasAnyDateTab = false;

  wb.eachSheet((ws) => {
    const raw = ws.name.trim();
    const lower = raw.toLowerCase();
    if (DATE_TAB_REGEX.test(raw)) hasAnyDateTab = true;
    else if (isTemplatePlaceholderTab(raw)) placeholderDateWs = ws;
    else if (lower === 'template') templateWs = ws;
    else if (lower === 'validation lists' || lower === 'validation list') validationWs = ws;
    else if (lower === 'master accounts' || lower === 'master account') masterWs = ws;
  });

  // 1. Ensure template sheet exists (preserve existing Template columns if already present)
  if (!templateWs) {
    templateWs = wb.addWorksheet('template');
    applyStandardSheetFormatting(templateWs);
  } else {
    const tplMap = getWorksheetColumnMap(templateWs);
    templateWs.getColumn(tplMap.trackingCol).hidden = true;
    if (tplMap.emailSubjCol) {
      templateWs.getColumn(tplMap.emailSubjCol).hidden = false;
    }
  }

  // 2. Ensure validation lists sheet exists
  if (!validationWs) {
    validationWs = wb.addWorksheet('validation lists');
    validationWs.getColumn(1).width = 26;
    validationWs.getColumn(2).width = 22;
    const hRow = validationWs.getRow(1);
    hRow.getCell(1).value = 'Requested Actions';
    hRow.getCell(2).value = 'Status Options';
    hRow.font = { bold: true };
    hRow.commit();

    const maxLen = Math.max(DEFAULT_ACTIONS.length, DEFAULT_STATUSES.length);
    for (let i = 0; i < maxLen; i++) {
      const r = validationWs.getRow(i + 2);
      if (DEFAULT_ACTIONS[i]) r.getCell(1).value = DEFAULT_ACTIONS[i];
      if (DEFAULT_STATUSES[i]) r.getCell(2).value = DEFAULT_STATUSES[i];
      r.commit();
    }
  }

  // 3. Ensure Master Accounts sheet exists
  if (!masterWs) {
    masterWs = wb.addWorksheet('Master Accounts');
    masterWs.getColumn(1).width = 28;
    masterWs.getColumn(2).width = 16;
    masterWs.getColumn(3).width = 18;
    const hRow = masterWs.getRow(1);
    hRow.getCell(1).value = 'Broker Name';
    hRow.getCell(2).value = 'Master#';
    hRow.getCell(3).value = 'Default Bidder#';
    hRow.font = { bold: true };
    hRow.commit();
  }

  // 4. Ensure at least one MM-DD-YYYY daily tab exists:
  //    If the workbook has "Rename to Current Date", rename it in-place to today's date so all 74 template rows & formulas stay intact!
  let createdDateTab = null;
  if (!hasAnyDateTab) {
    createdDateTab = todayStr;
    if (placeholderDateWs) {
      placeholderDateWs.name = todayStr;
      const colMap = getWorksheetColumnMap(placeholderDateWs);
      placeholderDateWs.getColumn(colMap.trackingCol).hidden = true;
      if (colMap.emailSubjCol) {
        placeholderDateWs.getColumn(colMap.emailSubjCol).hidden = false;
      }
    } else {
      const dailyWs = wb.addWorksheet(todayStr);
      cloneTemplateStructure(templateWs, dailyWs);
      placeWorksheetFirstInWorkbook(wb, dailyWs);
    }
  }

  // 5. Remove default empty "Sheet1" if it is completely blank and other sheets exist
  const sheet1 = wb.getWorksheet('Sheet1');
  if (sheet1 && wb.worksheets.length > 1) {
    let sheet1HasData = false;
    sheet1.eachRow({ includeEmpty: false }, (row) => {
      for (let c = 1; c <= 11; c++) {
        if (cellText(row, c) !== '') sheet1HasData = true;
      }
    });
    if (!sheet1HasData) {
      wb.removeWorksheet(sheet1.id);
    }
  }

  return createdDateTab;
}

async function promptCreateNewWorkbook() {
  closeToolsMenu();
  if (!parentDirHandle) {
    return showModalAlert({
      title: 'Select Agent Folder First',
      message: 'Please connect an Agent Folder before creating a new workbook.'
    });
  }

  const now = new Date();
  const monthName = now.toLocaleString('en-US', { month: 'long' });
  const defaultFileName = `Agent Bidder Support Tracker - ${monthName} ${now.getFullYear()}.xlsx`;

  const rawName = await showModalPrompt({
    title: '✨ Create New Monthly Workbook',
    badge: parentDirHandle.name,
    message: 'Enter a filename for the new Excel tracker workbook (.xlsx). It will be initialized with template, validation lists, Master Accounts, and today\'s daily tab.',
    inputLabel: 'Workbook Filename (.xlsx)',
    defaultValue: defaultFileName,
    confirmText: '✨ Create Workbook',
    validate: (val) => {
      let checkName = val.trim();
      if (!checkName.toLowerCase().endsWith('.xlsx')) checkName += '.xlsx';
      if (checkName.toLowerCase().endsWith('(view only).xlsx') || checkName.startsWith('~$')) {
        return 'Filename cannot start with "~$" or end with "(VIEW ONLY).xlsx".';
      }
      return null;
    }
  });
  if (!rawName) return;

  let fileName = rawName.trim();
  if (!fileName.toLowerCase().endsWith('.xlsx')) {
    fileName += '.xlsx';
  }

  try {
    let exists = false;
    try {
      await parentDirHandle.getFileHandle(fileName, { create: false });
      exists = true;
    } catch (_) {
      exists = false;
    }

    if (exists) {
      const overwrite = await showModalConfirm({
        title: '⚠️ Workbook Already Exists',
        badge: fileName,
        message: `"${fileName}" already exists in "${parentDirHandle.name}".\n\nDo you want to overwrite it with a fresh workbook?`,
        confirmText: 'Overwrite Workbook',
        variant: 'danger'
      });
      if (!overwrite) return;
    }

    originalArchiveState = null;
    const newWb = new ExcelJS.Workbook();
    bootstrapWorkbookSchemaInMemory(newWb);

    const newFileHandle = await parentDirHandle.getFileHandle(fileName, { create: true });
    workingFileHandle = newFileHandle;
    workbook = newWb;

    const saved = await saveWorkbookToDisk();
    if (saved) {
      showToast(`Bootstrapped new workbook: "${fileName}"`, 'success');
      await scanDirectoryAndLoadWorkbooks(fileName);
    }
  } catch (err) {
    console.error('Error creating new workbook:', err);
    await showModalAlert({
      title: 'Could Not Create Workbook',
      message: err.message
    });
  }
}

async function handleRepairCurrentWorkbookSchema() {
  closeToolsMenu();
  if (!workbook || !workingFileHandle) return;

  await loadWorkbookFromDisk(sheetSelect.value);
  const createdTab = bootstrapWorkbookSchemaInMemory(workbook);

  // Ensure any existing MM-DD-YYYY tabs have headers on Row 1 and proper tracking column visibility
  dateSheets.forEach(tabName => {
    const ws = workbook.getWorksheet(tabName);
    if (!ws) return;
    const row1 = ws.getRow(1);
    const hasHeader = [1, 2, 3].some(c => cellText(row1, c) !== '');
    if (!hasHeader) {
      const templateSheet = findTemplateSheet();
      if (templateSheet) {
        cloneTemplateStructure(templateSheet, ws);
      } else {
        applyStandardSheetFormatting(ws);
      }
    } else {
      const colMap = getWorksheetColumnMap(ws);
      ws.getColumn(colMap.trackingCol).hidden = true;
      if (colMap.emailSubjCol) {
        ws.getColumn(colMap.emailSubjCol).hidden = false;
      }
    }
  });

  const targetSelectTab = createdTab || sheetSelect.value;
  const saved = await saveWorkbookToDisk();
  if (saved) {
    await loadWorkbookFromDisk(targetSelectTab, {
      preserveStatus: saved === 'mirror-locked'
    });
    showToast(
      createdTab
        ? `Initialized daily tab "${createdTab}" with all formulas & validations intact.`
        : 'Workbook schema verified.',
      'success'
    );
  }
}

btnNewWorkbook.addEventListener('click', promptCreateNewWorkbook);
btnRepairSchema.addEventListener('click', handleRepairCurrentWorkbookSchema);
btnBannerBootstrap.addEventListener('click', handleRepairCurrentWorkbookSchema);

// --- SHEET CREATION, RENAMING & DELETION ---
btnAddBlankTab.addEventListener('click', async () => {
  closeToolsMenu();
  if (!workbook) return;

  const latestSheet = dateSheets.length > 0 ? dateSheets[0] : getTodayFormatted();
  const suggestedDate = dateSheets.includes(getTodayFormatted())
    ? addDaysToDateTab(latestSheet, 1)
    : getTodayFormatted();

  const rawInput = await showModalPrompt({
    title: '➕ New Blank Day Tab',
    badge: 'MM-DD-YYYY',
    message: 'Enter the date for the new blank daily worksheet tab (cloned from your template headers):',
    inputLabel: 'Daily Tab Date (MM-DD-YYYY)',
    defaultValue: suggestedDate,
    placeholder: 'MM-DD-YYYY',
    confirmText: '➕ Create Tab',
    validate: (val) => {
      if (!DATE_TAB_REGEX.test(val)) {
        return 'Please use exact MM-DD-YYYY format (e.g., 10-25-2026).';
      }
      if (workbook.getWorksheet(val)) {
        return `Worksheet tab "${val}" already exists in this workbook.`;
      }
      return null;
    }
  });
  if (!rawInput) return;
  const newTabName = rawInput.trim();

  await loadWorkbookFromDisk(sheetSelect.value);
  if (workbook.getWorksheet(newTabName)) {
    return showModalAlert({
      title: 'Tab Already Exists',
      message: `Worksheet tab "${newTabName}" already exists!`
    });
  }

  let placeholderWs = null;
  workbook.eachSheet(ws => {
    if (isTemplatePlaceholderTab(ws.name)) placeholderWs = ws;
  });

  if (placeholderWs && dateSheets.length === 0) {
    placeholderWs.name = newTabName;
  } else {
    const templateSheet = findTemplateSheet();
    const newSheet = workbook.addWorksheet(newTabName);
    placeWorksheetFirst(newSheet);

    if (templateSheet) {
      cloneTemplateStructure(templateSheet, newSheet);
    } else {
      applyStandardSheetFormatting(newSheet);
    }
  }

  const saved = await saveWorkbookToDisk();
  if (saved) {
    await loadWorkbookFromDisk(newTabName, {
      preserveStatus: saved === 'mirror-locked'
    });
    showToast(`Created blank daily tab "${newTabName}".`, 'success');
  }
});

btnRenameTab.addEventListener('click', async () => {
  closeToolsMenu();
  const currentTab = sheetSelect.value;
  if (!workbook || !currentTab || !DATE_TAB_REGEX.test(currentTab)) {
    return showModalAlert({
      title: 'No Daily Tab Selected',
      message: 'Select a valid daily MM-DD-YYYY tab to rename.'
    });
  }

  const rawNewName = await showModalPrompt({
    title: '✏️ Rename Current Day Tab',
    badge: currentTab,
    message: `Enter a new date name (MM-DD-YYYY) for worksheet tab "${currentTab}":`,
    inputLabel: 'New Tab Name (MM-DD-YYYY)',
    defaultValue: currentTab,
    placeholder: 'MM-DD-YYYY',
    confirmText: '✏️ Rename Tab',
    validate: (val) => {
      if (!DATE_TAB_REGEX.test(val)) {
        return 'New tab name must use exact MM-DD-YYYY format.';
      }
      if (val !== currentTab && workbook.getWorksheet(val)) {
        return `A worksheet named "${val}" already exists in this workbook.`;
      }
      return null;
    }
  });
  if (!rawNewName) return;
  const newTabName = rawNewName.trim();

  if (newTabName === currentTab) return;

  await loadWorkbookFromDisk(currentTab);
  if (workbook.getWorksheet(newTabName)) {
    return showModalAlert({
      title: 'Tab Already Exists',
      message: `A worksheet named "${newTabName}" already exists in this workbook.`
    });
  }

  const ws = workbook.getWorksheet(currentTab);
  if (!ws) {
    return showModalAlert({
      title: 'Worksheet Not Found',
      message: `Worksheet "${currentTab}" was not found.`
    });
  }

  ws.name = newTabName;
  const saved = await saveWorkbookToDisk();
  if (saved) {
    await loadWorkbookFromDisk(newTabName, {
      preserveStatus: saved === 'mirror-locked'
    });
    showToast(`Renamed tab "${currentTab}" → "${newTabName}".`, 'success');
  }
});

btnDeleteTab.addEventListener('click', async () => {
  closeToolsMenu();
  const currentTab = sheetSelect.value;
  if (!workbook || !currentTab || !DATE_TAB_REGEX.test(currentTab)) {
    return showModalAlert({
      title: 'No Daily Tab Selected',
      message: 'Select a valid daily MM-DD-YYYY tab to delete.'
    });
  }

  const rowsOnTab = allRows.filter(r => r.sheetName === currentTab);
  const confirmed = await showModalConfirm({
    title: `🗑️ Delete Daily Tab "${currentTab}"?`,
    badge: `${rowsOnTab.length} request(s)`,
    message:
      `Are you sure you want to permanently delete worksheet tab "${currentTab}"?\n\n` +
      `This tab currently contains ${rowsOnTab.length} request(s).\n` +
      `Reference tabs (template, validation lists, Master Accounts) will not be affected.`,
    confirmText: '🗑️ Delete Tab',
    variant: 'danger'
  });
  if (!confirmed) return;

  await loadWorkbookFromDisk(currentTab);
  const ws = workbook.getWorksheet(currentTab);
  if (!ws) {
    return showModalAlert({
      title: 'Worksheet Not Found',
      message: `Worksheet "${currentTab}" was not found.`
    });
  }

  if (workbook.worksheets.length <= 1) {
    const fallbackTemplate = workbook.addWorksheet('template');
    applyStandardSheetFormatting(fallbackTemplate);
  }

  workbook.removeWorksheet(ws.id);

  const saved = await saveWorkbookToDisk();
  if (saved) {
    await loadWorkbookFromDisk(null, {
      preserveStatus: saved === 'mirror-locked'
    });
    showToast(`Deleted worksheet tab "${currentTab}".`, 'success');
  }
});

// --- 4. DYNAMIC REFERENCE SHEET PARSING (validation lists & Master Accounts) ---
function parseValidationSheet(worksheet) {
  const actionsSet = new Set();
  const statusesColB = new Set();
  const statusesColE = new Set();

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const actionVal = cellText(row, 1);
    const statusValB = cellText(row, 2);
    const statusValE = cellText(row, 5);

    if (actionVal) actionsSet.add(actionVal);
    if (statusValB) statusesColB.add(statusValB);
    if (statusValE) statusesColE.add(statusValE);
  });

  validationActions = Array.from(actionsSet);
  validationStatuses = statusesColE.size > 0
    ? Array.from(statusesColE)
    : Array.from(statusesColB);
}

function parseMasterAccountsSheet(worksheet) {
  const seen = new Set();
  const row1 = worksheet.getRow(1);
  const h1 = cellText(row1, 1).toLowerCase();
  const h2 = cellText(row1, 2).toLowerCase();
  const h3 = cellText(row1, 3).toLowerCase();

  // In "Tracker Template - Make a Copy.xlsx", Col A is "Master Account" and Col B is "Account Name"
  const isMasterFirst =
    h1.includes('master') ||
    h2.includes('account name') ||
    (h2.includes('broker') && !h1.includes('broker'));
  const hasBidderInCol3 = h3.includes('bidder');

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    let broker = '';
    let master = '';
    let bidder = '';

    if (isMasterFirst) {
      master = cellText(row, 1).toUpperCase();
      broker = cellText(row, 2);
      bidder = hasBidderInCol3 ? cellText(row, 3).toUpperCase() : '';
    } else {
      broker = cellText(row, 1);
      master = cellText(row, 2).toUpperCase();
      bidder = cellText(row, 3).toUpperCase();
    }

    if (!broker && !master) return;
    const dedupeKey = `${broker.toLowerCase()}::${master.toLowerCase()}`;
    if (seen.has(dedupeKey)) return;
    seen.add(dedupeKey);

    masterAccountsMap.push({ broker, master, bidder });
  });
}

function populateReferenceControls() {
  const actionsToUse = validationActions.length > 0 ? validationActions : DEFAULT_ACTIONS;
  actionListEl.innerHTML = '';
  actionsToUse.forEach(action => {
    const opt = document.createElement('option');
    opt.value = action;
    actionListEl.appendChild(opt);
  });

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

  brokerListEl.innerHTML = '';
  const uniqueBrokers = Array.from(new Set(masterAccountsMap.map(m => m.broker).filter(Boolean)));
  uniqueBrokers.forEach(broker => {
    const opt = document.createElement('option');
    opt.value = broker;
    brokerListEl.appendChild(opt);
  });
}

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
  const colMap = getWorksheetColumnMap(worksheet);
  worksheet.getColumn(colMap.trackingCol).hidden = true;
  if (colMap.emailSubjCol) {
    worksheet.getColumn(colMap.emailSubjCol).hidden = false;
  }

  const defaultStatusVal = getDefaultStatus();

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
      if (legacyBackfillMap.has(cacheKey)) {
        key = legacyBackfillMap.get(cacheKey);
      } else {
        key = generateTrackingGuid();
        legacyBackfillMap.set(cacheKey, key);
      }
      row.getCell(colMap.trackingCol).value = key;
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
  renderList();
}

btnPrevSheet.addEventListener('click', () => {
  const idx = dateSheets.indexOf(sheetSelect.value);
  if (idx !== -1 && idx < dateSheets.length - 1) {
    sheetSelect.value = dateSheets[idx + 1];
    onActiveSheetChanged();
  }
});

btnNextSheet.addEventListener('click', () => {
  const idx = dateSheets.indexOf(sheetSelect.value);
  if (idx > 0) {
    sheetSelect.value = dateSheets[idx - 1];
    onActiveSheetChanged();
  }
});

sheetSelect.addEventListener('change', onActiveSheetChanged);

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
    const parts = [];
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

  listBody.innerHTML = '';
  const colSpan = isDeleteMode ? 12 : 11;
  if (sorted.length === 0) {
    listBody.innerHTML = `<tr><td colspan="${colSpan}" style="text-align:center; color:var(--text-muted); padding: 24px;">No matching requests found.</td></tr>`;
    const shadowRowEmpty = createShadowNewRowElement(colSpan, currentSheet);
    if (shadowRowEmpty) {
      listBody.appendChild(shadowRowEmpty);
    }
    syncSelectAllCheckboxes();
    updateStatusBarTelemetry();
    return;
  }

  sorted.forEach((item, idx) => {
    const tr = document.createElement('tr');
    const compId = getRowCompositeId(item.sheetName, item.rowNumber);
    const isCheckedForDelete = isDeleteMode && selectedForDelete.has(compId);
    if (isCheckedForDelete) {
      tr.classList.add('selected-for-delete');
    }

    const isComplete = isCompletedStatus(item.status);
    const noteText = item.notes || '—';
    const isLongNote = noteText.length > 45 || noteText.includes('\n');
    const shortGuid = item.key.length > 8 ? `${item.key.slice(0, 8)}...` : item.key;

    const deleteCheckboxCell = isDeleteMode
      ? `<td class="delete-cell">
           <input type="checkbox" class="chk-delete-row" data-comp="${escapeHtml(compId)}" ${isCheckedForDelete ? 'checked' : ''}>
         </td>`
      : '';

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
      ${deleteCheckboxCell}
      <td style="white-space:nowrap;">
        <button type="button" class="clickable-date-tab" data-date-tab="${escapeHtml(item.sheetName)}" title="Click to switch to daily tab ${escapeHtml(item.sheetName)}">
          📅 ${escapeHtml(item.sheetName)}
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
      <td style="font-family:monospace; font-size:12px; white-space:nowrap;">
        M: ${masterHtml}<br>B: ${bidderHtml}
      </td>
      <td>${actionHtml}</td>
      <td style="font-size:12px;">
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
        <button class="btn-edit" data-sheet="${escapeHtml(item.sheetName)}" data-row="${item.rowNumber}">Edit</button>
      </td>
    `;
    listBody.appendChild(tr);
  });

  const shadowRow = createShadowNewRowElement(colSpan, currentSheet);
  if (shadowRow) {
    listBody.appendChild(shadowRow);
  }

  syncSelectAllCheckboxes();
  updateStatusBarTelemetry();
}

// --- DELETE SELECTION MODE & CONFIRMATION ---
function setDeleteMode(enabled) {
  closeToolsMenu();
  isDeleteMode = Boolean(enabled);
  selectedForDelete.clear();
  btnDeleteMode.innerHTML = isDeleteMode
    ? '<span>✕ Exit Batch Delete Mode</span>'
    : '<span>☑️ Batch Delete Requests Mode</span>';
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
  updateStatusBarTelemetry();
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
    setDeleteMode(false);
    await loadWorkbookFromDisk(currentTab, {
      preserveStatus: saveResult === 'mirror-locked'
    });
    showToast(`Deleted ${deletedCount} request(s) from workbook.`, 'success');
  }
});

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

// Delegate clicks for Shadow New Request Row, Date Tab jump, Key Chip, Like-Item chips, Collapsible Notes, and Edit button
listBody.addEventListener('click', (e) => {
  const shadowRow = e.target.closest('.shadow-new-row');
  if (shadowRow) {
    openNewRequestModal();
    return;
  }

  const dateTabBtn = e.target.closest('.clickable-date-tab');
  if (dateTabBtn && dateTabBtn.dataset.dateTab) {
    const targetTab = dateTabBtn.dataset.dateTab;
    if (dateSheets.includes(targetTab)) {
      sheetSelect.value = targetTab;
      scopeSelect.value = 'current';
      activeKeyFilter = null;
      activeFieldFilter = null;
      if (searchInput.value) {
        searchInput.value = '';
      }
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

  const editBtn = e.target.closest('.btn-edit');
  if (editBtn) {
    openEditModal(editBtn.dataset.sheet, Number(editBtn.dataset.row));
  }
});

document.getElementById('btnClearKeyFilter').addEventListener('click', clearAllSortAndFilters);
if (btnClearSortFilter) {
  btnClearSortFilter.addEventListener('click', clearAllSortAndFilters);
}

scopeSelect.addEventListener('change', renderList);
searchInput.addEventListener('input', renderList);

// --- 6. EDIT & CREATE MODAL + INPUT HYGIENE + CROSS-TAB STATUS PROPAGATION ---
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

btnNewEntry.addEventListener('click', openNewRequestModal);

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

  await loadWorkbookFromDisk(currentSheetName);
  const worksheet = workbook.getWorksheet(currentSheetName);
  if (!worksheet) {
    return showModalAlert({
      title: 'Worksheet Not Found',
      message: `Target worksheet "${currentSheetName}" was not found in the workbook.`
    });
  }

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
    targetRowNum = editingTarget.rowNumber;
    trackingKey = editingTarget.key;
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
    editModal.classList.remove('open');
    await loadWorkbookFromDisk(currentSheetName, {
      preserveStatus: saveResult === 'mirror-locked'
    });
  }
});

function getNextEmptyRow(worksheet) {
  const colMap = getWorksheetColumnMap(worksheet);
  let lastRow = 1;
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
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
      .some(val => val !== '');
    if (hasData && rowNumber > lastRow) lastRow = rowNumber;
  });
  return lastRow + 1;
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
      if (!isCompletedStatus(row.status)) {
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
    const isLatestComplete = isCompletedStatus(latestRow.status);

    if (existingOnTargetByKey.has(normKey)) {
      const targetRow = existingOnTargetByKey.get(normKey);
      const isTargetComplete = isCompletedStatus(targetRow.status);
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
      <div class="preview-section-title" style="color:#fbbf24;">
        ⚠️ Open on Older Tab, but Marked "Closed / Complete" on Later Tab (${plan.supersededByComplete.length})
      </div>
      <div style="font-size:12px; color:var(--text-muted); margin-bottom:6px;">
        These requests had an open status on an earlier sheet, but were marked <strong>${escapeHtml(plan.supersededByComplete[0].completedRow.status)}</strong> on a newer sheet. Check any item below to reopen and carry it over into <code>${escapeHtml(targetVal)}</code>:
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
                ${escapeHtml(openRow.broker)} • Was <strong>${escapeHtml(openRow.status)}</strong> on ${escapeHtml(openRow.sheetName)}, marked <strong>${escapeHtml(completedRow.status)}</strong> on ${escapeHtml(completedRow.sheetName)}
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

  try {
    await loadWorkbookFromDisk(sheetSelect.value);

    const plan = computeRolloverAnalysis(targetTabName, includeFuture);
    if (!plan.valid) return;

    const itemsToCarry = [...plan.toRollOver];
    plan.supersededByComplete.forEach(({ openRow, completedRow }) => {
      if (forcedKeys.has(openRow.key.toLowerCase())) {
        itemsToCarry.push(openRow);
        const completedWs = workbook.getWorksheet(completedRow.sheetName);
        if (completedWs) {
          const compColMap = getWorksheetColumnMap(completedWs);
          const cr = completedWs.getRow(completedRow.rowNumber);
          cr.getCell(compColMap.statusCol).value = openRow.status;
          cr.commit();
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
        const rNum = writeRowIdx++;
        const r = targetSheet.getRow(rNum);
        writeItemToWorksheetRow(r, rNum, item, targetColMap);
        existingOnTargetByKey.set(normKey, item);
        addedCount++;
      }
    }

    for (const entry of plan.alreadyOnTarget) {
      if (entry.reopenedOnSource) {
        const excelRow = targetSheet.getRow(entry.targetRow.rowNumber);
        excelRow.getCell(targetColMap.statusCol).value = entry.latestRow.status;
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

function placeWorksheetFirstInWorkbook(wb, targetSheet) {
  try {
    let order = 2;
    wb.eachSheet((ws) => {
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

function placeWorksheetFirst(targetSheet) {
  if (workbook) placeWorksheetFirstInWorkbook(workbook, targetSheet);
}

function cloneTemplateStructure(templateSheet, newSheet) {
  if (Array.isArray(templateSheet.views) && templateSheet.views.length > 0) {
    newSheet.views = JSON.parse(JSON.stringify(templateSheet.views));
  }
  if (templateSheet.pageSetup) {
    newSheet.pageSetup = JSON.parse(JSON.stringify(templateSheet.pageSetup));
  }

  if (Array.isArray(templateSheet.columns)) {
    templateSheet.columns.forEach((col, idx) => {
      if (!col) return;
      const targetCol = newSheet.getColumn(idx + 1);
      targetCol.width = col.width || 15;
      targetCol.hidden = col.hidden || false;
      if (col.style) {
        targetCol.style = JSON.parse(JSON.stringify(col.style));
      }
    });
  }

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

  const colMap = getWorksheetColumnMap(newSheet);
  newSheet.getColumn(colMap.trackingCol).hidden = true;
  if (colMap.emailSubjCol) {
    newSheet.getColumn(colMap.emailSubjCol).hidden = false;
  }

  // If cloning from the full 74-row Tracker Template, preserve rows 2..74 cell styles and Column C XLOOKUP formulas
  if (colMap.isTemplateLayout) {
    const maxTemplateRow = Math.max(templateSheet.rowCount || 74, 74);
    for (let r = 2; r <= maxTemplateRow; r++) {
      const srcRow = templateSheet.getRow(r);
      const dstRow = newSheet.getRow(r);
      if (srcRow.height) dstRow.height = srcRow.height;
      srcRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (colNumber <= 12 && cell.style) {
          dstRow.getCell(colNumber).style = JSON.parse(JSON.stringify(cell.style));
        }
      });
      setBrokerCellWithFormula(dstRow, r, '', '', colMap);
      dstRow.commit();
    }
  }
}

function writeItemToWorksheetRow(row, rowNumber, item, colMap) {
  const masterUpper = item.master ? item.master.toUpperCase() : '';
  row.getCell(1).value = item.subject;
  row.getCell(2).value = item.ticket ? item.ticket.toUpperCase() : '';
  setBrokerCellWithFormula(row, rowNumber, item.broker, masterUpper, colMap);
  row.getCell(4).value = masterUpper;
  row.getCell(5).value = item.bidder ? item.bidder.toUpperCase() : '';
  row.getCell(6).value = item.action;
  row.getCell(7).value = item.yard ? item.yard.toUpperCase() : '';
  row.getCell(8).value = item.lots;
  row.getCell(colMap.notesCol).value = item.notes;
  row.getCell(colMap.statusCol).value = item.status;
  row.getCell(colMap.trackingCol).value = item.key;
  row.commit();
}

function prepareWorkbookFormulasBeforeSave(wb) {
  if (!wb) return;
  wb.eachSheet((ws) => {
    const rawName = ws.name.trim();
    const lower = rawName.toLowerCase();
    const isDailyOrTemplate =
      lower === 'template' ||
      isTemplatePlaceholderTab(rawName) ||
      DATE_TAB_REGEX.test(rawName);
    if (!isDailyOrTemplate) return;

    const colMap = getWorksheetColumnMap(ws);
    ws.getColumn(colMap.trackingCol).hidden = true;
    if (colMap.emailSubjCol) {
      ws.getColumn(colMap.emailSubjCol).hidden = false;
    }

    if (!colMap.isTemplateLayout) return;

    const maxR = Math.max(ws.rowCount || 74, 74);
    for (let r = 2; r <= maxR; r++) {
      const row = ws.getRow(r);
      const masterVal = cellText(row, 4).toUpperCase();
      const brokerVal = cellText(row, 3);
      setBrokerCellWithFormula(row, r, brokerVal, masterVal, colMap);
    }
  });
}

async function serializeWorkbookBuffer() {
  prepareWorkbookFormulasBeforeSave(workbook);
  const rawBuffer = await workbook.xlsx.writeBuffer();
  return await repairSavedXlsxArchive(rawBuffer, originalArchiveState);
}

// --- 8. DUAL-WRITE MIRRORING + AUTO-RETRY & MANUAL RETRY ---
function clearMirrorRetryState() {
  pendingMirrorBuffer = null;
  mirrorRetryCount = 0;
  mirrorCountdownSec = 0;
  if (mirrorCountdownInterval) {
    clearInterval(mirrorCountdownInterval);
    mirrorCountdownInterval = null;
  }
  if (btnRetryMirror) {
    btnRetryMirror.style.display = 'none';
  }
}

function scheduleMirrorAutoRetry() {
  if (mirrorCountdownInterval) {
    clearInterval(mirrorCountdownInterval);
  }
  btnRetryMirror.style.display = 'inline-block';

  if (mirrorRetryCount >= MIRROR_MAX_AUTO_RETRIES) {
    btnRetryMirror.textContent = '🔄 Retry Mirror Sync';
    setStatus('⚠️ Saved Working File (Mirror locked — click Retry when free)', 'warning');
    return;
  }

  mirrorCountdownSec = MIRROR_RETRY_INTERVAL_SEC;
  btnRetryMirror.textContent = `🔄 Retry Mirror (${mirrorCountdownSec}s)`;

  mirrorCountdownInterval = setInterval(async () => {
    mirrorCountdownSec--;
    if (mirrorCountdownSec > 0) {
      btnRetryMirror.textContent = `🔄 Retry Mirror (${mirrorCountdownSec}s)`;
    } else {
      clearInterval(mirrorCountdownInterval);
      mirrorCountdownInterval = null;
      mirrorRetryCount++;
      await retryMirrorSyncNow(true);
    }
  }, 1000);
}

async function attemptMirrorWrite(buffer, isRetry = false) {
  const targetMirrorDir = supervisorDirHandle || parentDirHandle;
  if (!targetMirrorDir || !workingFileHandle) return true;

  const mirrorFileName = workingFileHandle.name.replace(/\.xlsx$/i, '') + ' (VIEW ONLY).xlsx';
  try {
    const mirrorHandle = await targetMirrorDir.getFileHandle(mirrorFileName, { create: true });
    const mirrorWritable = await mirrorHandle.createWritable();
    await mirrorWritable.write(buffer);
    await mirrorWritable.close();

    const wasRetrying = Boolean(pendingMirrorBuffer) || isRetry;
    clearMirrorRetryState();

    const mirrorLocationNote = supervisorDirHandle ? ` (Mirrored to ${supervisorDirHandle.name})` : '';
    setStatus(`Saved & Mirrored: ${workingFileHandle.name}${mirrorLocationNote}`, 'ready');
    if (wasRetrying) {
      showToast(`✅ Supervisor mirror "${mirrorFileName}" synced!`, 'success');
    }
    return true;
  } catch (mirrorErr) {
    console.warn(`Supervisor mirror file "${mirrorFileName}" is locked or busy:`, mirrorErr);
    pendingMirrorBuffer = buffer;
    setStatus('⚠️ Saved Working File (Mirror in use by Supervisor)', 'warning');
    scheduleMirrorAutoRetry();
    if (!isRetry) {
      showToast('⚠️ Saved Working File (Mirror in use by Supervisor — auto-retrying every 15s)', 'warning');
    }
    return false;
  }
}

async function retryMirrorSyncNow(isAuto = false) {
  if (!workingFileHandle || !workbook) return;
  btnRetryMirror.textContent = '⏳ Retrying Mirror...';

  let bufferToWrite = pendingMirrorBuffer;
  if (!bufferToWrite) {
    try {
      bufferToWrite = await serializeWorkbookBuffer();
    } catch (e) {
      console.error('Could not serialize buffer for mirror retry:', e);
      return;
    }
  }

  const ok = await attemptMirrorWrite(bufferToWrite, true);
  if (!ok && !isAuto) {
    showToast('⚠️ Supervisor (VIEW ONLY) file is still locked in Excel. Will keep retrying.', 'warning');
  }
}

btnRetryMirror.addEventListener('click', () => {
  mirrorRetryCount = 0; // Reset auto-retry counter on manual click
  retryMirrorSyncNow(false);
});

async function saveWorkbookToDisk() {
  if (!workingFileHandle) return false;

  setStatus('Saving to Excel...', '');
  let buffer;
  try {
    buffer = await serializeWorkbookBuffer();
  } catch (err) {
    console.error('Failed to serialize workbook buffer:', err);
    setStatus('Serialization Error', 'locked');
    await showModalAlert({
      title: 'Excel Serialization Error',
      message: 'Failed to generate Excel buffer: ' + err.message,
      variant: 'danger'
    });
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
    await showModalAlert({
      title: '🔒 Save Blocked by Windows (File Locked)',
      badge: workingFileHandle.name,
      message:
        `The working workbook ("${workingFileHandle.name}") is currently open and locked in Desktop Excel.\n\n` +
        'Please close the file in Desktop Excel and click Save again.',
      variant: 'danger'
    });
    return false;
  }

  // 2. Dual-Write Mirroring to (VIEW ONLY).xlsx with Auto-Retry & Manual Retry fallback
  const mirrorOk = await attemptMirrorWrite(buffer, false);
  return mirrorOk ? true : 'mirror-locked';
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
  if (statusTimestamp && (stateClass === 'ready' || stateClass === 'warning')) {
    const nowTime = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    statusTimestamp.textContent = `Last updated: ${nowTime}`;
  }
  updateStatusBarTelemetry();
}

function updateStatusBarTelemetry() {
  if (statusFolderInfo) {
    if (!parentDirHandle) {
      statusFolderInfo.textContent = '📂 Folder: Not connected';
    } else {
      const supName = supervisorDirHandle ? supervisorDirHandle.name : 'Same as Agent';
      statusFolderInfo.textContent = `📂 Agent: ${parentDirHandle.name} • 👁️ Mirror: ${supName}`;
    }
  }

  if (statusRowStats) {
    if (!workbook) {
      statusRowStats.textContent = '📊 0 requests';
    } else {
      const openTotal = allRows.filter(r => !isCompletedStatus(r.status)).length;
      statusRowStats.textContent = `📊 Showing ${currentlyVisibleRows.length} of ${allRows.length} requests (${openTotal} open across ${dateSheets.length} tab${dateSheets.length === 1 ? '' : 's'})`;
    }
  }
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

// --- 9. PWA SERVICE WORKER REGISTRATION ---
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js')
      .then((reg) => console.log('Service Worker registered with scope:', reg.scope))
      .catch((err) => console.warn('Service Worker registration failed:', err));
  });
}
