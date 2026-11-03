// --- STATE ---
let parentDirHandle = null;
let cachedDirHandle = null;
let supervisorDirHandle = null;
let cachedSupervisorDirHandle = null;
let templatesDirHandle = null;
let cachedTemplatesDirHandle = null;
let availableTemplates = []; // [{ name, handle, lastModified }]
let lastSavedWorkbookName = null;
let autoReconnectEnabled = localStorage.getItem('trackerui_auto_reconnect') !== 'false';
let autoReconnectGesturePending = false;

let availableWorkbooks = []; // [{ name, handle, lastModified, sortWeight }]
let workingFileHandle = null;
let workbook = null;
let allRows = []; // Flattened index across all MM-DD-YYYY tabs
let dateSheets = []; // Ordered list of MM-DD-YYYY sheet names (newest first)
let activeKeyFilter = null;
let activeFieldFilter = null; // { field, label, value } for Like-Item filtering across tabs
let sortState = { column: null, direction: 'asc' }; // Column sort state
let editingTarget = null; // { sheetName, rowNumber, key } or null for new (Modal)
let dockEditingTarget = null; // { sheetName, rowNumber, key } or null for new (Bottom Dock)
const DOCK_COLLAPSED_STORAGE_KEY = 'trackerui_dock_collapsed';
let isBottomDockCollapsed = localStorage.getItem(DOCK_COLLAPSED_STORAGE_KEY) === 'true';
let activeTicketDetailsTarget = null; // { key, ticket, sheetName, rowNumber }
let viewOnlyState = { fileName: '', fullPath: '', lastModified: 0, buffer: null, sheets: [], rows: [] };

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
const IDB_TEMPLATES_KEY = 'templatesFolderHandle';
const IDB_LAST_WORKBOOK_KEY = 'lastActiveWorkbookName';
const AUTO_RECONNECT_STORAGE_KEY = 'trackerui_auto_reconnect';
const THEME_STORAGE_KEY = 'trackerui_theme_mode'; // 'auto' | 'light' | 'dark'
const FONT_SCALE_STORAGE_KEY = 'trackerui_font_scale'; // '85' .. '125' (percent)
const FONT_SCALE_MIN = 85;
const FONT_SCALE_MAX = 125;
const FONT_SCALE_STEP = 5;

let themePreference = localStorage.getItem(THEME_STORAGE_KEY) || 'auto';
if (!['auto', 'light', 'dark'].includes(themePreference)) {
  themePreference = 'auto';
}
let fontScalePercent = parseInt(localStorage.getItem(FONT_SCALE_STORAGE_KEY) || '100', 10);
if (isNaN(fontScalePercent) || fontScalePercent < FONT_SCALE_MIN || fontScalePercent > FONT_SCALE_MAX) {
  fontScalePercent = 100;
}

// --- DOM ELEMENTS ---
const metaThemeColor = document.getElementById('metaThemeColor');
const statusPill = document.getElementById('statusPill');
const btnRetryMirror = document.getElementById('btnRetryMirror');
const workbookPickerWrap = document.getElementById('workbookPickerWrap');
const workbookSelect = document.getElementById('workbookSelect');
const btnReconnect = document.getElementById('btnReconnect');
const reconnectBtnLabel = document.getElementById('reconnectBtnLabel');
const btnLinkFile = document.getElementById('btnLinkFile');
const btnOpenViewOnly = document.getElementById('btnOpenViewOnly');
const btnRollover = document.getElementById('btnRollover');
const btnNewEntry = document.getElementById('btnNewEntry');

// Appearance & Popover Tools Menu Elements
const headerThemeGroup = document.getElementById('headerThemeGroup');
const menuThemeGroup = document.getElementById('menuThemeGroup');
const menuThemeHint = document.getElementById('menuThemeHint');
const fontSizeSlider = document.getElementById('fontSizeSlider');
const btnFontDecrease = document.getElementById('btnFontDecrease');
const btnFontIncrease = document.getElementById('btnFontIncrease');
const fontSizeValue = document.getElementById('fontSizeValue');
const btnResetFontSize = document.getElementById('btnResetFontSize');

const toolsMenuWrap = document.getElementById('toolsMenuWrap');
const btnToolsMenu = document.getElementById('btnToolsMenu');
const toolsMenuDropdown = document.getElementById('toolsMenuDropdown');
const menuSelectAgentFolder = document.getElementById('menuSelectAgentFolder');
const agentDirLabel = document.getElementById('agentDirLabel');
const btnResetAgentDir = document.getElementById('btnResetAgentDir');
const btnSupervisorDir = document.getElementById('btnSupervisorDir');
const supervisorDirLabel = document.getElementById('supervisorDirLabel');
const btnResetSupervisorDir = document.getElementById('btnResetSupervisorDir');
const btnTemplatesDir = document.getElementById('btnTemplatesDir');
const templatesDirLabel = document.getElementById('templatesDirLabel');
const btnResetTemplatesDir = document.getElementById('btnResetTemplatesDir');
const menuOpenViewOnly = document.getElementById('menuOpenViewOnly');
const chkAutoReconnect = document.getElementById('chkAutoReconnect');
const chkAutoReconnectBar = document.getElementById('chkAutoReconnectBar');
const btnNewWorkbook = document.getElementById('btnNewWorkbook');
const btnRepairSchema = document.getElementById('btnRepairSchema');
const btnAddBlankTab = document.getElementById('btnAddBlankTab');
const btnRenameTab = document.getElementById('btnRenameTab');
const btnDeleteTab = document.getElementById('btnDeleteTab');
const btnDeleteMode = document.getElementById('btnDeleteMode');
const deleteModeLabel = document.getElementById('deleteModeLabel');

// --- APPEARANCE ENGINE: THEME (LIGHT / DARK / AUTO) & APP-WIDE FONT SCALING ---
const systemLightMediaQuery = typeof window.matchMedia === 'function'
  ? window.matchMedia('(prefers-color-scheme: light)')
  : null;

function resolveEffectiveTheme(pref = themePreference) {
  if (pref === 'light') return 'light';
  if (pref === 'dark') return 'dark';
  return (systemLightMediaQuery && systemLightMediaQuery.matches) ? 'light' : 'dark';
}

function applyThemeToDOM() {
  const effective = resolveEffectiveTheme(themePreference);
  document.documentElement.setAttribute('data-theme', effective);
  document.documentElement.style.colorScheme = effective;

  if (metaThemeColor) {
    metaThemeColor.setAttribute('content', effective === 'light' ? '#ffffff' : '#0f172a');
  }

  const updateSegButtons = (container) => {
    if (!container) return;
    container.querySelectorAll('.theme-seg-btn').forEach(btn => {
      const active = btn.dataset.themeChoice === themePreference;
      btn.classList.toggle('active', active);
      btn.setAttribute('aria-pressed', String(active));
    });
  };
  updateSegButtons(headerThemeGroup);
  updateSegButtons(menuThemeGroup);

  if (menuThemeHint) {
    if (themePreference === 'auto') {
      const sysLabel = effective === 'light' ? 'Light' : 'Dark';
      menuThemeHint.textContent = `Auto (${sysLabel})`;
    } else {
      menuThemeHint.textContent = themePreference === 'light' ? 'Light Mode' : 'Dark Mode';
    }
  }
}

function setThemePreference(nextPref, { persist = true, announce = false } = {}) {
  if (!['auto', 'light', 'dark'].includes(nextPref)) return;
  themePreference = nextPref;
  if (persist) {
    localStorage.setItem(THEME_STORAGE_KEY, themePreference);
  }
  applyThemeToDOM();
  if (announce) {
    const effective = resolveEffectiveTheme(themePreference);
    const label = themePreference === 'auto'
      ? `Automatic (${effective === 'light' ? 'System Light' : 'System Dark'})`
      : (themePreference === 'light' ? 'Light Mode' : 'Dark Mode');
    showToast(`Theme set to ${label}.`, 'info');
  }
}

function applyFontScaleToDOM(percent, { persist = true } = {}) {
  const clamped = Math.max(FONT_SCALE_MIN, Math.min(FONT_SCALE_MAX, Math.round(percent / FONT_SCALE_STEP) * FONT_SCALE_STEP));
  fontScalePercent = clamped;
  const scaleRatio = +(clamped / 100).toFixed(2);
  document.documentElement.style.setProperty('--font-scale', String(scaleRatio));

  if (persist) {
    localStorage.setItem(FONT_SCALE_STORAGE_KEY, String(clamped));
  }
  if (fontSizeSlider) {
    fontSizeSlider.value = String(clamped);
  }
  if (fontSizeValue) {
    fontSizeValue.textContent = `${clamped}%`;
  }
  if (btnResetFontSize) {
    btnResetFontSize.style.display = clamped !== 100 ? 'inline-block' : 'none';
  }
  if (btnFontDecrease) {
    btnFontDecrease.disabled = clamped <= FONT_SCALE_MIN;
  }
  if (btnFontIncrease) {
    btnFontIncrease.disabled = clamped >= FONT_SCALE_MAX;
  }
}

// Apply saved/system appearance immediately before initial render
applyThemeToDOM();
applyFontScaleToDOM(fontScalePercent, { persist: false });

if (systemLightMediaQuery) {
  const handleSystemThemeChange = () => {
    if (themePreference === 'auto') {
      applyThemeToDOM();
    }
  };
  if (typeof systemLightMediaQuery.addEventListener === 'function') {
    systemLightMediaQuery.addEventListener('change', handleSystemThemeChange);
  } else if (typeof systemLightMediaQuery.addListener === 'function') {
    systemLightMediaQuery.addListener(handleSystemThemeChange);
  }
}

[headerThemeGroup, menuThemeGroup].forEach(groupEl => {
  if (!groupEl) return;
  groupEl.addEventListener('click', (e) => {
    const btn = e.target.closest('.theme-seg-btn');
    if (!btn || !btn.dataset.themeChoice) return;
    setThemePreference(btn.dataset.themeChoice, { persist: true, announce: false });
  });
});

if (fontSizeSlider) {
  fontSizeSlider.addEventListener('input', (e) => {
    applyFontScaleToDOM(Number(e.target.value), { persist: true });
  });
}
if (btnFontDecrease) {
  btnFontDecrease.addEventListener('click', () => {
    applyFontScaleToDOM(fontScalePercent - FONT_SCALE_STEP, { persist: true });
  });
}
if (btnFontIncrease) {
  btnFontIncrease.addEventListener('click', () => {
    applyFontScaleToDOM(fontScalePercent + FONT_SCALE_STEP, { persist: true });
  });
}
if (btnResetFontSize) {
  btnResetFontSize.addEventListener('click', () => {
    applyFontScaleToDOM(100, { persist: true });
  });
}

// File Paths Bar Elements
const pathTrackerText = document.getElementById('pathTrackerText');
const pathViewOnlyText = document.getElementById('pathViewOnlyText');
const pathTemplatesSep = document.getElementById('pathTemplatesSep');
const pathTemplatesItem = document.getElementById('pathTemplatesItem');
const pathTemplatesText = document.getElementById('pathTemplatesText');

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
const ticketDetailsModal = document.getElementById('ticketDetailsModal');
const viewOnlyModal = document.getElementById('viewOnlyModal');
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

// Modal Form Fields
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
const btnModalInsertTimestamp = document.getElementById('btnModalInsertTimestamp');
const brokerListEl = document.getElementById('brokerList');
const actionListEl = document.getElementById('actionList');

// Bottom Dock Editor DOM Elements
const bottomEditorDock = document.getElementById('bottomEditorDock');
const dockHeaderToggleZone = document.getElementById('dockHeaderToggleZone');
const dockModeBadge = document.getElementById('dockModeBadge');
const dockTargetLabel = document.getElementById('dockTargetLabel');
const dockKeyBadge = document.getElementById('dockKeyBadge');
const btnDockTicketDetails = document.getElementById('btnDockTicketDetails');
const btnDockOpenModal = document.getElementById('btnDockOpenModal');
const btnDockNewClear = document.getElementById('btnDockNewClear');
const btnDockToggleCollapse = document.getElementById('btnDockToggleCollapse');
const dockEntryForm = document.getElementById('dockEntryForm');
const dockSubject = document.getElementById('dockSubject');
const dockTicket = document.getElementById('dockTicket');
const dockBroker = document.getElementById('dockBroker');
const dockMaster = document.getElementById('dockMaster');
const dockBidder = document.getElementById('dockBidder');
const dockAction = document.getElementById('dockAction');
const dockYard = document.getElementById('dockYard');
const dockLots = document.getElementById('dockLots');
const dockStatus = document.getElementById('dockStatus');
const dockNotes = document.getElementById('dockNotes');
const btnDockInsertTimestamp = document.getElementById('btnDockInsertTimestamp');
const dockHintText = document.getElementById('dockHintText');
const btnDockCancelEdit = document.getElementById('btnDockCancelEdit');
const btnDockSave = document.getElementById('btnDockSave');

// Ticket Details Modal DOM Elements
const ticketDetailsTitle = document.getElementById('ticketDetailsTitle');
const ticketDetailsSubtitle = document.getElementById('ticketDetailsSubtitle');
const ticketDetailsKeyBadge = document.getElementById('ticketDetailsKeyBadge');
const ticketDetailsMetaGrid = document.getElementById('ticketDetailsMetaGrid');
const ticketTimelineHeader = document.getElementById('ticketTimelineHeader');
const ticketTimelineList = document.getElementById('ticketTimelineList');
const ticketQuickNoteForm = document.getElementById('ticketQuickNoteForm');
const ticketQuickNoteTargetBadge = document.getElementById('ticketQuickNoteTargetBadge');
const ticketQuickNoteInput = document.getElementById('ticketQuickNoteInput');
const btnTicketEditInDock = document.getElementById('btnTicketEditInDock');
const btnCloseTicketDetails = document.getElementById('btnCloseTicketDetails');

// View-Only Inspector Modal DOM Elements
const viewOnlyTitle = document.getElementById('viewOnlyTitle');
const viewOnlyPathSubtitle = document.getElementById('viewOnlyPathSubtitle');
const viewOnlySyncBadge = document.getElementById('viewOnlySyncBadge');
const btnDownloadViewOnly = document.getElementById('btnDownloadViewOnly');
const viewOnlySheetSelect = document.getElementById('viewOnlySheetSelect');
const viewOnlySearchInput = document.getElementById('viewOnlySearchInput');
const viewOnlyStatsText = document.getElementById('viewOnlyStatsText');
const viewOnlyBody = document.getElementById('viewOnlyBody');
const btnRefreshViewOnly = document.getElementById('btnRefreshViewOnly');
const btnCloseViewOnly = document.getElementById('btnCloseViewOnly');

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
const dialogSelectWrap = document.getElementById('dialogSelectWrap');
const dialogSelectLabel = document.getElementById('dialogSelectLabel');
const dialogSelect = document.getElementById('dialogSelect');
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
let activeDialogHasSelect = false;

function closeDialogModal(resultValue) {
  if (!dialogModal) return;
  dialogModal.classList.remove('open');
  dialogInputError.style.display = 'none';
  dialogInputError.textContent = '';
  if (dialogSelectWrap) dialogSelectWrap.style.display = 'none';
  const resolver = activeDialogResolver;
  activeDialogResolver = null;
  activeDialogValidator = null;
  activeDialogHasSelect = false;
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
  selectLabel = '',
  selectOptions = null, // [{ value, label }]
  defaultSelectValue = '',
  validate = null
}) {
  if (activeDialogResolver) {
    closeDialogModal(mode === 'prompt' ? null : false);
  }

  return new Promise((resolve) => {
    activeDialogResolver = resolve;
    activeDialogMode = mode;
    activeDialogValidator = typeof validate === 'function' ? validate : null;
    activeDialogHasSelect = Boolean(Array.isArray(selectOptions) && selectOptions.length > 0 && dialogSelect);

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

    if (activeDialogHasSelect) {
      dialogSelectWrap.style.display = 'flex';
      dialogSelectLabel.textContent = selectLabel || 'Select Option';
      dialogSelect.innerHTML = '';
      selectOptions.forEach(optObj => {
        const opt = document.createElement('option');
        opt.value = optObj.value;
        opt.textContent = optObj.label;
        dialogSelect.appendChild(opt);
      });
      if (defaultSelectValue) {
        dialogSelect.value = defaultSelectValue;
      }
    } else if (dialogSelectWrap) {
      dialogSelectWrap.style.display = 'none';
    }

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
    if (activeDialogHasSelect) {
      closeDialogModal({ value: val, selectValue: dialogSelect.value });
    } else {
      closeDialogModal(val);
    }
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
    if (ticketDetailsModal && ticketDetailsModal.classList.contains('open')) {
      ticketDetailsModal.classList.remove('open');
      return;
    }
    if (viewOnlyModal && viewOnlyModal.classList.contains('open')) {
      viewOnlyModal.classList.remove('open');
      return;
    }
    closeToolsMenu();
  }
});

// --- INDEXEDDB DIRECTORY HANDLE & SETTINGS PERSISTENCE ---
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

async function saveHandleToIDB(key, value) {
  try {
    const db = await openHandleDB();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readwrite');
      tx.objectStore(IDB_STORE).put(value, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn(`Could not persist (${key}) in IndexedDB:`, err);
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
    console.warn(`Could not read (${key}) from IndexedDB:`, err);
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
    console.warn(`Could not delete (${key}) from IndexedDB:`, err);
  }
}

function setAutoReconnectPreference(enabled) {
  autoReconnectEnabled = Boolean(enabled);
  localStorage.setItem(AUTO_RECONNECT_STORAGE_KEY, autoReconnectEnabled ? 'true' : 'false');
  if (chkAutoReconnect) chkAutoReconnect.checked = autoReconnectEnabled;
  if (chkAutoReconnectBar) chkAutoReconnectBar.checked = autoReconnectEnabled;
}

if (chkAutoReconnect) {
  chkAutoReconnect.checked = autoReconnectEnabled;
  chkAutoReconnect.addEventListener('change', (e) => {
    setAutoReconnectPreference(e.target.checked);
    showToast(
      autoReconnectEnabled
        ? '⚡ Auto-Reconnect on Startup enabled.'
        : 'Auto-Reconnect on Startup disabled.',
      'info'
    );
  });
}

if (chkAutoReconnectBar) {
  chkAutoReconnectBar.checked = autoReconnectEnabled;
  chkAutoReconnectBar.addEventListener('change', (e) => {
    setAutoReconnectPreference(e.target.checked);
    showToast(
      autoReconnectEnabled
        ? '⚡ Auto-Reconnect on Startup enabled.'
        : 'Auto-Reconnect on Startup disabled.',
      'info'
    );
  });
}

function setReconnectButtonText(text) {
  if (reconnectBtnLabel) {
    reconnectBtnLabel.textContent = text;
  } else if (btnReconnect) {
    btnReconnect.textContent = text;
  }
}

function updateAgentDirUI() {
  const setLabel = (txt) => {
    if (agentDirLabel) agentDirLabel.textContent = txt;
    else menuSelectAgentFolder.textContent = txt;
  };
  if (parentDirHandle) {
    setLabel(`Tracker Folder: ${parentDirHandle.name}`);
    menuSelectAgentFolder.classList.add('custom-active');
    if (btnResetAgentDir) btnResetAgentDir.style.display = 'inline-block';
  } else if (cachedDirHandle) {
    setLabel(`Reconnect Tracker: ${cachedDirHandle.name}`);
    menuSelectAgentFolder.classList.add('custom-active');
    if (btnResetAgentDir) btnResetAgentDir.style.display = 'inline-block';
  } else {
    setLabel('Tracker Workbooks Folder...');
    menuSelectAgentFolder.classList.remove('custom-active');
    if (btnResetAgentDir) btnResetAgentDir.style.display = 'none';
  }
  updateFilePathsBarUI();
  updateStatusBarTelemetry();
}

function updateSupervisorDirUI() {
  const setLabel = (txt) => {
    if (supervisorDirLabel) supervisorDirLabel.textContent = txt;
    else btnSupervisorDir.textContent = txt;
  };
  if (supervisorDirHandle) {
    setLabel(`Supervisor Folder: ${supervisorDirHandle.name}`);
    btnSupervisorDir.title = `Saving (VIEW ONLY).xlsx copies to "${supervisorDirHandle.name}". Click to change folder.`;
    btnSupervisorDir.classList.add('custom-active');
    btnResetSupervisorDir.style.display = 'inline-block';
  } else if (cachedSupervisorDirHandle) {
    setLabel(`Reconnect Supervisor: ${cachedSupervisorDirHandle.name}`);
    btnSupervisorDir.title = `Click to reconnect saved Supervisor folder "${cachedSupervisorDirHandle.name}"`;
    btnSupervisorDir.classList.add('custom-active');
    btnResetSupervisorDir.style.display = 'inline-block';
  } else {
    setLabel('Supervisor Folder: Same as Agent');
    btnSupervisorDir.title = 'Select a custom folder where (VIEW ONLY).xlsx copies are saved';
    btnSupervisorDir.classList.remove('custom-active');
    btnResetSupervisorDir.style.display = 'none';
  }
  updateFilePathsBarUI();
  updateStatusBarTelemetry();
}

function updateTemplatesDirUI() {
  const setLabel = (txt) => {
    if (templatesDirLabel) templatesDirLabel.textContent = txt;
    else btnTemplatesDir.textContent = txt;
  };
  if (templatesDirHandle) {
    const countSuffix = availableTemplates.length > 0 ? ` (${availableTemplates.length})` : '';
    setLabel(`Templates Folder: ${templatesDirHandle.name}${countSuffix}`);
    btnTemplatesDir.title = `Using templates from "${templatesDirHandle.name}". Click to change folder.`;
    btnTemplatesDir.classList.add('custom-active');
    btnResetTemplatesDir.style.display = 'inline-block';
  } else if (cachedTemplatesDirHandle) {
    setLabel(`Reconnect Templates: ${cachedTemplatesDirHandle.name}`);
    btnTemplatesDir.title = `Click to reconnect saved Templates folder "${cachedTemplatesDirHandle.name}"`;
    btnTemplatesDir.classList.add('custom-active');
    btnResetTemplatesDir.style.display = 'inline-block';
  } else {
    setLabel('Templates Folder: Built-in Default');
    btnTemplatesDir.title = 'Select folder containing master .xlsx templates for new monthly workbooks';
    btnTemplatesDir.classList.remove('custom-active');
    btnResetTemplatesDir.style.display = 'none';
  }
  updateFilePathsBarUI();
  updateStatusBarTelemetry();
}

function updateFilePathsBarUI() {
  const agentFolderName = parentDirHandle
    ? parentDirHandle.name
    : (cachedDirHandle ? `${cachedDirHandle.name} (cached)` : null);

  const supFolderName = supervisorDirHandle
    ? supervisorDirHandle.name
    : (cachedSupervisorDirHandle
        ? `${cachedSupervisorDirHandle.name} (cached)`
        : (parentDirHandle ? parentDirHandle.name : (cachedDirHandle ? `${cachedDirHandle.name} (cached)` : null)));

  const wbName = workingFileHandle
    ? workingFileHandle.name
    : (lastSavedWorkbookName || null);

  if (pathTrackerText) {
    if (parentDirHandle && workingFileHandle) {
      pathTrackerText.textContent = `${parentDirHandle.name}\\${workingFileHandle.name}`;
    } else if (parentDirHandle) {
      pathTrackerText.textContent = `${parentDirHandle.name}\\(No workbook selected)`;
    } else if (agentFolderName && wbName) {
      pathTrackerText.textContent = `${agentFolderName}\\${wbName}`;
    } else if (agentFolderName) {
      pathTrackerText.textContent = `${agentFolderName}\\`;
    } else {
      pathTrackerText.textContent = 'Not connected';
    }
  }

  if (pathViewOnlyText) {
    if ((supervisorDirHandle || parentDirHandle) && workingFileHandle) {
      const mirrorDir = supervisorDirHandle ? supervisorDirHandle.name : parentDirHandle.name;
      const mirrorFile = workingFileHandle.name.replace(/\.xlsx$/i, '') + ' (VIEW ONLY).xlsx';
      pathViewOnlyText.textContent = `${mirrorDir}\\${mirrorFile}`;
    } else if (supFolderName && wbName) {
      const mirrorFile = wbName.replace(/\.xlsx$/i, '') + ' (VIEW ONLY).xlsx';
      pathViewOnlyText.textContent = `${supFolderName}\\${mirrorFile}`;
    } else if (supFolderName) {
      pathViewOnlyText.textContent = `${supFolderName}\\*(VIEW ONLY).xlsx`;
    } else {
      pathViewOnlyText.textContent = 'Not connected';
    }
  }

  const tplName = templatesDirHandle
    ? `${templatesDirHandle.name}\\ (${availableTemplates.length} template${availableTemplates.length === 1 ? '' : 's'})`
    : (cachedTemplatesDirHandle ? `${cachedTemplatesDirHandle.name}\\ (cached)` : null);

  if (pathTemplatesItem && pathTemplatesSep && pathTemplatesText) {
    if (tplName) {
      pathTemplatesSep.style.display = 'inline';
      pathTemplatesItem.style.display = 'inline-flex';
      pathTemplatesText.textContent = tplName;
    } else {
      pathTemplatesSep.style.display = 'none';
      pathTemplatesItem.style.display = 'none';
    }
  }

  const canOpenViewOnly = Boolean((parentDirHandle || supervisorDirHandle) && workingFileHandle);
  if (btnOpenViewOnly) {
    btnOpenViewOnly.style.display = canOpenViewOnly ? 'inline-block' : 'none';
  }
  if (menuOpenViewOnly) {
    menuOpenViewOnly.disabled = !canOpenViewOnly;
  }
}

async function scanTemplatesDirectory() {
  availableTemplates = [];
  if (!templatesDirHandle) {
    updateTemplatesDirUI();
    return;
  }
  try {
    for await (const entry of templatesDirHandle.values()) {
      if (entry.kind !== 'file') continue;
      const name = entry.name;
      const lower = name.toLowerCase();
      if (lower.endsWith('.xlsx') && !name.startsWith('~$') && !lower.endsWith('(view only).xlsx')) {
        const f = await entry.getFile();
        availableTemplates.push({
          name,
          handle: entry,
          lastModified: f.lastModified
        });
      }
    }
    availableTemplates.sort((a, b) => a.name.localeCompare(b.name));
  } catch (err) {
    console.warn('Could not scan templates directory:', err);
  }
  updateTemplatesDirUI();
}

async function attemptReconnectFolders({ requestPermissionIfNeeded = true, silent = false } = {}) {
  if (!cachedDirHandle) return false;
  try {
    let perm = await cachedDirHandle.queryPermission({ mode: 'readwrite' });
    if (perm !== 'granted' && requestPermissionIfNeeded) {
      perm = await cachedDirHandle.requestPermission({ mode: 'readwrite' });
    }
    if (perm !== 'granted') {
      return false;
    }

    parentDirHandle = cachedDirHandle;
    autoReconnectGesturePending = false;
    btnReconnect.style.display = 'none';
    btnLinkFile.style.display = 'none';
    btnNewWorkbook.disabled = false;

    if (cachedSupervisorDirHandle && !supervisorDirHandle) {
      try {
        let supPerm = await cachedSupervisorDirHandle.queryPermission({ mode: 'readwrite' });
        if (supPerm !== 'granted' && requestPermissionIfNeeded) {
          supPerm = await cachedSupervisorDirHandle.requestPermission({ mode: 'readwrite' });
        }
        if (supPerm === 'granted') {
          supervisorDirHandle = cachedSupervisorDirHandle;
        }
      } catch (supErr) {
        console.warn('Could not reconnect supervisor folder:', supErr);
      }
    }

    if (cachedTemplatesDirHandle && !templatesDirHandle) {
      try {
        let tplPerm = await cachedTemplatesDirHandle.queryPermission({ mode: 'read' });
        if (tplPerm !== 'granted' && requestPermissionIfNeeded) {
          tplPerm = await cachedTemplatesDirHandle.requestPermission({ mode: 'read' });
        }
        if (tplPerm === 'granted') {
          templatesDirHandle = cachedTemplatesDirHandle;
          await scanTemplatesDirectory();
        }
      } catch (tplErr) {
        console.warn('Could not reconnect templates folder:', tplErr);
      }
    }

    updateAgentDirUI();
    updateSupervisorDirUI();
    updateTemplatesDirUI();

    await scanDirectoryAndLoadWorkbooks(lastSavedWorkbookName);
    return true;
  } catch (err) {
    if (!silent) {
      console.error('Reconnect folder error:', err);
      await showModalAlert({
        title: 'Could Not Reconnect Folder',
        message: err.message
      });
    }
    return false;
  }
}

async function initCachedDirectoryHandles() {
  if (!('indexedDB' in window)) return;
  const [savedAgentHandle, savedSupHandle, savedTplHandle, savedWbName] = await Promise.all([
    getHandleFromIDB(IDB_AGENT_KEY),
    getHandleFromIDB(IDB_SUPERVISOR_KEY),
    getHandleFromIDB(IDB_TEMPLATES_KEY),
    getHandleFromIDB(IDB_LAST_WORKBOOK_KEY)
  ]);

  if (savedWbName && typeof savedWbName === 'string') {
    lastSavedWorkbookName = savedWbName;
  }
  if (savedSupHandle) {
    cachedSupervisorDirHandle = savedSupHandle;
  }
  if (savedTplHandle) {
    cachedTemplatesDirHandle = savedTplHandle;
  }
  if (savedAgentHandle) {
    cachedDirHandle = savedAgentHandle;
    setReconnectButtonText(`Reconnect (${savedAgentHandle.name})`);
    btnReconnect.style.display = 'inline-flex';
  }

  updateAgentDirUI();
  updateSupervisorDirUI();
  updateTemplatesDirUI();

  if (savedAgentHandle && autoReconnectEnabled) {
    const connectedImmediately = await attemptReconnectFolders({
      requestPermissionIfNeeded: false,
      silent: true
    });
    if (!connectedImmediately) {
      // Browser requires a user gesture after full restart before requestPermission() can fire.
      autoReconnectGesturePending = true;
      setReconnectButtonText(`Auto-Reconnect (${savedAgentHandle.name})`);
      setStatus('Click anywhere to auto-reconnect folder', 'warning');

      const onFirstUserGesture = async (evt) => {
        if (!autoReconnectGesturePending || !autoReconnectEnabled || parentDirHandle) {
          document.removeEventListener('click', onFirstUserGesture);
          return;
        }
        // Ignore if user is explicitly unchecking the auto-reconnect toggle, changing theme/font, or picking a new folder
        if (
          evt.target.closest('#chkAutoReconnect') ||
          evt.target.closest('#chkAutoReconnectBar') ||
          evt.target.closest('#btnLinkFile') ||
          evt.target.closest('#menuSelectAgentFolder') ||
          evt.target.closest('#btnResetAgentDir') ||
          evt.target.closest('#headerThemeGroup') ||
          evt.target.closest('#menuThemeGroup') ||
          evt.target.closest('.menu-fontsize-control')
        ) {
          return;
        }
        document.removeEventListener('click', onFirstUserGesture);
        autoReconnectGesturePending = false;
        await attemptReconnectFolders({ requestPermissionIfNeeded: true, silent: true });
      };
      document.addEventListener('click', onFirstUserGesture);
    }
  }
}

initCachedDirectoryHandles();

// --- 1. DIRECTORY SELECTION, SUPERVISOR EXPORT FOLDER, TEMPLATES FOLDER & WORKBOOK DISCOVERY ---
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
    autoReconnectGesturePending = false;
    await saveHandleToIDB(IDB_AGENT_KEY, dirHandle);
    btnReconnect.style.display = 'none';
    btnLinkFile.style.display = 'none';
    btnNewWorkbook.disabled = false;
    updateAgentDirUI();
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

if (btnResetAgentDir) {
  btnResetAgentDir.addEventListener('click', async () => {
    closeToolsMenu();
    parentDirHandle = null;
    cachedDirHandle = null;
    workingFileHandle = null;
    workbook = null;
    availableWorkbooks = [];
    lastSavedWorkbookName = null;
    autoReconnectGesturePending = false;
    await deleteHandleFromIDB(IDB_AGENT_KEY);
    await deleteHandleFromIDB(IDB_LAST_WORKBOOK_KEY);
    workbookPickerWrap.style.display = 'none';
    btnReconnect.style.display = 'none';
    btnLinkFile.style.display = 'inline-flex';
    btnNewEntry.style.display = 'none';
    btnRollover.style.display = 'none';
    if (bottomEditorDock) bottomEditorDock.style.display = 'none';
    sheetTabsBar.style.display = 'none';
    schemaBanner.style.display = 'none';
    setStatus('No Folder Linked', '');
    updateAgentDirUI();
    listBody.innerHTML = '<tr><td colspan="12" class="empty-table-cell">Click "Select Agent Folder" to connect your tracker directory.</td></tr>';
    showToast('Cleared saved Tracker Workbooks folder.', 'info');
  });
}

btnReconnect.addEventListener('click', async () => {
  if (!cachedDirHandle) return;
  const ok = await attemptReconnectFolders({ requestPermissionIfNeeded: true, silent: false });
  if (!ok) {
    await showModalAlert({
      title: 'Permission Required',
      message: 'Read/Write permission was not granted for the folder.'
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

if (btnTemplatesDir) {
  btnTemplatesDir.addEventListener('click', async () => {
    closeToolsMenu();
    try {
      if (typeof window.showDirectoryPicker !== 'function') {
        return showModalAlert({
          title: 'Browser Not Supported',
          message: 'Your browser does not support showDirectoryPicker. Please use Chrome or Edge over localhost or HTTPS.'
        });
      }

      if (cachedTemplatesDirHandle && !templatesDirHandle) {
        let perm = await cachedTemplatesDirHandle.queryPermission({ mode: 'read' });
        if (perm !== 'granted') {
          perm = await cachedTemplatesDirHandle.requestPermission({ mode: 'read' });
        }
        if (perm === 'granted') {
          templatesDirHandle = cachedTemplatesDirHandle;
          await scanTemplatesDirectory();
          showToast(`Reconnected Templates folder: "${templatesDirHandle.name}" (${availableTemplates.length} template(s))`, 'success');
          return;
        }
      }

      const tplDir = await window.showDirectoryPicker({ mode: 'read' });
      templatesDirHandle = tplDir;
      cachedTemplatesDirHandle = tplDir;
      await saveHandleToIDB(IDB_TEMPLATES_KEY, tplDir);
      await scanTemplatesDirectory();
      showToast(`Templates folder set to "${tplDir.name}" (${availableTemplates.length} .xlsx template(s) found)`, 'success');
    } catch (err) {
      if (err.name !== 'AbortError') {
        await showModalAlert({
          title: 'Templates Folder Error',
          message: 'Could not select Templates folder: ' + err.message
        });
      }
    }
  });
}

if (btnResetTemplatesDir) {
  btnResetTemplatesDir.addEventListener('click', async () => {
    closeToolsMenu();
    templatesDirHandle = null;
    cachedTemplatesDirHandle = null;
    availableTemplates = [];
    await deleteHandleFromIDB(IDB_TEMPLATES_KEY);
    updateTemplatesDirUI();
    showToast('Templates folder reset to built-in default schema.', 'info');
  });
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
  lastSavedWorkbookName = workingFileHandle.name;
  await saveHandleToIDB(IDB_LAST_WORKBOOK_KEY, workingFileHandle.name);
  updateFilePathsBarUI();

  await loadWorkbookFromDisk();
}

workbookSelect.addEventListener('change', async () => {
  const idx = Number(workbookSelect.value);
  if (availableWorkbooks[idx]) {
    workingFileHandle = availableWorkbooks[idx].handle;
    lastSavedWorkbookName = workingFileHandle.name;
    await saveHandleToIDB(IDB_LAST_WORKBOOK_KEY, workingFileHandle.name);
    clearTransientFiltersOnNavigation({ resetScopeToCurrent: true });
    dockEditingTarget = null;
    clearMirrorRetryState();
    setDeleteMode(false);
    updateFilePathsBarUI();
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

  // Pass 2: Sort Daily MM-DD-YYYY Tabs oldest -> newest before indexing so un-keyed cross-day rows inherit prior Tracking IDs
  dailySheetObjs.sort((a, b) => parseDateTab(a.rawName) - parseDateTab(b.rawName));
  dailySheetObjs.forEach(({ worksheet, rawName }) => {
    indexWorksheetRows(worksheet, rawName);
  });

  // Also enrich masterAccountsMap from existing daily rows so Broker <-> Master # auto-fill works even without a Master Accounts tab
  const seenAccountKeys = new Set(
    masterAccountsMap.map(m => `${m.broker.toLowerCase()}::${m.master.toLowerCase()}`)
  );
  allRows.forEach(r => {
    if (!r.broker && !r.master) return;
    const k = `${r.broker.toLowerCase()}::${r.master.toLowerCase()}`;
    if (!seenAccountKeys.has(k)) {
      seenAccountKeys.add(k);
      masterAccountsMap.push({
        broker: r.broker,
        master: r.master.toUpperCase(),
        bidder: r.bidder ? r.bidder.toUpperCase() : ''
      });
    }
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
  btnRollover.style.display = 'inline-flex';
  btnRepairSchema.disabled = false;
  btnAddBlankTab.disabled = false;
  btnRenameTab.disabled = dateSheets.length === 0;
  btnDeleteTab.disabled = dateSheets.length === 0;
  btnDeleteMode.disabled = allRows.length === 0;
  if (bottomEditorDock) {
    bottomEditorDock.style.display = 'block';
    syncBottomDockStateAfterLoad();
  }
  updateFilePathsBarUI();

  // Check if workbook is a fresh template with "Rename to Current Date" or missing reference tabs (Master Accounts is optional)
  const missingParts = [];
  if (dateSheets.length === 0 && !placeholderDateSheet) missingParts.push('Daily MM-DD-YYYY Tab');
  if (!hasTemplate) missingParts.push('template');
  if (!hasValidation) missingParts.push('validation lists');

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
  let placeholderDateWs = null;
  let hasAnyDateTab = false;

  wb.eachSheet((ws) => {
    const raw = ws.name.trim();
    const lower = raw.toLowerCase();
    if (DATE_TAB_REGEX.test(raw)) hasAnyDateTab = true;
    else if (isTemplatePlaceholderTab(raw)) placeholderDateWs = ws;
    else if (lower === 'template') templateWs = ws;
    else if (lower === 'validation lists' || lower === 'validation list') validationWs = ws;
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

  // Note: Master Accounts sheet is intentionally excluded when creating/bootstrapping monthly workbooks.

  // 3. Ensure at least one MM-DD-YYYY daily tab exists:
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

  // 4. Remove default empty "Sheet1" if it is completely blank and other sheets exist
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

  if (templatesDirHandle) {
    await scanTemplatesDirectory();
  }

  const now = new Date();
  const monthName = now.toLocaleString('en-US', { month: 'long' });
  const defaultFileName = `Agent Bidder Support Tracker - ${monthName} ${now.getFullYear()}.xlsx`;

  let selectOptions = null;
  if (availableTemplates.length > 0) {
    selectOptions = [
      ...availableTemplates.map((t, idx) => ({
        value: `tpl::${idx}`,
        label: `📋 From Templates Folder: ${t.name}`
      })),
      {
        value: 'builtin',
        label: '✨ Built-in Standard Schema (template, validation lists & today\'s tab — no Master Accounts)'
      }
    ];
  }

  const promptResult = await showModalPrompt({
    title: '✨ Create New Monthly Workbook',
    badge: parentDirHandle.name,
    message: availableTemplates.length > 0
      ? `Choose a template from "${templatesDirHandle.name}" or the built-in schema, and enter a filename for the new monthly tracker (.xlsx):`
      : 'Enter a filename for the new Excel tracker workbook (.xlsx). It will be initialized with template, validation lists, and today\'s daily tab (Master Accounts excluded).',
    selectLabel: 'Workbook Blueprint / Template Source',
    selectOptions,
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
  if (!promptResult) return;

  const rawName = typeof promptResult === 'object' ? promptResult.value : promptResult;
  const chosenSource = typeof promptResult === 'object' ? promptResult.selectValue : 'builtin';

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

    const newWb = new ExcelJS.Workbook();
    if (chosenSource && chosenSource.startsWith('tpl::')) {
      const tplIdx = parseInt(chosenSource.replace('tpl::', ''), 10);
      const tplEntry = availableTemplates[tplIdx];
      if (tplEntry && tplEntry.handle) {
        const tplFile = await tplEntry.handle.getFile();
        const tplBuffer = await tplFile.arrayBuffer();
        originalArchiveState = await captureOriginalArchiveState(tplBuffer);
        const sanitizedTpl = sanitizeXlsxBufferForExcelJS(tplBuffer);
        await newWb.xlsx.load(sanitizedTpl);
      } else {
        originalArchiveState = null;
      }
    } else {
      originalArchiveState = null;
    }

    bootstrapWorkbookSchemaInMemory(newWb);

    const newFileHandle = await parentDirHandle.getFileHandle(fileName, { create: true });
    workingFileHandle = newFileHandle;
    workbook = newWb;

    const saved = await saveWorkbookToDisk();
    if (saved) {
      showToast(`Created new monthly workbook: "${fileName}"`, 'success');
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
  const newTabName = (typeof rawInput === 'object' ? rawInput.value : rawInput).trim();

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
    clearTransientFiltersOnNavigation({ resetScopeToCurrent: true });
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
  const newTabName = (typeof rawNewName === 'object' ? rawNewName.value : rawNewName).trim();

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
    clearTransientFiltersOnNavigation({ resetScopeToCurrent: true });
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
      `Reference tabs (template, validation lists) will not be affected.`,
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
    clearTransientFiltersOnNavigation({ resetScopeToCurrent: true });
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

function populateSelectWithStatuses(selectEl) {
  if (!selectEl) return;
  const statusesToUse = validationStatuses.length > 0 ? validationStatuses : DEFAULT_STATUSES;
  const currentSelected = selectEl.value;
  selectEl.innerHTML = '';
  statusesToUse.forEach(st => {
    const opt = document.createElement('option');
    opt.value = st;
    opt.textContent = st;
    selectEl.appendChild(opt);
  });
  if (currentSelected && statusesToUse.includes(currentSelected)) {
    selectEl.value = currentSelected;
  }
}

function populateReferenceControls() {
  const actionsToUse = validationActions.length > 0 ? validationActions : DEFAULT_ACTIONS;
  actionListEl.innerHTML = '';
  actionsToUse.forEach(action => {
    const opt = document.createElement('option');
    opt.value = action;
    actionListEl.appendChild(opt);
  });

  populateSelectWithStatuses(fStatus);
  populateSelectWithStatuses(dockStatus);

  brokerListEl.innerHTML = '';
  const uniqueBrokers = Array.from(new Set(masterAccountsMap.map(m => m.broker).filter(Boolean)));
  uniqueBrokers.forEach(broker => {
    const opt = document.createElement('option');
    opt.value = broker;
    brokerListEl.appendChild(opt);
  });
}

function applyBrokerAutoFillToPair(brokerInputEl, masterInputEl, bidderInputEl) {
  const typedBroker = brokerInputEl.value.trim().toLowerCase();
  if (!typedBroker) return;

  const match = masterAccountsMap.find(m => m.broker.toLowerCase() === typedBroker);
  if (match) {
    if (!masterInputEl.value.trim() && match.master) {
      masterInputEl.value = match.master;
    }
    if (!bidderInputEl.value.trim() && match.bidder) {
      bidderInputEl.value = match.bidder;
    }
  }
}

function applyMasterReverseAutoFillToPair(masterInputEl, brokerInputEl, bidderInputEl) {
  const typedMaster = masterInputEl.value.trim().toLowerCase();
  if (!typedMaster) return;

  const match = masterAccountsMap.find(m => m.master.toLowerCase() === typedMaster);
  if (match) {
    if (!brokerInputEl.value.trim() && match.broker) {
      brokerInputEl.value = match.broker;
    }
    if (!bidderInputEl.value.trim() && match.bidder) {
      bidderInputEl.value = match.bidder;
    }
  }
}

fBroker.addEventListener('input', () => applyBrokerAutoFillToPair(fBroker, fMaster, fBidder));
fBroker.addEventListener('change', () => applyBrokerAutoFillToPair(fBroker, fMaster, fBidder));
fMaster.addEventListener('input', () => applyMasterReverseAutoFillToPair(fMaster, fBroker, fBidder));
fMaster.addEventListener('change', () => applyMasterReverseAutoFillToPair(fMaster, fBroker, fBidder));

if (dockBroker && dockMaster && dockBidder) {
  dockBroker.addEventListener('input', () => applyBrokerAutoFillToPair(dockBroker, dockMaster, dockBidder));
  dockBroker.addEventListener('change', () => applyBrokerAutoFillToPair(dockBroker, dockMaster, dockBidder));
  dockMaster.addEventListener('input', () => applyMasterReverseAutoFillToPair(dockMaster, dockBroker, dockBidder));
  dockMaster.addEventListener('change', () => applyMasterReverseAutoFillToPair(dockMaster, dockBroker, dockBidder));
}

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

  const ticketA = normalizeTicketId(rowA.ticket);
  const ticketB = normalizeTicketId(rowB.ticket);

  // Explicit, different non-empty Ticket Numbers indicate distinct requests
  if (ticketA && ticketB && ticketA !== ticketB) {
    return false;
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

  // Case A: Same valid Ticket Number + compatible Lot Numbers & Account
  if (ticketA && ticketB && ticketA === ticketB) {
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

  listBody.innerHTML = '';
  const colSpan = isDeleteMode ? 12 : 11;
  if (sorted.length === 0) {
    listBody.innerHTML = `<tr><td colspan="${colSpan}" class="empty-table-cell">No matching requests found.</td></tr>`;
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
    tr.dataset.sheet = item.sheetName;
    tr.dataset.row = String(item.rowNumber);

    const compId = getRowCompositeId(item.sheetName, item.rowNumber);
    const isCheckedForDelete = isDeleteMode && selectedForDelete.has(compId);
    if (isCheckedForDelete) {
      tr.classList.add('selected-for-delete');
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

  syncSelectAllCheckboxes();
  updateStatusBarTelemetry();
}

// --- DELETE SELECTION MODE & CONFIRMATION ---
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
  thDeleteSelect.style.display = isDeleteMode ? 'table-cell' : 'none';
  updateDeleteSelectionUI();
  renderList();
}

function updateDeleteSelectionUI() {
  const count = selectedForDelete.size;
  deleteSelectedCountEl.textContent = String(count);
  btnConfirmDeleteSelected.textContent = `Delete Selected (${count})`;
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
      <strong>Future-Dated Tab Detected:</strong> Your workbook's latest sheet is
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

function buildConsolidatedRolloverItem(entries, baseRow) {
  const mergedNotes = mergeChronologicalActionNotes(entries);
  const canonicalRow = entries[entries.length - 1] || baseRow;
  const pickLatestIfUnkeyedMerge = (primaryVal, fieldName) => {
    if (primaryVal && String(primaryVal).trim() !== '') return primaryVal;
    const allSameKey = entries.every(
      e => e.key && baseRow.key && e.key.trim().toLowerCase() === baseRow.key.trim().toLowerCase()
    );
    if (allSameKey) return primaryVal;
    for (let i = entries.length - 1; i >= 0; i--) {
      const candidate = entries[i][fieldName];
      if (candidate && String(candidate).trim() !== '') return candidate;
    }
    return primaryVal;
  };

  return {
    ...baseRow,
    key: canonicalRow.key || baseRow.key,
    subject: pickLatestIfUnkeyedMerge(baseRow.subject, 'subject'),
    ticket: pickLatestIfUnkeyedMerge(baseRow.ticket, 'ticket'),
    broker: pickLatestIfUnkeyedMerge(baseRow.broker, 'broker'),
    master: pickLatestIfUnkeyedMerge(baseRow.master, 'master'),
    bidder: pickLatestIfUnkeyedMerge(baseRow.bidder, 'bidder'),
    action: pickLatestIfUnkeyedMerge(baseRow.action, 'action'),
    yard: pickLatestIfUnkeyedMerge(baseRow.yard, 'yard'),
    lots: pickLatestIfUnkeyedMerge(baseRow.lots, 'lots'),
    notes: mergedNotes,
    _entries: entries,
    _sourceTabCount: entries.length,
    _firstSheet: entries[0] ? entries[0].sheetName : baseRow.sheetName
  };
}

function unifyHistoricalGroupTrackingIds(entries, canonicalKey) {
  if (!Array.isArray(entries) || !canonicalKey || !workbook) return;
  const normCanonical = canonicalKey.trim().toLowerCase();

  entries.forEach(histRow => {
    if (!histRow || !histRow.sheetName || !histRow.rowNumber) return;
    if (histRow.key && histRow.key.trim().toLowerCase() === normCanonical) return;

    const ws = workbook.getWorksheet(histRow.sheetName);
    if (!ws) return;
    const colMap = getWorksheetColumnMap(ws);

    // Ensure no other distinct row on the same sheet already holds canonicalKey
    let collision = false;
    ws.eachRow({ includeEmpty: false }, (r, rNum) => {
      if (rNum === 1 || rNum === histRow.rowNumber) return;
      if (cellText(r, colMap.trackingCol).trim().toLowerCase() === normCanonical) {
        collision = true;
      }
    });
    if (collision) return;

    const excelRow = ws.getRow(histRow.rowNumber);
    excelRow.getCell(colMap.trackingCol).value = canonicalKey;
    excelRow.commit();
  });
}

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
        reopenedOnSource: isTargetComplete && !isLatestComplete,
        hasMissingPriorNotes,
        mergedTargetNotes,
        groupEntries: group.entries
      });
    } else if (!isLatestComplete) {
      toRollOver.push(consolidatedLatest);
    } else if (group.lastOpenRow) {
      const consolidatedOpen = buildConsolidatedRolloverItem(group.entries, group.lastOpenRow);
      supersededByComplete.push({
        openRow: consolidatedOpen,
        completedRow: latestRow,
        groupEntries: group.entries
      });
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
    rolloverPreviewBox.innerHTML = `<div style="color:var(--warning-text);">Please enter a valid target date in exact <code>MM-DD-YYYY</code> format.</div>`;
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
    : `<span class="badge badge-pending">New Tab (Will be created &amp; cloned from template)</span>`;

  let html = `
    <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px; padding-bottom:8px; border-bottom:1px solid var(--border);">
      <div><strong>Target:</strong> <code>${escapeHtml(targetVal)}</code></div>
      ${tabStateBadge}
    </div>
  `;

  html += `<div>
    <div class="preview-section-title">Unfinished Unique Requests to Roll Over (${plan.toRollOver.length})</div>`;
  if (plan.toRollOver.length === 0) {
    html += `<div style="color:var(--text-muted); font-size:0.78125rem;">No new unfinished requests found across source sheets.</div>`;
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
            <strong>${escapeHtml(item.subject)}</strong>
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
            <strong>${escapeHtml(targetRow.subject)}</strong>
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
              <strong>${escapeHtml(openRow.subject)}</strong>
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
        if (entry.reopenedOnSource) {
          excelRow.getCell(targetColMap.statusCol).value = entry.latestRow.status;
          reopenedCount++;
        }
        if (entry.hasMissingPriorNotes && entry.mergedTargetNotes) {
          excelRow.getCell(targetColMap.notesCol).value = entry.mergedTargetNotes;
          notesSyncedCount++;
        }
        excelRow.commit();
      }
    }

    const saveResult = await saveWorkbookToDisk();
    if (saveResult) {
      rolloverModal.classList.remove('open');
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
    updateFilePathsBarUI();
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
      statusFolderInfo.textContent = 'Folder: Not connected';
    } else {
      const supName = supervisorDirHandle ? supervisorDirHandle.name : 'Same as Agent';
      statusFolderInfo.textContent = `Agent: ${parentDirHandle.name} • Mirror: ${supName}`;
    }
  }

  if (statusRowStats) {
    if (!workbook) {
      statusRowStats.textContent = '0 requests';
    } else {
      const openTotal = allRows.filter(r => !isCompletedStatus(r.status)).length;
      statusRowStats.textContent = `Showing ${currentlyVisibleRows.length} of ${allRows.length} requests (${openTotal} open across ${dateSheets.length} tab${dateSheets.length === 1 ? '' : 's'})`;
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

// --- 9. APP VERSION & PWA SERVICE WORKER REGISTRATION ---
const APP_VERSION = 'v1.12.0';

if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js')
      .then((reg) => {
        console.log(`Service Worker (${APP_VERSION}) registered with scope:`, reg.scope);
        reg.addEventListener('updatefound', () => {
          const newWorker = reg.installing;
          if (newWorker) {
            newWorker.addEventListener('statechange', () => {
              if (newWorker.state === 'installed' && navigator.serviceWorker.controller) {
                showToast('TrackerUI updated. Reload to apply the latest version.', 'info');
              }
            });
          }
        });
      })
      .catch((err) => console.warn('Service Worker registration failed:', err));
  });
}
