// ============================================================================
// core-ui.js — Shared App State, DOM Elements, Theme/Text Scaling, Dialogs,
//              Tools Menu, Status Bar Telemetry, Toasts & PWA Update Safeguards
// ============================================================================

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

// In-App Workbook Scratch Sheet Tabs State (v1.14.0)
const DAILY_TRACKER_WS_TAB = '__DAILY_TRACKER__';
const SCRATCH_OPEN_TABS_STORAGE_KEY = 'trackerui_open_scratch_tabs_v1';
let scratchSheets = []; // [{ name, nonEmptyRows, maxCol }] custom sheets in active workbook
let openScratchTabs = []; // Ordered sheet names opened as tabs in #workspaceTabsBar
let activeWorkspaceTab = DAILY_TRACKER_WS_TAB; // DAILY_TRACKER_WS_TAB or scratch sheet name
const scratchGridDimensionsMap = new Map(); // sheetName -> { rows, cols }
const scratchActiveCellMap = new Map(); // sheetName -> { row, col }
const scratchPendingEdits = new Map(); // sheetName -> Map<'row::col', rawValue>
let scratchAutoSaveTimer = null;
let scratchIsSaving = false;
let scratchInlineEditState = null; // { sheetName, row, col, originalRaw, inputEl, tdEl }

// Mirror Retry & Primary Save Fail-Safe State
let pendingMirrorBuffer = null;
let mirrorCountdownInterval = null;
let mirrorCountdownSec = 0;
let mirrorRetryCount = 0;
const MIRROR_RETRY_INTERVAL_SEC = 15;
const MIRROR_MAX_AUTO_RETRIES = 20;
let pendingPrimarySaveState = null; // { profileId, workbookName, savedAt, buffer }

// Multi-Department Tracker Workspace Profiles & Dynamic Column Schema State (v1.17.0)
const DEFAULT_TRACKER_PROFILE_ID = 'default';
const TRACKER_PROFILES_STORAGE_KEY = 'trackerui_tracker_profiles_v1';
const ACTIVE_PROFILE_STORAGE_KEY = 'trackerui_active_profile_id_v1';
const SCHEMA_OVERRIDES_STORAGE_KEY = 'trackerui_schema_overrides_v1';
const ACTION_CATEGORY_OVERRIDES_STORAGE_KEY = 'trackerui_action_categories_v1';

function loadSavedTrackerProfiles() {
  try {
    const raw = localStorage.getItem(TRACKER_PROFILES_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        const sanitized = parsed
          .filter(p => p && typeof p.id === 'string' && p.id.trim() !== '')
          .map(p => ({
            id: p.id.trim(),
            name: String(p.name || 'Daily Tracker').trim() || 'Daily Tracker',
            mirrorEnabled: p.mirrorEnabled !== false
          }));
        if (!sanitized.some(p => p.id === DEFAULT_TRACKER_PROFILE_ID)) {
          sanitized.unshift({ id: DEFAULT_TRACKER_PROFILE_ID, name: 'Daily Tracker', mirrorEnabled: true });
        }
        return sanitized;
      }
    }
  } catch (e) {
    console.warn('Could not parse saved tracker profiles:', e);
  }
  return [{ id: DEFAULT_TRACKER_PROFILE_ID, name: 'Daily Tracker', mirrorEnabled: true }];
}

let trackerProfiles = loadSavedTrackerProfiles();
let activeProfileId = (() => {
  const saved = localStorage.getItem(ACTIVE_PROFILE_STORAGE_KEY);
  if (saved && trackerProfiles.some(p => p.id === saved)) return saved;
  return DEFAULT_TRACKER_PROFILE_ID;
})();
const trackerProfileRuntimeCache = new Map(); // profileId -> runtime snapshot

function loadSavedJsonMapFromStorage(storageKey) {
  try {
    const raw = localStorage.getItem(storageKey);
    if (raw) {
      const obj = JSON.parse(raw);
      if (obj && typeof obj === 'object' && !Array.isArray(obj)) return obj;
    }
  } catch (_) { /* ignore */ }
  return {};
}

const FORM_LAYOUTS_STORAGE_KEY = 'trackerui_form_layouts_v1';
let schemaOverridesStore = loadSavedJsonMapFromStorage(SCHEMA_OVERRIDES_STORAGE_KEY);
let actionCategoryOverridesStore = loadSavedJsonMapFromStorage(ACTION_CATEGORY_OVERRIDES_STORAGE_KEY);
let formLayoutsStore = loadSavedJsonMapFromStorage(FORM_LAYOUTS_STORAGE_KEY);

function saveFormLayoutsStore() {
  try {
    localStorage.setItem(FORM_LAYOUTS_STORAGE_KEY, JSON.stringify(formLayoutsStore || {}));
  } catch (_) { /* ignore */ }
}

// Active Workbook Dynamic Schema Descriptor (defaults to standard 11-col Broker layout)
let activeSchemaConfig = {
  isStandardSchema: true,
  mode: 'auto',
  columns: [],
  primaryCol: 1,
  ticketCol: 2,
  secondaryCol: 3,
  statusCol: 10,
  notesCol: 9,
  trackingCol: 11
};

// Row Selection & Batch / Delete Mode State
let isDeleteMode = false;
let currentlyVisibleRows = [];
const selectedForDelete = new Set(); // Stores "${sheetName}::${rowNumber}" for permanent row selection (Batch Status & Batch Delete)
let rolloverSourceMode = 'previous'; // 'previous' (Immediately preceding sheet only) | 'all' (All prior daily tabs)

// Dynamic Reference Data & Dev Mock Mode
let masterAccountsMap = []; // [{ broker, master, bidder, yard? }]
let hasMasterAccountsSheet = false;
let validationActions = [];
let validationStatuses = [];
let validationListsByHeader = new Map(); // lowerHeader -> string[]
let actionCategoryMap = new Map(); // lowerAction -> parentCategory

const DEV_MOCK_STORAGE_KEY = 'trackerui_dev_mock_accounts';
function isLocalhostDevEnvironment() {
  if (typeof window === 'undefined' || !window.location) return false;
  const host = String(window.location.hostname || '').toLowerCase();
  return (
    host === 'localhost' ||
    host === '127.0.0.1' ||
    host === '[::1]' ||
    window.location.protocol === 'file:'
  );
}
const storedDevMockPref = localStorage.getItem(DEV_MOCK_STORAGE_KEY);
let useDevMockAccounts = storedDevMockPref !== null
  ? storedDevMockPref === 'true'
  : isLocalhostDevEnvironment();

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
  'completed',
  'new',
  'waiting',
  'overdue',
  'closed',
  'resolved',
  'done',
  'cancelled',
  'canceled'
]);

const COMPLETED_STATUS_SET = new Set([
  'complete',
  'completed',
  'closed',
  'resolved',
  'done',
  'cancelled',
  'canceled'
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

function isProtectedReferenceSheetName(sheetName) {
  const norm = String(sheetName || '').trim().toLowerCase();
  return (
    norm === 'template' ||
    norm === 'validation list' ||
    norm === 'validation lists' ||
    norm === 'master account' ||
    norm === 'master accounts'
  );
}

function isCustomScratchSheetName(sheetName, wsState = 'visible') {
  const trimmed = String(sheetName || '').trim();
  if (!trimmed) return false;
  if (wsState === 'veryHidden' || wsState === 'hidden') return false;
  if (DATE_TAB_REGEX.test(trimmed)) return false;
  if (isTemplatePlaceholderTab(trimmed)) return false;
  if (isProtectedReferenceSheetName(trimmed)) return false;
  return true;
}

function isCompletedStatus(status) {
  const norm = String(status || '').trim().toLowerCase();
  if (!norm) return false;
  if (COMPLETED_STATUS_SET.has(norm)) return true;
  return (
    norm.startsWith('complete') ||
    norm.startsWith('closed') ||
    norm.startsWith('resolved')
  );
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
const btnRetryPrimarySave = document.getElementById('btnRetryPrimarySave');
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
const menuOpenMasterRollup = document.getElementById('menuOpenMasterRollup');
const chkEnableMirrorCopy = document.getElementById('chkEnableMirrorCopy');
const chkAutoReconnect = document.getElementById('chkAutoReconnect');
const chkAutoReconnectBar = document.getElementById('chkAutoReconnectBar');
const btnNewWorkbook = document.getElementById('btnNewWorkbook');
const btnRepairSchema = document.getElementById('btnRepairSchema');
const btnOpenSchemaMapper = document.getElementById('btnOpenSchemaMapper');
const menuAddTrackerProfile = document.getElementById('menuAddTrackerProfile');
const menuRenameTrackerProfile = document.getElementById('menuRenameTrackerProfile');
const menuNewScratchSheet = document.getElementById('menuNewScratchSheet');
const chkDevMockAccounts = document.getElementById('chkDevMockAccounts');
const devMockStateBadge = document.getElementById('devMockStateBadge');
const btnAddBlankTab = document.getElementById('btnAddBlankTab');
const btnRolloverPrevSheet = document.getElementById('btnRolloverPrevSheet');
const btnRenameTab = document.getElementById('btnRenameTab');
const btnDeleteTab = document.getElementById('btnDeleteTab');
const btnDeleteMode = document.getElementById('btnDeleteMode');
const deleteModeLabel = document.getElementById('deleteModeLabel');
const btnExportConfigJson = document.getElementById('btnExportConfigJson');
const btnImportConfigJson = document.getElementById('btnImportConfigJson');
const configJsonFileInput = document.getElementById('configJsonFileInput');
const btnCheckForUpdates = document.getElementById('btnCheckForUpdates');
const appVersionTag = document.getElementById('appVersionTag');
const pwaUpdateBanner = document.getElementById('pwaUpdateBanner');
const btnPwaReloadNow = document.getElementById('btnPwaReloadNow');
const btnPwaDismissUpdate = document.getElementById('btnPwaDismissUpdate');

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

// Row Selection, Batch Action Dock & Delete Mode DOM Elements
const deleteBanner = document.getElementById('deleteBanner');
const thDeleteSelect = document.getElementById('thDeleteSelect');
const chkSelectAllHeader = document.getElementById('chkSelectAllHeader');
const chkSelectAllBanner = document.getElementById('chkSelectAllBanner');
const deleteSelectedCountEl = document.getElementById('deleteSelectedCount');
const btnConfirmDeleteSelected = document.getElementById('btnConfirmDeleteSelected');
const btnCancelDeleteMode = document.getElementById('btnCancelDeleteMode');

const batchActionDock = document.getElementById('batchActionDock');
const batchSelectedCountEl = document.getElementById('batchSelectedCount');
const batchStatusSelect = document.getElementById('batchStatusSelect');
const btnApplyBatchStatus = document.getElementById('btnApplyBatchStatus');
const btnBatchDeleteSelected = document.getElementById('btnBatchDeleteSelected');
const btnClearBatchSelection = document.getElementById('btnClearBatchSelection');

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
const btnModalGrabTicket = document.getElementById('btnModalGrabTicket');
const btnModalPullPrior = document.getElementById('btnModalPullPrior');
const modalStandardFieldsGrid = document.getElementById('modalStandardFieldsGrid');
const modalDynamicFieldsGrid = document.getElementById('modalDynamicFieldsGrid');
const brokerListEl = document.getElementById('brokerList');
const actionListEl = document.getElementById('actionList');

// Bottom Dock Editor DOM Elements
const bottomEditorDock = document.getElementById('bottomEditorDock');
const dockHeaderToggleZone = document.getElementById('dockHeaderToggleZone');
const dockModeBadge = document.getElementById('dockModeBadge');
const dockTargetLabel = document.getElementById('dockTargetLabel');
const dockKeyBadge = document.getElementById('dockKeyBadge');
const btnDockPullPrior = document.getElementById('btnDockPullPrior');
const btnDockSyncForward = document.getElementById('btnDockSyncForward');
const btnDockTicketDetails = document.getElementById('btnDockTicketDetails');
const btnDockOpenModal = document.getElementById('btnDockOpenModal');
const btnDockToggleLayout = document.getElementById('btnDockToggleLayout');
const btnDockResetLayout = document.getElementById('btnDockResetLayout');
const btnDockNewClear = document.getElementById('btnDockNewClear');
const btnDockToggleCollapse = document.getElementById('btnDockToggleCollapse');
const dockEntryForm = document.getElementById('dockEntryForm');
const dockStandardFieldsGrid = document.getElementById('dockStandardFieldsGrid');
const dockDynamicFieldsGrid = document.getElementById('dockDynamicFieldsGrid');
const dockSubject = document.getElementById('dockSubject');
const dockTicket = document.getElementById('dockTicket');
const btnDockGrabTicket = document.getElementById('btnDockGrabTicket');
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
const btnTicketSyncForward = document.getElementById('btnTicketSyncForward');
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
const viewOnlyHead = document.getElementById('viewOnlyHead');
const viewOnlyBody = document.getElementById('viewOnlyBody');
const btnRefreshViewOnly = document.getElementById('btnRefreshViewOnly');
const btnCloseViewOnly = document.getElementById('btnCloseViewOnly');

// Rollover Modal DOM Elements
const rolloverForm = document.getElementById('rolloverForm');
const rolloverTodayBadge = document.getElementById('rolloverTodayBadge');
const rolloverFutureBanner = document.getElementById('rolloverFutureBanner');
const rolloverTargetDate = document.getElementById('rolloverTargetDate');
const rolloverPresetBtns = document.getElementById('rolloverPresetBtns');
const rolloverSourceWrap = document.getElementById('rolloverSourceWrap');
const rolloverSourceGroup = document.getElementById('rolloverSourceGroup');
const btnRolloverSourcePrev = document.getElementById('btnRolloverSourcePrev');
const btnRolloverSourceAll = document.getElementById('btnRolloverSourceAll');
const rolloverSourceDesc = document.getElementById('rolloverSourceDesc');
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

// Workspace Tabs Bar & Editable Scratch Sheet DOM Elements
const workspaceTabsBar = document.getElementById('workspaceTabsBar');
const workspaceTabsList = document.getElementById('workspaceTabsList');
const btnAddTrackerProfile = document.getElementById('btnAddTrackerProfile');
const scratchSheetMenuWrap = document.getElementById('scratchSheetMenuWrap');
const btnOpenScratchMenu = document.getElementById('btnOpenScratchMenu');
const scratchAvailableCountBadge = document.getElementById('scratchAvailableCountBadge');
const scratchSheetDropdown = document.getElementById('scratchSheetDropdown');
const scratchSheetMenuList = document.getElementById('scratchSheetMenuList');
const btnNewScratchSheetFromBar = document.getElementById('btnNewScratchSheetFromBar');
const dailyTrackerWorkspace = document.getElementById('dailyTrackerWorkspace');
const scratchSheetWorkspace = document.getElementById('scratchSheetWorkspace');
const scratchActiveSheetBadge = document.getElementById('scratchActiveSheetBadge');
const scratchStatsMeta = document.getElementById('scratchStatsMeta');
const scratchSaveStateBadge = document.getElementById('scratchSaveStateBadge');
const chkScratchStripedRows = document.getElementById('chkScratchStripedRows');
const scratchSearchInput = document.getElementById('scratchSearchInput');
const scratchSearchMatchCount = document.getElementById('scratchSearchMatchCount');
const chkScratchFilterRows = document.getElementById('chkScratchFilterRows');
const scratchFillMenuWrap = document.getElementById('scratchFillMenuWrap');
const btnScratchApplyFill = document.getElementById('btnScratchApplyFill');
const scratchFillActiveBar = document.getElementById('scratchFillActiveBar');
const btnScratchFillDropdownToggle = document.getElementById('btnScratchFillDropdownToggle');
const scratchFillDropdown = document.getElementById('scratchFillDropdown');
const btnScratchNoFill = document.getElementById('btnScratchNoFill');
const scratchFillPastelGrid = document.getElementById('scratchFillPastelGrid');
const scratchFillStandardGrid = document.getElementById('scratchFillStandardGrid');
const btnScratchTextToColumns = document.getElementById('btnScratchTextToColumns');
const btnScratchUndo = document.getElementById('btnScratchUndo');
const btnScratchRedo = document.getElementById('btnScratchRedo');
const btnScratchAddRows = document.getElementById('btnScratchAddRows');
const btnScratchAddCol = document.getElementById('btnScratchAddCol');
const btnScratchRenameSheet = document.getElementById('btnScratchRenameSheet');
const btnScratchDeleteSheet = document.getElementById('btnScratchDeleteSheet');
const btnScratchSaveNow = document.getElementById('btnScratchSaveNow');
const scratchCellCoord = document.getElementById('scratchCellCoord');
const scratchFormulaInput = document.getElementById('scratchFormulaInput');
const scratchSelectionStats = document.getElementById('scratchSelectionStats');
const scratchGridWrap = document.getElementById('scratchGridWrap');
const scratchGridTable = document.getElementById('scratchGridTable');
const scratchGridHead = document.getElementById('scratchGridHead');
const scratchGridBody = document.getElementById('scratchGridBody');
const scratchContextMenu = document.getElementById('scratchContextMenu');
const scratchCtxRangeLabel = document.getElementById('scratchCtxRangeLabel');
const scratchCtxQuickSwatches = document.getElementById('scratchCtxQuickSwatches');
const textToColumnsModal = document.getElementById('textToColumnsModal');
const textToColumnsForm = document.getElementById('textToColumnsForm');
const ttcSourceBadge = document.getElementById('ttcSourceBadge');
const ttcSubtitle = document.getElementById('ttcSubtitle');
const chkTtcComma = document.getElementById('chkTtcComma');
const chkTtcTab = document.getElementById('chkTtcTab');
const chkTtcSemicolon = document.getElementById('chkTtcSemicolon');
const chkTtcPipe = document.getElementById('chkTtcPipe');
const chkTtcHyphen = document.getElementById('chkTtcHyphen');
const chkTtcSpace = document.getElementById('chkTtcSpace');
const chkTtcOther = document.getElementById('chkTtcOther');
const ttcOtherInput = document.getElementById('ttcOtherInput');
const chkTtcTrim = document.getElementById('chkTtcTrim');
const chkTtcConsecutive = document.getElementById('chkTtcConsecutive');
const ttcOverwriteBanner = document.getElementById('ttcOverwriteBanner');
const ttcPreviewTitle = document.getElementById('ttcPreviewTitle');
const ttcPreviewBox = document.getElementById('ttcPreviewBox');
const btnCancelTextToColumns = document.getElementById('btnCancelTextToColumns');
const btnConfirmTextToColumns = document.getElementById('btnConfirmTextToColumns');

// Dynamic Column Role Mapper Modal DOM Elements (v1.17.0)
const schemaMapperModal = document.getElementById('schemaMapperModal');
const schemaMapperSubtitle = document.getElementById('schemaMapperSubtitle');
const schemaMapperModeBadge = document.getElementById('schemaMapperModeBadge');
const schemaMapperForm = document.getElementById('schemaMapperForm');
const schemaModeSelect = document.getElementById('schemaModeSelect');
const schemaPrimaryColSelect = document.getElementById('schemaPrimaryColSelect');
const schemaStatusColSelect = document.getElementById('schemaStatusColSelect');
const schemaNotesColSelect = document.getElementById('schemaNotesColSelect');
const schemaTicketColSelect = document.getElementById('schemaTicketColSelect');
const schemaSecondaryColSelect = document.getElementById('schemaSecondaryColSelect');
const schemaTrackingColInfo = document.getElementById('schemaTrackingColInfo');
const schemaDetectedColsPreview = document.getElementById('schemaDetectedColsPreview');
const btnResetSchemaMapper = document.getElementById('btnResetSchemaMapper');
const btnCancelSchemaMapper = document.getElementById('btnCancelSchemaMapper');
const btnSaveSchemaMapper = document.getElementById('btnSaveSchemaMapper');

// Supervisor Multi-Workbook Master Rollup Modal DOM Elements (v1.17.0)
const masterRollupModal = document.getElementById('masterRollupModal');
const masterRollupTitle = document.getElementById('masterRollupTitle');
const masterRollupSubtitle = document.getElementById('masterRollupSubtitle');
const masterRollupWorkerBadge = document.getElementById('masterRollupWorkerBadge');
const btnPickMasterRollupFolder = document.getElementById('btnPickMasterRollupFolder');
const btnExportMasterRollupXlsx = document.getElementById('btnExportMasterRollupXlsx');
const masterRollupSummaryGrid = document.getElementById('masterRollupSummaryGrid');
const masterRollupCategoryBar = document.getElementById('masterRollupCategoryBar');
const masterRollupSourceMode = document.getElementById('masterRollupSourceMode');
const masterRollupWorkbookFilter = document.getElementById('masterRollupWorkbookFilter');
const masterRollupCategoryFilter = document.getElementById('masterRollupCategoryFilter');
const masterRollupStatusFilter = document.getElementById('masterRollupStatusFilter');
const masterRollupDedupCheck = document.getElementById('masterRollupDedupCheck');
const masterRollupSearchInput = document.getElementById('masterRollupSearchInput');
const masterRollupStatsText = document.getElementById('masterRollupStatsText');
const masterRollupHead = document.getElementById('masterRollupHead');
const masterRollupBody = document.getElementById('masterRollupBody');
const btnRefreshMasterRollup = document.getElementById('btnRefreshMasterRollup');
const btnCloseMasterRollup = document.getElementById('btnCloseMasterRollup');

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
    if (schemaMapperModal && schemaMapperModal.classList.contains('open')) {
      schemaMapperModal.classList.remove('open');
      return;
    }
    if (masterRollupModal && masterRollupModal.classList.contains('open')) {
      masterRollupModal.classList.remove('open');
      return;
    }
    closeToolsMenu();
  }
});

function colNumberToExcelLetters(colNum) {
  let n = Math.max(1, Math.floor(colNum || 1));
  let letters = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

// --- TICKET EXTRACTION ([XXXX-XX-XXXX]), DEDUPLICATION & CATEGORY PIPELINE ---
const BRACKETED_TICKET_REGEX = /\[([A-Za-z0-9]{2,12}(?:-[A-Za-z0-9]{1,12})+)\]/g;
const TICKET_PLACEHOLDER_IGNORE_SET = new Set([
  'n/a', 'na', 'none', 'null', 'tbd', 'pending', '-', '—', '--', '0', 'unknown'
]);

function extractBracketedTicketsFromSubject(subjectText) {
  const raw = String(subjectText || '');
  if (!raw) return [];
  const seen = new Set();
  const extracted = [];
  let match;
  BRACKETED_TICKET_REGEX.lastIndex = 0;
  while ((match = BRACKETED_TICKET_REGEX.exec(raw)) !== null) {
    const token = String(match[1] || '').trim().toUpperCase();
    if (token && !seen.has(token)) {
      seen.add(token);
      extracted.push(token);
    }
  }
  return extracted;
}

function splitNormalizedTicketTokens(rawTicketStr) {
  if (!rawTicketStr) return [];
  const seen = new Set();
  const out = [];
  String(rawTicketStr)
    .split(/[,;]+/)
    .forEach(part => {
      const cleaned = part.trim().replace(/^#+\s*/, '').replace(/^\[|\]$/g, '').trim();
      if (!cleaned) return;
      if (TICKET_PLACEHOLDER_IGNORE_SET.has(cleaned.toLowerCase())) return;
      const upper = cleaned.toUpperCase();
      if (!seen.has(upper)) {
        seen.add(upper);
        out.push(upper);
      }
    });
  return out;
}

function mergeCommaDelimitedTickets(existingStr, incomingListOrStr) {
  const existingTokens = splitNormalizedTicketTokens(existingStr);
  const incomingTokens = Array.isArray(incomingListOrStr)
    ? incomingListOrStr.flatMap(t => splitNormalizedTicketTokens(t))
    : splitNormalizedTicketTokens(incomingListOrStr);

  const seen = new Set(existingTokens);
  const added = [];
  incomingTokens.forEach(tok => {
    if (!seen.has(tok)) {
      seen.add(tok);
      existingTokens.push(tok);
      added.push(tok);
    }
  });

  return {
    merged: existingTokens.join(', '),
    added,
    allTokens: existingTokens
  };
}

// Supervisor Pipeline: Subtype -> Parent Category Lookup
const BUILTIN_ACTION_CATEGORY_RULES = [
  { keywords: ['title', 'poa', 'lien', 'ownership', 'mv-907', 'salvage', 'bos', 'bill of sale'], category: 'Titles & Ownership' },
  { keywords: ['account', 'unlock', 'password', 'login', 'bidder', 'linking', 'registration', 'kyc', 'license'], category: 'Account & Access' },
  { keywords: ['bid limit', 'deposit', 'qualification', 'buying power', 'eligibility'], category: 'Bidding & Eligibility' },
  { keywords: ['wire', 'payment', 'invoice', 'fee', 'storage', 'waiver', 'refund', 'credit', 'relist', 'billing'], category: 'Billing, Fees & Payments' },
  { keywords: ['gate pass', 'pickup', 'release', 'transport', 'yard', 'sublot', 'keys', 'forklift'], category: 'Yard & Release Ops' },
  { keywords: ['cancel', 'dispute', 'condition', 'arbitration', 'damage', 'lot'], category: 'Lot & Dispute Review' }
];

function resolveActionParentCategory(actionStr) {
  const clean = String(actionStr || '').trim();
  if (!clean || clean === '—') return 'Uncategorized';
  const lower = clean.toLowerCase();

  if (actionCategoryOverridesStore && actionCategoryOverridesStore[lower]) {
    return actionCategoryOverridesStore[lower];
  }
  if (actionCategoryMap && actionCategoryMap.has(lower)) {
    return actionCategoryMap.get(lower);
  }
  for (const rule of BUILTIN_ACTION_CATEGORY_RULES) {
    if (rule.keywords.some(kw => lower.includes(kw))) {
      return rule.category;
    }
  }
  return 'General Support';
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
  const activeProfile = trackerProfiles.find(p => p.id === activeProfileId) || trackerProfiles[0];
  const profilePrefix = trackerProfiles.length > 1 && activeProfile
    ? `[${activeProfile.name}] `
    : '';
  const mirrorEnabled = activeProfile ? activeProfile.mirrorEnabled !== false : true;

  if (statusFolderInfo) {
    if (!parentDirHandle) {
      statusFolderInfo.textContent = `${profilePrefix}Folder: Not connected`;
    } else {
      const supName = !mirrorEnabled
        ? 'Off'
        : (supervisorDirHandle ? supervisorDirHandle.name : 'Same as Agent');
      statusFolderInfo.textContent = `${profilePrefix}Agent: ${parentDirHandle.name} • Mirror: ${supName}`;
    }
  }

  if (statusRowStats) {
    if (!workbook) {
      statusRowStats.textContent = '0 requests';
    } else if (activeWorkspaceTab !== DAILY_TRACKER_WS_TAB) {
      const openTotal = allRows.filter(r => !isCompletedStatus(r.status)).length;
      const meta = scratchSheets.find(s => s.name === activeWorkspaceTab);
      const sRows = meta ? meta.nonEmptyRows : 0;
      statusRowStats.textContent = `Scratch Tab: ${activeWorkspaceTab} (${sRows} row${sRows === 1 ? '' : 's'}) • ${activeProfile ? activeProfile.name : 'Daily Tracker'}: ${openTotal} open`;
    } else {
      const openTotal = allRows.filter(r => !isCompletedStatus(r.status)).length;
      const schemaNote = !activeSchemaConfig.isStandardSchema
        ? ` • ${activeSchemaConfig.columns.length} custom cols`
        : '';
      statusRowStats.textContent = `Showing ${currentlyVisibleRows.length} of ${allRows.length} requests (${openTotal} open across ${dateSheets.length} tab${dateSheets.length === 1 ? '' : 's'}${schemaNote})`;
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

// --- 9. APP VERSION & PWA SERVICE WORKER REGISTRATION & CACHE SAFEGUARDS ---
const APP_VERSION = 'v1.17.1';
let activeSwRegistration = null;
let lastSwUpdateCheckAt = 0;
const SW_UPDATE_CHECK_THROTTLE_MS = 45000;

function showPwaUpdateBanner() {
  if (pwaUpdateBanner) {
    pwaUpdateBanner.style.display = 'flex';
  }
}

async function forceClearAppCacheAndReload() {
  closeToolsMenu();
  showToast('Refreshing application cache and loading latest version...', 'info');
  try {
    if ('serviceWorker' in navigator) {
      const regs = await navigator.serviceWorker.getRegistrations();
      await Promise.all(
        regs.map(async (r) => {
          try {
            await r.update();
          } catch (_) { /* ignore */ }
          if (r.waiting) {
            r.waiting.postMessage({ type: 'SKIP_WAITING' });
          }
        })
      );
    }
    if ('caches' in window) {
      const keys = await caches.keys();
      await Promise.all(keys.map((k) => caches.delete(k)));
    }
  } catch (err) {
    console.warn('Error clearing PWA cache:', err);
  }
  window.location.reload();
}

function checkServiceWorkerUpdateOnResume() {
  if (!activeSwRegistration) return;
  const now = Date.now();
  if (now - lastSwUpdateCheckAt < SW_UPDATE_CHECK_THROTTLE_MS) return;
  lastSwUpdateCheckAt = now;
  activeSwRegistration.update().catch(() => { /* ignore offline errors */ });
}

if (btnCheckForUpdates) {
  btnCheckForUpdates.addEventListener('click', () => {
    forceClearAppCacheAndReload();
  });
}

if (appVersionTag) {
  appVersionTag.addEventListener('click', () => {
    forceClearAppCacheAndReload();
  });
}

if (btnPwaReloadNow) {
  btnPwaReloadNow.addEventListener('click', () => {
    forceClearAppCacheAndReload();
  });
}

if (btnPwaDismissUpdate) {
  btnPwaDismissUpdate.addEventListener('click', () => {
    if (pwaUpdateBanner) {
      pwaUpdateBanner.style.display = 'none';
    }
  });
}

if ('serviceWorker' in navigator) {
  const hadExistingController = Boolean(navigator.serviceWorker.controller);

  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (hadExistingController) {
      showPwaUpdateBanner();
    }
  });

  window.addEventListener('load', () => {
    navigator.serviceWorker.register('./sw.js', { updateViaCache: 'none' })
      .then((reg) => {
        activeSwRegistration = reg;
        lastSwUpdateCheckAt = Date.now();
        console.log(`Service Worker (${APP_VERSION}) registered with scope:`, reg.scope);

        if (reg.waiting && navigator.serviceWorker.controller) {
          showPwaUpdateBanner();
        }

        reg.addEventListener('updatefound', () => {
          const newWorker = reg.installing;
          if (newWorker) {
            newWorker.addEventListener('statechange', () => {
              if (newWorker.state === 'installed' && navigator.serviceWorker.controller) {
                showPwaUpdateBanner();
              }
            });
          }
        });
      })
      .catch((err) => console.warn('Service Worker registration failed:', err));
  });

  window.addEventListener('focus', checkServiceWorkerUpdateOnResume);
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') {
      checkServiceWorkerUpdateOnResume();
    }
  });
}
