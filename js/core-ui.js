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

function normalizeActionCategoryOverridesMap(rawObj) {
  const out = {};
  if (!rawObj || typeof rawObj !== 'object' || Array.isArray(rawObj)) return out;
  Object.entries(rawObj).forEach(([k, v]) => {
    const normKey = String(k || '').replace(/\s+/g, ' ').trim().toLowerCase();
    const normVal = String(v || '').replace(/\s+/g, ' ').trim();
    if (normKey && normVal) {
      out[normKey] = normVal;
    }
  });
  return out;
}

const FORM_LAYOUTS_STORAGE_KEY = 'trackerui_form_layouts_v1';
const TABLE_COL_WIDTHS_STORAGE_KEY = 'trackerui_table_col_widths_v1';
const ROW_LAST_ACCESSED_STORAGE_KEY = 'trackerui_row_last_accessed_v1';
const ROW_LAST_ACCESSED_MAX_ENTRIES = 5000;
let schemaOverridesStore = loadSavedJsonMapFromStorage(SCHEMA_OVERRIDES_STORAGE_KEY);
let actionCategoryOverridesStore = normalizeActionCategoryOverridesMap(
  loadSavedJsonMapFromStorage(ACTION_CATEGORY_OVERRIDES_STORAGE_KEY)
);
let formLayoutsStore = loadSavedJsonMapFromStorage(FORM_LAYOUTS_STORAGE_KEY);
let tableColWidthsStore = loadSavedJsonMapFromStorage(TABLE_COL_WIDTHS_STORAGE_KEY);
let rowLastAccessedStore = loadSavedJsonMapFromStorage(ROW_LAST_ACCESSED_STORAGE_KEY);

function saveTableColWidthsStore() {
  try {
    localStorage.setItem(TABLE_COL_WIDTHS_STORAGE_KEY, JSON.stringify(tableColWidthsStore || {}));
  } catch (_) { /* ignore */ }
}

function saveRowLastAccessedStore() {
  try {
    const keys = Object.keys(rowLastAccessedStore || {});
    if (keys.length > ROW_LAST_ACCESSED_MAX_ENTRIES) {
      const sorted = keys
        .map(k => [k, Number(rowLastAccessedStore[k]) || 0])
        .sort((a, b) => b[1] - a[1])
        .slice(0, ROW_LAST_ACCESSED_MAX_ENTRIES);
      rowLastAccessedStore = Object.fromEntries(sorted);
    }
    localStorage.setItem(ROW_LAST_ACCESSED_STORAGE_KEY, JSON.stringify(rowLastAccessedStore || {}));
  } catch (_) { /* ignore */ }
}

function recordRowLastAccessed(trackingKey, tsMs = Date.now()) {
  const normKey = String(trackingKey || '').trim().toLowerCase();
  if (!normKey) return 0;
  const stamp = Number.isFinite(Number(tsMs)) ? Number(tsMs) : Date.now();
  if (!rowLastAccessedStore || typeof rowLastAccessedStore !== 'object') {
    rowLastAccessedStore = {};
  }
  rowLastAccessedStore[normKey] = stamp;
  saveRowLastAccessedStore();
  return stamp;
}

function getRowLastAccessedMs(trackingKey) {
  const normKey = String(trackingKey || '').trim().toLowerCase();
  if (!normKey || !rowLastAccessedStore || typeof rowLastAccessedStore !== 'object') return 0;
  const val = Number(rowLastAccessedStore[normKey]);
  return Number.isFinite(val) && val > 0 ? val : 0;
}

function formatRowLastAccessedLabel(tsMs) {
  const num = Number(tsMs);
  if (!Number.isFinite(num) || num <= 0) {
    return { ms: 0, short: '—', full: 'Not accessed yet' };
  }
  const dt = new Date(num);
  if (isNaN(dt.getTime())) {
    return { ms: 0, short: '—', full: 'Not accessed yet' };
  }
  const now = new Date();
  const timeShort = dt.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  const isSameDay =
    dt.getFullYear() === now.getFullYear() &&
    dt.getMonth() === now.getMonth() &&
    dt.getDate() === now.getDate();
  const short = isSameDay
    ? timeShort
    : `${dt.getMonth() + 1}/${dt.getDate()} ${timeShort}`;
  const full = `Last accessed: ${dt.toLocaleString([], {
    month: '2-digit',
    day: '2-digit',
    year: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
    second: '2-digit'
  })}`;
  return { ms: num, short, full };
}

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
  actionCol: 6,
  statusCol: 10,
  notesCol: 9,
  emailSubjCol: null,
  trackingCol: 11
};

// Row Selection & Batch / Delete Mode State
let isDeleteMode = false;
let currentlyVisibleRows = [];
const selectedForDelete = new Set(); // Stores "${sheetName}::${rowNumber}" for permanent row selection (Batch Status, Copy/Paste & Batch Delete)
let copiedTrackerRowsClipboard = []; // Cloned row items copied via checkbox selection for manual cross-tab paste
let rolloverSourceMode = 'previous'; // 'previous' (Immediately preceding sheet only) | 'all' (All prior daily tabs)

// Dynamic Reference Data & Dev Mock Mode
let masterAccountsMap = []; // [{ broker, master }]
let hasMasterAccountsSheet = false;
let validationActions = [];
let validationStatuses = [];
let validationYards = []; // [{ value, name, label }]
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

// Guard against VS Code Live Server auto-reloading the browser page when TrackerUI
// auto-saves a working/test .xlsx file inside the project workspace (e.g. assets/Testing/*.xlsx).
if (typeof window !== 'undefined' && isLocalhostDevEnvironment() && typeof window.WebSocket === 'function') {
  const NativeWebSocket = window.WebSocket;
  const shouldSuppressLiveServerReload = () => {
    const now = Date.now();
    if (window.__trackerDiskWriteInFlight) return true;
    if (typeof window.__lastTrackerWorkbookWriteAt === 'number' && now - window.__lastTrackerWorkbookWriteAt < 5000) {
      return true;
    }
    if (typeof window.__lastTrackerScratchEditAt === 'number' && now - window.__lastTrackerScratchEditAt < 6000) {
      return true;
    }
    return false;
  };

  function GuardedWebSocket(url, protocols) {
    const wsInstance = protocols !== undefined
      ? new NativeWebSocket(url, protocols)
      : new NativeWebSocket(url);
    const urlStr = String(url || '');
    if (/\/ws(?:\?|$)/i.test(urlStr)) {
      const origAddEventListener = wsInstance.addEventListener.bind(wsInstance);
      wsInstance.addEventListener = function (type, listener, options) {
        if (type === 'message' && typeof listener === 'function') {
          const wrapped = function (evt) {
            if (evt && evt.data === 'reload' && shouldSuppressLiveServerReload()) {
              console.info('[TrackerUI] Suppressed Live Server page reload triggered by in-workspace .xlsx auto-save.');
              return;
            }
            return listener.call(this, evt);
          };
          return origAddEventListener(type, wrapped, options);
        }
        return origAddEventListener(type, listener, options);
      };

      let assignedOnMessage = null;
      Object.defineProperty(wsInstance, 'onmessage', {
        configurable: true,
        enumerable: true,
        get() {
          return assignedOnMessage;
        },
        set(fn) {
          if (assignedOnMessage) {
            wsInstance.removeEventListener('message', assignedOnMessage);
          }
          if (typeof fn !== 'function') {
            assignedOnMessage = null;
            return;
          }
          assignedOnMessage = function (evt) {
            if (evt && evt.data === 'reload' && shouldSuppressLiveServerReload()) {
              console.info('[TrackerUI] Suppressed Live Server page reload triggered by in-workspace .xlsx auto-save.');
              return;
            }
            return fn.call(this, evt);
          };
          origAddEventListener('message', assignedOnMessage);
        }
      });
    }
    return wsInstance;
  }
  GuardedWebSocket.prototype = NativeWebSocket.prototype;
  GuardedWebSocket.CONNECTING = NativeWebSocket.CONNECTING;
  GuardedWebSocket.OPEN = NativeWebSocket.OPEN;
  GuardedWebSocket.CLOSING = NativeWebSocket.CLOSING;
  GuardedWebSocket.CLOSED = NativeWebSocket.CLOSED;
  window.WebSocket = GuardedWebSocket;
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
  'Waiting',
  'Waiting on Buyer',
  'On Hold',
  'Overdue',
  'Escalated',
  'Complete'
];

const KNOWN_STATUS_VALUES = new Set([
  'pending',
  'in progress',
  'waiting on buyer',
  'awaiting buyer',
  'escalated',
  'complete',
  'completed',
  'new',
  'waiting',
  'on hold',
  'hold',
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

const ROLLOVER_ACTIVE_STATUS_WHITELIST = new Set([
  'pending',
  'in progress',
  'waiting',
  'waiting on buyer',
  'on hold',
  'hold',
  'overdue',
  'escalated',
  'new',
  'open',
  'reopened'
]);

const SMART_OVERDUE_MAX_AGE_DAYS = 5;
const SMART_OVERDUE_MIN_CARRIES = 7;

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
    norm === 'master accounts' ||
    norm === 'yard list' ||
    norm === 'yard lists' ||
    norm === 'yards' ||
    norm === 'yard locations' ||
    norm === 'trackerconfig' ||
    norm === 'tracker config'
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

function isWaitingOrHoldStatus(status) {
  const norm = String(status || '').trim().toLowerCase();
  if (!norm) return false;
  return (
    norm === 'waiting' ||
    norm === 'waiting on buyer' ||
    norm === 'awaiting buyer' ||
    norm === 'on hold' ||
    norm === 'hold' ||
    /^(waiting|on\s*hold|awaiting)\b/i.test(norm)
  );
}

function isOverdueStatus(status) {
  const norm = String(status || '').trim().toLowerCase();
  if (!norm) return false;
  return norm === 'overdue' || /^overdue\b/i.test(norm);
}

function isNewStatus(status) {
  return String(status || '').replace(/\s+/g, ' ').trim().toLowerCase() === 'new';
}

function resolveInProgressStatusLabel() {
  if (typeof validationStatuses !== 'undefined' && Array.isArray(validationStatuses)) {
    const match = validationStatuses.find(s => /^in[\s\-_]*progress$/i.test(String(s || '').trim()));
    if (match) return String(match).trim();
  }
  return 'In Progress';
}

function isCompletedStatus(status) {
  const norm = String(status || '').trim().toLowerCase();
  if (!norm) return false;
  if (
    ROLLOVER_ACTIVE_STATUS_WHITELIST.has(norm) ||
    isWaitingOrHoldStatus(norm) ||
    isOverdueStatus(norm)
  ) {
    return false;
  }
  if (COMPLETED_STATUS_SET.has(norm)) return true;
  return (
    norm.startsWith('complete') ||
    norm.startsWith('closed') ||
    norm.startsWith('resolved')
  );
}

function isRolloverActiveStatus(status) {
  const norm = String(status || '').trim().toLowerCase();
  if (!norm) return true;
  if (
    ROLLOVER_ACTIVE_STATUS_WHITELIST.has(norm) ||
    isWaitingOrHoldStatus(norm) ||
    isOverdueStatus(norm)
  ) {
    return true;
  }
  return !isCompletedStatus(norm);
}

function getStatusBadgeClass(status) {
  if (isCompletedStatus(status)) return 'badge-complete';
  if (isOverdueStatus(status)) return 'badge-overdue';
  const norm = String(status || '').trim().toLowerCase();
  if (/escalat/i.test(norm)) return 'badge-escalated';
  if (isWaitingOrHoldStatus(norm)) return 'badge-waiting';
  if (/^(in\s*progress|new)$/i.test(norm)) return 'badge-progress';
  return 'badge-pending';
}

function evaluateRequestOverdueMetrics(entries, referenceTabOrTime = null, currentStatus = '') {
  const list = Array.isArray(entries) ? entries.filter(Boolean) : [];
  let firstTime = 0;
  let firstSheet = '';
  let latestTime = 0;
  let latestEntryStatus = '';
  const distinctSheets = new Set();

  list.forEach(entry => {
    if (!entry) return;
    if (entry.sheetName) {
      distinctSheets.add(entry.sheetName);
    }
    const t = entry.sheetName ? parseDateTab(entry.sheetName) : 0;
    if (t > 0) {
      if (!firstTime || t < firstTime) {
        firstTime = t;
        firstSheet = entry.sheetName;
      }
      if (t >= latestTime) {
        latestTime = t;
        if (entry.status && String(entry.status).trim() !== '') {
          latestEntryStatus = String(entry.status).trim();
        }
      }
    } else if (!latestTime && entry.status && String(entry.status).trim() !== '') {
      latestEntryStatus = String(entry.status).trim();
    }
  });

  if (!latestEntryStatus && list.length > 0) {
    const tail = list[list.length - 1];
    if (tail && tail.status) {
      latestEntryStatus = String(tail.status).trim();
    }
  }

  let refTime = 0;
  let refTabName = '';
  if (typeof referenceTabOrTime === 'number' && Number.isFinite(referenceTabOrTime)) {
    refTime = referenceTabOrTime;
  } else if (typeof referenceTabOrTime === 'string' && DATE_TAB_REGEX.test(referenceTabOrTime.trim())) {
    refTabName = referenceTabOrTime.trim();
    refTime = parseDateTab(refTabName);
  }
  if (!refTime) {
    refTime = latestTime;
  }
  if (refTabName && (!latestTime || refTime >= latestTime)) {
    distinctSheets.add(refTabName);
  }

  const ageDays = (firstTime > 0 && refTime > 0 && refTime >= firstTime)
    ? Math.round((refTime - firstTime) / (1000 * 60 * 60 * 24))
    : 0;
  const carryCount = Math.max(list.length, distinctSheets.size);

  const resolvedStatus = String(currentStatus || latestEntryStatus || '').trim();
  const isWaitingOrHold = isWaitingOrHoldStatus(resolvedStatus);
  const qualifiesForAutoOverdue = !isWaitingOrHold && (
    ageDays > SMART_OVERDUE_MAX_AGE_DAYS || carryCount >= SMART_OVERDUE_MIN_CARRIES
  );

  return {
    ageDays,
    carryCount,
    firstSheet,
    isWaitingOrHold,
    qualifiesForAutoOverdue
  };
}

function mergeRequiredOperationalStatuses(baseStatuses) {
  const list = Array.isArray(baseStatuses)
    ? baseStatuses.map(s => String(s || '').trim()).filter(Boolean)
    : [];
  const required = ['Waiting', 'On Hold', 'Overdue'];
  const result = [...list];

  required.forEach(reqStatus => {
    const exists = result.some(s => s.toLowerCase() === reqStatus.toLowerCase());
    if (!exists) {
      const completeIdx = result.findIndex(s => isCompletedStatus(s));
      if (completeIdx !== -1) {
        result.splice(completeIdx, 0, reqStatus);
      } else {
        result.push(reqStatus);
      }
    }
  });

  return result;
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
const HIDE_COMPLETED_STORAGE_KEY = 'trackerui_hide_completed_v1';
const HIDE_UUID_KEY_STORAGE_KEY = 'trackerui_hide_uuid_key_v1';
const THEME_STORAGE_KEY = 'trackerui_theme_mode'; // 'auto' | 'light' | 'dark'
const FONT_SCALE_STORAGE_KEY = 'trackerui_font_scale'; // '85' .. '125' (percent)
const FONT_SCALE_MIN = 85;
const FONT_SCALE_MAX = 125;
const FONT_SCALE_STEP = 5;

let hideCompletedRequests = localStorage.getItem(HIDE_COMPLETED_STORAGE_KEY) === 'true';
let hideUuidKey = localStorage.getItem(HIDE_UUID_KEY_STORAGE_KEY) === 'true';

function setHideCompletedPreference(enabled, { persist = true } = {}) {
  hideCompletedRequests = Boolean(enabled);
  if (persist) {
    try {
      localStorage.setItem(HIDE_COMPLETED_STORAGE_KEY, String(hideCompletedRequests));
    } catch (_) { /* ignore */ }
  }
  return hideCompletedRequests;
}

function setHideUuidKeyPreference(enabled, { persist = true } = {}) {
  hideUuidKey = Boolean(enabled);
  if (persist) {
    try {
      localStorage.setItem(HIDE_UUID_KEY_STORAGE_KEY, String(hideUuidKey));
    } catch (_) { /* ignore */ }
  }
  if (typeof document !== 'undefined' && document.body) {
    document.body.classList.toggle('hide-tracker-uuid-keys', hideUuidKey);
  }
  if (typeof syncHideUuidButtonUI === 'function') {
    syncHideUuidButtonUI();
  }
  return hideUuidKey;
}

if (typeof document !== 'undefined' && document.body && hideUuidKey) {
  document.body.classList.add('hide-tracker-uuid-keys');
}

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
const btnResetTableColWidths = document.getElementById('btnResetTableColWidths');
const btnExportConfigJson = document.getElementById('btnExportConfigJson');
const btnImportConfigJson = document.getElementById('btnImportConfigJson');
const configJsonFileInput = document.getElementById('configJsonFileInput');
const btnCheckForUpdates = document.getElementById('btnCheckForUpdates');
const appVersionTag = document.getElementById('appVersionTag');
const pwaUpdateBanner = document.getElementById('pwaUpdateBanner');
const btnPwaReloadNow = document.getElementById('btnPwaReloadNow');
const btnPwaDismissUpdate = document.getElementById('btnPwaDismissUpdate');
const multiTabLockBanner = document.getElementById('multiTabLockBanner');
const multiTabLockDesc = document.getElementById('multiTabLockDesc');
const btnTakeOverSession = document.getElementById('btnTakeOverSession');

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

// Sheet Navigation & Controls
const btnPrevSheet = document.getElementById('btnPrevSheet');
const btnNextSheet = document.getElementById('btnNextSheet');
const sheetSelect = document.getElementById('sheetSelect');
const sheetTabsBar = document.getElementById('sheetTabsBar');
const scopeSelect = document.getElementById('scopeSelect');
const btnToggleHideCompleted = document.getElementById('btnToggleHideCompleted');
const hideCompletedLabel = document.getElementById('hideCompletedLabel');
const btnToggleHideUuid = document.getElementById('btnToggleHideUuid');
const hideUuidLabel = document.getElementById('hideUuidLabel');
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
const btnBatchCopyRows = document.getElementById('btnBatchCopyRows');
const btnBatchCopyToTab = document.getElementById('btnBatchCopyToTab');
const batchStatusSep = document.getElementById('batchStatusSep');
const batchStatusControlGroup = document.getElementById('batchStatusControlGroup');
const batchStatusSelect = document.getElementById('batchStatusSelect');
const btnApplyBatchStatus = document.getElementById('btnApplyBatchStatus');
const btnBatchDeleteSelected = document.getElementById('btnBatchDeleteSelected');
const btnClearBatchSelection = document.getElementById('btnClearBatchSelection');
const pasteRowsToolbarGroup = document.getElementById('pasteRowsToolbarGroup');
const btnPasteTrackerRows = document.getElementById('btnPasteTrackerRows');
const pasteCopiedCountEl = document.getElementById('pasteCopiedCount');
const btnClearCopiedRows = document.getElementById('btnClearCopiedRows');

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
const yardListEl = document.getElementById('yardList');

// Bottom Dock Editor DOM Elements
const bottomEditorDock = document.getElementById('bottomEditorDock');
const dockHeaderToggleZone = document.getElementById('dockHeaderToggleZone');
const dockModeBadge = document.getElementById('dockModeBadge');
const dockTargetLabel = document.getElementById('dockTargetLabel');
const dockKeyBadge = document.getElementById('dockKeyBadge');
const dockDraftBadge = document.getElementById('dockDraftBadge');
const btnDockResumeDraft = document.getElementById('btnDockResumeDraft');
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

// Workbook Profile Bar, Workspace Sheet Tabs Bar & Editable Scratch Sheet DOM Elements
const trackerProfilesBar = document.getElementById('trackerProfilesBar');
const trackerProfilesList = document.getElementById('trackerProfilesList');
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
const chkScratchAutoIncrement = document.getElementById('chkScratchAutoIncrement');
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
const scratchCellFormatSelect = document.getElementById('scratchCellFormatSelect');
const scratchTextFormatGroup = document.getElementById('scratchTextFormatGroup');
const btnScratchFormatBold = document.getElementById('btnScratchFormatBold');
const btnScratchFormatItalic = document.getElementById('btnScratchFormatItalic');
const btnScratchFormatUnderline = document.getElementById('btnScratchFormatUnderline');
const btnScratchAlignLeft = document.getElementById('btnScratchAlignLeft');
const btnScratchAlignCenter = document.getElementById('btnScratchAlignCenter');
const btnScratchAlignRight = document.getElementById('btnScratchAlignRight');
const btnScratchClearTextFormat = document.getElementById('btnScratchClearTextFormat');
const btnScratchTextToColumns = document.getElementById('btnScratchTextToColumns');
const btnScratchAutoFilter = document.getElementById('btnScratchAutoFilter');
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
const scratchFormulaNotice = document.getElementById('scratchFormulaNotice');
const btnAckScratchFormulaNotice = document.getElementById('btnAckScratchFormulaNotice');
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
const chkTtcNewline = document.getElementById('chkTtcNewline');
const chkTtcOther = document.getElementById('chkTtcOther');
const ttcOtherInput = document.getElementById('ttcOtherInput');
const chkTtcTrim = document.getElementById('chkTtcTrim');
const chkTtcConsecutive = document.getElementById('chkTtcConsecutive');
const radTtcAcross = document.getElementById('radTtcAcross');
const radTtcDown = document.getElementById('radTtcDown');
const ttcOverwriteBanner = document.getElementById('ttcOverwriteBanner');
const ttcPreviewTitle = document.getElementById('ttcPreviewTitle');
const ttcPreviewBox = document.getElementById('ttcPreviewBox');
const btnCancelTextToColumns = document.getElementById('btnCancelTextToColumns');
const btnConfirmTextToColumns = document.getElementById('btnConfirmTextToColumns');
const scratchAutoFilterModal = document.getElementById('scratchAutoFilterModal');
const scratchAutoFilterForm = document.getElementById('scratchAutoFilterForm');
const afSheetBadge = document.getElementById('afSheetBadge');
const afPresetModeSelect = document.getElementById('afPresetModeSelect');
const afReturnStartCol = document.getElementById('afReturnStartCol');
const afReturnEndCol = document.getElementById('afReturnEndCol');
const afCriteriaCol = document.getElementById('afCriteriaCol');
const afDestStartCol = document.getElementById('afDestStartCol');
const afStartRowInput = document.getElementById('afStartRowInput');
const afMaxRowInput = document.getElementById('afMaxRowInput');
const afThreeBucketWrap = document.getElementById('afThreeBucketWrap');
const afYellowKeywordInput = document.getElementById('afYellowKeywordInput');
const afIncludeCriteriaInGreen = document.getElementById('afIncludeCriteriaInGreen');
const afSingleCustomWrap = document.getElementById('afSingleCustomWrap');
const afCustomOperatorSelect = document.getElementById('afCustomOperatorSelect');
const afCustomValueInput = document.getElementById('afCustomValueInput');
const afCustomColorSelect = document.getElementById('afCustomColorSelect');
const afFormulaPreviewBox = document.getElementById('afFormulaPreviewBox');
const btnCancelAutoFilter = document.getElementById('btnCancelAutoFilter');
const btnConfirmAutoFilter = document.getElementById('btnConfirmAutoFilter');

// Dynamic Full-Column Schema Manager Modal DOM Elements (v1.19.7)
const schemaMapperModal = document.getElementById('schemaMapperModal');
const schemaMapperSubtitle = document.getElementById('schemaMapperSubtitle');
const schemaMapperModeBadge = document.getElementById('schemaMapperModeBadge');
const schemaMapperForm = document.getElementById('schemaMapperForm');
const schemaModeSelect = document.getElementById('schemaModeSelect');
const schemaTrackingColInfo = document.getElementById('schemaTrackingColInfo');
const schemaRolesSummaryBar = document.getElementById('schemaRolesSummaryBar');
const schemaColSearchInput = document.getElementById('schemaColSearchInput');
const btnSchemaShowAllCols = document.getElementById('btnSchemaShowAllCols');
const chkSchemaAllTable = document.getElementById('chkSchemaAllTable');
const chkSchemaAllEditor = document.getElementById('chkSchemaAllEditor');
const schemaAllColumnsTableBody = document.getElementById('schemaAllColumnsTableBody');
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
    if (manualImageLightbox && manualImageLightbox.style.display !== 'none') {
      closeManualImageLightbox();
      return;
    }
    if (userManualModal && userManualModal.classList.contains('open')) {
      closeUserManualModal();
      return;
    }
    closeToolsMenu();
  }
});

// --- INTERACTIVE IN-APP USER MANUAL & REFERENCE GUIDE MODAL ---
const btnOpenUserManual = document.getElementById('btnOpenUserManual');
const menuOpenUserManual = document.getElementById('menuOpenUserManual');
const userManualModal = document.getElementById('userManualModal');
const btnCloseUserManual = document.getElementById('btnCloseUserManual');
const btnCloseUserManualHeader = document.getElementById('btnCloseUserManualHeader');
const userManualSearchInput = document.getElementById('userManualSearchInput');
const btnClearManualSearch = document.getElementById('btnClearManualSearch');
const userManualTocList = document.getElementById('userManualTocList');
const userManualContentPane = document.getElementById('userManualContentPane');
const userManualEmptySearch = document.getElementById('userManualEmptySearch');
const manualImageLightbox = document.getElementById('manualImageLightbox');
const manualLightboxImg = document.getElementById('manualLightboxImg');
const manualLightboxCaption = document.getElementById('manualLightboxCaption');
const btnCloseManualLightbox = document.getElementById('btnCloseManualLightbox');

function openUserManualModal(targetSectionId = '') {
  closeToolsMenu();
  if (!userManualModal) return;
  userManualModal.classList.add('open');
  if (targetSectionId) {
    scrollToManualSection(targetSectionId);
  }
  if (userManualSearchInput) {
    setTimeout(() => userManualSearchInput.focus(), 40);
  }
}

function closeUserManualModal() {
  closeManualImageLightbox();
  if (userManualModal) {
    userManualModal.classList.remove('open');
  }
}

function openManualImageLightbox(src, captionText) {
  if (!manualImageLightbox || !manualLightboxImg) return;
  manualLightboxImg.src = src || '';
  if (manualLightboxCaption) {
    manualLightboxCaption.textContent = captionText || 'Screenshot Preview';
  }
  manualImageLightbox.style.display = 'flex';
}

function closeManualImageLightbox() {
  if (!manualImageLightbox) return;
  manualImageLightbox.style.display = 'none';
}

function setActiveManualTocItem(sectionId) {
  if (!userManualTocList) return;
  const buttons = userManualTocList.querySelectorAll('.manual-toc-item');
  buttons.forEach(btn => {
    btn.classList.toggle('active', btn.dataset.manualTarget === sectionId);
  });
}

function scrollToManualSection(sectionId) {
  if (!userManualContentPane || !sectionId) return;
  const targetEl = document.getElementById(sectionId);
  if (!targetEl) return;
  if (targetEl.style.display === 'none') {
    filterUserManualSections('');
  }
  setActiveManualTocItem(sectionId);
  if (typeof targetEl.scrollIntoView === 'function') {
    targetEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
  } else {
    userManualContentPane.scrollTop = targetEl.offsetTop;
  }
}

function filterUserManualSections(rawQuery) {
  if (!userManualContentPane) return;
  const q = String(rawQuery || '').trim().toLowerCase();
  if (userManualSearchInput && userManualSearchInput.value !== rawQuery && rawQuery === '') {
    userManualSearchInput.value = '';
  }
  if (btnClearManualSearch) {
    btnClearManualSearch.style.display = q ? 'inline-flex' : 'none';
  }

  const sections = userManualContentPane.querySelectorAll('.manual-section');
  let visibleCount = 0;
  let firstVisibleId = '';

  sections.forEach(sec => {
    const text = String(sec.textContent || '').toLowerCase();
    const matches = !q || text.includes(q);
    sec.style.display = matches ? '' : 'none';
    if (matches) {
      visibleCount++;
      if (!firstVisibleId) firstVisibleId = sec.id;
    }
    if (userManualTocList && sec.id) {
      const tocBtn = userManualTocList.querySelector(`.manual-toc-item[data-manual-target="${sec.id}"]`);
      if (tocBtn) {
        tocBtn.style.display = matches ? '' : 'none';
      }
    }
  });

  if (userManualEmptySearch) {
    userManualEmptySearch.style.display = visibleCount === 0 ? 'block' : 'none';
  }
  if (firstVisibleId) {
    setActiveManualTocItem(firstVisibleId);
  }
}

if (btnOpenUserManual) {
  btnOpenUserManual.addEventListener('click', () => openUserManualModal());
}
if (menuOpenUserManual) {
  menuOpenUserManual.addEventListener('click', () => openUserManualModal());
}
if (btnCloseUserManual) {
  btnCloseUserManual.addEventListener('click', () => closeUserManualModal());
}
if (btnCloseUserManualHeader) {
  btnCloseUserManualHeader.addEventListener('click', () => closeUserManualModal());
}
if (userManualModal) {
  userManualModal.addEventListener('click', (e) => {
    if (e.target === userManualModal) closeUserManualModal();
  });
}
if (userManualTocList) {
  userManualTocList.addEventListener('click', (e) => {
    const btn = e.target.closest ? e.target.closest('.manual-toc-item') : null;
    if (btn && btn.dataset && btn.dataset.manualTarget) {
      scrollToManualSection(btn.dataset.manualTarget);
    }
  });
}
if (userManualSearchInput) {
  userManualSearchInput.addEventListener('input', () => {
    filterUserManualSections(userManualSearchInput.value);
  });
}
if (btnClearManualSearch) {
  btnClearManualSearch.addEventListener('click', () => {
    filterUserManualSections('');
    if (userManualSearchInput) userManualSearchInput.focus();
  });
}
if (userManualContentPane) {
  userManualContentPane.addEventListener('click', (e) => {
    const img = e.target.closest ? e.target.closest('.manual-screenshot') : null;
    if (img) {
      const fig = img.closest ? img.closest('.manual-figure') : null;
      const capEl = fig ? fig.querySelector('figcaption') : null;
      const caption = capEl ? capEl.textContent.replace(/\s*—\s*Click image to enlarge/i, '').trim() : (img.alt || 'Screenshot Preview');
      openManualImageLightbox(img.getAttribute('src') || img.src, caption);
    }
  });
  userManualContentPane.addEventListener('scroll', () => {
    if (!userManualTocList) return;
    const paneTop = userManualContentPane.getBoundingClientRect().top;
    const sections = userManualContentPane.querySelectorAll('.manual-section');
    let activeId = '';
    sections.forEach(sec => {
      if (sec.style.display === 'none') return;
      const rect = sec.getBoundingClientRect();
      if (rect.top - paneTop <= 90) {
        activeId = sec.id;
      }
    });
    if (!activeId && sections.length > 0) {
      const firstVis = Array.from(sections).find(s => s.style.display !== 'none');
      if (firstVis) activeId = firstVis.id;
    }
    if (activeId) setActiveManualTocItem(activeId);
  }, { passive: true });
}
if (btnCloseManualLightbox) {
  btnCloseManualLightbox.addEventListener('click', () => closeManualImageLightbox());
}
if (manualImageLightbox) {
  manualImageLightbox.addEventListener('click', (e) => {
    if (e.target === manualImageLightbox || e.target === manualLightboxImg) {
      closeManualImageLightbox();
    }
  });
}

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
  const clean = String(actionStr || '').replace(/\s+/g, ' ').trim();
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

function renderInlineCopyBtnHtml(rawValue, label = 'Value') {
  const trimmed = String(rawValue || '').trim();
  if (!trimmed || trimmed === '—') return '';
  return `<button type="button" class="btn-inline-copy" data-copy-val="${escapeHtml(trimmed)}" data-copy-label="${escapeHtml(label)}" title="Copy ${escapeHtml(label)} (${escapeHtml(trimmed)}) to clipboard" aria-label="Copy ${escapeHtml(label)} ${escapeHtml(trimmed)}"><svg class="copy-icon-svg" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="5.5" y="5.5" width="8" height="8.5" rx="1.3"></rect><path d="M3.5 10.5H3a1 1 0 0 1-1-1v-6a1 1 0 0 1 1-1h5.5a1 1 0 0 1 1 1v.5"></path></svg></button>`;
}

async function copyTextToClipboard(rawText, label = 'Value', btnEl = null) {
  const text = String(rawText || '').trim();
  if (!text) return false;

  let copied = false;
  if (typeof navigator !== 'undefined' && navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
    try {
      await navigator.clipboard.writeText(text);
      copied = true;
    } catch (_) {
      copied = false;
    }
  }

  if (!copied && typeof document !== 'undefined') {
    try {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.setAttribute('readonly', '');
      ta.style.position = 'fixed';
      ta.style.top = '-9999px';
      ta.style.left = '-9999px';
      document.body.appendChild(ta);
      ta.select();
      copied = document.execCommand('copy');
      ta.remove();
    } catch (_) {
      copied = false;
    }
  }

  if (!copied) {
    showToast(`Could not copy ${label} to clipboard.`, 'warning');
    return false;
  }

  if (btnEl) {
    btnEl.classList.add('copied');
    btnEl.dataset.copied = 'true';
    if (btnEl._copyTimer) clearTimeout(btnEl._copyTimer);
    btnEl._copyTimer = setTimeout(() => {
      btnEl.classList.remove('copied');
      delete btnEl.dataset.copied;
    }, 1200);
  }

  showToast(`📋 Copied ${label}: ${text}`, 'info');
  return true;
}

if (typeof document !== 'undefined') {
  document.addEventListener('click', (e) => {
    const copyBtn = e.target && e.target.closest ? e.target.closest('.btn-inline-copy') : null;
    if (!copyBtn) return;
    if (typeof listBody !== 'undefined' && listBody && listBody.contains(copyBtn)) return;
    e.stopPropagation();
    e.preventDefault();
    copyTextToClipboard(copyBtn.dataset.copyVal, copyBtn.dataset.copyLabel || 'Value', copyBtn);
  });
}

// --- 9. APP VERSION, MULTI-TAB INSTANCE LOCKING, CHROME CHECK & PWA SERVICE WORKER ---
const APP_VERSION = 'v1.19.10';

function syncAppVersionDOM() {
  if (typeof document === 'undefined') return;
  document.title = `TrackerUI ${APP_VERSION}`;
  if (appVersionTag) {
    appVersionTag.textContent = APP_VERSION;
    appVersionTag.title = `TrackerUI ${APP_VERSION} — Click to check for updates & refresh app cache`;
  }
  if (typeof document.querySelectorAll === 'function') {
    document.querySelectorAll('[data-app-version]').forEach((el) => {
      if (el) el.textContent = APP_VERSION;
    });
  }
}
syncAppVersionDOM();

// Cross-Tab Session Lock & Concurrent Save Protection
const TAB_INSTANCE_ID = `tab_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;
const INSTANCE_WEB_LOCK_NAME = 'trackerui_active_session_lock_v1';
const INSTANCE_CHANNEL_NAME = 'trackerui_instance_sync_v1';
const INSTANCE_STORAGE_HEARTBEAT_KEY = 'trackerui_instance_heartbeat_v1';
const INSTANCE_STORAGE_TAKEOVER_KEY = 'trackerui_instance_takeover_v1';
const INSTANCE_HEARTBEAT_INTERVAL_MS = 4000;
const INSTANCE_HEARTBEAT_STALE_MS = 10000;

let isPrimaryInstanceLockHolder = false;
let isLockedByOtherTab = false;
let activeWebLockReleaser = null;
let instanceLockWaitingQueued = false;
let instanceHeartbeatTimer = null;
let instanceBroadcastChannel = null;
let pendingTakeoverTargetId = null;
let pendingTakeoverAt = 0;
let lastLoadedDiskModifiedMs = 0;
let saveWorkbookMutexPromise = Promise.resolve();

function broadcastInstanceMessage(payload) {
  const msg = { ...payload, senderTabId: TAB_INSTANCE_ID, sentAt: Date.now() };
  if (instanceBroadcastChannel) {
    try {
      instanceBroadcastChannel.postMessage(msg);
    } catch (_) { /* ignore */ }
  }
  return msg;
}

function writeInstanceHeartbeat() {
  if (!isPrimaryInstanceLockHolder) return;
  try {
    localStorage.setItem(
      INSTANCE_STORAGE_HEARTBEAT_KEY,
      JSON.stringify({ tabId: TAB_INSTANCE_ID, ts: Date.now() })
    );
  } catch (_) { /* ignore */ }
}

function clearOwnInstanceHeartbeat() {
  try {
    const raw = localStorage.getItem(INSTANCE_STORAGE_HEARTBEAT_KEY);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    if (parsed && parsed.tabId === TAB_INSTANCE_ID) {
      localStorage.removeItem(INSTANCE_STORAGE_HEARTBEAT_KEY);
    }
  } catch (_) { /* ignore */ }
}

function startInstanceHeartbeat() {
  stopInstanceHeartbeat();
  writeInstanceHeartbeat();
  instanceHeartbeatTimer = setInterval(() => {
    if (isPrimaryInstanceLockHolder) {
      writeInstanceHeartbeat();
    } else {
      stopInstanceHeartbeat();
    }
  }, INSTANCE_HEARTBEAT_INTERVAL_MS);
}

function stopInstanceHeartbeat() {
  if (instanceHeartbeatTimer) {
    clearInterval(instanceHeartbeatTimer);
    instanceHeartbeatTimer = null;
  }
}

function releaseActiveWebLock() {
  stopInstanceHeartbeat();
  clearOwnInstanceHeartbeat();
  if (typeof activeWebLockReleaser === 'function') {
    const rel = activeWebLockReleaser;
    activeWebLockReleaser = null;
    try {
      rel();
    } catch (_) { /* ignore */ }
  }
}

function enterMultiTabStandbyLockMode(reasonText = '') {
  isPrimaryInstanceLockHolder = false;
  isLockedByOtherTab = true;
  stopInstanceHeartbeat();

  if (typeof scratchAutoSaveTimer !== 'undefined' && scratchAutoSaveTimer) {
    clearTimeout(scratchAutoSaveTimer);
    scratchAutoSaveTimer = null;
  }
  if (typeof mirrorCountdownInterval !== 'undefined' && mirrorCountdownInterval) {
    clearInterval(mirrorCountdownInterval);
    mirrorCountdownInterval = null;
  }

  if (multiTabLockBanner) {
    multiTabLockBanner.style.display = 'flex';
  }
  if (multiTabLockDesc) {
    multiTabLockDesc.textContent =
      reasonText ||
      'TrackerUI is already open in another browser tab or window. Background saves and disk writes are paused here to prevent Excel file race conditions.';
  }
  if (statusPill) {
    setStatus('🔒 Standby (Open in Another Tab)', 'locked');
  }
}

async function exitMultiTabStandbyLockMode({ reloadWorkbook = false } = {}) {
  const wasLocked = isLockedByOtherTab;
  isLockedByOtherTab = false;
  isPrimaryInstanceLockHolder = true;
  startInstanceHeartbeat();

  if (multiTabLockBanner) {
    multiTabLockBanner.style.display = 'none';
  }

  if (
    reloadWorkbook &&
    wasLocked &&
    typeof loadWorkbookFromDisk === 'function' &&
    typeof workingFileHandle !== 'undefined' &&
    workingFileHandle
  ) {
    try {
      const targetTab = (typeof sheetSelect !== 'undefined' && sheetSelect) ? sheetSelect.value : null;
      await loadWorkbookFromDisk(targetTab, {
        skipFlushPendingScratch: true,
        skipStagedRecoveryCheck: true
      });
      showToast('⚡ Active session acquired — synced latest workbook from disk.', 'success');
    } catch (err) {
      console.warn('Could not reload workbook after lock acquisition:', err);
      setStatus(`Synced: ${workingFileHandle.name}`, 'ready');
    }
  } else if (wasLocked) {
    if (typeof workingFileHandle !== 'undefined' && workingFileHandle) {
      setStatus(`Synced: ${workingFileHandle.name}`, 'ready');
    } else if (typeof autoReconnectGesturePending !== 'undefined' && autoReconnectGesturePending && typeof cachedDirHandle !== 'undefined' && cachedDirHandle) {
      setStatus('Click anywhere to auto-reconnect folder', 'warning');
    } else if (typeof parentDirHandle !== 'undefined' && !parentDirHandle) {
      setStatus('No Folder Linked', '');
    }
  }
}

function queueWebLockWaiter() {
  if (instanceLockWaitingQueued) return;
  if (typeof navigator === 'undefined' || !navigator.locks || typeof navigator.locks.request !== 'function') {
    return;
  }
  instanceLockWaitingQueued = true;

  navigator.locks.request(INSTANCE_WEB_LOCK_NAME, async () => {
    instanceLockWaitingQueued = false;

    if (
      pendingTakeoverTargetId &&
      pendingTakeoverTargetId !== TAB_INSTANCE_ID &&
      Date.now() - pendingTakeoverAt < 3000
    ) {
      // Another tab explicitly initiated takeover; yield immediately and re-queue
      setTimeout(() => queueWebLockWaiter(), 120);
      return;
    }

    await new Promise((resolve) => {
      activeWebLockReleaser = resolve;
      exitMultiTabStandbyLockMode({ reloadWorkbook: true });
      broadcastInstanceMessage({ type: 'LOCK_CLAIMED' });
    });
  }).catch((err) => {
    instanceLockWaitingQueued = false;
    console.warn('Web Lock waiter error:', err);
  });
}

function tryAcquireWebLockImmediately() {
  if (typeof navigator === 'undefined' || !navigator.locks || typeof navigator.locks.request !== 'function') {
    return false;
  }
  navigator.locks.request(INSTANCE_WEB_LOCK_NAME, { ifAvailable: true }, async (lock) => {
    if (!lock) {
      enterMultiTabStandbyLockMode(
        'TrackerUI is already active in another tab or window. Background saves are paused here to prevent Excel race conditions.'
      );
      queueWebLockWaiter();
      return;
    }

    await new Promise((resolve) => {
      activeWebLockReleaser = resolve;
      exitMultiTabStandbyLockMode({ reloadWorkbook: false });
      broadcastInstanceMessage({ type: 'LOCK_CLAIMED' });
    });
  }).catch((err) => {
    console.warn('Initial Web Lock request error:', err);
  });
  return true;
}

async function takeOverActiveSessionInThisTab() {
  if (btnTakeOverSession) btnTakeOverSession.disabled = true;
  pendingTakeoverTargetId = TAB_INSTANCE_ID;
  pendingTakeoverAt = Date.now();

  try {
    broadcastInstanceMessage({
      type: 'FORCE_TAKEOVER',
      targetTabId: TAB_INSTANCE_ID
    });
    try {
      localStorage.setItem(
        INSTANCE_STORAGE_TAKEOVER_KEY,
        JSON.stringify({ targetTabId: TAB_INSTANCE_ID, ts: Date.now() })
      );
    } catch (_) { /* ignore */ }

    const hasWebLocks = typeof navigator !== 'undefined' && navigator.locks && typeof navigator.locks.request === 'function';
    if (hasWebLocks) {
      queueWebLockWaiter();
      // Wait up to 1.4s for the previous holder to release its Web Lock cleanly
      const start = Date.now();
      while (isLockedByOtherTab && Date.now() - start < 1400) {
        await new Promise((r) => setTimeout(r, 80));
      }
      if (isLockedByOtherTab) {
        // Fallback promotion if previous tab was unresponsive or closed without releasing
        await exitMultiTabStandbyLockMode({ reloadWorkbook: true });
      }
    } else {
      await exitMultiTabStandbyLockMode({ reloadWorkbook: true });
    }
  } finally {
    if (btnTakeOverSession) btnTakeOverSession.disabled = false;
  }
}

async function ensureActiveTabLockBeforeWrite(actionLabel = 'save changes') {
  if (!isLockedByOtherTab) return true;

  const confirmTakeOver = await showModalConfirm({
    title: '🔒 TrackerUI Active in Another Tab',
    badge: 'Standby Mode',
    message:
      `Another TrackerUI tab or window currently holds the active workspace lock.\n\n` +
      `To prevent race conditions and accidental overwrites before you ${actionLabel}, take over the active session here and reload the latest file from disk?`,
    confirmText: '⚡ Take Over & Reload',
    cancelText: 'Keep in Standby',
    variant: 'danger'
  });

  if (confirmTakeOver) {
    await takeOverActiveSessionInThisTab();
  }
  // Always return false so stale standby data in memory is never written over disk without fresh reload
  return false;
}

function handleRemoteInstanceMessage(msg) {
  if (!msg || !msg.type || msg.senderTabId === TAB_INSTANCE_ID) return;

  if (msg.type === 'FORCE_TAKEOVER') {
    pendingTakeoverTargetId = msg.targetTabId || null;
    pendingTakeoverAt = Date.now();
    if (isPrimaryInstanceLockHolder || activeWebLockReleaser) {
      releaseActiveWebLock();
      enterMultiTabStandbyLockMode('Control was transferred to another TrackerUI tab. Click "Take Over Active Session Here" to resume editing in this tab.');
      queueWebLockWaiter();
      showToast('🔒 Active session moved to another TrackerUI tab.', 'warning');
    }
  } else if (msg.type === 'WORKBOOK_SAVED' && isLockedByOtherTab) {
    if (multiTabLockDesc && msg.fileName) {
      const savedTime = msg.lastModified ? new Date(msg.lastModified).toLocaleTimeString() : 'just now';
      multiTabLockDesc.textContent = `"${msg.fileName}" was updated in the active tab at ${savedTime}. Take over here to reload fresh data before editing.`;
    }
  }
}

function initMultiTabInstanceLock() {
  if (typeof BroadcastChannel === 'function') {
    try {
      instanceBroadcastChannel = new BroadcastChannel(INSTANCE_CHANNEL_NAME);
      instanceBroadcastChannel.onmessage = (evt) => handleRemoteInstanceMessage(evt.data);
    } catch (_) {
      instanceBroadcastChannel = null;
    }
  }

  if (typeof window !== 'undefined' && typeof window.addEventListener === 'function') {
    window.addEventListener('storage', (evt) => {
      if (!evt || !evt.newValue) return;
      if (evt.key === INSTANCE_STORAGE_TAKEOVER_KEY) {
        try {
          const data = JSON.parse(evt.newValue);
          if (data && data.targetTabId && data.targetTabId !== TAB_INSTANCE_ID) {
            handleRemoteInstanceMessage({
              type: 'FORCE_TAKEOVER',
              senderTabId: data.targetTabId,
              targetTabId: data.targetTabId
            });
          }
        } catch (_) { /* ignore */ }
      }
    });

    const hasUnsavedInFlightWork = () => {
      if (window.__trackerDiskWriteInFlight) return true;
      if (typeof scratchIsSaving !== 'undefined' && scratchIsSaving) return true;
      if (typeof scratchInlineEditState !== 'undefined' && scratchInlineEditState) return true;
      if (typeof hasAnyPendingScratchMutations === 'function' && hasAnyPendingScratchMutations()) return true;
      return false;
    };

    window.addEventListener('beforeunload', (evt) => {
      if (hasUnsavedInFlightWork()) {
        evt.preventDefault();
        evt.returnValue = '';
      }
    });

    const cleanupOnClose = () => {
      releaseActiveWebLock();
      if (instanceBroadcastChannel) {
        try {
          instanceBroadcastChannel.close();
        } catch (_) { /* ignore */ }
      }
    };
    window.addEventListener('pagehide', cleanupOnClose);
  }

  const usedWebLocks = tryAcquireWebLockImmediately();
  if (!usedWebLocks) {
    // Fallback for environments without navigator.locks (e.g., file:// protocol)
    let hasFreshOtherTab = false;
    try {
      const rawHb = localStorage.getItem(INSTANCE_STORAGE_HEARTBEAT_KEY);
      if (rawHb) {
        const parsed = JSON.parse(rawHb);
        if (parsed && parsed.tabId !== TAB_INSTANCE_ID && Date.now() - Number(parsed.ts || 0) < INSTANCE_HEARTBEAT_STALE_MS) {
          hasFreshOtherTab = true;
        }
      }
    } catch (_) { /* ignore */ }

    if (hasFreshOtherTab) {
      enterMultiTabStandbyLockMode();
    } else {
      exitMultiTabStandbyLockMode({ reloadWorkbook: false });
    }
  }
}

if (btnTakeOverSession) {
  btnTakeOverSession.addEventListener('click', () => {
    takeOverActiveSessionInThisTab();
  });
}

initMultiTabInstanceLock();

function evaluateDesktopChromeCompatibility(navObj = (typeof navigator !== 'undefined' ? navigator : null)) {
  if (!navObj) {
    return { isSupported: true, isGoogleChrome: true, isMobileMode: false, detectedLabel: 'Unknown' };
  }
  const ua = String(navObj.userAgent || '');
  const uaData = navObj.userAgentData || null;

  // 1. Mobile / Tablet Mode Check (via User-Agent Client Hints + UA string; avoids false-positives on Windows touch laptops)
  let isMobileMode = false;
  if (uaData && typeof uaData.mobile === 'boolean') {
    isMobileMode = uaData.mobile;
  }
  if (/Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini|\bMobile\b/i.test(ua)) {
    isMobileMode = true;
  } else if (navObj.platform === 'MacIntel' && Number(navObj.maxTouchPoints) > 1 && !/\bChrome\//i.test(ua)) {
    isMobileMode = true;
  }

  // 2. Browser Type Check (Google Chrome only; excludes Edge, Opera, Brave, Vivaldi, Firefox, Safari, etc.)
  const isBrave = Boolean(navObj.brave);
  let isGoogleChrome = false;
  let browserName = 'Non-Chrome Browser';

  if (isBrave) {
    browserName = 'Brave';
  } else if (/\bEdg(?:e|A|iOS)?\//i.test(ua)) {
    browserName = 'Microsoft Edge';
  } else if (/\b(?:OPR|Opera)\b/i.test(ua)) {
    browserName = 'Opera';
  } else if (/\bVivaldi\b/i.test(ua)) {
    browserName = 'Vivaldi';
  } else if (/\bSamsungBrowser\b/i.test(ua)) {
    browserName = 'Samsung Internet';
  } else if (/\b(?:Firefox|FxiOS)\//i.test(ua)) {
    browserName = 'Mozilla Firefox';
  } else if (/\bCriOS\//i.test(ua)) {
    browserName = 'Chrome on iOS';
  } else if (uaData && Array.isArray(uaData.brands) && uaData.brands.length > 0) {
    const brands = uaData.brands.map(b => String((b && b.brand) || ''));
    const hasGoogleChromeBrand = brands.some(b => /Google Chrome/i.test(b));
    const hasOtherBrand = brands.some(b => /Microsoft Edge|Opera|Brave|Vivaldi|Yandex/i.test(b));
    if (hasGoogleChromeBrand && !hasOtherBrand) {
      isGoogleChrome = true;
      browserName = 'Google Chrome';
    } else {
      browserName = brands.find(b => !/Not.?A.?Brand|Chromium/i.test(b)) || brands[0] || 'Chromium Browser';
    }
  } else if (/\bChrome\/\d+/i.test(ua)) {
    isGoogleChrome = true;
    browserName = 'Google Chrome';
  } else if (/\bSafari\/\d+/i.test(ua)) {
    browserName = 'Apple Safari';
  }

  const modeLabel = isMobileMode ? 'Mobile Mode' : 'Desktop Mode';
  const detectedLabel = `${browserName} (${modeLabel})`;
  return {
    isSupported: isGoogleChrome && !isMobileMode,
    isGoogleChrome,
    isMobileMode,
    detectedLabel
  };
}

function checkDesktopChromeEnvironmentAndAlert(navObj = (typeof navigator !== 'undefined' ? navigator : null)) {
  const result = evaluateDesktopChromeCompatibility(navObj);
  if (!result.isSupported && typeof window !== 'undefined' && typeof window.alert === 'function') {
    window.alert(
      `⚠️ Browser Compatibility Notice\n\n` +
      `This application was developed specifically for Google Chrome in Desktop Mode.\n` +
      `Detected browser: ${result.detectedLabel}\n\n` +
      `Direct local folder read/write (File System Access API) and desktop workspace features may not work properly in other browsers or mobile mode. Click OK to acknowledge.`
    );
  }
  return result;
}

if (typeof navigator !== 'undefined' && (navigator.userAgent || navigator.userAgentData)) {
  checkDesktopChromeEnvironmentAndAlert(navigator);
}
let activeSwRegistration = null;
let lastSwUpdateCheckAt = 0;
let pwaUpdatePending = false;
const SW_UPDATE_CHECK_THROTTLE_MS = 45000;

function showPwaUpdateBanner() {
  pwaUpdatePending = true;
  if (pwaUpdateBanner) {
    pwaUpdateBanner.style.display = 'flex';
  }
  if (appVersionTag) {
    appVersionTag.classList.add('update-available');
    appVersionTag.textContent = `${APP_VERSION} • Update Ready`;
    appVersionTag.title = `A newer version of TrackerUI is available — click to refresh and apply the update`;
  }
  if (btnToolsMenu) {
    btnToolsMenu.classList.add('has-update-indicator');
  }
  if (btnCheckForUpdates) {
    btnCheckForUpdates.classList.add('menu-item-update-ready');
  }
}

async function forceClearAppCacheAndReload() {
  closeToolsMenu();
  showToast('Refreshing application cache and loading latest version...', 'info');
  try {
    if (typeof commitActiveScratchInlineEdit === 'function') {
      commitActiveScratchInlineEdit();
    }
    if (
      typeof hasAnyPendingScratchMutations === 'function' &&
      hasAnyPendingScratchMutations() &&
      typeof flushPendingScratchEditsToDisk === 'function'
    ) {
      await flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: true });
    }
    if (typeof saveWorkbookMutexPromise !== 'undefined' && saveWorkbookMutexPromise) {
      await saveWorkbookMutexPromise.catch(() => {});
    }
  } catch (flushErr) {
    console.warn('Could not flush pending edits before cache reload:', flushErr);
  }
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
