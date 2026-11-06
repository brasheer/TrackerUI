# TrackerUI — Changelog & Feature Tracker

This file tracks all requested features, architectural suggestions (from [`newfeatures.txt`](newfeatures.txt), [`--Project--- TrackerUI.md`](--Project---%20TrackerUI.md), and user sessions), implementation statuses, active bug fixes, and version release history.

> **Automated Tracking Rule**: Whenever a feature or suggestion is implemented, modified, or decided against, this file is automatically updated by the coding assistant per [`AGENTS.md`](AGENTS.md).

---

## 1. Feature & Suggestion Status Matrix (`newfeatures.txt` & `--Project--- TrackerUI.md`)

### ✅ Implemented Features
| Feature / Suggestion | Source | Status | Implemented In | Implementation Details |
|---|---|---|---|---|
| **Click-to-Filter ("Like-Item List") on Master #, Bidder #, Lot #, Yard #, and all core fields** | `newfeatures.txt` (L3–6) | ✅ Implemented | `v1.8.0` (`6c81c77`) | Clicking `Master #`, `Bidder #`, individual `Lot #` tokens, `Yard #`, `Subject`, `Ticket #`, `Broker`, `Action`, `Status`, or `Tracking ID` filters across all daily tabs and shows the Like-Item banner (`#filterBanner`). |
| **Click Date in Table to Open Date Tab** | `newfeatures.txt` (L5) | ✅ Implemented | `v1.8.0` (`6c81c77`) | Clicking the `📅 MM-DD-YYYY` button (`.clickable-date-tab`) on any row switches `#sheetSelect` to that daily tab, resets view scope to `Current Day Tab Only`, and clears active filters. |
| **Sortable Column Headers + Clear Sort/Filter Option** | `newfeatures.txt` (L8) | ✅ Implemented | `v1.8.0` (`6c81c77`) | All table headers (including separate `Master`/`Bidder` and `Yard`/`Lots` sub-headers) support 3-state sorting (`▲ Asc` → `▼ Desc` → `↕ Default`). Added `✕ Clear Sort / Filter` button (`#btnClearSortFilter`) in the controls bar and banner. |
| **Natural Filter Auto-Clearing on List View Navigation** | `newfeatures.txt` (L12) + User Request | ✅ Implemented | `v1.11.0` | Clicking any Day Tab pill in `#sheetTabsBar`, `◀`/`▶` steppers, `#sheetSelect`, `#scopeSelect`, `📅 Date Tab` cell, or switching workbooks naturally clears active Like-Item/Tracking ID filters and search queries (`clearTransientFiltersOnNavigation`). Typing in `#searchInput` and clicking Like-Item chips also mutually clear each other. |
| **Docked Bottom Request Editor Panel (Pinned, Collapsible & Excluding UUID from Editable Inputs)** | `newfeatures.txt` (L11–13) + User Request | ✅ Implemented | `v1.11.0` / `v1.11.2` | Added persistent `#bottomEditorDock` pinned outside the scrollable `.table-wrap` pane (`100dvh` flex layout with sticky table column headers) and made it collapsible (`#btnDockToggleCollapse`, persisted in `localStorage`). Clicking any request row (or `Edit`) auto-populates and expands the dock for immediate inline editing without exposing the immutable UUIDv4 as an editable field. Includes `🕒 +Timestamp` button, `＋ New / Clear` mode, and `🗔 Popup Editor` fallback. |
| **Update "Create Monthly Workbook" (Exclude Master Account List & Support Templates Folder)** | `newfeatures.txt` (L15, L57) | ✅ Implemented | `v1.11.0` | Updated `bootstrapWorkbookSchemaInMemory()` and `promptCreateNewWorkbook()` so newly created monthly workbooks exclude `Master Accounts` as a required/created tab (initializing only `template`, `validation lists`, and today's `MM-DD-YYYY` sheet, or cloning directly from a template in the configured Templates Folder). Bidirectional Broker ↔ Master # auto-fill learns from existing workbook rows as well as optional `Master Accounts` sheets. |
| **Auto-Reconnect Folders on Startup, Disable Auto-Reconnect Checkbox & Persistent Full File Paths Bar** | `newfeatures.txt` (L53) | ✅ Implemented | `v1.11.0` | Added automatic folder & workbook reconnection on startup when browser permission is already granted, a persistent checkbox (`#chkAutoReconnect` / `#chkAutoReconnectBar`) to disable auto-reconnect, and `#filePathsBar` displaying the full relative paths (`<AgentFolder>\<Workbook.xlsx>`, `<SupervisorFolder>\<Workbook (VIEW ONLY).xlsx>`, and `<TemplatesFolder>`) until explicitly cleared/updated. |
| **Separate Directory Pickers for Templates, Current Workbooks & Supervisor View-Only Trackers** | `newfeatures.txt` (L57) | ✅ Implemented | `v1.11.0` | Added 3 distinct folder pickers persisted in `IndexedDB`: **Tracker Workbooks Folder (Agent)**, **Supervisor `(VIEW ONLY)` Folder**, and **Templates Folder** (`templatesFolderHandle`). |
| **`👁️ Open View-Only Tracker` Button & Read-Only ArrayBuffer Inspector** | `newfeatures.txt` (L42, L51) + User Request | ✅ Implemented | `v1.11.0` / `v1.12.0` | Added `#menuOpenViewOnly` inside the **Tools & Setup** menu (moved out of the primary header in `v1.12.0` to keep the top bar clean) to inspect `<Workbook Name> (VIEW ONLY).xlsx` in-app via a non-locking read-only `ArrayBuffer` snapshot (`#viewOnlyModal`), complete with sheet selector, search filter, and one-click `.xlsx` copy download. |
| **Ticket Details & Aggregated Timestamped Notes Modal + Creation Date / Age Tracking** | `newfeatures.txt` (L33–35, L59) | ✅ Implemented | `v1.11.0` | Added `Details` button on every row and in the Bottom Editor Dock opening `#ticketDetailsModal`, which aggregates all daily tab entries for a ticket chronologically, displays original creation tab & day span, highlights `[MM/DD/YYYY HH:MM]` timestamps, and provides an inline **Append Timestamped Note** form. |
| **Light Mode Color Palette & `Light` / `Auto` / `Dark` Theme Switcher** | `newfeatures.txt` (L39) + User Request | ✅ Implemented | `v1.12.0` | Added complete semantic Light Mode and Dark Mode token palettes in `css/styles.css`, automatic OS theme detection via `matchMedia('(prefers-color-scheme: light)')`, and synced `Light` / `Auto` / `Dark` segmented controls in both the top header (`#headerThemeGroup`) and **Tools & Setup** menu (`#menuThemeGroup`), persisted in `localStorage` (`trackerui_theme_mode`). |
| **Whole-App Text Size Slider & Professional UI / Menu Cleanup** | User Request | ✅ Implemented | `v1.12.0` | Converted all typography across `css/styles.css` to scalable `rem` units driven by `html { font-size: calc(16px * var(--font-scale, 1)); }` and added an interactive **Text Size** slider (`85%`–`125%`, `A−`/`A+` steppers, `Reset`) inside **Tools & Setup**, persisted in `localStorage` (`trackerui_font_scale`). Replaced header/menu emoji clutter with crisp inline SVG icons and refined control styling. |
| **Consolidated Unique-Request Rollover with Chronological Action Notes Appending** | User Request | ✅ Implemented | `v1.12.0` | Upgraded `indexWorksheetRows()` (chronological oldest → newest indexing + prior-day Tracking ID inheritance) and `computeRolloverAnalysis()` (`areRowsSameUniqueRequest` + `mergeChronologicalActionNotes`) so each unique request across prior days is rolled forward **once** (never duplicating per day) with all prior days' `Action Notes` chronologically appended and deduplicated (`[MM-DD-YYYY] ...`), while unifying hidden Column K/L `Tracking ID`s across consolidated rows. |
| **Batch Status Update Architecture (Permanent Checkboxes & Floating Batch Action Dock)** | `--Project--- TrackerUI.md` (§1, L4–12) | ✅ Implemented | `v1.13.0` | Made the left-edge row selection checkbox column (`#thDeleteSelect` / `.chk-delete-row`) permanent on the main table and added a floating `#batchActionDock` (status dropdown + `Apply` + `Delete` + `Clear`) that renders only when ≥1 rows are ticked. `handleApplyBatchStatusUpdate()` updates `allRows` optimistically, re-renders the DOM immediately, strictly preserves immutable Column K/L UUIDv4s, and executes a single `saveWorkbookToDisk()` dual-write. |
| **Auto-Fill Graceful Degradation & Mock Testing Data (`mockMasterAccounts.js`)** | `--Project--- TrackerUI.md` (§2, L14–18) | ✅ Implemented | `v1.13.0` | Guarded `setBrokerCellWithFormula()` against writing `'Master Accounts'` `_xlfn.XLOOKUP` formulas when that sheet is absent; created standalone `js/mockMasterAccounts.js` dummy dataset (10 sample accounts including `"Auto Import Tbilisi LLC"`, Yard `110`); added `localhost` auto-fallback when no `Master Accounts` tab exists plus an explicit `#chkDevMockAccounts` toggle inside **Tools & Setup → Workbook & Schema**. |
| **Rollover from Previous Sheet & `"Completed"` Status Filtering** | `--Project--- TrackerUI.md` (L25, L27) | ✅ Implemented | `v1.13.0` | Fixed `isCompletedStatus()` (`COMPLETED_STATUS_SET`) to recognize `'completed'` (with trailing `'d'`), `'resolved'`, `'done'`, and `'cancelled'`/`'canceled'` so completed items never roll over. Added **Source Sheet Scope** (`Previous Sheet Only` vs. `All Prior Sheets`) in `#rolloverModal`, `#btnRolloverPrevSheet` in **Tools & Setup**, a carry-over option in **New Day Tab** (`#btnAddBlankTab`), and a one-click rollover CTA on empty daily tabs. |
| **Fix Dark Mode `#workbookSelect` White-on-White Dropdown Text** | `--Project--- TrackerUI.md` (L29) | ✅ Implemented | `v1.13.0` | Styled `.workbook-picker-wrap select` with `background-color: var(--surface-muted); color: var(--text)` and added global `select option, select optgroup { background-color: var(--surface); color: var(--text); }` in `css/styles.css` so native Windows Chromium `<option>` menus remain legible without hovering. |
| **PWA Cache-Busting, Auto-Update Banner & Self-Healing Cache Reset** | User Request | ✅ Implemented | `v1.13.1` | Updated `sw.js` (`tracker-ui-v17`) to bypass HTTP disk cache during install (`{ cache: 'reload' }`) and revalidate same-origin assets (`{ cache: 'no-cache' }`); registered `./sw.js` with `{ updateViaCache: 'none' }` and automatic `focus`/`visibilitychange` update checks; added version query strings (`?v=1.13.1`), `#pwaUpdateBanner`, and `#btnCheckForUpdates` (`Check for Updates & Clear Cache`) in **Tools & Setup**. |
| **In-App Workbook Scratch Sheet Tabs & Editable Spreadsheet Grid (Title Tracking)** | User Request | ✅ Implemented | `v1.14.0` | Added `#workspaceTabsBar` (`Daily Tracker` + closable custom scratch sheet tabs persisted per workbook in `localStorage`) and `#scratchSheetMenuWrap` (`Scratch Sheets ▾` picker & `+ Create New Scratch Sheet...`, hiding `template`, `validation lists`, and `Master Accounts`). Added `#scratchSheetWorkspace` featuring an Excel-style editable grid, formula/value bar, arrow/Tab/Enter keyboard navigation, multi-cell TSV clipboard paste (`Ctrl+V`), search/row filtering, carrier tracking number string preservation, 1.2s debounced auto-save + `Ctrl+S` dual-write mirroring, and sheet rename/delete. |
| **Modular Multi-File JavaScript Architecture (Zero-Build Domain Split)** | User Request | ✅ Implemented | `v1.14.1` | Split monolithic 7,992-line `js/app.js` into 6 focused, function-named scripts (`core-ui.js`, `workbook-io.js`, `tracker-table.js`, `request-editors.js`, `scratch-sheets.js`, `rollover.js`) alongside `mockMasterAccounts.js`, preserving zero-build `localhost`/`https://`/`file://` compatibility and synced Service Worker caching (`tracker-ui-v19`). |

---

### ⏳ Partially Implemented / In Progress
| Feature / Suggestion | Source | Status | What's Done vs. What Remains |
|---|---|---|---|
| **Dynamic Schema Mapping Configuration** | `newfeatures.txt` (L21–23) | ⏳ Partial | **Done**: `getWorksheetColumnMap()` and `parseMasterAccountsSheet()` auto-detect the 11-column Standard layout vs. 12-column Tracker Template layout (`Status` in Col I, `Action Taken/Notes` in Col J, `Sent Email Subject` in Col K, `Tracking ID` in Col L).<br>**Remaining**: User-facing column mapping configuration UI for arbitrary custom spreadsheet schemas. |
| **IndexedDB Local Fail-Safe Cache** | `newfeatures.txt` (L25–27) | ⏳ Partial | **Done**: `IndexedDB` (`TrackerUI_DirectoryDB`) persists Agent, Supervisor, and Templates folder handles plus `lastActiveWorkbookName`; `pendingMirrorBuffer` caches unsynced Supervisor `(VIEW ONLY)` saves in memory with 15s auto-retry.<br>**Remaining**: Offline IndexedDB mutation staging queue if the primary working `.xlsx` write fails or disconnects. |

---

### 📋 Planned / Backlog & Parked Scope
| Feature / Suggestion | Source | Status | Notes |
|---|---|---|---|
| **Mass Report Filter (380,000-Row Client-Side `Set` Filter)** | `--Project--- TrackerUI.md` (§3, L20) | ⏸️ Parked | Explicitly parked ("Do not implement yet"). Uses `O(1)` `Set` lookup and lightweight `{ header: 1 }` array-of-arrays parsing to avoid V8 memory crashes on massive reports. |
| **Smart Scratchpad (Real-Time Regex Lot # & Title Tracking Extractor)** | `--Project--- TrackerUI.md` (§3, L21) | ⏸️ Parked | Explicitly parked ("Do not implement yet"). Split-pane raw text input using real-time Regex to extract 8-digit lot numbers and alphanumeric title tracking numbers into a structured table with eventual UUID ticket linking. |
| **Chart.js Metric View for Current Data (Phase 1)** | `newfeatures.txt` (L40) | ⏸️ On Hold | Explicitly deferred per user direction ("We can hold off on Chart.JS view for now") while focusing on core agent workflow & testing branch. |
| **Config Excel Sheet or `.ini` / `.json` File** | `newfeatures.txt` (L55) | 📋 Backlog | Store workspace/settings configuration in an external `.ini`/`.json` file or hidden config sheet inside the workbook. |
| **Supervisor Pipeline: Subtype → Parent Type Lookup (Phase 2)** | `newfeatures.txt` (L44) | 📋 Backlog | Action subtype to parent category lookup mapping for supervisor reporting. |
| **Web Worker Background ExcelJS Parsing & Multi-File Master Tracker Loop (Phase 3)** | `newfeatures.txt` (L29–31, L46–48) | 📋 Backlog | Offload heavy ExcelJS parsing/search to a Web Worker and aggregate multiple agent workbooks into a Master Tracker. |

---

### 🚫 Decided Against / Superseded
*Use this section to record any suggested features or architectural ideas that were evaluated and intentionally declined or superseded, along with the reason.*

| Feature / Suggestion | Date Decided | Reason / Alternative Chosen |
|---|---|---|
| **Requiring / Creating `Master Accounts` Sheet in New Monthly Workbooks** | `2026-11` | **Decided Against** (`newfeatures.txt` L15): Creating a new monthly workbook no longer creates or requires a `Master Accounts` tab. Bidirectional Broker ↔ Master # auto-fill automatically learns from existing rows across daily tabs (while still reading `Master Accounts` if present in a legacy/template file). |
| **Single-File `showOpenFilePicker` Workflow** | `2026-10` | **Superseded** by `window.showDirectoryPicker({ mode: 'readwrite' })` so the app can automatically discover monthly workbooks and dual-write `<Workbook Name> (VIEW ONLY).xlsx` mirrors without prompting on every save. |
| **Using Ticket Number (Col B) or Broker Hash as Primary Key** | `2026-10` | **Decided Against** because Ticket Numbers are optional metadata that may be added or edited later. Pure immutable UUIDv4 (`crypto.randomUUID()`) in hidden Column K/L is used instead. |
| **Bundlers / NPM Build Frameworks (Vite, Webpack, React)** | `2026-10` | **Decided Against** to keep TrackerUI a zero-build, standalone PWA that runs directly via VS Code Live Server (`localhost`) and GitHub Pages (`https://`). |

---

## 2. Version Release History

### `v1.14.1` (Current — `meerkat`)
- **Modular Multi-File JavaScript Refactor (`js/*.js`)**:
  - Split the 7,992-line monolithic `js/app.js` into 6 focused, function-named files for easier manual editing and significantly smaller context footprints during feature iterations:
    - [`js/core-ui.js`](js/core-ui.js): Shared state, DOM references, Light/Auto/Dark theme & text-size scaling, modal dialog engine, Tools popover, status bar telemetry, toasts, and PWA Service Worker update checks.
    - [`js/workbook-io.js`](js/workbook-io.js): IndexedDB folder handles, OpenXML archive preservation/repair, ExcelJS load/save, dual-write mirroring & auto-retry, schema bootstrapping, daily tab CRUD, and reference sheet auto-fill.
    - [`js/tracker-table.js`](js/tracker-table.js): Worksheet row indexing & UUIDv4 backfill, sheet stepper/tab strip, request table rendering, global search, sorting, Like-Item filtering, batch status updates, and batch delete.
    - [`js/request-editors.js`](js/request-editors.js): Docked bottom request editor, popup request modal, cross-tab status propagation, ticket details timeline modal, and Supervisor View-Only Mirror inspector.
    - [`js/scratch-sheets.js`](js/scratch-sheets.js): In-app scratch sheet workspace tabs, editable spreadsheet grid, formula/value bar, keyboard navigation, and multi-cell TSV clipboard paste.
    - [`js/rollover.js`](js/rollover.js): Unique request identity matching, chronological notes merging, smart rollover engine, and startup initialization.
  - Updated [`index.html`](index.html) (`?v=1.14.1`), [`sw.js`](sw.js) (`CACHE_NAME = 'tracker-ui-v19'` & `ASSETS_TO_CACHE`), and [`AGENTS.md`](AGENTS.md).

### `v1.14.0`
- **In-App Workbook Scratch Sheet Tabs & Editable Spreadsheet Grid (`#workspaceTabsBar` / `#scratchSheetWorkspace`)**:
  - Added a top-level in-app **Workspace Tab Bar** (`#workspaceTabsBar`) featuring the pinned **Daily Tracker** tab alongside open custom/scratch worksheet tabs (persisted per workbook in `localStorage` and closable via `×` without removing the sheet from Excel).
  - Added the **`Scratch Sheets ▾`** picker (`#scratchSheetMenuWrap`), surfacing only custom/scratch worksheets in the active workbook while hiding internal reference and daily tabs (`template`, `validation lists`, `Master Accounts`, and `MM-DD-YYYY` sheets).
  - Added **Create New Scratch Sheet...** (`#btnNewScratchSheetFromBar` & `#menuNewScratchSheet` in **Tools & Setup**) with choice of a **Title Tracking Starter Layout** (`Lot #`, `Yard #`, `Broker / Master #`, `Title Status`, `Carrier / Tracking #`, `Date Updated`, `Notes`) or a blank grid, plus scratch sheet **Rename** and **Delete Sheet** actions.
  - Built an Excel-style **Editable Spreadsheet Grid** (`#scratchSheetWorkspace`): supports inline double-click/type-to-edit cells, coordinate & `fx` formula/value bar (`#scratchFormulaInput`), `Arrow` / `Tab` / `Enter` / `F2` / `Delete` navigation, multi-row/column TSV clipboard paste (`Ctrl+V` from Excel/emails/carrier portals), live cell search + row filter, `+10 Rows` / `+ Col` expansion, automatic string preservation for 11+ digit carrier tracking numbers (preventing scientific notation), and 1.2s debounced auto-save + `Ctrl+S` dual-write mirroring via `saveWorkbookToDisk()`.
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v18'` in `sw.js` and updated asset version query strings to `?v=1.14.0`.

### `v1.13.1`
- **PWA Cache-Busting, Background Update Banner & Self-Healing Cache Reset**:
  - Updated [`sw.js`](sw.js) (`CACHE_NAME = 'tracker-ui-v17'`) to fetch local app shell assets with `{ cache: 'reload' }` during installation and `{ cache: 'no-cache' }` on same-origin GET requests so Chrome never copies stale JavaScript/CSS from its HTTP disk cache into the Service Worker cache.
  - Added `{ ignoreSearch: true }` offline fallback matching in [`sw.js`](sw.js) alongside versioned asset URLs (`css/styles.css?v=1.13.1`, `js/mockMasterAccounts.js?v=1.13.1`, `js/app.js?v=1.13.1`) in [`index.html`](index.html) so HTML and JS can never desynchronize.
  - Registered `./sw.js` with `{ updateViaCache: 'none' }` in [`js/app.js`](js/app.js) and added throttled `focus` / `visibilitychange` update checks (`checkServiceWorkerUpdateOnResume`) so installed desktop PWAs automatically detect new deployments when brought to the foreground.
  - Added a top **Update Ready** banner (`#pwaUpdateBanner` with **Reload Now**) plus a one-click **Check for Updates & Clear Cache** action (`#btnCheckForUpdates` in **Tools & Setup** and clickable `#appVersionTag` badge) via `forceClearAppCacheAndReload()`.

### `v1.13.0`
- **Batch Status Update Architecture (`#batchActionDock` & Permanent Selection Checkboxes)**:
  - Made the left-edge selection checkbox column (`#thDeleteSelect` / `.chk-delete-row`) permanent on the request table so rows can be multi-selected at any time without entering a separate mode.
  - Added a floating bottom **Batch Action Dock** (`#batchActionDock`) rendered whenever ≥1 rows are checked, providing bulk status selection (`#batchStatusSelect`), **Apply Status** (`#btnApplyBatchStatus`), **Delete** (`#btnBatchDeleteSelected`), and **Clear** (`#btnClearBatchSelection`).
  - Implemented `handleApplyBatchStatusUpdate()` on top of ExcelJS: preserves every row's immutable hidden UUIDv4 (`Tracking ID` in Col K/L), updates `allRows` and re-renders the DOM optimistically, and executes a single `saveWorkbookToDisk()` dual-write to the working `.xlsx` and Supervisor `(VIEW ONLY).xlsx` mirror.
- **Auto-Fill Graceful Degradation & Mock Testing Dataset (`js/mockMasterAccounts.js`)**:
  - Guarded `setBrokerCellWithFormula()` so 12-column template layouts never write `_xlfn.XLOOKUP(..., 'Master Accounts'!...)` formulas when the optional `'Master Accounts'` worksheet is absent, falling back cleanly to plain string Broker cells while keeping all form inputs editable.
  - Added standalone `js/mockMasterAccounts.js` (`window.MOCK_MASTER_ACCOUNTS`) containing 10 sample Broker / Master # / Bidder # / Yard # records (including `"Auto Import Tbilisi LLC"`, `MA-104820`, `BD-882104`, Yard `110`) plus bidirectional Yard auto-fill when empty.
  - Added automatic `localhost` mock fallback when a workbook has no `Master Accounts` tab and an explicit **Mock Master Accounts (Dev)** toggle (`#chkDevMockAccounts`) inside **Tools & Setup → Workbook & Schema**.
- **Rollover from Previous Sheet & `"Completed"` Status Fix**:
  - Upgraded `isCompletedStatus()` (`COMPLETED_STATUS_SET`) to match `'completed'` (with trailing `'d'`), `'complete'`, `'closed'`, `'resolved'`, `'done'`, and `'cancelled'`/`'canceled'` so `"Completed"` requests no longer roll forward.
  - Added **Source Sheet Scope** (`Previous Sheet Only` vs. `All Prior Sheets`) to `#rolloverModal`, a **Rollover from Previous Sheet...** action (`#btnRolloverPrevSheet`) in **Tools & Setup**, a carry-over mode when creating a **New Day Tab** (`#btnAddBlankTab`), and a one-click rollover CTA on empty daily tabs.
- **Dark Mode `#workbookSelect` Dropdown Contrast Fix**:
  - Updated `.workbook-picker-wrap select` (`background-color: var(--surface-muted); color: var(--text)`) and added global `select option, select optgroup { background-color: var(--surface); color: var(--text); }` rules in `css/styles.css` to eliminate white-on-white dropdown text on Windows Chromium.
- **PWA Service Worker Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v16'` in `sw.js` and added `./js/mockMasterAccounts.js` to offline precache assets.

### `v1.12.0`
- **Light Mode Palette & `Light` / `Auto` / `Dark` Theme Switcher**:
  - Added full high-contrast Light Mode and Dark Mode semantic color schemes in `css/styles.css`, including `@media (prefers-color-scheme: light)` automatic system detection when in `Auto` mode and explicit `[data-theme="light" | "dark"]` overrides.
  - Added synced `Light` / `Auto` / `Dark` segmented theme controls in both the top header (`#headerThemeGroup`) and **Tools & Setup** (`#menuThemeGroup`), persisted in `localStorage` (`trackerui_theme_mode`) with live OS preference listeners and dynamic PWA `<meta name="theme-color">` updates.
- **Professional UI / Menu Cleanup & Whole-App Text Size Slider**:
  - Redesigned buttons, badges, and the **Tools & Setup** popover menu with clean inline SVG icons, consistent hierarchy, and refined enterprise styling.
  - Converted typography across `css/styles.css` to `rem` units scaled from `html { font-size: calc(16px * var(--font-scale, 1)); }` and added an interactive **Text Size** slider (`85%`–`125%`, `A−` / `A+` steppers, `Reset`) inside **Tools & Setup**, persisted in `localStorage` (`trackerui_font_scale`).
  - Moved **Open View-Only Mirror...** (`#menuOpenViewOnly`) out of the top header and into **Tools & Setup → Directories & Mirroring**.
- **Consolidated Unique-Request Rollover & Chronological Action Notes Appending**:
  - Indexed daily `MM-DD-YYYY` sheets chronologically (oldest → newest) so carried-over rows without a hidden Column K `Tracking ID` automatically inherit the prior sheet's `Tracking ID` instead of generating separate per-day UUIDs.
  - Upgraded `computeRolloverAnalysis()` and `rolloverForm` (`areRowsSameUniqueRequest` + `mergeChronologicalActionNotes`) so each unique request across multiple prior days rolls forward **once** rather than duplicating one row per prior day, chronologically appending and deduplicating all prior days' `Action Notes` (`[MM-DD-YYYY] ...`), enriching existing target rows if prior notes were missing, and unifying hidden Column K/L `Tracking ID`s across consolidated rows.

### `v1.11.2`
- **Viewport-Pinned & Collapsible Bottom Request Editor Dock (`#bottomEditorDock`)**:
  - Restructured the workspace into a `100dvh` flex-column layout so `.table-wrap` scrolls independently from top to bottom (with sticky `thead#tableHead th` column headers) while `#bottomEditorDock` stays permanently pinned above the bottom status bar.
  - Added collapse/expand support (`#btnDockToggleCollapse` `▾ Collapse` / `▴ Expand Editor` and clickable `#dockHeaderToggleZone`), persisting the collapsed preference in `localStorage` (`trackerui_dock_collapsed`) and automatically expanding when selecting a row to edit or clicking `+ New Request`.
  - Removed window scroll-jumping (`scrollIntoView`) when clicking rows so list scroll position stays steady.

### `v1.11.1` — Commit `02c8ac7`
- **Fixed Console `ReferenceError` Crashes Reported in Chrome Testing (`errorlogs`)**:
  - Resolved `Uncaught ReferenceError: btnTicketEditInDock is not defined` on startup (which had halted top-level script attachment for listeners below line 4182). Added `#btnTicketEditInDock` (`✏️ Edit in Bottom Dock`) to `#ticketDetailsModal` and declared its DOM reference.
  - Resolved `ReferenceError: dockTitleText is not defined` (and `dockTargetMeta`) in `setBottomDockNewMode()` / `loadRowIntoBottomDock()`, aligning the Bottom Editor Dock with `#dockTargetLabel`, `#dockKeyBadge`, `#dockHintText`, and `.editing-mode` / `.edit-mode` styles so workbook loading and day-tab switching complete cleanly.
  - Aligned Ticket Details Modal (`renderTicketDetailsModal`) and Supervisor View-Only Inspector (`openViewOnlyTrackerModal`, `renderViewOnlyTable`, `#btnCloseViewOnly`) with their declared `viewOnlyState` object, DOM element IDs, and CSS classes, plus added `sanitizeXlsxBufferForExcelJS()` before loading View-Only snapshots.
  - Wired `#btnModalInsertTimestamp` in the popup editor modal to insert `[MM/DD/YYYY HH:MM]` timestamps.
  - Added standard `<meta name="mobile-web-app-capable" content="yes">` alongside `apple-mobile-web-app-capable` in `index.html` to eliminate the Chrome deprecation warning.

### `v1.11.0` — Commit `00c24ff`
- **Natural Filter Auto-Clearing**: Clicking Day Tab pills (`#sheetTabsBar`), `◀`/`▶` steppers, `#sheetSelect`, `#scopeSelect`, `📅 Date Tab` cells, or switching workbooks naturally clears active Like-Item/Tracking ID filters and search queries so agents never have to manually clear filters when changing views.
- **Docked Bottom Request Editor (`#bottomEditorDock`)**: Clicking any row in the table auto-populates the bottom editor panel for fast inline editing (with the immutable Tracking ID excluded from editable inputs), plus a `🕒 +Timestamp` note helper and `＋ New / Clear` mode.
- **Updated Monthly Workbook Creation & Templates Folder**: Excluded `Master Accounts` from new monthly workbook bootstrapping (`newfeatures.txt` L15) and added support for cloning from a custom **Templates Folder** (`templatesFolderHandle`).
- **Auto-Reconnect on Startup & Persistent File Paths Bar (`#filePathsBar`)**: Automatically reconnects persisted folders and the last active workbook on startup when permission is granted, adds a checkbox to disable auto-reconnect (`#chkAutoReconnect`), and displays full relative paths for the active working tracker, `(VIEW ONLY)` mirror, and templates folder.
- **`👁️ Open View-Only Tracker` Inspector (`#viewOnlyModal`)**: Added header & Tools menu button to inspect the supervisor `(VIEW ONLY).xlsx` file via a non-locking read-only `ArrayBuffer` snapshot or download a copy.
- **Ticket Details & Aggregated Timestamped Notes Modal (`#ticketDetailsModal`)**: Added `📋 Details` view showing ticket metadata, original creation date & days tracked, a chronological timeline of notes across all daily tabs, and an inline **Append Timestamped Note** form.

### `v1.10.0` — Commit `e58c9c9`
- Added visible application version badge (`v1.10.0`) in the top header (`#appVersionTag`) and bottom status bar.
- Added Service Worker (`sw.js`) update detection toast notifying users when a new version is deployed from GitHub.

### `v1.9.0` — Commit `bc1bec9`
- Fixed Desktop Excel workbook corruption warning when saving workbooks containing dynamic array formulas (`FILTER`, `LET`), `XLOOKUP`, external links (`[1]`), and cross-sheet Data Validation dropdowns (`<extLst>`).
- Added raw OpenXML ZIP post-processor (`captureOriginalArchiveState`, `repairSavedXlsxArchive`) to preserve `xl/metadata.xml`, `xl/externalLinks/*`, and shared string spill results.
- Added automatic detection for 12-column Tracker Template layout (`Status` in Col I, `Action Taken/Notes` in Col J, `Sent Email Subject` in Col K, hidden `Tracking ID` in Col L).

### `v1.8.0` — Commit `6c81c77`
- Implemented **Like-Item Filtering**: clicking `Master #`, `Bidder #`, `Lot Numbers` (individual lot tokens), `Yard #`, `Subject`, `Ticket #`, `Broker`, `Action`, `Status`, or `Tracking ID` filters all matching rows across all daily tabs.
- Implemented **Clickable Date Tab Jump**: clicking `📅 MM-DD-YYYY` in any row switches directly to that daily sheet.
- Implemented **3-State Column Header Sorting** (`▲ Asc`, `▼ Desc`, `↕ Unsorted`) across all columns, including split `Master / Bidder` and `Yard / Lots` sub-headers, plus a `✕ Clear Sort / Filter` button.
- Fixed `docProps/core.xml` namespace parsing error in ExcelJS during Rollover (`sanitizeXlsxBufferForExcelJS`).

### `v1.7.0` — Commit `fe5bd7e`
- Replaced all native browser `alert()`, `confirm()`, and `prompt()` calls with a custom in-app modal dialog engine (`#dialogModal`).
- Added a subtle shadow **"＋ Add new support request"** row (`.shadow-new-row`) at the bottom of the request table.

### `v1.6.0` — Commits `0d0b6e7`, `9e7ac98`
- Added persistent bottom workspace status bar (`#statusBar`) displaying sync status, connected Agent/Supervisor folder names, open/total request counts, and last-updated timestamp.
- Ordered the horizontal `#sheetTabsBar` chronologically (oldest left → newest right) to match the `◀` / `▶` stepper buttons.

### `v1.5.0` — Commits `e499a87`, `adda6a3`
- Upgraded to directory-based dual-write I/O (`showDirectoryPicker`) with custom Supervisor `(VIEW ONLY)` export folder and 15-second background auto-retry when the mirror file is locked.
- Added pure UUIDv4 tracking in hidden Column K, automatic global search across all daily tabs, batch delete selection mode, and smart multi-sheet rollover with future-date handling.

