# TrackerUI — Changelog & Feature Tracker

This file tracks all requested features, architectural suggestions (from [`newfeatures.txt`](newfeatures.txt) and user sessions), implementation statuses, and version release history.

> **Automated Tracking Rule**: Whenever a feature or suggestion is implemented, modified, or decided against, this file is automatically updated by the coding assistant per [`AGENTS.md`](AGENTS.md).

---

## 1. Feature & Suggestion Status Matrix (`newfeatures.txt`)

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
| **`👁️ Open View-Only Tracker` Button & Read-Only ArrayBuffer Inspector** | `newfeatures.txt` (L42, L51) | ✅ Implemented | `v1.11.0` | Added `#btnOpenViewOnly` in the top header and Tools menu to inspect `<Workbook Name> (VIEW ONLY).xlsx` in-app via a non-locking read-only `ArrayBuffer` snapshot (`#viewOnlyModal`), complete with sheet selector, search filter, and one-click `.xlsx` copy download. |
| **Ticket Details & Aggregated Timestamped Notes Modal + Creation Date / Age Tracking** | `newfeatures.txt` (L33–35, L59) | ✅ Implemented | `v1.11.0` | Added `📋 Details` button on every row and in the Bottom Editor Dock opening `#ticketDetailsModal`, which aggregates all daily tab entries for a ticket chronologically, displays original creation tab & day span, highlights `[MM/DD/YYYY HH:MM]` timestamps, and provides an inline **Append Timestamped Note** form. |

---

### ⏳ Partially Implemented / In Progress
| Feature / Suggestion | Source | Status | What's Done vs. What Remains |
|---|---|---|---|
| **Dynamic Schema Mapping Configuration** | `newfeatures.txt` (L21–23) | ⏳ Partial | **Done**: `getWorksheetColumnMap()` and `parseMasterAccountsSheet()` auto-detect the 11-column Standard layout vs. 12-column Tracker Template layout (`Status` in Col I, `Action Taken/Notes` in Col J, `Sent Email Subject` in Col K, `Tracking ID` in Col L).<br>**Remaining**: User-facing column mapping configuration UI for arbitrary custom spreadsheet schemas. |
| **IndexedDB Local Fail-Safe Cache** | `newfeatures.txt` (L25–27) | ⏳ Partial | **Done**: `IndexedDB` (`TrackerUI_DirectoryDB`) persists Agent, Supervisor, and Templates folder handles plus `lastActiveWorkbookName`; `pendingMirrorBuffer` caches unsynced Supervisor `(VIEW ONLY)` saves in memory with 15s auto-retry.<br>**Remaining**: Offline IndexedDB mutation staging queue if the primary working `.xlsx` write fails or disconnects. |
| **CSS Variables & Light Mode Toggle** | `newfeatures.txt` (L39) | ⏳ Partial | **Done**: `:root` CSS variables are implemented in `css/styles.css`.<br>**Remaining**: Light mode color palette and header theme toggle button. |

---

### 📋 Planned / Backlog & On Hold
| Feature / Suggestion | Source | Status | Notes |
|---|---|---|---|
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

### `v1.11.2` (Current)
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

