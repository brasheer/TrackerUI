# TrackerUI — Changelog & Release History

This document records the version release history, major feature rollouts, bug fixes, and architectural milestones for **TrackerUI**.

> For upcoming features, backlog items, and architectural proposals, see [`FUTURE_FEATURES.md`](FUTURE_FEATURES.md).

---

## Version Release History

### `v1.15.0` (Current — `main`)
- **Excel-Like Scratch Spreadsheets (`js/scratch-sheets.js`, `css/styles.css`, `index.html`)**:
  - **Multi-Cell Range Highlighting**: Added rectangular cell range selection (`anchor` + `focus`) via click-and-drag, `Shift + Click`, `Shift + Arrow` keys, `Ctrl/Cmd + Arrow` data-boundary jumps (`Ctrl + Shift + Arrow` to highlight blocks), `Ctrl/Cmd + A` (data bounds on first press, whole grid on second), and clicking/dragging row numbers, column letters, or the top-left `#` corner header.
  - **Smooth Edge Auto-Scroll & Grid Auto-Expansion**: Dragging near or past the edges of `#scratchGridWrap` smoothly scrolls the viewport at 60fps (`requestAnimationFrame`) and automatically expands blank rows/columns when dragging past the bottom or right sheet boundary.
  - **Multi-Cell Copy (`Ctrl+C`), Cut (`Ctrl+X`), Paste (`Ctrl+V` + Repeat-Fill) & Bulk Clear**: Added full rectangular TSV copy/cut with dashed copied-range outlines, multi-cell block paste, Excel-style 1-cell repeat-fill when pasting a single copied value onto a multi-cell highlight, and bulk `Delete`/`Backspace` range clearing.
  - **50-Step Undo / Redo Stack (`Ctrl+Z` / `Ctrl+Y`)**: Added per-sheet 50-step Undo/Redo history (`#btnScratchUndo`, `#btnScratchRedo`, `Ctrl+Z`, `Ctrl+Y` / `Ctrl+Shift+Z`) covering single-cell edits, formula bar edits, multi-cell pastes, cuts, and range clears.
  - **Interactive Column Width Resize & Selection Quick-Stats**: Added drag-to-resize handles (`.scratch-col-resizer`) on column headers with double-click auto-fit persisted directly to Excel column widths, plus a live `#scratchSelectionStats` pill (`R × C`, `Count`, numeric `Sum`, and `Avg`) in the formula bar when 2+ cells are highlighted.
- **Excel-Style Alternating Shaded Tracker Rows (`css/styles.css`)**:
  - Added alternating bright/gray row shading (`--row-stripe`) to tracker request rows (`tbody#listBody tr[data-sheet]`) and the Supervisor View-Only Mirror table (`.view-only-table-wrap tbody tr`) in both Light Mode (`#ffffff` / `#f1f5f9`) and Dark Mode (`#0f172a` / `#172338`), alongside crisp hover (`--row-hover`) and active bottom-dock editing row (`--row-active` + left accent bar) states.
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v20'` in [`sw.js`](sw.js) and updated asset version query strings to `?v=1.15.0` in [`index.html`](index.html).

### `v1.14.1`
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
  - Resolved `Uncaught ReferenceError: btnTicketEditInDock is not defined` on startup. Added `#btnTicketEditInDock` (`✏️ Edit in Bottom Dock`) to `#ticketDetailsModal` and declared its DOM reference.
  - Resolved `ReferenceError: dockTitleText is not defined` (and `dockTargetMeta`) in `setBottomDockNewMode()` / `loadRowIntoBottomDock()`, aligning the Bottom Editor Dock with `#dockTargetLabel`, `#dockKeyBadge`, `#dockHintText`, and `.editing-mode` / `.edit-mode` styles so workbook loading and day-tab switching complete cleanly.
  - Aligned Ticket Details Modal (`renderTicketDetailsModal`) and Supervisor View-Only Inspector (`openViewOnlyTrackerModal`, `renderViewOnlyTable`, `#btnCloseViewOnly`) with their declared `viewOnlyState` object, DOM element IDs, and CSS classes, plus added `sanitizeXlsxBufferForExcelJS()` before loading View-Only snapshots.
  - Wired `#btnModalInsertTimestamp` in the popup editor modal to insert `[MM/DD/YYYY HH:MM]` timestamps.
  - Added standard `<meta name="mobile-web-app-capable" content="yes">` alongside `apple-mobile-web-app-capable` in `index.html` to eliminate the Chrome deprecation warning.

### `v1.11.0` — Commit `00c24ff`
- **Natural Filter Auto-Clearing**: Clicking Day Tab pills (`#sheetTabsBar`), `◀`/`▶` steppers, `#sheetSelect`, `#scopeSelect`, `📅 Date Tab` cells, or switching workbooks naturally clears active Like-Item/Tracking ID filters and search queries so agents never have to manually clear filters when changing views.
- **Docked Bottom Request Editor (`#bottomEditorDock`)**: Clicking any row in the table auto-populates the bottom editor panel for fast inline editing (with the immutable Tracking ID excluded from editable inputs), plus a `🕒 +Timestamp` note helper and `＋ New / Clear` mode.
- **Updated Monthly Workbook Creation & Templates Folder**: Excluded `Master Accounts` from new monthly workbook bootstrapping and added support for cloning from a custom **Templates Folder** (`templatesFolderHandle`).
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
