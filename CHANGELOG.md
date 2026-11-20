# TrackerUI — Changelog & Release History

This document records the version release history, major feature rollouts, bug fixes, and architectural milestones for **TrackerUI**.

> For upcoming features, backlog items, and architectural proposals, see [`FUTURE_FEATURES.md`](FUTURE_FEATURES.md).

---

## Version Release History

### `v1.19.5` (Current — `main` / `meerkat`)
- **🕒 Row Bottom Metadata Bar: UUIDv4 & Last-Accessed Timestamp (`js/core-ui.js`, `js/tracker-table.js`, `js/request-editors.js`, `css/styles.css`)**:
  - Moved the **Tracking ID (UUIDv4)** chip (now expanded to show the full 36-character UUIDv4) and the new **Last Accessed** timestamp (`2:16 PM` for today, `11/19 2:16 PM` for earlier days, with full date-time tooltip) into a slim bottom metadata bar spanning underneath each daily tracker row (for both Standard and Custom Department schemas), freeing up horizontal table space so the **Subject** column stays wider and flatter.
  - Both **Accessed** and **ID** remain 1-click sortable via the rightmost **`Accessed ↕ / ID ↕`** header controls and fully searchable via global search (`buildItemSearchHaystack`), with last-accessed timestamps persisted locally per immutable Tracking ID (`trackerui_row_last_accessed_v1`) and updated live on row click / Edit / Details without modifying `.xlsx` columns.
- **📐 Compact Row-Number & Checkbox Columns (`css/styles.css`)**:
  - Squished the left-edge `#` row-number column and checkbox selection column to minimal width (`1%` / tight padding / compact `0.85rem` checkbox, spanning cleanly across each row + bottom bar) and added a clean `1px` vertical right border on the checkbox column to separate selection controls from request data.
- **🛡️ Unsaved New Request Draft Autosave & Switch Warning (`index.html`, `js/core-ui.js`, `js/request-editors.js`, `js/tracker-table.js`, `css/styles.css`)**:
  - Automatically saves in-progress **New Request** entries in the bottom editor dock (Standard & Custom schemas) to `localStorage` (`trackerui_dock_new_draft_v1`), displaying a live **`● Draft Autosaved`** badge (`#dockDraftBadge`) and a 1-click **`📝 Resume Draft`** button (`#btnDockResumeDraft`) when viewing another row.
  - Prompts a confirmation warning (`confirmSwitchAwayFromUnsavedNewDraft`) before loading another row when an unsaved new request is being filled out (keeping the draft safely auto-saved even if you choose to switch rows), and confirms before discarding an active draft via `＋ Clear / New`.
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v45'` in [`sw.js`](sw.js) and synced `?v=1.19.5` asset query strings across [`index.html`](index.html) and [`js/core-ui.js`](js/core-ui.js).

### `v1.19.4` (`main` / `meerkat`)
- **🔁 Rollover Status Transitions & Waiting/Hold Overdue Exemption (`js/core-ui.js`, `js/rollover.js`)**:
  - Automatically transitions requests with `"New"` status to `"In Progress"` (`resolveInProgressStatusLabel()`) when carried forward to a subsequent daily tab, displaying `New → In Progress` badges in the Rollover preview and reporting transition counts on completion.
  - Exempts `Waiting`, `Waiting on Buyer`, and `On Hold` requests from `evaluateRequestOverdueMetrics()` (`qualifiesForAutoOverdue`) so aged or frequently carried waiting items remain `Waiting` / `On Hold` with purple styling rather than mutating to `Overdue`.
- **🔄 Always-Live Bidirectional Master Buyer Account Sync (`js/workbook-io.js`)**:
  - Updated `applyBrokerAutoFillToPair()` and `applyMasterReverseAutoFillToPair()` via `syncAccountAuxiliaryFields()` so changing `Broker Name` or `Master #` to a known account immediately updates the paired field and default `Bidder #` without requiring fields to be cleared first (while preserving custom edits when re-validating the same account).
- **📋 Email HTML `<table>` Paste Priority, Multi-Line Preservation & Transpose Actions (`js/scratch-sheets.js`, `index.html`)**:
  - Updated `parseClipboardHtmlTableToMatrix()` and `applyScratchPasteFromText()` to reconstruct cell values and multi-line `<br>` / `<p>` / `<div>` content (plus `colspan`/`rowspan` and inline/`<style>` styling) directly from `text/html` `<table>` payloads before falling back to RFC-4180 multi-line TSV parsing.
  - Added **`Paste Special → Transpose`** and **`Transpose Selection`** to `#scratchContextMenu` (`transpose2DMatrix()`, `transposeScratchSelection()`) with relative formula shifting and full Undo/Redo support.
  - Fixed `applyScratchCellChanges()` grid expansion detection (`needsExpand` / Row-1 header refresh) and preserved `scrollTop`/`scrollLeft` in `renderActiveScratchSheetGrid()`, plus added a `localhost` Live Server `/ws` reload guard in `js/core-ui.js` (`GuardedWebSocket`) so auto-saving test `.xlsx` workbooks inside `assets/Testing/` never triggers a dev-server browser page reload.
- **🟦 Explicit Fill Handle `Auto-increment` Toggle (`index.html`, `js/core-ui.js`, `js/scratch-sheets.js`, `css/styles.css`)**:
  - Added the persistent **`Auto-increment`** toolbar checkbox (`#chkScratchAutoIncrement`, defaulting unchecked so dragging copies values safely; hold `Ctrl` while dragging to invert) and `BigInt` series support for 11–30 digit tracking numbers while still saving 11+ digit strings as text in `.xlsx`.
- **✂️ Data to Columns: Newline (`\n`) Delimiter & Split Direction (`index.html`, `js/core-ui.js`, `js/scratch-sheets.js`, `css/styles.css`)**:
  - Added the **`Newline (\n)`** delimiter (`#chkTtcNewline`, auto-detected before sample joining) and **`Split Across (Columns →)` / `Split Down (Rows ↓)`** toggle (`#radTtcAcross` / `#radTtcDown`) to `#textToColumnsModal`, with downward overwrite detection and Undo/Redo support.
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v42'` in [`sw.js`](sw.js) and synced `?v=1.19.4` asset query strings across [`index.html`](index.html) and [`js/core-ui.js`](js/core-ui.js).

### `v1.19.3` (`main` / `meerkat`)
- **🔧 OpenXML Dynamic Array & View-Only Mirror Repair Fix (`js/workbook-io.js`, `js/scratch-sheets.js`)**:
  - Scoped `repairSavedXlsxArchive()`'s `t="s"` → `t="str"` conversion strictly to legacy reference sheets and active multi-cell `ref="Start:End"` spill rectangles (with XML entity escaping), preventing static shared-string cells on Scratch Sheets from triggering Desktop Excel repair prompts.
  - Updated dynamic array post-processing and `recalculateScratchWorksheetFormulas()` to preserve cached results for unsupported Excel functions (`=LET`/`=XLOOKUP`), purge stale cached spill cells when an array anchor is edited to an in-app evaluated formula, and write `<f t="array" ref="<cell>">` + `cm="1"` + `xl/metadata.xml` (`XLDAPR`) for `=FILTER`, `=SORT`, `=UNIQUE`, `=LET`, and `=XLOOKUP` cells so Excel never prepends legacy `@` implicit intersection (`=@FILTER(...)`).
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v40'` in [`sw.js`](sw.js) and updated asset version query strings to `?v=1.19.3` in [`index.html`](index.html) and [`js/core-ui.js`](js/core-ui.js).

### `v1.19.2` (`main` / `meerkat`)
- **📖 Comprehensive User Manual Update & Excel Function Reference (`USER_MANUAL.md`, `index.html`)**:
  - Updated all 15 sections across both [`USER_MANUAL.md`](USER_MANUAL.md) and the interactive in-app `#userManualModal` (adding `Section 12: Scratch Sheet Excel Functions, Formula Syntax & Dynamic Array Reference` covering `=FILTER`, `=UNIQUE`, `=SORT`, `=IF`, `=AND`, `=OR`, `=NOT`, `=ISBLANK`, `=ISEMPTY`, `=ISNUMBER`, `=ISTEXT`, `=COUNTIF`, `=COUNTA`, `=COUNTBLANK`, `=SUM`, `=AVERAGE`, `=MIN`, `=MAX`, `=TRIM`, `=UPPER`, `=LOWER`, `=LEN`, `=LEFT`, `=RIGHT`, `=MID`, `=SEARCH`, `=FIND`, `=TODAY()`, `*`/`+` boolean array criteria, `$` absolute/relative coordinates, cross-sheet `Sheet1!Range` references, and spill error codes).
  - Documented `v1.18.1`–`v1.19.1` features including 1080p responsive space optimization vs. ultrawide `#sheetTabsBar`, Overdue (red fill + outline) & Waiting/On Hold (purple fill + outline) highlights, Left-Hand Row Numbers (`#`), Integrated Split-Pill Copy Buttons, Multi-Tab Session Standby Lock & Fail-Safe Primary Save staging, Manual Row Copy & Paste across daily tabs, and Scratch Sheet Data Types + Corner Fill Handle auto-increment.
- **📸 Refreshed High-Resolution Manual Screenshots (`assets/images/manual/01`–`16`)**:
  - Updated all existing screenshots (`01`–`10`) to `v1.19.2` and added 6 new visual walkthrough screenshots (`11`–`16`) covering Overdue/Waiting highlights, Manual Row Copy/Paste, the `Auto-Filter Lists` builder modal, Live `=FILTER` Red/Yellow/Green side lists with Corner Fill Handle, Multi-Tab Standby & Retry banners, and the View-Only Mirror Inspector.
  - Fixed all image references in [`USER_MANUAL.md`](USER_MANUAL.md) and embedded all 16 screenshots in `#userManualModal`.
- **🗺️ Roadmap Addition (`FUTURE_FEATURES.md`)**:
  - Added **Simple Text Editing & Cell Formatting Controls for Scratch Sheets** (`Ctrl+B` Bold/Unbold toggle to override auto-bolded Row 1 headers, `Ctrl+I`, `Ctrl+U`, Left/Center/Right text alignment, and Clear Formatting) under Section 1 (`📋 Planned`).
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v39'` (with all 16 manual screenshots added to `ASSETS_TO_CACHE`) in [`sw.js`](sw.js) and synced `?v=1.19.2` asset query strings across [`index.html`](index.html) and [`js/core-ui.js`](js/core-ui.js).

### `v1.19.1` (`main` / `meerkat`)
- **📐 1920x1080 Screen-Space Optimization (`css/styles.css`)**:
  - Added `@media (max-width: 1920px), (max-height: 1080px)` rules that hide `#sheetTabsBar` (`.sheet-tabs-bar`) and tighten vertical padding across the header, profile/sheet bars, controls toolbar, table cells, bottom editor dock, and status bar on `1920x1080` and smaller displays while keeping the quick tab strip visible on ultrawide monitors (e.g., `3440x1440`).
- **🚨 Overdue & Waiting / On Hold Row Backgrounds + Outlines (`css/styles.css`, `js/tracker-table.js`, `js/rollover.js`)**:
  - Updated `tbody#listBody tr.overdue-outline-red` to render a red background tint (`var(--danger-soft)`, deepening to `var(--danger-soft-hover)` on hover) alongside its red outline, matching `.preview-item.overdue-outline-red` in the Rollover modal.
  - Added matching purple/violet background + outline styling (`.waiting-outline-purple` using `var(--waiting)` / `var(--waiting-soft)` / `var(--waiting-soft-hover)`) for **Waiting**, **Waiting on Buyer**, and **On Hold** rows in both the main tracker table and the Rollover preview pane.
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v38'` in [`sw.js`](sw.js) and updated asset version query strings to `?v=1.19.1` in [`index.html`](index.html).

### `v1.19.0` (`main` / `meerkat`)
- **📋 Excel Copy + `Enter` to Paste & Relative Formula Shifting (`js/scratch-sheets.js`, `index.html`)**:
  - Pressing `Ctrl+C`, navigating to a destination cell, and pressing **`Enter`** now immediately pastes the copied cell/range and clears the dashed copy marquee (while `Ctrl+V` keeps the copy marquee active for repeated pasting).
  - Internal Scratch Sheet copy/paste now carries raw formulas (`rawEditMatrix`) and automatically shifts relative cell references (e.g., `A2` → `A3`, respecting `$A$2` / `$A2` / `A$2` anchors).
- **🟦 Corner Fill Handle Square, Cell Data Types & Date/Number Auto-Increment (`js/scratch-sheets.js`, `js/core-ui.js`, `css/styles.css`, `index.html`)**:
  - Added an interactive Excel-style bottom-right **Fill Handle** square (`.scratch-fill-handle`) on the active selection: drag down/up/left/right (or double-click to auto-fill down adjacent data columns) to copy or extrapolate values and styles across cells.
  - Added a **Cell Data Type / Format** selector (`#scratchCellFormatSelect`: `Type: Auto`, `Type: Text (@)`, `Type: Date (MM/DD/YYYY)`, `Type: Number (123)`, `Type: Currency ($)`) persisted via OpenXML `numFmt`.
  - Dates (`MM/DD/YYYY` / `YYYY-MM-DD`) and Numbers (`<= 10` digits) automatically increment by `+1` day / `+1` (or by the step delta across 2+ selected source cells); explicit **`Type: Text (@)`** cells and `11+` digit tracking numbers never increment, and holding `Ctrl` while dragging forces exact copy without incrementing.
- **⚡ Live `=FILTER` Array Functions & 1-Click Auto-Filter Side Lists (`js/scratch-sheets.js`, `js/core-ui.js`, `css/styles.css`, `index.html`)**:
  - Added an in-app Scratch Sheet formula & dynamic array spill engine (`recalculateScratchWorksheetFormulas`) supporting **`=FILTER`**, **`=UNIQUE`**, **`=SORT`**, **`=IF`**, **`=COUNTIF`**, **`=COUNTA`**, **`=COUNTBLANK`**, **`=SUM`**, **`=AVERAGE`**, **`=MIN`**, **`=MAX`**, **`=ISBLANK`**, **`=ISNUMBER`**, **`=SEARCH`**, **`=FIND`**, **`=TRIM`**, **`=UPPER`**, **`=LOWER`**, **`=LEN`**, **`=LEFT`**, **`=RIGHT`**, **`=MID`**, and **`=TODAY()`** with element-wise boolean `*` (AND) / `+` (OR) broadcasting.
  - Added the **`⚡ Auto-Filter Lists`** toolbar button (`#btnScratchAutoFilter`), context menu action, and **`#scratchAutoFilterModal`** builder (plus a `Title Tracking + Auto-Filtered Red / Yellow / Green Side Lists` starter layout in `Create New Scratch Sheet...`) to generate color-coded **Red** (Empty tracking), **Yellow** (`"Duplicate"`), and **Green** (Occupied / Available tracking) side lists ready for 1-click Outlook/Gmail copy & paste.
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v37'` in [`sw.js`](sw.js) and updated asset version query strings to `?v=1.19.0` in [`index.html`](index.html).

### `v1.18.4` (`main` / `meerkat`)
- **🏷️ Centralized App Version Management (`index.html`, `js/core-ui.js`)**:
  - Added `syncAppVersionDOM()` in `js/core-ui.js` wired to `data-app-version` attributes so `APP_VERSION` (`v1.18.4`) serves as the single source of truth for `document.title`, `#appVersionTag` (text & tooltip), the `Check for Updates & Clear Cache` menu item, User Manual badge, and status bar footer.
- **📊 Excel-Styled HTML Table & Color-Preserving Scratch Sheet Copy/Paste (`js/scratch-sheets.js`, `css/styles.css`)**:
  - Updated Scratch Sheet copy/cut (`Ctrl+C`/`Ctrl+X` and right-click Copy/Cut) to export both `text/plain` (TSV) and a native Excel-compatible `text/html` `<table>` (`serializeScratchSelectionToExcelHtml`, including cell borders, Calibri/Aptos 11pt typography, column widths, bold header shading, cell fill colors via inline CSS + `bgcolor`, and font colors) for clean pasting into Outlook, Gmail, Teams, and Word.
  - Updated Scratch Sheet paste (`Ctrl+V` and right-click Paste) to preserve cell fill colors, font colors, and bold formatting across internal Scratch Sheet copies and external Excel/HTML tables (`parseClipboardHtmlTableToMatrix`, supporting Excel `<style>` class rules as well as inline styles).
- **📋 Integrated Split-Pill Copy-to-Clipboard & Per-Column Mapper Toggles (`index.html`, `css/styles.css`, `js/workbook-io.js`, `js/tracker-table.js`, `js/request-editors.js`)**:
  - Redesigned `.cell-copy-group` + `.btn-inline-copy` into a full-width **split-pill** control sharing a unified hover outline: value text uses the full pill width with `white-space: nowrap` on code pills (preventing 3-digit Yard # wrapping like `340`), while `.btn-inline-copy` overlays the right edge transparently until hovered with an opaque backing (`opacity: 0.92` on pill hover, `1` on direct button hover); clicking the left segment filters Like-Items and clicking the copy overlay copies to clipboard.
  - Enabled split-pill copy buttons across **Email Subject, Ticket #, Master #, Bidder #, Yard #, and Lot #** (each lot token individually filterable and copyable) on the Standard Daily Tracker and `#ticketDetailsModal`—excluding Broker Name, Requested Action, Action Notes, Status, and Tracking ID/UUID.
  - Added per-column **`📋 Copy`** checkboxes inside **Configure Column Mapping (`#schemaMapperModal`)** (`inferDefaultCopyableForCustomColumn`, persisted per workbook/profile via `copyCols`) so users can customize which columns display inline copy pills on Custom Department Trackers.
- **📐 Reclaimed Vertical Screen Space (`index.html`, `css/styles.css`, `js/core-ui.js`, `js/workbook-io.js`)**:
  - Removed the redundant `#filePathsBar` strip beneath `#trackerProfilesBar` (paths and `Auto-Reconnect Folders on Startup` remain accessible inside `⚙️ Tools & Setup ▾` with full-path hover tooltips and in the bottom status bar) to give the tracker table more vertical room.
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v36'` in [`sw.js`](sw.js) and updated asset version query strings to `?v=1.18.4` in [`index.html`](index.html).

### `v1.18.3`
- **🔀 Unified Feature Merge (`main` + `meerkat`) (`index.html`, `css/styles.css`, `js/*`)**:
  - Merged all `main` (`v1.18.1`–`v1.18.2`) enhancements—**Manual Row Copy & Paste**, **Optional Primary Field on Custom Schemas**, **Chronological Left-to-Right Daily Tab Ordering**, **Mid-Month Status vs. Retroactive UUIDv4 Auto-Healing**, and **Clean Validation List Dropdown Sync**—with all `meerkat` enhancements (**Smart Overdue Mutation & `.overdue-outline-red` Alerting**, **`Waiting` & `On Hold` Rollover Whitelist**, **Inline Copy Buttons**, **Multi-Tab Instance Locking & Save Mutex**, and **Template Sort Deprioritization**).
- **🔢 Left-Hand Row Numbers (`index.html`, `js/tracker-table.js`, `js/workbook-io.js`, `css/styles.css`)**:
  - Added read-only Excel-style left-hand line/row numbers (`#` header `.tracker-row-num-th` and `<th scope="row" class="tracker-row-num-cell">`) showing sequential row index `1..N` (with a tooltip displaying the underlying Excel sheet and row number) across both Standard and Custom Department tracker tables, preserving checkbox batch selection, Like-Item filtering, and click-to-edit row behavior.
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v30'` in [`sw.js`](sw.js) and updated asset version query strings to `?v=1.18.3` in [`index.html`](index.html).

### `v1.18.2`
- **🔧 Secondary / Custom Tracker Bug Fixes (`main` — `js/workbook-io.js`, `js/tracker-table.js`, `js/request-editors.js`, `js/rollover.js`)**:
  - **Optional First Field on Custom Department Schemas**: Removed the hardcoded HTML `required` constraint on the first column (`primaryCol`) in `buildDynamicFieldsHtmlForScope` and updated `saveRequestToWorkbook` and `areRowsSameUniqueRequest` so custom department trackers can save and match requests even when the first column is left blank.
  - **Chronological Left-to-Right Daily Tab Ordering (`orderWorkbookSheetsChronologically`)**: Replaced `placeWorksheetFirstInWorkbook` with `orderWorkbookSheetsChronologically`, ordering all `MM-DD-YYYY` daily tabs chronologically from left to right (oldest date on the left → newest date on the right) while keeping all non-date sheets (`Scratch`, `Master Accounts`, `validation lists`, `template`) to their right.
  - **Mid-Month Status Column vs. Retroactive UUIDv4 Auto-Healing**: Fixed the collision when a `Status` column (e.g., Column J) is added midway through the month where earlier sheets had a retroactive UUIDv4 `Tracking ID` in that same column:
    - `findPrimarySchemaHeaderWorksheet` and `findTemplateSheet` now select the daily sheet with the most visible non-tracking columns (`syncWorkbookSheetsWithActiveSchema` syncs `template` and earlier tabs so `Status` in Column J stays visible and `Tracking ID` in Column K stays hidden).
    - `indexWorksheetRows`, `writeItemToWorksheetRow`, `buildConsolidatedRolloverItem`, and `unifyHistoricalGroupTrackingIds` automatically detect any UUIDv4 sitting in `statusCol` (or other visible columns), migrate it to `trackingCol` (Column K), and replace the visible cell with the row's clean status instead of displaying a UUID in the Status column.
  - **Clean Validation List Dropdown Sync & UUID Filtering**: Stopped `loadWorkbookFromDisk` from scraping past sheet values (and retroactive UUIDs) into `validationStatuses`. Added `extractDataValidationOptionsByCol`, `buildCustomSchemaDataValidationsExtLst`, and expanded `<extLst>` / `<dataValidations>` preservation in `captureOriginalArchiveState` and `repairSavedXlsxArchive` so new daily tabs on secondary trackers inherit Excel validation list dropdowns cleanly.
- **🔒 Multi-Instance Cross-Tab Session Lock, Save Mutex & External Disk Guard (`meerkat` — `index.html`, `css/styles.css`, `js/core-ui.js`, `js/workbook-io.js`, `js/scratch-sheets.js`)**:
  - Added browser-wide session coordination (`initMultiTabInstanceLock`) via Web Locks API (`navigator.locks`), `BroadcastChannel`, and `localStorage` heartbeat fallback so opening TrackerUI in a second tab places the duplicate instance into **Standby Mode** (`#multiTabLockBanner`), pauses background auto-saves/mirror retries, and prevents concurrent Excel file race conditions.
  - Added **`⚡ Take Over Active Session Here`** (`#btnTakeOverSession`) and `ensureActiveTabLockBeforeWrite` to safely transfer the lock and reload the latest `.xlsx` from disk before any write, plus automatic lock handoff when the primary tab closes.
  - Serialized `saveWorkbookToDisk` through `saveWorkbookMutexPromise` and added a pre-write `lastLoadedDiskModifiedMs` disk check to warn before overwriting files modified externally.
- **📁 Multi-Day Testing Workbook & Template Sort Deprioritization (`assets/Testing/`, `js/workbook-io.js`)**:
  - Updated `extractMonthYearWeight` so template filenames (`template`, `make a copy`) always sort below working/monthly `.xlsx` workbooks, and preserved `lastSavedWorkbookName` when re-selecting an Agent folder.
  - Generated [`assets/Testing/Agent Bidder Support Tracker - November 2026.xlsx`](assets/Testing/Agent%20Bidder%20Support%20Tracker%20-%20November%202026.xlsx) (and `(VIEW ONLY).xlsx` mirror) with 8 daily tabs (`11-06-2026` through `11-16-2026`), 158 `Master Accounts`, updated `Validation List`, populated `Scratch` sheet, and realistic business test cases covering Smart Overdue Mutation (`> 5 days` & `7+ carries`), `Waiting` / `Waiting on Buyer` / `On Hold` rollovers, `🎯 Grab Ticket`, `📥 Pull Prior`, `🔄 Sync Forward`, and inline copy buttons (excluded from git via `.gitignore`).
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v29'` in [`sw.js`](sw.js) and updated asset version query strings to `?v=1.18.2` in [`index.html`](index.html).

### `v1.18.1`
- **📋 Manual Row Copy & Paste Across Daily Tabs (`main` — `#btnBatchCopyRows`, `#btnBatchCopyToTab`, `#btnPasteTrackerRows`, `js/tracker-table.js`)**:
  - Added manual checkbox-based **Copy & Paste** for daily request rows so entries on secondary/custom department workbooks (even those without a Status column or `validation lists` sheet) can be carried forward to the next day manually.
  - **Two-Step Copy → Switch Tab → Paste**: Check 1+ rows and click **`📋 Copy Row(s)`** in the floating Batch Action Bar (or press `Ctrl+C`), switch to any target `MM-DD-YYYY` tab, and click **`📥 Paste Row(s) (N)`** in the top Controls bar (or press `Ctrl+V`).
  - **One-Step Direct Copy to Next Day / Tab**: Check 1+ rows and click **`📅 Copy to Next Day / Tab...`** in the floating Batch Action Bar to copy and paste selected rows directly to tomorrow's/today's tab (automatically creating the new `MM-DD-YYYY` tab with your custom or standard headers if it doesn't exist yet).
  - Preserves all custom department column values (`customValues`) and each row's hidden UUIDv4 `Tracking ID` across daily tabs, and automatically hides the `Set Status` control group in the Batch Bar when a custom workbook has no Status column.
- **🚨 Smart Overdue Mutation & `.overdue-outline-red` Alerting (`meerkat` — `js/core-ui.js`, `js/rollover.js`, `js/tracker-table.js`, `css/styles.css`)**:
  - Added `evaluateRequestOverdueMetrics` and automated **Overdue** escalation during rollover whenever an active unfinished request is **> 5 days old** or has **carried over 7+ daily tabs**, plus red row/preview outline styling (`.overdue-outline-red`) and `.badge-overdue` pills.
- **⏳ First-Class `Waiting`, `On Hold`, and `Overdue` Statuses (`meerkat` — `js/core-ui.js`, `js/workbook-io.js`, `js/request-editors.js`, `js/rollover.js`, `index.html`, `css/styles.css`)**:
  - Added **`Waiting`**, **`On Hold`**, and **`Overdue`** to `DEFAULT_STATUSES`, `#batchStatusSelect`, `#dockStatus`, `#fStatus`, and dynamic custom status dropdowns (`mergeRequiredOperationalStatuses`), with dedicated `.badge-waiting`, `.badge-progress`, `.badge-escalated`, and `.badge-overdue` badge styling across the Tracker Table, Ticket Details timeline, and View-Only Inspector.
  - Added explicit `ROLLOVER_ACTIVE_STATUS_WHITELIST` / `isRolloverActiveStatus` preservation in `js/rollover.js` so `Waiting`, `Waiting on Buyer`, `On Hold`, and `Overdue` entries always roll forward cleanly.
- **📋 Inline Copy Buttons for Ticket # & Master # (`meerkat` — `js/core-ui.js`, `js/tracker-table.js`, `js/request-editors.js`, `css/styles.css`)**:
  - Added isolated 1-click SVG copy buttons (`.btn-inline-copy` via `renderInlineCopyBtnHtml` and `copyTextToClipboard`) beside **Ticket #** and **Master #** chips in the Tracker Table (including custom schema Ticket/Account columns) and `#ticketDetailsModal`, using `e.stopPropagation()` so copying never triggers Like-Item filtering or row selection.
- **🗺️ Roadmap Consolidation (`FUTURE_FEATURES.md`)**:
  - Merged Gemini Pro backlog proposals into [`FUTURE_FEATURES.md`](FUTURE_FEATURES.md), marked the newly shipped features accordingly, and removed temporary export drafts.
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v28'` in [`sw.js`](sw.js) and updated asset version query strings to `?v=1.18.1` in [`index.html`](index.html).

### `v1.18.0`
- **📖 Interactive In-App User Manual & Reference Guide (`#userManualModal`, `USER_MANUAL.md`, `index.html`, `css/styles.css`, `js/core-ui.js`)**:
  - Added a dedicated **`📖 Manual`** button (`#btnOpenUserManual`) in the top header bar alongside **`User Manual & Reference Guide...`** (`#menuOpenUserManual`) inside **`⚙️ Tools & Setup ▾`**.
  - Built an interactive two-column in-app documentation modal (`#userManualModal`) featuring a sticky 14-section Table of Contents sidebar (`#userManualTocList`) with smooth-scroll navigation and scroll-spy active section highlighting, live instant keyword search (`#userManualSearchInput`), "Where to Click" reference tables, keyboard shortcuts, and 10 embedded high-resolution workspace screenshots (`assets/images/manual/*.png`).
  - Added a full-screen click-to-zoom screenshot lightbox (`#manualImageLightbox`) and a standalone Markdown reference guide ([`USER_MANUAL.md`](USER_MANUAL.md)).
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v27'` in [`sw.js`](sw.js) (precaching `USER_MANUAL.md` and all 10 manual screenshots for offline access) and updated asset version query strings to `?v=1.18.0` in [`index.html`](index.html).

### `v1.17.3`
- **🎨 High-Contrast Active Workbook Profile Tab Hover Fix (`css/styles.css`)**:
  - Fixed an issue where hovering over the active workbook tab (`.wb-profile-tab.active`) caused button text, icons, and badges to blend into the button background and become illegible in both Light and Dark modes.
  - Root cause: generic `button:hover:not(:disabled)` specificity `(0, 2, 1)` overridden the tab's background with `--surface-hover` while retaining `--accent-contrast` text color (`#082f49` on `#2d3d56` in Dark mode; `#ffffff` on `#f1f5f9` in Light mode).
  - Explicitly scoped inactive tab hover (`.wb-profile-tab:hover:not(.active)`), elevated active tab hover specificity (`.wb-profile-tab.active:hover` / `button.wb-profile-tab.active:hover`), and paired `var(--accent-hover)` with `var(--accent-contrast)` for crisp, high-contrast, accessible readability.
  - Applied the same specificity and contrast guard across `.ws-app-tab`, `.sheet-tab-pill`, and `.rollover-source-btn`.
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v26'` in [`sw.js`](sw.js) and updated asset version query strings to `?v=1.17.3` in [`index.html`](index.html).

### `v1.17.2`
- **🗂️ Separated Workbook Tabs Bar vs. In-Workbook Sheet Tabs Bar & Filled Active Tab Colors (`#trackerProfilesBar`, `#workspaceTabsBar`, `css/styles.css`, `js/scratch-sheets.js`)**:
  - Separated Department / Workbook Tracker Tabs into their own top-level **`#trackerProfilesBar`** (`Workbooks` bar above `#filePathsBar`) while keeping **`#workspaceTabsBar`** (`Sheets` bar below `#filePathsBar`) strictly for views inside the active workbook (`Daily Requests` + open `Scratch Sheets`), preserving clear parent-to-child hierarchy.
  - Replaced subtle bottom-underline active tab indicators with unmistakable **filled active tab backgrounds** across both bars (`.wb-profile-tab.active` uses solid `--accent` fill; `.ws-app-tab.active` uses distinct `--row-active` color fill).
- **✕ Acknowledgeable / Closeable Scratch Sheet Formula Warning (`#btnAckScratchFormulaNotice`, `js/scratch-sheets.js`)**:
  - Added an **`Acknowledge ✕`** button on `#scratchFormulaNotice` that dismisses the warning and persists acknowledgment in `localStorage` (`trackerui_scratch_formula_notice_ack_v1`).
- **🌐 Desktop Google Chrome Environment Check (`evaluateDesktopChromeCompatibility`, `checkDesktopChromeEnvironmentAndAlert`, `js/core-ui.js`)**:
  - Added startup browser and viewport mode detection using `navigator.userAgentData` (`brands` & `mobile` Client Hints) with `navigator.userAgent` fallback, firing a native `window.alert(...)` prompt whenever opened outside Desktop Google Chrome (e.g., Firefox, Edge, Safari, Brave, Opera, or Mobile Mode) so users explicitly acknowledge that TrackerUI was built for Chrome in Desktop Mode.
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v25'` in [`sw.js`](sw.js) and updated asset version query strings to `?v=1.17.2` in [`index.html`](index.html).

### `v1.17.1`
- **📐 Polished 12-Column Editor Grid & Lockable Form Layout Customizer (`#btnDockToggleLayout`, `js/request-editors.js`, `css/styles.css`)**:
  - Upgraded the Bottom Editor Dock (`#dockStandardFieldsGrid`, `#dockDynamicFieldsGrid`) and Popup Modal dynamic grid (`#modalDynamicFieldsGrid`) to a baseline-aligned **12-column CSS grid** with single-line label badges (`<span class="field-col-tag">Col X</span>`), eliminating multi-line label wrapping (`Ticket Number` + `🎯 Grab Ticket`), misaligned inputs, and trailing row whitespace, while pinning `.dock-form-actions` (`Save to Workbook`) so it is never clipped behind scrollbars.
  - Smart custom-schema defaults automatically place mid-sheet `Notes` columns (e.g., `Col E` of `A..H`) at the bottom of the form while preserving exact Excel column saves, and auto-balance row column spans (`Subject` / `Inbound Email Subject`, `Ticket Number`, and short ID fields) to fill all 12 columns cleanly.
  - Added a **Locked-by-Default Form Layout Mode** (`🔒 Layout` → `🔓 Lock Layout` + `↺ Reset Layout`): when unlocked, each field displays compact controls to reorder fields (`◀` `▶`), adjust column width (`−` / `+` from `2/12` to `12/12` or drag the right border), and resize textarea height (`↕−` / `↕+` rows `1..6`), persisted per schema in `localStorage` (`trackerui_form_layouts_v1`) and portable `.json` workspace backups.
- **⚠️ Scratch Sheet Values-Only Formula Notice & OpenXML Formula Preservation (`#scratchFormulaNotice`, `js/scratch-sheets.js`, `js/workbook-io.js`)**:
  - Added an explicit notice banner (`#scratchFormulaNotice`) and updated the `Cell` input bar on scratch sheets clarifying that Excel formulas (`=SUM`, `=XLOOKUP`, `=LET`, cross-sheet references, etc.) **do not evaluate or update in-app**, and explaining why Excel 365 formulas display internal OpenXML tokens such as `_xlpm.data` or `_xlfn.` until opened in Desktop Excel.
  - Preserved exact OpenXML `val.formula` prefixes (`_xlfn.`) without partial stripping in `extractScratchCellInfo` and extended external workbook reference (`[1]`) archive preservation across scratch sheets in `repairSavedXlsxArchive`.
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v24'` in [`sw.js`](sw.js) and updated asset version query strings to `?v=1.17.1` in [`index.html`](index.html).

### `v1.17.0`
- **🎯 Grab Ticket Number from Email Subject (`[XXXX-XX-XXXX]`, `js/core-ui.js`, `js/request-editors.js`)**:
  - Added **`🎯 Grab Ticket`** buttons (`#btnDockGrabTicket`, `#btnModalGrabTicket`, and dynamic custom ticket inputs) to extract bracketed hyphenated ticket IDs (e.g., `[ABCD-12-3456]`, `[1234-56-7890]`, including multiple matches) from the Subject, normalize to uppercase, and merge comma-delimited into Ticket Number without duplicates.
- **📥 Sync from Previous Entry ("Pull from Prior Tab", `js/request-editors.js`)**:
  - Added **`📥 Pull Prior` / `📥 Pull from Prior Tab`** (`#btnDockPullPrior`, `#btnModalPullPrior`) to look backward across older `MM-DD-YYYY` tabs (`parseDateTab < currentDate`), match prior instances by UUIDv4 key, Ticket #, composite identity, or Subject, non-destructively populate only currently blank fields, and merge Ticket Numbers comma-delimited without duplicates.
- **🔄 Forward Batch Sync Across Daily Tabs (`js/request-editors.js`)**:
  - Added **`🔄 Sync Forward`** in the Bottom Editor Dock (`#btnDockSyncForward`) and Ticket Details modal (`#btnTicketSyncForward`) to propagate newly added data onto rolled-over instances across newer daily tabs (`parseDateTab > sourceDate`).
  - Strictly non-destructive: fills only blank fields on forward tabs while comma-merging Ticket Numbers (`TICKET-A, TICKET-B`) and presenting a confirmation diff preview before saving.
- **🗂️ Multi-Department Tracker Tabs & Dynamic Column Schema Engine (`#workspaceTabsBar`, `#schemaMapperModal`)**:
  - Generalized `#workspaceTabsBar` into named **Tracker Workspace Profiles** (`＋ Tracker Tab` / `#btnAddTrackerProfile`, double-click to rename, `×` to close secondary tabs), each persisting its own `agentFolderHandle`, `supervisorFolderHandle`, `templatesFolderHandle`, active workbook, and optional `(VIEW ONLY).xlsx` mirror toggle (`#chkEnableMirrorCopy`) in `IndexedDB`.
  - Added automatic Row-1 header discovery (`resolveWorkbookSchemaConfig`) and **`Configure Column Mapping...`** (`#schemaMapperModal`) supporting custom department workbooks with arbitrary column counts/headers, automatic `validation lists` column dropdown binding, hidden immutable UUIDv4 `Tracking ID` column (`maxVisibleCol + 1`), dynamic `#tableHead` / `#listBody` rendering, dynamic Dock & Modal editor grids (`#dockDynamicFieldsGrid`, `#modalDynamicFieldsGrid`), custom-schema New Workbook creation (`custom_cols`), and Smart Rollover support.
- **🛡️ IndexedDB Local Fail-Safe Primary Save Cache (`#btnRetryPrimarySave`, `js/workbook-io.js`)**:
  - When Desktop Excel locks the primary working `.xlsx` file during save, TrackerUI now stages the serialized workbook buffer in `IndexedDB` (`TrackerUI_DirectoryDB`), keeps in-memory edits intact, displays **`💾 Retry Primary Save`** in `#statusBar`, and offers 1-click fail-safe recovery on reload.
- **📊 Supervisor Multi-Workbook Master Rollup, Category Pipeline & Background Web Worker (`#masterRollupModal`, `js/workbook-io.js`)**:
  - Added **`Supervisor Multi-Workbook Rollup...`** (`#menuOpenMasterRollup`) scanning multiple `.xlsx` / `(VIEW ONLY).xlsx` files via a background `Worker` (`parseWorkbookBufferInBackground` with graceful local async fallback), consolidating rolled-over rows per unique request, classifying **Subtype → Parent Category** (`resolveActionParentCategory` + `validation lists` category column), rendering KPI cards + clickable category breakdown pills + unified cross-workbook search/filter table, and exporting multi-sheet `.xlsx` rollups (`#btnExportMasterRollupXlsx`).
- **⚙️ Portable Workspace Config (`.json`) Export & Import (`#btnExportConfigJson`, `#btnImportConfigJson`)**:
  - Added 1-click export and import of workspace preferences, Multi-Department Tracker Profiles, custom column schema overrides, and category mappings as portable `.json` files in **Tools & Setup**.
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v23'` in [`sw.js`](sw.js) and updated asset version query strings to `?v=1.17.0` in [`index.html`](index.html).

### `v1.16.1`
- **Fixed Scratch Sheet `Enter` / Continuous Typing Focus & Auto-Save Interruption (`js/scratch-sheets.js`)**:
  - Fixed a race condition where pressing `Enter` to commit a cell started a 1.2s background auto-save timer (`scheduleScratchAutoSave`) that fired `commitActiveScratchInlineEdit()` *while* typing in the next cell below—destroying the active inline `<input>`, dropping DOM focus to `document.body` during `.xlsx` buffer serialization, and causing subsequent keystrokes to overwrite partial cell text.
  - Background auto-save (`isUserActivelyInteractingWithScratch`) now defers whenever an inline cell editor, formula input, or pointer drag is active, resets its debounce timer on every inline keystroke, tracks `scratchMutationVersion` across async disk writes so in-flight saves never drop newer edits, restores `#scratchGridWrap` focus cleanly, and moves the selection down 1 row (`Shift+Enter` up 1 row) when pressing `Enter` on the grid (`F2` or double-click edits existing text in place).
- **Roadmap Update (`FUTURE_FEATURES.md`)**:
  - Documented architecture and viability plan for **Multi-Department Tracker Tabs & Dynamic Column Schema Engine** (side-by-side department tracker profiles with dynamic Row-1 header discovery and semantic role mapping).
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v22'` in [`sw.js`](sw.js) and updated asset version query strings to `?v=1.16.1` in [`index.html`](index.html).

### `v1.16.0`
- **Scratch Sheet Right-Click Context Menu (`#scratchContextMenu`, `js/scratch-sheets.js`, `css/styles.css`)**:
  - Added an Excel-style right-click context menu on `#scratchGridWrap` supporting **Cut (`Ctrl+X`)**, **Copy (`Ctrl+C`)**, **Paste (`Ctrl+V`)** (with system Clipboard API + internal clipboard fallback), quick Excel default **Fill Color** swatches, **Data to Columns...**, **Clear Contents (`Del`)**, and **Clear Fill Color**.
  - Preserves active multi-cell highlights when right-clicking inside a selection and moves selection when right-clicking outside it (while preserving native browser right-click inside open inline cell text inputs).
- **Excel Default Cell Fill Colors (`#scratchFillMenuWrap`)**:
  - Added a split **Fill** toolbar button (1-click apply of active color + `▾` palette) and right-click swatch picker featuring 24 Microsoft Excel default colors (**Excel Cell Styles & Light Headers** + **Standard & Bold Header Colors**) plus **No Fill**.
  - Reads existing OpenXML solid ARGB and Office Theme/tint fills (`extractScratchCellVisualStyle`), computes contrasting text colors, persists cell fills directly to the `.xlsx` workbook & `(VIEW ONLY)` mirror, and integrates with the 50-step Undo/Redo stack (`Ctrl+Z` / `Ctrl+Y`).
- **App-Only Striped Scratch Rows (`#chkScratchStripedRows`)**:
  - Added an app-only **Striped rows** toggle (persisted in `localStorage`) that renders alternating shaded rows in Light and Dark themes without writing background fills to the Excel file.
- **Data to Columns ("Text to Columns") Wizard (`#textToColumnsModal`)**:
  - Added **Data to Columns** in the scratch toolbar and right-click menu to split delimited text (`Comma`, `Tab`, `Semicolon`, `Pipe`, `Hyphen`, `Space`, or `Other` custom delimiter, with CSV quote handling, consecutive delimiter merging, and whitespace trimming) across adjacent columns with auto-detection, overwrite warnings, live multi-column preview, automatic grid expansion, and 1-step Undo (`Ctrl+Z`).
- **PWA Cache Bump**:
  - Incremented `CACHE_NAME` to `'tracker-ui-v21'` in [`sw.js`](sw.js) and updated asset version query strings to `?v=1.16.0` in [`index.html`](index.html).

### `v1.15.0`
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
