# TrackerUI — Future Features & Strategic Roadmap

This document maintains the roadmap, feature requests, planned enhancements, and architectural backlog for **TrackerUI**.

---

## 1. Planned Enhancements & Active Backlog (📋 Planned)

### 🛡️ Queue #1 — Rolling Daily, Pre-Action & Version-Upgrade Workbook Backups (`_Backups/`)
- **Status**: 📋 Planned *(Queued Next)*
- **Proposal**:
  - Maintain a bounded rolling backup ring buffer inside a configurable **Backups Directory** (defaulting to a `_Backups/` subfolder of the Agent Folder, or persisted in `IndexedDB` + `.json` config via `backupsDirName` / `backupsSubPath`; excluded from normal workbook scanning) so storage never balloons:
    - **1 Daily Snapshot:** Capture one clean `.xlsx` snapshot per workbook on the first load/save of each calendar day (using `SHA-256` content hashing to skip unchanged copies) and prune anything older than the last **5–7 days**.
    - **App-Version-Upgrade Guard:** Automatically capture a safety snapshot before the first write after an `APP_VERSION` update (e.g., `_pre-v1.19.x.xlsx`) to guarantee instant recovery if a TrackerUI update ever introduces a workbook/OpenXML regression.
    - **Pre-Destructive-Action Slots:** Keep a 2–3 slot ring buffer (`_pre-rollover`, `_pre-delete`) before Smart Rollover, Batch Delete, or Day Tab deletion.
    - **In-App Restore UI:** Add **`Restore Workbook from Backup...`** inside **`⚙️ Tools & Setup ▾`** listing available snapshots with timestamps and file sizes for 1-click restoration.

---

### 🆔 Queue #2 — Agent Identity (`Agent Name` + `UUIDv4`) & Hidden `TrackerConfig` Sheet
- **Status**: 📋 Planned *(Queued Next — Promoted from Parked)*
- **Proposal**:
  - **Agent Profile Identity in TrackerUI**: Add an Agent Identity section inside **`⚙️ Tools & Setup ▾`** (or Workspace Settings) where the user can enter/update their **Agent Name** and automatically generate (or view/copy) a persistent **Agent UUIDv4** (`crypto.randomUUID()`), persisted in browser storage and included in `.json` config exports.
  - **Hidden `TrackerConfig` Workbook Sheet**: Create and maintain a hidden (`ws.state = 'hidden'` / `'veryHidden'`) `TrackerConfig` worksheet inside each tracker `.xlsx` workbook (automatically excluded from Daily Tracker tabs, Scratch Sheet tabs, Rollover, and deletion) storing structured key-value rows:
    - `agentName` & `agentUuid` (immutable owner identity for the workbook)
    - Associated workbook-level settings (e.g., `profileName`, `schemaVersion` / `appVersion`, active column mapping overrides, custom form layout JSON, and backup/mirror preferences) so critical tracker settings travel directly with the `.xlsx` file even across devices or browsers.
  - **Auto-Sync on Load/Save**: When opening a workbook with an existing `TrackerConfig` tab, read its metadata safely without overwriting local unsaved edits, and sync updated identity/workbook settings back to `TrackerConfig` on save.

---

### 📂 Queue #3 — Team Root Directory Auto-Scaffolding & Startup `.json` Config
- **Status**: 📋 Planned *(Queued Next)*
- **Proposal**:
  - **1-Click Team Root Auto-Scaffolding (OneDrive → SharePoint Portable)**: Once a user grants `showDirectoryPicker({ mode: 'readwrite' })` on a single synced **Team Root Folder** (OneDrive or a locally synced SharePoint Document Library), TrackerUI automatically creates the canonical folder hierarchy via `getDirectoryHandle(name, { create: true })` so migrating storage locations requires zero manual folder setup:
    - `_Config/` (team presets, schema contract, startup `trackerui-config.json`)
    - `_Templates/` (master monthly `.xlsx` templates)
    - `_ViewOnly_JSON/` (fast per-agent JSON mirrors: `mirror_<agentUuid>.json`)
    - `_SharedQueue/` (conflict-free per-ticket queue & claim files)
    - `Agents/<AgentName>/` (primary working `.xlsx` with hidden `TrackerConfig`, `_Backups/`, and `_State/`)
  - **Workspace `.json` Config & Startup Auto-Check**: Extend `Export Workspace Config (.json)` / `Import Workspace Config (.json)` to store **Agent Name**, **Agent UUIDv4**, and per-profile folder names/relative subpaths (`agentSubPath`, `supervisorSubPath`, `templatesSubPath`, `backupsSubPath`, `lastWorkbookName`), with an automatic startup check (`_Config/trackerui-config.json`) that resolves subfolders from the granted Team Root handle or displays a guided 1-click **"Connect Expected Folder"** prompt.

---

### 💾 Queue #4 — JSON View-Only Mirrors, Shared Request Queue & NoSQL Chrome-Wipe Recovery (Separate Aggregator Repo)
- **Status**: 📋 Planned *(Queued After Queues #1–#3)*
- **Proposal**:
  - **Per-Agent `.json` View-Only Mirrors (`_ViewOnly_JSON/mirror_<agentUuid>.json`)**:
    - Replace legacy `(VIEW ONLY).xlsx` mirror writes with fast (~2–5ms), lock-free per-agent `.json` mirrors containing normalized rows, `agentUuid`, `agentName`, `schemaVersion`, `appVersion`, `lastAccessed`, `carryCount`, and active schema metadata.
    - Keep `(VIEW ONLY).xlsx` behind an optional **"Legacy Excel Mirror"** toggle during migration so supervisors have zero downtime while the standalone Aggregator is built.
  - **Conflict-Free Shared Request Queue (`_SharedQueue/`)**:
    - Avoid multi-writer OneDrive/SharePoint sync collisions by using a file-per-record NoSQL layout: `_SharedQueue/items/<trackingUuid>.json` for shared/queued tickets and `_SharedQueue/claims/<trackingUuid>__<agentUuid>.json` for agent claims (resolved deterministically by earliest `claimedAt` + `agentUuid` tiebreaker if two agents claim inside the same sync window).
  - **Private Browser State NoSQL JSON Export/Restore (`_State/state_<agentUuid>.json`)**:
    - Full document-style JSON snapshot of `IndexedDB` metadata + `localStorage` state (agent identity, tracker profiles, column mappings, form layouts, action category overrides, row last-accessed timestamps, Row-1 unbold preferences, open scratch tabs, UI theme/font scale, and folder subpath descriptors) to recover cleanly after a Chrome site-data or profile wipe.
  - **Separate Aggregator & DB Manager Project + Shared `CONTRACT.md`**:
    - Build the **Supervisor Aggregator & DB Manager** in its own separate project and Git repository to isolate reporting/analytics changes from the production agent client.
    - Maintain a concise, versioned **`CONTRACT.md`** (defining canonical folder paths, `schemaVersion` payloads like `trackerui.mirror.v1` / `trackerui.queue.v1` / `trackerui.state.v1`, read/write ownership boundaries, and a brief contract-change log) so both projects stay mutually aware without needing to read each other's codebases.

---

---

| Feature / Suggestion | Subsystem | Status | Description & Architecture |
|---|---|---|---|
| **1. Rolling Daily, Pre-Action & Version-Upgrade Backups (`_Backups/`)** | *Workbook I/O / Safety* | 📋 Planned *(Queued #1)* | Capped 5–7 day `.xlsx` ring buffer + `SHA-256` dedup + automatic pre-version-upgrade & pre-rollover/delete safety snapshots with an in-app **Restore from Backup** picker. |
| **2. Agent Identity (`Agent Name` + `UUIDv4`) & Hidden `TrackerConfig` Sheet** | *Identity / Workbook I/O* | 📋 Planned *(Queued #2)* | Enter Agent Name & generate persistent `UUIDv4` in TrackerUI; persist identity & workbook-level settings in a hidden `TrackerConfig` Excel sheet so settings travel with the workbook. |
| **3. Team Root Auto-Scaffolding & Startup `.json` Config** | *Config / Onboarding* | 📋 Planned *(Queued #3)* | Auto-create canonical OneDrive/SharePoint folder tree from a single Team Root handle (`{ create: true }`), store agent identity + subpaths in `.json`, and auto-check `_Config/trackerui-config.json` on startup. |
| **4. JSON Mirrors, Shared Queue, Wipe Recovery & Separate Aggregator Repo** | *Sync / Persistence / Multi-Repo* | 📋 Planned *(Queued #4)* | Lock-free per-agent `.json` mirrors (`_ViewOnly_JSON/`), conflict-free per-ticket `_SharedQueue/`, Chrome-wipe `IndexedDB`+`localStorage` JSON snapshots (`_State/`), and a separate Aggregator repo synced via `CONTRACT.md`. |
| **Custom Ticket ID Schemas** | *Core UI / Editors* | 📋 Planned | Settings input to define custom ticket formats via Regular Expressions, removing hardcoded validations for greater queue flexibility. |
| **In-Cell Validation Dropdowns** | *Scratch Sheets* | 📋 Planned | Map IndexedDB validation lists to custom columns using isolated `<select>` nodes that bypass the grid's multi-cell selection event bubbling. |
| **Charts & Visual Analytics Panel (Phase 1 & 2)** | *Daily Tracker / Scratch Sheets / Supervisor Rollup* | 📋 Planned | Native SVG (zero-dependency) or offline-cached Chart.js visual cards for Daily Request Volume/Status breakdown, Title Tracking Red/Yellow/Green bucket progress, and Multi-Workbook Supervisor Category trends, with optional image/HTML export. |
| **Supervisor Dashboard Auto-Refresh** | *Supervisor View* | 📋 Planned | A 3-to-5-minute active polling loop (`setInterval`) to traverse the directory and re-parse the master aggregated list, keeping the dashboard current. |

---

## 2. Newly & Recently Implemented Features (`v1.17.0` – `v1.19.10`)

### 🟠 Orange Pulsing Update Alert, Completed Green Rows, Toggle Switches, Compact Headers, Duplicate UUID Reuse & `TrackerConfig` Guard (`v1.19.10`)
- **Status**: ✅ Implemented in `v1.19.10`
- **Delivered**:
  - **Orange Pulsing Update Alert (`index.html`, `css/styles.css`, `js/core-ui.js`)**: Friendly warm-orange pulsing version badge (`#appVersionTag.update-available`), Tools menu indicator, and orange pulsing `#pwaUpdateBanner` so agents notice available updates without alarm-red styling.
  - **Completed Green Rows & "Hiding Completed" Switch (`index.html`, `css/styles.css`, `js/core-ui.js`, `js/tracker-table.js`, `js/rollover.js`, `js/workbook-io.js`)**: Added `.completed-outline-green` (`--success` perimeter outline & `--success-soft` green background) on completed entries and Rollover target previews, plus a persistent **`Hide Completed`** / **`Hiding Completed`** sliding pill switch (`role="switch"`) displaying `"Hiding Completed (N)"` when active.
  - **Compact `Master # / Bidder #` Header, Status-Folded `Accessed` Time & "Hide ID" Switch (`index.html`, `css/styles.css`, `js/core-ui.js`, `js/tracker-table.js`, `js/workbook-io.js`)**: Dropped `"Broker "` from Col D/E table headers, stacked `Master #` / `Bidder #` sort controls vertically in a compact column (`.col-master-bidder`), folded transparent corner Last Accessed timestamps directly beneath Status badges inside `.tracker-status-cell` (with `Accessed ↕` stacked in `th.tracker-status-th`), and added a persistent **`Hide ID`** / **`Hiding ID`** switch (`#btnToggleHideUuid`) that collapses the bottom UUID bar completely when active.
  - **Duplicate Request Check & Shared UUIDv4 Assignment (`js/request-editors.js`)**: Allows duplicate/follow-up request entries (e.g., duplicate email subjects) to be saved normally while automatically reusing the existing matching request's UUIDv4 Tracking ID (`findExistingMatchingRequestForNewEntry()`).
  - **1-Click Copyable UUIDv4 Tracking IDs (`js/tracker-table.js`, `js/request-editors.js`, `css/styles.css`)**: Integrated split-pill copy button next to each row's UUIDv4 chip plus 1-click copy on `#dockKeyBadge`, `#modalKeyDisplay`, and `#ticketDetailsKeyBadge`.
  - **Resizable Tracker Table Columns (`index.html`, `css/styles.css`, `js/core-ui.js`, `js/tracker-table.js`, `js/workbook-io.js`)**: Draggable right-edge resize handles + double-click auto-fit on `#tableHead` across Standard & Custom Department schemas, persisted per profile/schema in `localStorage` and `.json` config export/import, plus **Reset Table Column Widths** in Tools & Setup.
  - **Forward-Compat `TrackerConfig` Protection (`js/core-ui.js`)**: Added `TrackerConfig` / `Tracker Config` to `isProtectedReferenceSheetName()` so older cached clients safely ignore upcoming hidden `TrackerConfig` sheets.

### ✍️ Simple Text Editing & Cell Formatting Controls for Scratch Sheets (`v1.19.9`)
- **Status**: ✅ Implemented in `v1.19.9`
- **Delivered**:
  - Added **Bold (`Ctrl+B`)**, **Italic (`Ctrl+I`)**, **Underline (`Ctrl+U`)**, **Text Alignment (`Left`, `Center`, `Right`)**, and **Clear Text Formatting / Unbold** controls to both the Scratch Sheet toolbar (`#scratchTextFormatGroup`) and `#scratchContextMenu` (`.scratch-ctx-format-section`).
  - Persists `cell.font` (`bold`, `italic`, `underline`) and `cell.alignment.horizontal` through 50-step Undo/Redo, HTML clipboard copy/paste (`serializeScratchSelectionToExcelHtml` / `parseClipboardHtmlTableToMatrix`), and `.xlsx` dual-write saves, including an explicit Row-1 Bold/Unbold override (`trackerui_scratch_row1_unbold_v1` / `.unbold-cell`).

### 🧱 Modular Scratch Sheet & Spreadsheet Engine Split (`v1.19.8`)
- **Status**: ✅ Implemented in `v1.19.8`
- **Delivered**:
  - Extracted the pure OpenXML styling, cell type/date parsing, `applyRawEditToExcelCell()`, `extractScratchCellInfo()`, `measureScratchWorksheetBounds()`, formula AST tokenizer & dynamic array spill evaluator (`recalculateScratchWorksheetFormulas()`), fill-handle series extrapolation, RFC-4180 TSV / Excel HTML `<table>` clipboard matrix parsers & serializers, and delimited line tokenizer into [`js/spreadsheet-engine.js`](js/spreadsheet-engine.js), loaded right before [`js/scratch-sheets.js`](js/scratch-sheets.js).
  - Kept UI rendering, pointer range selection, inline editing, Undo/Redo, per-profile runtime snapshots, fill palette UI, context menu, and modals in [`js/scratch-sheets.js`](js/scratch-sheets.js), while extending style snapshots with `alignment`, `italic`, `underline`, and `buildStyleSnapshotForTextFormatting()` in preparation for Scratch Sheet formatting controls.

### 🧩 Comprehensive Full-Column Schema Mapping & Custom Field Validation (`v1.19.7`)
- **Status**: ✅ Implemented in `v1.19.7`
- **Delivered**:
  - **Unlimited Full-Column Schema Manager (`index.html`, `css/styles.css`, `js/core-ui.js`, `js/workbook-io.js`)**: Redesigned `#schemaMapperModal` into a unified table listing every workbook column (across both Standard and Custom Department schemas) with **Table** / **Editor** visibility toggles, custom **Display Labels** (non-destructive to Excel Row-1 headers), **Semantic Role** assignment (adding `Requested Action`), **Data Type**, **Dropdown / Validation Source**, **Input Control (`<datalist>` vs. `<select>`)**, and **`📋 Copy`** toggles, plus column search and Select-All controls.
  - **End-to-End Standard & Custom Visibility, Labels & Validation Wiring (`js/tracker-table.js`, `js/request-editors.js`)**: Wired custom column visibility, display labels, strict `<select>` mirrors, extra Standard columns (e.g., Col K `Sent Email Subject`), and non-destructive hidden-column preservation across the Daily Table, Bottom Dock, Popup Editor, and Ticket Details modal.

### 🛡️ Comprehensive Defensive Code Audit Fixes, Advanced Column Settings & Yard # Dropdown (`v1.19.6`)
- **Status**: ✅ Implemented in `v1.19.6`
- **Delivered**:
  - **Advanced Column Mapping Settings (`index.html`, `css/styles.css`, `js/workbook-io.js`, `js/request-editors.js`)**: Added the expandable **`⚙️ Advanced Column Settings ▾`** table inside `#schemaMapperModal` allowing per-column **Data Type** (`Auto`, `Uppercase Code / ID`, `Number`, `Currency`, `Date`, `Multi-line Notes`), **Dropdown / Validation Source** (`Auto-Detect`, `None`, `Yard # / Name List`, `Master Accounts`, `Validation Lists`, or custom comma-separated lists), **Input Control** (`Searchable Input + Dropdown` vs. `Strict <select>`), and **`📋 Copy`** toggles across all columns.
  - **Standalone `Yard # / Yard Name` Dropdown & Strict `Broker ↔ Master #` Decoupling (`index.html`, `js/mockMasterAccounts.js`, `js/workbook-io.js`, `js/request-editors.js`)**: Decoupled `Bidder #` and `Yard #` from `Broker Name` ↔ `Master #` auto-fill, and wired `<datalist id="yardList">` to the existing `Yard #` (`Col G`) field so linked `Yard List` tabs or `validation lists` columns provide searchable `Yard # — Yard Name` autocomplete without adding a column.
  - **String Normalization, Casing Deduplication & Formula Parity (`js/core-ui.js`, `js/workbook-io.js`, `js/request-editors.js`, `js/scratch-sheets.js`)**: Normalized `actionCategoryOverridesStore` keys, case-insensitively deduplicated `validation lists` and status `<select>` options, merged Master Rollup parent categories case-insensitively, and added `==`, matrix `=IF` false-branches, Date-vs-String comparisons, and `"="` / `"<>"` `=COUNTIF` shorthand.
  - **Tiered Priority Role Inference, UUIDv4 Scrubbing & 11-Col Validation Bootstrap (`js/workbook-io.js`, `js/request-editors.js`, `index.html`)**: Added 3-tier priority matching + `usedSet` so `"Title Status"` never collides with `Subject`, added `buildStandardSchemaDataValidationsExtLst()` (`J2:J500` & `F2:F500`), gated Column C `XLOOKUP` XML rewriting, and scrubbed UUIDv4 strings from visible columns in View-Only Inspector and Master Rollup.
  - **Scratch `Type: Text (@)` Formula Guard, Non-Blocking Auto-Save Lock Toast & Per-Profile Scratch State (`js/core-ui.js`, `js/workbook-io.js`, `js/scratch-sheets.js`, `js/rollover.js`)**: Checked `Type: Text (@)` before `=` formula parsing, stripped `<rPh>` phonetic XML tags, made background scratch auto-save locks non-blocking (`silentLockAlert`), isolated open Scratch Sheets & Undo/Redo stacks per Tracker Profile (`captureScratchRuntimeSnapshot` / `restoreScratchRuntimeSnapshot`), and fixed Rollover target-tab `New → In Progress` & superseded force-reopen preservation.

### 🕒 Row Last-Accessed Timestamp, Compact `#`/Checkbox Columns & Unsaved New Request Draft Autosave/Warning (`v1.19.5`)
- **Status**: ✅ Implemented in `v1.19.5`
- **Delivered**:
  - **Sortable `Accessed` Column (`js/core-ui.js`, `js/tracker-table.js`)**: Records and displays when each request row was last clicked/opened (persisted per Tracking ID in `localStorage` without altering the `.xlsx` schema), with live DOM updates and full timestamp tooltips.
  - **Compact `#` & Checkbox Columns (`css/styles.css`)**: Squished row-number and checkbox columns to minimal width with a `1px` right separator on the checkbox column.
  - **New Request Draft Autosave & Overwrite Guard (`index.html`, `js/request-editors.js`, `js/tracker-table.js`)**: Auto-saves in-progress new requests in `localStorage` (`#dockDraftBadge`, `#btnDockResumeDraft`) and warns before clicking another row overwrites an unsaved new request.

### 🔄 Always-Live Bidirectional Master Buyer Account Sync, Rollover Status Rules & Scratch Sheet Workflow Upgrades (`v1.19.4`)
- **Status**: ✅ Implemented in `v1.19.4`
- **Delivered**:
  - **Always-Live Bidirectional Master Buyer Account Sync (`js/workbook-io.js`)**: Entering or switching a known `Broker Name` or `Master #` immediately updates the paired field and default `Bidder #` without requiring fields to be cleared first.
  - **Rollover Status Transitions & Waiting/Hold Exemption (`js/core-ui.js`, `js/rollover.js`)**: Automatically transitions `New` → `In Progress` when carrying forward to a subsequent daily tab, and exempts `Waiting`, `Waiting on Buyer`, and `On Hold` from auto-`Overdue` aging mutation.
  - **Scratch Sheet Email `<table>` Paste, Transpose, `Auto-increment` Toggle & Newline / Split-Down (`index.html`, `js/scratch-sheets.js`, `css/styles.css`)**: Prioritizes `text/html` `<table>` reconstruction (preserving `<br>` / `<p>` multi-line cells, `colspan`/`rowspan`, and styles), adds `Paste Special → Transpose` & `Transpose Selection`, adds the explicit `#chkScratchAutoIncrement` toolbar checkbox (supporting 11+ digit `BigInt` series stored as text), and adds `Newline (\n)` + `Split Across / Split Down` to `#textToColumnsModal`.

### 📐 1080p Screen-Space Optimization & Overdue / Waiting Row Backgrounds (`v1.19.1`)
- **Status**: ✅ Implemented in `v1.19.1`
- **Delivered**: Added `@media (max-width: 1920px), (max-height: 1080px)` CSS rules hiding `#sheetTabsBar` and compacting vertical toolbar/table/dock spacing on 1080p and smaller screens (while retaining `#sheetTabsBar` on ultrawide monitors), updated `overdue-outline-red` table rows with `var(--danger-soft)` red fills, and added `waiting-outline-purple` (`var(--waiting-soft)` + `var(--waiting)`) styling for **Waiting / Waiting on Buyer / On Hold** rows in both the main table and Rollover preview pane.

### ⚡ Scratch Sheet Excel Fill Handle, Cell Types, Copy+Enter Paste & Live `=FILTER` Array Lists (`v1.19.0`)
- **Status**: ✅ Implemented in `v1.19.0`
- **Delivered**:
  - **Copy + `Enter` to Paste**: `Ctrl+C` → move to destination → `Enter` pastes and clears the copy marquee (`Ctrl+V` retains marquee for repeated pasting), with relative formula reference shifting.
  - **Cell Data Types / Formats (`#scratchCellFormatSelect`) & Corner Fill Handle (`.scratch-fill-handle`)**: Supports `Auto`, `Text (@)`, `Date (MM/DD/YYYY)`, `Number (123)`, and `Currency ($)` (`numFmt`), dragging or double-clicking the corner fill square to copy or auto-increment Dates & Numbers (`+1` or multi-cell delta; `Ctrl` or `Text (@)` skips incrementing).
  - **Live `=FILTER` Dynamic Array Engine & Auto-Filter Side Lists (`#scratchAutoFilterModal`)**: Evaluates `=FILTER`, `=UNIQUE`, `=SORT`, `=IF`, `=COUNTIF`, `=COUNTA`, `=SUM`, etc. live in-app with spill bounds and 1-click Red (Empty) / Yellow (`"Duplicate"`) / Green (Occupied) Title Tracking side lists.

### 🔢 Left-Hand Row Numbers (`v1.18.3` on `meerkat`)
- **Status**: ✅ Implemented in `v1.18.3`
- **Delivered**: Injected read-only Excel-style `<th scope="row" class="tracker-row-num-cell">` line numbers (`1..N`) along the left edge of the Daily Tracker table across both Standard and Custom Department views (showing underlying Excel row and sheet name on hover), preserving batch checkbox selection, Like-Item filtering, and click-to-edit row interactions.

---

### 🔒 Multi-Tab Instance Locking, Save Mutex & External Modification Guard (`v1.18.2` on `meerkat`)
- **Status**: ✅ Implemented in `v1.18.2`
- **Delivered**: Added cross-tab session coordination (`initMultiTabInstanceLock` using Web Locks API `navigator.locks` + `BroadcastChannel` + `localStorage` heartbeat fallback) in `js/core-ui.js`, placing duplicate tabs into **Read-Only Standby Mode** (`#multiTabLockBanner`, pausing background auto-saves and mirror retries) with a 1-click **`⚡ Take Over Active Session Here`** button (`#btnTakeOverSession`), plus an in-tab async `saveWorkbookMutexPromise` queue and `lastLoadedDiskModifiedMs` external file modification guard in `js/workbook-io.js` and `js/scratch-sheets.js`.

---

### 🔧 Secondary / Custom Tracker Auto-Healing & Chronological Ordering (`v1.18.2` on `main`)
- **Status**: ✅ Implemented in `v1.18.2`
- **Delivered**:
  - Removed mandatory first field requirement on custom schemas;
  - Added chronological left-to-right daily tab ordering (`orderWorkbookSheetsChronologically`);
  - Fixed mid-month Status column vs. retroactive UUIDv4 collision by auto-migrating legacy UUIDs to Column K and restoring clean statuses;
  - Cleaned validation list dropdown synchronization and prevented UUID scraping into status dropdowns.

---

### 🚨 Smart Overdue Mutation & `.overdue-outline-red` Alerting (`v1.18.1` on `meerkat`)
- **Status**: ✅ Implemented in `v1.18.1`
- **Delivered**: Added first-class **`Overdue`** status badges (`.badge-overdue`) and automated aging evaluation (`evaluateRequestOverdueMetrics` in `js/core-ui.js`, `js/rollover.js`, `js/tracker-table.js`). Unfinished active items that are **> 5 days old** or **carried over 7+ times** automatically update to `"Overdue"` during rollover and render a red row/preview outline (`.overdue-outline-red`).

---

### ⏳ "Waiting" & "On Hold" Rollover Whitelist (`v1.18.1` on `meerkat`)
- **Status**: ✅ Implemented in `v1.18.1`
- **Delivered**: Added **`Waiting`** and **`On Hold`** to `DEFAULT_STATUSES`, status dropdowns (`#batchStatusSelect`, `#dockStatus`, `#fStatus`, dynamic custom status columns, and `mergeRequiredOperationalStatuses`), distinct purple/violet status badges (`.badge-waiting`), and explicit `ROLLOVER_ACTIVE_STATUS_WHITELIST` / `isRolloverActiveStatus` preservation in `js/rollover.js` so Waiting / Waiting on Buyer / On Hold items are never dropped.

---

### 📋 Inline Copy Buttons (`v1.18.1` → Enhanced Split-Pill in `v1.18.3`)
- **Status**: ✅ Implemented in `v1.18.1` / Upgraded in `v1.18.3`
- **Delivered**: Injected isolated inline SVG copy icon buttons (`.btn-inline-copy` via `renderInlineCopyBtnHtml` & `copyTextToClipboard`) styled as an **integrated split-pill** (`.cell-copy-group`: clicking the left value segment filters; clicking the integrated right copy segment copies to clipboard, sharing one unified hover outline) across **Email Subject, Ticket #, Master #, Bidder #, Yard #, and Lot #** on the Standard Daily Tracker (excluding Broker Name, Requested Action, Action Notes, Status, and Tracking ID) and `#ticketDetailsModal`, plus per-column **`📋 Copy`** checkboxes inside **Configure Column Mapping (`#schemaMapperModal`)** for Custom Department Trackers.

---

### 📋 Manual Row Copy & Paste Across Daily Tabs (`v1.18.1` on `main`)
- **Status**: ✅ Implemented in `v1.18.1`
- **Delivered**: Manual checkbox-based row copy & paste (`#btnBatchCopyRows`, `#btnBatchCopyToTab`, `#btnPasteTrackerRows`) allowing 1-click or two-step copying of requests across daily tabs on any tracker profile.

---

### 📖 Interactive In-App User Manual & Reference Guide (`v1.18.0`)
- **Status**: ✅ Implemented in `v1.18.0`
- **Delivered**: Interactive two-column documentation modal (`#userManualModal`, `USER_MANUAL.md`) with live search, sticky Table of Contents, 10 high-resolution screenshots, and lightbox viewer.

---

### 🎯 Grab Ticket Number from Email Subject (`[XXXX-XX-XXXX]`)
- **Status**: ✅ Implemented in `v1.17.0`
- **Delivered**: Manual `🎯 Grab Ticket` button (`#btnDockGrabTicket`, `#btnModalGrabTicket`, and dynamic custom ticket fields) extracts bracketed hyphenated ticket patterns from Subject, normalizes to uppercase, and merges comma-delimited without duplicates.

---

### 🔄 Forward Batch Sync Across Daily Tabs
- **Status**: ✅ Implemented in `v1.17.0`
- **Delivered**: `🔄 Sync Forward` button in `#bottomEditorDock` and `#ticketDetailsModal` (`handleForwardBatchSync`). Propagates newly added details forward across newer `MM-DD-YYYY` tabs non-destructively with confirmation diff preview.

---

### 📥 Sync from Previous Entry ("Pull from Prior Tab")
- **Status**: ✅ Implemented in `v1.17.0`
- **Delivered**: `📥 Pull Prior` / `📥 Pull from Prior Tab` button in `#bottomEditorDock` and `#editModal` (`handlePullFromPriorTab`). Scans earlier `MM-DD-YYYY` tabs for matching requests and fills blank fields while comma-merging Ticket Numbers.

---

### 🗂️ Multi-Department Tracker Tabs & Dynamic Column Schema Engine
- **Status**: ✅ Implemented in `v1.17.0`
- **Delivered**: Multiple Tracker Profiles in `#trackerProfilesBar`, dynamic Row-1 header discovery & role mapper (`#schemaMapperModal`), and dynamic table/editor grid generation.

---

## 3. Completed Architectural Backlog (`v1.17.0`)

| Feature / Suggestion | Source | Status | Implementation Summary |
|---|---|---|---|
| **Dynamic Schema Mapping Configuration** | `TrackerUI Architectural Recs` | ✅ Implemented in `v1.17.0` | Full Row-1 header discovery (`resolveWorkbookSchemaConfig`), semantic role mapper modal (`#schemaMapperModal`), series-aware override persistence (`getWorkbookSeriesKey`), and dynamic table/form generation. |
| **IndexedDB Local Fail-Safe Cache** | `TrackerUI Architectural Recs` | ✅ Implemented in `v1.17.0` | Stages serialized `.xlsx` buffers in `IndexedDB` (`stagedPrimarySave::*`) whenever Desktop Excel locks the primary working file, keeps in-memory edits live, surfaces `💾 Retry Primary Save` (`#btnRetryPrimarySave`), and prompts recovery on reload. |
| **Supervisor Pipeline: Subtype → Parent Type Lookup** | `Supervisor Pipeline (Phase 2)` | ✅ Implemented in `v1.17.0` | `resolveActionParentCategory()` maps action subtypes via custom JSON overrides, `validation lists` category columns, and built-in operational category rules across `#ticketDetailsModal` and `#masterRollupModal`. |
| **Web Worker Background ExcelJS Parsing** | `Deep Processing (Phase 3)` | ✅ Implemented in `v1.17.0` | `parseWorkbookBufferInBackground()` offloads multi-workbook `.xlsx` parsing to a dedicated `Blob` Web Worker importing ExcelJS (with automatic async main-thread fallback if blocked/offline). |
| **Multi-File Master Tracker Aggregation Loop** | `Deep Processing (Phase 3)` | ✅ Implemented in `v1.17.0` | `#masterRollupModal` (`#menuOpenMasterRollup`) scans multiple `.xlsx` / `(VIEW ONLY).xlsx` workbooks in a folder, deduplicates rolled-over requests per unique ID, renders KPI & category breakdowns, and exports `.xlsx` rollups (`#btnExportMasterRollupXlsx`). |
| **External Config File (`.json`)** | `TrackerUI Architecture` | ✅ Implemented in `v1.17.0` | `Export Workspace Config (.json)` (`#btnExportConfigJson`) and `Import Workspace Config (.json)...` (`#btnImportConfigJson`) in **Tools & Setup**. |

---

---
## 4. Parked & Deferred Scope

| Concept | Status | Rationale |
|---|---|---|
| **Cross-Queue Ticket Claiming (Tombstoning)** | ⏸️ Parked | Manual dual-write transfer of tickets between agent workbooks. Intentionally parked to prioritize core workflow stability before introducing cross-file mutation. |
| **Hidden UUID Configuration Sheets (`TrackerConfig`)** | 📋 Promoted to Section 1 | Promoted to **Queue #2** in Section 1 for Agent Name, Agent UUIDv4, and associated workbook-level settings stored directly in a hidden `TrackerConfig` tab. |
| **Nested Agent JSON Config Storage** | ⏸️ Parked | Saving a `settings.json` file inside a nested agent directory for state persistence. Parked for now (manual `.json` export/import is implemented). |
| **Mass Report Filter (380,000-Row Client-Side `Set` Filter)** | ⏸️ Parked | Explicitly parked. High memory overhead; requires `{ header: 1 }` stream parsing to avoid V8 memory crashes. |
| **Smart Scratchpad (Real-Time Regex Lot # & Title Extractor)** | ⏸️ Parked | Split-pane scratchpad with automated regex scraping for 8-digit lot numbers. Parked in favor of Excel-like in-app scratch sheet tabs (`v1.15.0`). |
| **Chart.js Metric View for Current Data** | 📋 Promoted to Section 1 | Superseded by the **Charts & Visual Analytics Panel (Phase 1 & 2)** proposal in Section 1 above (preferring lightweight zero-dependency SVG charts first, with optional Chart.js UMD if needed). |

---

## 5. Architecture Decisions & Decided Against

| Decision | Date | Reason / Alternative Chosen |
|---|---|---|
| **Excluding `Master Accounts` Tab from New Monthly Workbooks** | `2026-11` | Creating a new monthly workbook no longer requires or creates a `Master Accounts` sheet. Bidirectional Broker ↔ Master # auto-fill automatically learns from existing entries across daily tabs. |
| **Pure Immutable UUIDv4 vs. Ticket # Primary Key** | `2026-10` | Ticket numbers are optional metadata that may be omitted or updated later. Hidden Column K/L strictly holds a pure 36-character UUIDv4 (`crypto.randomUUID()`) as the immutable row key. |
| **Directory-Based I/O (`showDirectoryPicker`) vs. File Picker** | `2026-10` | Directory handles allow automatic monthly workbook discovery, template cloning, and background dual-write mirroring (`<Workbook> (VIEW ONLY).xlsx`) without repeated file-dialog prompts. |
| **Zero-Build PWA Structure vs. Bundlers / Frameworks** | `2026-10` | No Vite, Webpack, npm, or React. Keeps the app directly runnable on `localhost` (VS Code Live Server) and GitHub Pages (`https://`). |
