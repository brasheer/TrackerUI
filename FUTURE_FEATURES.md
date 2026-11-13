# TrackerUI — Future Features & Strategic Roadmap

This document maintains the roadmap, feature requests, planned enhancements, and architectural backlog for **TrackerUI**.

---

## 1. Planned Enhancements (📋 Planned)

### 📋 Custom Field Validation Assignment
- **Status**: 📋 Planned
- **Proposal**:
  - Allow users to explicitly assign validation dropdowns and list sources to any desired field/column in both Custom Department and Standard tracker profiles.
  - Rather than relying solely on automated header substring matching against the `validation lists` sheet or existing in-sheet XML, provide a configuration UI (e.g., inside Column Mapping / Schema Setup) where users can select which columns get dropdown validation, choose the source range or validation sheet column (or specify a custom comma-separated list of options), and have those validations reflected in both the in-app forms/grids (`<select>` / `<datalist>`) and written into the `.xlsx` worksheet data validations (`<x14:dataValidations>` / `<dataValidations>`).

---

### 🧩 Comprehensive Full-Column Schema Mapping
- **Status**: 📋 Planned
- **Proposal**:
  - Expand the Column Mapping modal (`#schemaMapperModal`) to manage and configure **all** columns in a workbook, eliminating the current constraint where only 6 semantic roles (Primary/Subject, Ticket #, Secondary/Account, Requested Action, Notes, Status) can be mapped.
  - Provide a full interactive column list for the active workbook showing every detected column (Col A, Col B, Col C, etc.), allowing users to inspect, re-label, assign semantic roles, set visibility/hidden flags, configure custom validation sources, and assign field types across all columns without limit.

---

### 🔍 1-Click Copyable UUIDv4 Tracking IDs
- **Status**: 📋 Planned
- **Proposal**:
  - Make the UUIDv4 Tracking ID chips in the table (`🔍 8f4a92c1...`) and badges in the Bottom Editor Dock (`#dockKeyBadge`), Edit Modal (`#modalKeyDisplay`), and Ticket Details modal 1-click copy-able.
  - Clicking the chip or an adjacent copy icon will copy the full 36-character UUIDv4 to the system clipboard (`navigator.clipboard.writeText`) and trigger a quick feedback toast (e.g., `📋 Copied Tracking ID to clipboard: 8f4a92c1-...`).

---

## 2. Recently Delivered Milestones (`v1.17.0` – `v1.18.2`)

### 🔧 Secondary / Custom Tracker Auto-Healing & Chronological Ordering
- **Status**: ✅ Implemented in `v1.18.2`
- **Delivered**:
  - Removed mandatory first field requirement on custom schemas;
  - Added chronological left-to-right daily tab ordering (`orderWorkbookSheetsChronologically`);
  - Fixed mid-month Status column vs. retroactive UUIDv4 collision by auto-migrating legacy UUIDs to Column K and restoring clean statuses;
  - Cleaned validation list dropdown synchronization and prevented UUID scraping into status dropdowns.

---

### 📋 Manual Row Copy & Paste Across Daily Tabs
- **Status**: ✅ Implemented in `v1.18.1`
- **Delivered**: Manual checkbox-based row copy & paste (`#btnBatchCopyRows`, `#btnBatchCopyToTab`, `#btnPasteTrackerRows`) allowing 1-click or two-step copying of requests across daily tabs on any tracker profile.

---

### 📖 Interactive In-App User Manual & Reference Guide
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

## 4. Parked & Deferred Scope

| Concept | Status | Rationale |
|---|---|---|
| **Mass Report Filter (380,000-Row Client-Side `Set` Filter)** | ⏸️ Parked | Explicitly parked. High memory overhead; requires `{ header: 1 }` stream parsing to avoid V8 memory crashes. |
| **Smart Scratchpad (Real-Time Regex Lot # & Title Extractor)** | ⏸️ Parked | Split-pane scratchpad with automated regex scraping for 8-digit lot numbers. Parked in favor of Excel-like in-app scratch sheet tabs (`v1.15.0`). |
| **Chart.js Metric View for Current Data** | ⏸️ On Hold | Deferred to prioritize agent daily workflow, table responsiveness, and scratch sheet capabilities. |

---

## 5. Architecture Decisions & Decided Against

| Decision | Date | Reason / Alternative Chosen |
|---|---|---|
| **Excluding `Master Accounts` Tab from New Monthly Workbooks** | `2026-11` | Creating a new monthly workbook no longer requires or creates a `Master Accounts` sheet. Bidirectional Broker ↔ Master # auto-fill automatically learns from existing entries across daily tabs. |
| **Pure Immutable UUIDv4 vs. Ticket # Primary Key** | `2026-10` | Ticket numbers are optional metadata that may be omitted or updated later. Hidden Column K/L strictly holds a pure 36-character UUIDv4 (`crypto.randomUUID()`) as the immutable row key. |
| **Directory-Based I/O (`showDirectoryPicker`) vs. File Picker** | `2026-10` | Directory handles allow automatic monthly workbook discovery, template cloning, and background dual-write mirroring (`<Workbook> (VIEW ONLY).xlsx`) without repeated file-dialog prompts. |
| **Zero-Build PWA Structure vs. Bundlers / Frameworks** | `2026-10` | No Vite, Webpack, npm, or React. Keeps the app directly runnable on `localhost` (VS Code Live Server) and GitHub Pages (`https://`). |
