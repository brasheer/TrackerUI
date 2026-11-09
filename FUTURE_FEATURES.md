# TrackerUI — Future Features & Strategic Roadmap

This document maintains the roadmap, feature requests, planned enhancements, and architectural backlog for **TrackerUI**.

---

## 1. Newly Implemented Features (`v1.17.0` on `meerkat`)

### 🎯 Grab Ticket Number from Email Subject (`[XXXX-XX-XXXX]`)
- **Status**: ✅ Implemented in `v1.17.0`
- **Delivered**: Manual `🎯 Grab Ticket` button (`#btnDockGrabTicket`, `#btnModalGrabTicket`, and dynamic custom ticket fields) extracts one or more bracketed hyphenated ticket patterns (`[ABCD-12-3456]`, `[1234-56-7890]`) from Subject, normalizes to uppercase, and merges comma-delimited without duplicates (`extractBracketedTicketsFromSubject` + `mergeCommaDelimitedTickets`).

---

### 🔄 Forward Batch Sync Across Daily Tabs
- **Status**: ✅ Implemented in `v1.17.0`
- **Delivered**: `🔄 Sync Forward` button in both `#bottomEditorDock` (`#btnDockSyncForward`) and `#ticketDetailsModal` (`#btnTicketSyncForward`) (`handleForwardBatchSync`). Propagates newly added details forward across newer `MM-DD-YYYY` tabs non-destructively (only populating blank fields while merging Ticket Numbers comma-delimited without duplicates) with a confirmation diff preview before saving.

---

### 📥 Sync from Previous Entry ("Pull from Prior Tab")
- **Status**: ✅ Implemented in `v1.17.0`
- **Delivered**: `📥 Pull Prior` / `📥 Pull from Prior Tab` button in `#bottomEditorDock` (`#btnDockPullPrior`) and `#editModal` (`#btnModalPullPrior`) (`handlePullFromPriorTab`). Scans earlier `MM-DD-YYYY` tabs for matching requests (by Tracking ID, Ticket #, composite identity, or Subject), fills only blank fields on the current form, and merges Ticket Numbers comma-delimited.

---

### 🗂️ Multi-Department Tracker Tabs & Dynamic Column Schema Engine
- **Status**: ✅ Implemented in `v1.17.0`
- **Delivered**:
  1. **Multiple Tracker Profiles in `#workspaceTabsBar`**: `＋ Tracker Tab` (`#btnAddTrackerProfile` & `#menuAddTrackerProfile`) creates independent workspace profiles with isolated `IndexedDB` folder handles, instant runtime snapshot switching, double-click tab renaming, and per-profile `(VIEW ONLY).xlsx` mirror toggle (`#chkEnableMirrorCopy`).
  2. **Dynamic Header Discovery & Role Mapper (`#schemaMapperModal`)**: Auto-detects Standard 11/12-column Broker layouts vs. arbitrary custom department Row-1 headers, auto-infers Primary, Status, Notes, Ticket #, and Secondary columns, assigns hidden UUIDv4 `Tracking ID` to `maxVisibleCol + 1`, binds matching `validation lists` columns, and provides `Configure Column Mapping...` (`#btnOpenSchemaMapper`).
  3. **Schema-Driven Table, Editors, View-Only & Rollover**: Dynamically drives `#tableHead`, `#listBody`, `#dockDynamicFieldsGrid`, `#modalDynamicFieldsGrid`, `#ticketDetailsModal`, `#viewOnlyModal`, and Smart Rollover (`js/rollover.js`).

---

## 2. Completed Architectural Backlog (`v1.17.0` on `meerkat`)

| Feature / Suggestion | Source | Status | Implementation Summary |
|---|---|---|---|
| **Dynamic Schema Mapping Configuration** | `TrackerUI Architectural Recs` | ✅ Implemented in `v1.17.0` | Full Row-1 header discovery (`resolveWorkbookSchemaConfig`), semantic role mapper modal (`#schemaMapperModal`), series-aware override persistence (`getWorkbookSeriesKey`), and dynamic table/form generation. |
| **IndexedDB Local Fail-Safe Cache** | `TrackerUI Architectural Recs` | ✅ Implemented in `v1.17.0` | Stages serialized `.xlsx` buffers in `IndexedDB` (`stagedPrimarySave::*`) whenever Desktop Excel locks the primary working file, keeps in-memory edits live, surfaces `💾 Retry Primary Save` (`#btnRetryPrimarySave`), and prompts recovery on reload. |
| **Supervisor Pipeline: Subtype → Parent Type Lookup** | `Supervisor Pipeline (Phase 2)` | ✅ Implemented in `v1.17.0` | `resolveActionParentCategory()` maps action subtypes via custom JSON overrides, `validation lists` category columns, and built-in operational category rules across `#ticketDetailsModal` and `#masterRollupModal`. |
| **Web Worker Background ExcelJS Parsing** | `Deep Processing (Phase 3)` | ✅ Implemented in `v1.17.0` | `parseWorkbookBufferInBackground()` offloads multi-workbook `.xlsx` parsing to a dedicated `Blob` Web Worker importing ExcelJS (with automatic async main-thread fallback if blocked/offline). |
| **Multi-File Master Tracker Aggregation Loop** | `Deep Processing (Phase 3)` | ✅ Implemented in `v1.17.0` | `#masterRollupModal` (`#menuOpenMasterRollup`) scans multiple `.xlsx` / `(VIEW ONLY).xlsx` workbooks in a folder, deduplicates rolled-over requests per unique ID, renders KPI & category breakdowns, and exports `.xlsx` rollups (`#btnExportMasterRollupXlsx`). |
| **External Config File (`.json`)** | `TrackerUI Architecture` | ✅ Implemented in `v1.17.0` | `Export Workspace Config (.json)` (`#btnExportConfigJson`) and `Import Workspace Config (.json)...` (`#btnImportConfigJson`) in **Tools & Setup**. |

---

## 3. Parked & Deferred Scope

| Concept | Status | Rationale |
|---|---|---|
| **Mass Report Filter (380,000-Row Client-Side `Set` Filter)** | ⏸️ Parked | Explicitly parked. High memory overhead; requires `{ header: 1 }` stream parsing to avoid V8 memory crashes. |
| **Smart Scratchpad (Real-Time Regex Lot # & Title Extractor)** | ⏸️ Parked | Split-pane scratchpad with automated regex scraping for 8-digit lot numbers. Parked in favor of Excel-like in-app scratch sheet tabs (`v1.15.0`). |
| **Chart.js Metric View for Current Data** | ⏸️ On Hold | Deferred to prioritize agent daily workflow, table responsiveness, and scratch sheet capabilities. |

---

## 4. Architecture Decisions & Decided Against

| Decision | Date | Reason / Alternative Chosen |
|---|---|---|
| **Excluding `Master Accounts` Tab from New Monthly Workbooks** | `2026-11` | Creating a new monthly workbook no longer requires or creates a `Master Accounts` sheet. Bidirectional Broker ↔ Master # auto-fill automatically learns from existing entries across daily tabs. |
| **Pure Immutable UUIDv4 vs. Ticket # Primary Key** | `2026-10` | Ticket numbers are optional metadata that may be omitted or updated later. Hidden Column K/L strictly holds a pure 36-character UUIDv4 (`crypto.randomUUID()`) as the immutable row key. |
| **Directory-Based I/O (`showDirectoryPicker`) vs. File Picker** | `2026-10` | Directory handles allow automatic monthly workbook discovery, template cloning, and background dual-write mirroring (`<Workbook> (VIEW ONLY).xlsx`) without repeated file-dialog prompts. |
| **Zero-Build PWA Structure vs. Bundlers / Frameworks** | `2026-10` | No Vite, Webpack, npm, or React. Keeps the app directly runnable on `localhost` (VS Code Live Server) and GitHub Pages (`https://`). |
