# TrackerUI — Future Features & Strategic Roadmap

This document maintains the roadmap, feature requests, planned enhancements, and architectural backlog for **TrackerUI**.

---

## 1. Upcoming & Planned Features

### 🎯 Grab Ticket Number from Email Subject (`[XXXX-XX-XXXX]`)
- **Status**: 📋 Planned
- **Trigger**: Manual button click (`Grab from Subject` / `🎯 Grab Ticket`) located beside the **Ticket Number (Col B)** and **Subject (Col A)** fields in both the **Bottom Editor Dock** and the **Modal Editor**.
- **Email Subject Syntax**: Matches ticket numbers formatted like `[XXXX-XX-XXXX]` (e.g. `[ABCD-12-3456]`, `[1234-56-7890]`, or general hyphenated ticket patterns in brackets `\[([A-Za-z0-9]{2,8}(?:-[A-Za-z0-9]{1,6})+)\]`).
- **Comma-Delimited Support & Deduplication**:
  - If the email subject contains multiple bracketed ticket numbers, extracts all of them into a clean, comma-delimited list (e.g. `ABCD-12-3456, EFGH-78-9012`).
  - If the Ticket Number field already contains one or more ticket numbers, appends newly extracted tickets without duplicating existing ones.
  - Normalizes extracted ticket numbers to uppercase.

---

### 🔄 Forward Batch Sync Across Daily Tabs
- **Status**: 📋 Planned
- **Trigger**: Dedicated **`Sync Forward Across Tabs`** button available in both:
  1. The **Bottom Editor Dock** (when an existing request row is loaded for editing).
  2. The **Ticket Details & History** modal (`#ticketDetailsModal`).
- **Use Case**: When an older entry is updated (for example, adding a ticket number that was missing when the request was first logged), propagating this new information to matching rolled-over rows across subsequent daily tabs (`parseDateTab > currentDate`).
- **Safety Checks & Overwrite Protection**:
  - **Non-destructive fill**: Only fills fields that are currently blank or empty on forward tabs so existing data is never accidentally wiped.
  - **Comma-delimited ticket merging**: If a forward entry already has a ticket number, new tickets are merged comma-delimited without duplicates (e.g., if forward has `TICKET-A` and updated entry has `TICKET-B`, forward becomes `TICKET-A, TICKET-B`).
  - **Confirmation preview**: Summarizes which forward tabs and rows will be updated before committing changes to the workbook and mirror.

---

### 📥 Sync from Previous Entry ("Pull from Prior Tab")
- **Status**: 📋 Planned
- **Trigger**: Button beside request editing inputs to look backward at earlier daily tabs (`parseDateTab < currentDate`).
- **Behavior**: Inspects earlier instances of the same unique request (matched by Tracking ID or Subject + Account + Lots) and pulls missing metadata (Broker, Master #, Bidder #, Yard, Action, Lots, Ticket #) into empty fields on the current entry.

---

### 🗂️ Multi-Department Tracker Tabs & Dynamic Column Schema Engine
- **Status**: 📋 Planned (Viable — High Value)
- **Use Case**: Agents working across multiple departments need to keep 2+ distinct **Daily Tracker** workspaces open side-by-side in `#workspaceTabsBar`, where secondary department trackers have completely different Excel column counts, headers, and dropdown validation lists rather than the hardcoded 11/12-column `STANDARD_HEADERS` layout.
- **Core Architectural Requirements**:
  1. **Multiple Tracker Profiles in `#workspaceTabsBar`**:
     - Generalize `DAILY_TRACKER_WS_TAB` into named **Tracker Workspace Profiles** (e.g., *Broker Support* + *Department B Tracker*), each persisting its own `agentFolderHandle`, optional `supervisorFolderHandle` (with per-profile toggle if a department does not use a `(VIEW ONLY)` mirror), `templatesFolderHandle`, and active `workingFileHandle` in `IndexedDB`.
  2. **Dynamic Header Discovery & Semantic Role Mapping (Replacing `STANDARD_HEADERS`)**:
     - Parse Row 1 of the workbook's `template` (or latest `MM-DD-YYYY` daily sheet) to build a dynamic column descriptor array: `[{ colIndex, key, headerLabel, inputType, dropdownOptions, width, required }]`.
     - Require mapping only **3 lightweight semantic roles** (auto-inferred by header name keywords with a 1-time **Column Mapper UI** override stored per profile/workbook):
       - **Immutable Primary Key (`Tracking ID`)**: Automatically assigned to the first hidden column (`maxVisibleCol + 1`) so UUIDv4 immutability, cross-day identity matching, and safe edits work identically regardless of column count.
       - **Status Column** *(e.g., `Status`, `Request Status`, `State`)*: Powers status badge colors, `#batchActionDock` bulk updates, and unfinished-item Rollover filtering.
       - **Notes / Timeline Column** *(optional, e.g., `Action Notes`, `Notes`, `Comments`)*: Powers `+ Insert Timestamp`, `#ticketDetailsModal` history, and chronological `[MM-DD-YYYY]` rollover note merging.
  3. **Schema-Driven Table, Search, Sort & Editor Forms**:
     - Dynamically generate `#tableHead` (`<th>` sort headers) and `#listBody` (`<td>` cells + click-to-filter Like-Items) from the active Tracker Tab's column descriptors.
     - Dynamically render `#dockEntryForm` (Bottom Editor Dock) and `#entryForm` (Popup Modal) inputs from the same schema, binding any column matching a header/column in `validation lists` to a `<select>` or `<datalist>`.

---

## 2. In-Progress & Architectural Backlog

| Feature / Suggestion | Source | Status | Description & Remaining Work |
|---|---|---|---|
| **Dynamic Schema Mapping Configuration** | `TrackerUI Architectural Recs` | ⏳ Partially Implemented | **Done**: `getWorksheetColumnMap()` and `parseMasterAccountsSheet()` auto-detect the standard 11-column layout vs. the 12-column Tracker Template layout (`Status` in Col I, `Notes` in Col J, `Sent Email Subject` in Col K, `Tracking ID` in Col L).<br>**Remaining**: Full dynamic Row-1 header discovery, semantic role mapper UI, and dynamic table/form generation to power **Multi-Department Tracker Tabs** (see Section 1 above). |
| **IndexedDB Local Fail-Safe Cache** | `TrackerUI Architectural Recs` | ⏳ Partially Implemented | **Done**: `IndexedDB` (`TrackerUI_DirectoryDB`) persists Agent, Supervisor, and Templates directory handles; in-memory `pendingMirrorBuffer` caches unsynced Supervisor saves with 15s auto-retry.<br>**Remaining**: Offline mutation staging queue in IndexedDB to buffer primary `.xlsx` edits if file handles temporarily freeze. |
| **Supervisor Pipeline: Subtype → Parent Type Lookup** | `Supervisor Pipeline (Phase 2)` | 📋 Backlog | Action subtype to parent category classification mapping for supervisor rollups and reporting. |
| **Web Worker Background ExcelJS Parsing** | `Deep Processing (Phase 3)` | 📋 Backlog | Offload heavy OpenXML unzipping and historical tab cross-search to a Web Worker to maintain 60fps main-thread UI performance during large lookups. |
| **Multi-File Master Tracker Aggregation Loop** | `Deep Processing (Phase 3)` | 📋 Backlog | Scan and aggregate multiple agent workbooks within a supervisor directory into a unified Master Tracker view. |
| **External Config File (`.json` / `.ini`) or Hidden Config Sheet** | `TrackerUI Architecture` | 📋 Backlog | Store application settings, custom column maps, and workspace preferences in an external file or dedicated hidden workbook tab. |

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
