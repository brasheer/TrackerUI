# TrackerUI — Workspace Agent Rules & Architecture Invariants

These rules apply to all AI models and coding assistants working in the `TrackerUI` repository. Follow these constraints strictly to keep the project consistent.

---

## 1. Modular Zero-Build PWA Structure (No Bundlers, No Inline Monoliths)
- **Keep CSS and JS Separated**: Maintain the modular layout:
  - `index.html`: Application shell, popover menu, banners, and modal markup only.
  - `css/styles.css`: All styles, CSS custom properties, and responsive rules.
  - `js/app.js`: All application state, File System Access API I/O, ExcelJS manipulation, schema bootstrapping, search, delete mode, and rollover logic.
  - `manifest.json` & `sw.js`: PWA manifest and offline Service Worker.
  - `assets/icons/`: App icons (`icon.svg`, `icon-192.png`, `icon-512.png`).
- **Do NOT inline** `<style>` or `<script>` blocks back into `index.html`.
- **Do NOT introduce** build steps, bundlers (Vite, Webpack, Rollup), npm runtime dependencies, or UI frameworks. The app must run directly via VS Code Live Server (`localhost`) and GitHub Pages (`https://`).
- **Service Worker Cache**: Whenever `index.html`, `css/styles.css`, or `js/app.js` changes, increment `CACHE_NAME` in `sw.js`.

---

## 2. Directory-Based Dual-Write I/O, Custom Supervisor Folder & Mirror Auto-Retry
- **Agent Folder (`showDirectoryPicker`)**: Always use `window.showDirectoryPicker({ mode: 'readwrite' })` rather than `showOpenFilePicker`.
- **Workbook Filtering**: When scanning the agent folder for working `.xlsx` files, always exclude:
  - Excel temporary lock files starting with `~$`
  - Mirror files ending with `(VIEW ONLY).xlsx` (case-insensitive)
- **IndexedDB Handle Persistence**: Persist both `agentFolderHandle` and `supervisorFolderHandle` in `IndexedDB` (`TrackerUI_DirectoryDB` -> `handles`) so the user can reconnect with 1 click on reload.
- **Dual-Write Mirroring & Retry (`saveWorkbookToDisk` / `attemptMirrorWrite`)**:
  1. Write the serialized `workbook.xlsx.writeBuffer()` to the primary working `.xlsx` file inside a `try/catch` that alerts if Desktop Excel has the working file locked.
  2. Write the identical buffer to `<Workbook Name> (VIEW ONLY).xlsx` in `supervisorDirHandle || parentDirHandle` inside a **separate, non-blocking `try/catch`**.
  3. If a supervisor has the `(VIEW ONLY)` copy locked in Desktop Excel, never fail the agent's save—log a warning pill (`⚠️ Saved Working File (Mirror in use by Supervisor)`), start a 15-second background auto-retry countdown (`scheduleMirrorAutoRetry`), and display the manual **`🔄 Retry Mirror`** button (`#btnRetryMirror`).

---

## 3. Workbook Schema, Bootstrapping & Pure Immutable GUIDs (Column K)
Daily tabs must use `MM-DD-YYYY` naming and follow this exact 11-column schema:
- **Col A (1)**: `Subject`
- **Col B (2)**: `Ticket Number` (optional metadata, normalized to uppercase on save)
- **Col C (3)**: `Broker Name`
- **Col D (4)**: `Master #` (normalized to uppercase on save)
- **Col E (5)**: `Bidder #` (normalized to uppercase on save)
- **Col F (6)**: `Requested Action`
- **Col G (7)**: `Yard #` (normalized to uppercase on save)
- **Col H (8)**: `Lot Numbers`
- **Col I (9)**: `Action Notes`
- **Col J (10)**: `Request Status`
- **Col K (11)**: `Tracking ID` (**Hidden column**: `worksheet.getColumn(11).hidden = true`)

### Schema Bootstrapping (`bootstrapWorkbookSchemaInMemory`)
- Support creating a brand-new monthly `.xlsx` file in the Agent Folder as well as non-destructively initializing/repairing blank or incomplete workbooks with:
  - `template` sheet (11 styled headers, column widths, hidden Column K)
  - `validation lists` sheet (default `Requested Actions` in Col A and `Status Options` in Col B)
  - `Master Accounts` sheet (`Broker Name`, `Master#`, `Default Bidder#`)
  - Today's `MM-DD-YYYY` daily tab cloned from `template`.

### Primary Key Rules (Column K)
- **Pure UUIDv4 Only**: New rows must always generate a lowercase 36-character UUIDv4 via `crypto.randomUUID()` (`generateTrackingGuid()`). Never use Ticket Number (Col B) or composite broker hashes as a primary key for new rows.
- **Strict Immutability on Edit**: Editing an existing row must strictly preserve its existing Column K value (`editingTarget.key`), even if Column B (`Ticket Number`) is added, modified, or cleared.
- **Deterministic Legacy Backfill**: If an existing row on a `MM-DD-YYYY` sheet has data in Columns A–J but a blank Column K, assign a UUIDv4 cached in `legacyBackfillMap` (keyed by `fileName::sheetName::rowNumber`) and write it to Cell 11 so it persists on the next save.
- **8-Character UI Chip**: Render the Tracking ID chip in the table using the first 8 characters (`🔍 8f4a92c1...`) with the full ID in `title` and `data-key`.

---

## 4. Reference Sheet Parsing & Bidirectional Auto-Fill
- **`validation lists` / `Validation List`** (case-insensitive):
  - Column A (Row 2+): Populates `<datalist id="actionList">`.
  - Column B (Row 2+): Populates `<select id="fStatus">`, falling back to `['Pending', 'In Progress', 'Waiting on Buyer', 'Escalated', 'Complete']`.
- **`Master Accounts` / `Master Account`** (case-insensitive):
  - Column A (Row 2+): `Broker Name`
  - Column B (Row 2+): `Master#`
  - Column C (Row 2+): `Default Bidder#` (optional)
  - Typing a known Broker Name in `#fBroker` auto-fills `#fMaster` (and `#fBidder` if empty); typing a known `Master#` in `#fMaster` reverse-fills `#fBroker` (and `#fBidder` if empty).

---

## 5. UI Navigation, Sheet Management, Search, Delete Mode & Smart Rollover
- **Popover Tools & Setup Menu (`#toolsMenuWrap`)**: Keep secondary administrative actions (folder switching, supervisor folder config, workbook creation/schema repair, tab create/rename/delete, and batch delete mode) inside the `⚙️ Tools & Setup ▾` popover menu so the primary header stays uncluttered.
- **Day Tab Stepper & Quick Tab Bar**: Maintain both the `◀` / `▶` stepper buttons flanking `#sheetSelect` and the horizontal `#sheetTabsBar` pill strip showing request counts per `MM-DD-YYYY` tab.
- **Protected Sheet Renaming & Deletion**: Only allow renaming/deleting `MM-DD-YYYY` daily tabs; never delete reference sheets (`template`, `validation lists`, `Master Accounts`) and always ensure at least one worksheet remains in the workbook.
- **Automatic Global Search**: Typing a non-empty query in `#searchInput` must automatically search across **all** `MM-DD-YYYY` sheets regardless of the `#scopeSelect` dropdown value, returning to `#scopeSelect` filtering when the search box is cleared.
- **Batch Delete Mode**: Deleting rows uses the checkbox selection mode (`#btnDeleteMode` / `#deleteBanner`), requires user confirmation, and sorts target row numbers **descending** (`b.rowNumber - a.rowNumber`) per worksheet before calling `ws.spliceRows(rowNumber, 1)` so row indices never shift during deletion.
- **Idempotent Multi-Sheet Rollover & Future Date Handling**:
  - Scans all prior `MM-DD-YYYY` sheets chronologically (oldest to newest) to resolve the latest state of each Tracking ID.
  - Detects future-dated tabs and provides preset targets (`Latest Tab`, `Next Day`, `Today`) plus an option to include unfinished items from future-dated tabs.
  - Surfaces older open requests that were marked `Complete` on a later sheet so the user can optionally reopen and carry them over.
