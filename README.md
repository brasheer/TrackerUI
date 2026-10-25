# Bidder Support Tracker (TrackerUI)

A zero-build, standalone Progressive Web App (PWA) designed to read, search, update, and manage agent support logs directly from Excel workbooks (`.xlsx`) stored locally or synced via OneDrive.

---

## Key Features

- **Directory-Based Dual-Write Architecture (`showDirectoryPicker`)**:
  - Connect your agent directory (e.g., `.\trackers\[Agent Name]\`) with 1-click folder reconnection persisted via `IndexedDB`.
  - Automatically scans for monthly working `.xlsx` workbooks while ignoring Excel lock files (`~$*`) and supervisor mirror files (`*(VIEW ONLY).xlsx`).
  - **Custom Supervisor Export Folder**: Choose a dedicated directory where `<Workbook Name> (VIEW ONLY).xlsx` copies are written (persisted in `IndexedDB`), or default to the same directory as the working workbook. If a supervisor has the `(VIEW ONLY)` mirror open in Desktop Excel, the agent's save still succeeds with a non-blocking warning pill.
- **Automatic Global Search Across All Sheets**:
  - Typing in the search bar automatically queries across **all** daily `MM-DD-YYYY` sheets in the workbook (displaying a `🌐 Searching All Tabs` indicator) and returns to your selected view scope (`Current Day Tab Only` or `All Daily Tabs`) as soon as the search box is cleared.
- **Batch Delete Selection Mode**:
  - Click **`🗑️ Delete Requests`** to reveal selection checkboxes next to every visible request (plus a "Select All" toggle).
  - Requires explicit confirmation listing the selected items before safely splicing rows from the workbook (in descending row order to preserve row indexing).
- **Pure Immutable GUIDs in Hidden Column K (`Tracking ID`)**:
  - Uses `crypto.randomUUID()` to assign a permanent lowercase 36-character UUIDv4 to every request in hidden Column K, completely decoupled from optional Ticket numbers (Col B).
  - Automatically backfills UUIDs for legacy rows missing a Column K identifier.
  - Displays a compact 8-character chip (`🔍 8f4a92c1...`) with full UUID hover tooltip and 1-click cross-workbook history tracing.
- **Dynamic Reference Sheet Parsing & Auto-Fill**:
  - **`validation lists` (or `Validation List`)**: Populates `Requested Action` (Col A) autocomplete and `Request Status` (Col B) dropdown options dynamically.
  - **`Master Accounts` (or `Master Account`)**: Maps `Broker Name` (Col A), `Master#` (Col B), and optional `Default Bidder#` (Col C), providing bidirectional auto-fill in the New/Edit Request modal.
- **Smart Multi-Sheet & Idempotent Rollover Engine (with Future Date Handling)**:
  - Scans **all** prior `MM-DD-YYYY` sheets in chronological order (oldest to newest) to resolve the latest historical state of every unique request UUID across the entire workbook.
  - **Future-Dated Tab Handling**: Detects when a workbook contains tabs dated ahead of today's calendar date, offering 1-click presets for `Latest Tab`, `Next Day`, or `Today`, plus an option to include open items from future-dated tabs.
  - **Superseded Request Override**: Highlights requests that were open on an older tab (`In Progress`, `Escalated`, `Pending`) but marked `Complete` on a later tab, allowing 1-click reopening and carryover.
  - **Cross-Tab Status Propagation**: Editing a request's status on an older tab prompts to sync the updated status to any newer tabs where that same Tracking ID exists.

---

## Directory Structure

```
TrackerUI/
├── assets/
│   ├── icons/
│   │   ├── icon-192.png       # 192x192 PWA icon
│   │   ├── icon-512.png       # 512x512 PWA icon & maskable icon
│   │   └── icon.svg           # Scalable vector icon
│   └── images/
├── css/
│   └── styles.css             # Dark-themed UI styles, banners, modals & toasts
├── data/                      # Folder for sample data or template files
├── js/
│   └── app.js                 # Directory I/O, dual-write, search, delete & rollover engine
├── .gitignore                 # Configured for OneDrive/Google Drive & OS temp files
├── AGENTS.md                  # Always-on workspace rules for AI coding assistants
├── index.html                 # App shell and modal markup
├── manifest.json              # PWA manifest
├── README.md                  # Project documentation
└── sw.js                      # Service Worker for offline shell caching
```

---

## Workbook Structure

### Daily Sheet Columns (`MM-DD-YYYY`)

| Col | Header | Description |
|---|---|---|
| **A** | `Subject` | Request subject line |
| **B** | `Ticket Number` | Optional ticket identifier (normalized to uppercase) |
| **C** | `Broker Name` | Broker associated with the request (linked to `Master Accounts`) |
| **D** | `Master #` | Master account number (auto-filled & normalized to uppercase) |
| **E** | `Bidder #` | Bidder number (normalized to uppercase) |
| **F** | `Requested Action` | Action requested (populated from `validation lists` Col A) |
| **G** | `Yard #` | Yard or branch location (normalized to uppercase) |
| **H** | `Lot Numbers` | Specific lot numbers |
| **I** | `Action Notes` | Detailed notes (collapsible preview in UI) |
| **J** | `Request Status` | Populated from `validation lists` Col B (falls back to `Pending`, `In Progress`, `Waiting on Buyer`, `Escalated`, `Complete`) |
| **K** | `Tracking ID` | **Hidden column**: Pure immutable UUIDv4 (`crypto.randomUUID()`) preserved across rollovers |

### Specialized Reference Sheets
- `template`: Blueprint sheet used to clone headers, column widths, and cell formatting when creating new daily tabs.
- `validation lists` / `Validation List`:
  - **Column A** (Row 2+): `Requested Action` autocomplete options.
  - **Column B** (Row 2+): `Request Status` select options.
- `Master Accounts` / `Master Account`:
  - **Column A** (Row 2+): `Broker Name`
  - **Column B** (Row 2+): `Master#`
  - **Column C** (Row 2+): `Default Bidder#` (optional)

---

## Getting Started

### Prerequisites
The **File System Access API** (`window.showDirectoryPicker`) requires:
1. A Chromium-based browser (**Google Chrome**, **Microsoft Edge**, **Brave**, or **Opera**).
2. A secure origin (`http://localhost` or `https://`). Opening `index.html` directly as `file:///` is not permitted by browser security policies.

### Running Locally
- **VS Code Live Server**: Right-click [`index.html`](index.html) and select **"Open with Live Server"**.
- **Python HTTP Server**: `python -m http.server 8080` -> `http://localhost:8080`.
- **Node.js**: `npx serve .`

---

## AI Assistant Governance ([`AGENTS.md`](AGENTS.md))
This repository includes an always-on [`AGENTS.md`](AGENTS.md) rules file that enforces the modular file layout, `showDirectoryPicker` dual-write architecture, immutable UUIDv4 primary keys, and ExcelJS conventions across all AI models.

---

## AI Assistant Governance ([`AGENTS.md`](AGENTS.md))
This repository includes an always-on [`AGENTS.md`](AGENTS.md) rules file that enforces the modular file layout, `showDirectoryPicker` dual-write architecture, immutable UUIDv4 primary keys, and ExcelJS conventions across all AI models.

