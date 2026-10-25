# Bidder Support Tracker (TrackerUI)

A lightweight, standalone Progressive Web App (PWA) designed to read, search, update, and manage agent support logs directly from an Excel workbook (`.xlsx`) stored locally or synced via OneDrive.

---

## Features

- **Direct In-Browser Excel Sync**: Powered by [ExcelJS](https://github.com/exceljs/exceljs) and the native browser **File System Access API**. Edits write back directly to the local/OneDrive `.xlsx` file without needing an external database or server backend.
- **Daily Rollover Engine**: One-click rollover to copy all unfinished requests (status $\neq$ `Complete`) from the most recent day tab to a newly created day tab (`MM-DD-YYYY`), automatically cloning column widths, headers, and styles.
- **Cross-Day History Tracking**: Each request maintains a unique persistent key in hidden **Column K**. Clicking any Key chip filters the entire workbook history to show the full multi-day lifecycle of that request.
- **Search & Scope Filters**: Filter records by current day tab or search across all historical tabs simultaneously.
- **Excel Lock Detection**: Alerts when the target `.xlsx` file is currently open and locked by desktop Excel, preventing silent corruption or lost edits.
- **Progressive Web App (PWA)**: Installable as a desktop app in Chrome, Edge, and other Chromium browsers with offline caching via Service Worker (`sw.js`).

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
│   └── styles.css             # Dark-themed UI styles
├── data/                      # Folder for sample data or template files
├── js/
│   └── app.js                 # PWA logic, ExcelJS I/O, search & rollover
├── .gitignore                 # Configured for OneDrive/Google Drive & OS temp files
├── index.html                 # App shell and modal markup
├── manifest.json              # PWA manifest
├── README.md                  # Project documentation
└── sw.js                      # Service Worker for offline shell caching
```

---

## Workbook Structure

The workbook uses daily sheets with date-formatted tab names (`MM-DD-YYYY`), alongside optional template and lookup sheets:

### Daily Sheet Columns

| Col | Header | Description |
|---|---|---|
| **A** | `Subject` | Request subject line |
| **B** | `Ticket Number` | Ticket identifier |
| **C** | `Broker Name` | Broker associated with the request |
| **D** | `Master #` | Master account number |
| **E** | `Bidder #` | Bidder number |
| **F** | `Requested Action` | Action requested |
| **G** | `Yard #` | Yard or branch location |
| **H** | `Lot Numbers` | Specific lot numbers |
| **I** | `Action Notes` | Detailed notes (collapsible preview in UI) |
| **J** | `Request Status` | `Pending`, `In Progress`, `Waiting on Buyer`, `Escalated`, or `Complete` |
| **K** | `Key` | **Hidden column**: Persistent tracking key preserved across rollovers |

### Specialized Sheets
- `template`: (Optional) Blueprint sheet used to inherit headers, column widths, and cell formatting when generating new daily tabs. If omitted, the most recent daily tab is used as the template.
- `Validation List`: Reference sheet for dropdown lookups.
- `Master Account`: Reference sheet for master account mappings.

---

## Getting Started

### Prerequisites
The **File System Access API** (`window.showOpenFilePicker`) requires:
1. A Chromium-based browser (**Google Chrome**, **Microsoft Edge**, **Brave**, or **Opera**).
2. A secure origin (`http://localhost` or `https://`). Opening `index.html` directly as `file:///` is not permitted by browser security policies.

---

### Running Locally

#### Option 1: VS Code Live Server (Recommended)
1. Install the **Live Server** extension in VS Code.
2. Right-click [`index.html`](index.html) and select **"Open with Live Server"**.

#### Option 2: Python HTTP Server
Run from the project root in PowerShell:
```powershell
python -m http.server 8080
```
Then visit `http://localhost:8080` in your browser.

#### Option 3: Node.js `serve`
```powershell
npx serve .
```

---

### Hosting on GitHub Pages

1. Push your repository to GitHub:
   ```powershell
   git remote add origin https://github.com/<your-username>/TrackerUI.git
   git push -u origin main
   ```
2. In your GitHub repository:
   - Navigate to **Settings** > **Pages**.
   - Under **Build and deployment** > **Source**, choose **Deploy from a branch**.
   - Select branch **`main`** and folder **`/ (root)`**, then click **Save**.
3. Access your live PWA at:
   ```
   https://<your-username>.github.io/TrackerUI/
   ```
4. Click the install icon in your browser's address bar to install TrackerUI to your desktop.

---

## OneDrive Integration & Best Practices

- When opening an Excel file synced via OneDrive, keep desktop Excel closed while actively editing and saving requests in the PWA to avoid Windows file locks.
- If desktop Excel has the file open, the app detects the lock and instructs you to close Excel before retrying the save.
- All OneDrive and Google Drive temporary lock files (`~$*`, `*.tmp`, `*.cloud`, conflict copies) are automatically excluded by [`.gitignore`](.gitignore).
