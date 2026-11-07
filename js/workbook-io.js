// ============================================================================
// workbook-io.js — Folder Handles (IndexedDB), Workbook Discovery, OpenXML
//                  Archive Repair, ExcelJS Load/Save, Mirror Auto-Retry,
//                  Schema Bootstrapping, Daily Tab CRUD & Account Auto-Fill
// ============================================================================

// --- INDEXEDDB DIRECTORY HANDLE & SETTINGS PERSISTENCE ---
function openHandleDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(IDB_NAME, 1);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(IDB_STORE)) {
        db.createObjectStore(IDB_STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

async function saveHandleToIDB(key, value) {
  try {
    const db = await openHandleDB();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readwrite');
      tx.objectStore(IDB_STORE).put(value, key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn(`Could not persist (${key}) in IndexedDB:`, err);
  }
}

async function getHandleFromIDB(key) {
  try {
    const db = await openHandleDB();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readonly');
      const req = tx.objectStore(IDB_STORE).get(key);
      req.onsuccess = () => resolve(req.result || null);
      req.onerror = () => reject(req.error);
    });
  } catch (err) {
    console.warn(`Could not read (${key}) from IndexedDB:`, err);
    return null;
  }
}

async function deleteHandleFromIDB(key) {
  try {
    const db = await openHandleDB();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readwrite');
      tx.objectStore(IDB_STORE).delete(key);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (err) {
    console.warn(`Could not delete (${key}) from IndexedDB:`, err);
  }
}

function setAutoReconnectPreference(enabled) {
  autoReconnectEnabled = Boolean(enabled);
  localStorage.setItem(AUTO_RECONNECT_STORAGE_KEY, autoReconnectEnabled ? 'true' : 'false');
  if (chkAutoReconnect) chkAutoReconnect.checked = autoReconnectEnabled;
  if (chkAutoReconnectBar) chkAutoReconnectBar.checked = autoReconnectEnabled;
}

if (chkAutoReconnect) {
  chkAutoReconnect.checked = autoReconnectEnabled;
  chkAutoReconnect.addEventListener('change', (e) => {
    setAutoReconnectPreference(e.target.checked);
    showToast(
      autoReconnectEnabled
        ? '⚡ Auto-Reconnect on Startup enabled.'
        : 'Auto-Reconnect on Startup disabled.',
      'info'
    );
  });
}

if (chkAutoReconnectBar) {
  chkAutoReconnectBar.checked = autoReconnectEnabled;
  chkAutoReconnectBar.addEventListener('change', (e) => {
    setAutoReconnectPreference(e.target.checked);
    showToast(
      autoReconnectEnabled
        ? '⚡ Auto-Reconnect on Startup enabled.'
        : 'Auto-Reconnect on Startup disabled.',
      'info'
    );
  });
}

function setReconnectButtonText(text) {
  if (reconnectBtnLabel) {
    reconnectBtnLabel.textContent = text;
  } else if (btnReconnect) {
    btnReconnect.textContent = text;
  }
}

function updateAgentDirUI() {
  const setLabel = (txt) => {
    if (agentDirLabel) agentDirLabel.textContent = txt;
    else menuSelectAgentFolder.textContent = txt;
  };
  if (parentDirHandle) {
    setLabel(`Tracker Folder: ${parentDirHandle.name}`);
    menuSelectAgentFolder.classList.add('custom-active');
    if (btnResetAgentDir) btnResetAgentDir.style.display = 'inline-block';
  } else if (cachedDirHandle) {
    setLabel(`Reconnect Tracker: ${cachedDirHandle.name}`);
    menuSelectAgentFolder.classList.add('custom-active');
    if (btnResetAgentDir) btnResetAgentDir.style.display = 'inline-block';
  } else {
    setLabel('Tracker Workbooks Folder...');
    menuSelectAgentFolder.classList.remove('custom-active');
    if (btnResetAgentDir) btnResetAgentDir.style.display = 'none';
  }
  updateFilePathsBarUI();
  updateStatusBarTelemetry();
}

function updateSupervisorDirUI() {
  const setLabel = (txt) => {
    if (supervisorDirLabel) supervisorDirLabel.textContent = txt;
    else btnSupervisorDir.textContent = txt;
  };
  if (supervisorDirHandle) {
    setLabel(`Supervisor Folder: ${supervisorDirHandle.name}`);
    btnSupervisorDir.title = `Saving (VIEW ONLY).xlsx copies to "${supervisorDirHandle.name}". Click to change folder.`;
    btnSupervisorDir.classList.add('custom-active');
    btnResetSupervisorDir.style.display = 'inline-block';
  } else if (cachedSupervisorDirHandle) {
    setLabel(`Reconnect Supervisor: ${cachedSupervisorDirHandle.name}`);
    btnSupervisorDir.title = `Click to reconnect saved Supervisor folder "${cachedSupervisorDirHandle.name}"`;
    btnSupervisorDir.classList.add('custom-active');
    btnResetSupervisorDir.style.display = 'inline-block';
  } else {
    setLabel('Supervisor Folder: Same as Agent');
    btnSupervisorDir.title = 'Select a custom folder where (VIEW ONLY).xlsx copies are saved';
    btnSupervisorDir.classList.remove('custom-active');
    btnResetSupervisorDir.style.display = 'none';
  }
  updateFilePathsBarUI();
  updateStatusBarTelemetry();
}

function updateTemplatesDirUI() {
  const setLabel = (txt) => {
    if (templatesDirLabel) templatesDirLabel.textContent = txt;
    else btnTemplatesDir.textContent = txt;
  };
  if (templatesDirHandle) {
    const countSuffix = availableTemplates.length > 0 ? ` (${availableTemplates.length})` : '';
    setLabel(`Templates Folder: ${templatesDirHandle.name}${countSuffix}`);
    btnTemplatesDir.title = `Using templates from "${templatesDirHandle.name}". Click to change folder.`;
    btnTemplatesDir.classList.add('custom-active');
    btnResetTemplatesDir.style.display = 'inline-block';
  } else if (cachedTemplatesDirHandle) {
    setLabel(`Reconnect Templates: ${cachedTemplatesDirHandle.name}`);
    btnTemplatesDir.title = `Click to reconnect saved Templates folder "${cachedTemplatesDirHandle.name}"`;
    btnTemplatesDir.classList.add('custom-active');
    btnResetTemplatesDir.style.display = 'inline-block';
  } else {
    setLabel('Templates Folder: Built-in Default');
    btnTemplatesDir.title = 'Select folder containing master .xlsx templates for new monthly workbooks';
    btnTemplatesDir.classList.remove('custom-active');
    btnResetTemplatesDir.style.display = 'none';
  }
  updateFilePathsBarUI();
  updateStatusBarTelemetry();
}

function updateFilePathsBarUI() {
  const agentFolderName = parentDirHandle
    ? parentDirHandle.name
    : (cachedDirHandle ? `${cachedDirHandle.name} (cached)` : null);

  const supFolderName = supervisorDirHandle
    ? supervisorDirHandle.name
    : (cachedSupervisorDirHandle
        ? `${cachedSupervisorDirHandle.name} (cached)`
        : (parentDirHandle ? parentDirHandle.name : (cachedDirHandle ? `${cachedDirHandle.name} (cached)` : null)));

  const wbName = workingFileHandle
    ? workingFileHandle.name
    : (lastSavedWorkbookName || null);

  if (pathTrackerText) {
    if (parentDirHandle && workingFileHandle) {
      pathTrackerText.textContent = `${parentDirHandle.name}\\${workingFileHandle.name}`;
    } else if (parentDirHandle) {
      pathTrackerText.textContent = `${parentDirHandle.name}\\(No workbook selected)`;
    } else if (agentFolderName && wbName) {
      pathTrackerText.textContent = `${agentFolderName}\\${wbName}`;
    } else if (agentFolderName) {
      pathTrackerText.textContent = `${agentFolderName}\\`;
    } else {
      pathTrackerText.textContent = 'Not connected';
    }
  }

  if (pathViewOnlyText) {
    if ((supervisorDirHandle || parentDirHandle) && workingFileHandle) {
      const mirrorDir = supervisorDirHandle ? supervisorDirHandle.name : parentDirHandle.name;
      const mirrorFile = workingFileHandle.name.replace(/\.xlsx$/i, '') + ' (VIEW ONLY).xlsx';
      pathViewOnlyText.textContent = `${mirrorDir}\\${mirrorFile}`;
    } else if (supFolderName && wbName) {
      const mirrorFile = wbName.replace(/\.xlsx$/i, '') + ' (VIEW ONLY).xlsx';
      pathViewOnlyText.textContent = `${supFolderName}\\${mirrorFile}`;
    } else if (supFolderName) {
      pathViewOnlyText.textContent = `${supFolderName}\\*(VIEW ONLY).xlsx`;
    } else {
      pathViewOnlyText.textContent = 'Not connected';
    }
  }

  const tplName = templatesDirHandle
    ? `${templatesDirHandle.name}\\ (${availableTemplates.length} template${availableTemplates.length === 1 ? '' : 's'})`
    : (cachedTemplatesDirHandle ? `${cachedTemplatesDirHandle.name}\\ (cached)` : null);

  if (pathTemplatesItem && pathTemplatesSep && pathTemplatesText) {
    if (tplName) {
      pathTemplatesSep.style.display = 'inline';
      pathTemplatesItem.style.display = 'inline-flex';
      pathTemplatesText.textContent = tplName;
    } else {
      pathTemplatesSep.style.display = 'none';
      pathTemplatesItem.style.display = 'none';
    }
  }

  const canOpenViewOnly = Boolean((parentDirHandle || supervisorDirHandle) && workingFileHandle);
  if (btnOpenViewOnly) {
    btnOpenViewOnly.style.display = canOpenViewOnly ? 'inline-block' : 'none';
  }
  if (menuOpenViewOnly) {
    menuOpenViewOnly.disabled = !canOpenViewOnly;
  }
}

async function scanTemplatesDirectory() {
  availableTemplates = [];
  if (!templatesDirHandle) {
    updateTemplatesDirUI();
    return;
  }
  try {
    for await (const entry of templatesDirHandle.values()) {
      if (entry.kind !== 'file') continue;
      const name = entry.name;
      const lower = name.toLowerCase();
      if (lower.endsWith('.xlsx') && !name.startsWith('~$') && !lower.endsWith('(view only).xlsx')) {
        const f = await entry.getFile();
        availableTemplates.push({
          name,
          handle: entry,
          lastModified: f.lastModified
        });
      }
    }
    availableTemplates.sort((a, b) => a.name.localeCompare(b.name));
  } catch (err) {
    console.warn('Could not scan templates directory:', err);
  }
  updateTemplatesDirUI();
}

async function attemptReconnectFolders({ requestPermissionIfNeeded = true, silent = false } = {}) {
  if (!cachedDirHandle) return false;
  try {
    let perm = await cachedDirHandle.queryPermission({ mode: 'readwrite' });
    if (perm !== 'granted' && requestPermissionIfNeeded) {
      perm = await cachedDirHandle.requestPermission({ mode: 'readwrite' });
    }
    if (perm !== 'granted') {
      return false;
    }

    parentDirHandle = cachedDirHandle;
    autoReconnectGesturePending = false;
    btnReconnect.style.display = 'none';
    btnLinkFile.style.display = 'none';
    btnNewWorkbook.disabled = false;

    if (cachedSupervisorDirHandle && !supervisorDirHandle) {
      try {
        let supPerm = await cachedSupervisorDirHandle.queryPermission({ mode: 'readwrite' });
        if (supPerm !== 'granted' && requestPermissionIfNeeded) {
          supPerm = await cachedSupervisorDirHandle.requestPermission({ mode: 'readwrite' });
        }
        if (supPerm === 'granted') {
          supervisorDirHandle = cachedSupervisorDirHandle;
        }
      } catch (supErr) {
        console.warn('Could not reconnect supervisor folder:', supErr);
      }
    }

    if (cachedTemplatesDirHandle && !templatesDirHandle) {
      try {
        let tplPerm = await cachedTemplatesDirHandle.queryPermission({ mode: 'read' });
        if (tplPerm !== 'granted' && requestPermissionIfNeeded) {
          tplPerm = await cachedTemplatesDirHandle.requestPermission({ mode: 'read' });
        }
        if (tplPerm === 'granted') {
          templatesDirHandle = cachedTemplatesDirHandle;
          await scanTemplatesDirectory();
        }
      } catch (tplErr) {
        console.warn('Could not reconnect templates folder:', tplErr);
      }
    }

    updateAgentDirUI();
    updateSupervisorDirUI();
    updateTemplatesDirUI();

    await scanDirectoryAndLoadWorkbooks(lastSavedWorkbookName);
    return true;
  } catch (err) {
    if (!silent) {
      console.error('Reconnect folder error:', err);
      await showModalAlert({
        title: 'Could Not Reconnect Folder',
        message: err.message
      });
    }
    return false;
  }
}

async function initCachedDirectoryHandles() {
  if (!('indexedDB' in window)) return;
  const [savedAgentHandle, savedSupHandle, savedTplHandle, savedWbName] = await Promise.all([
    getHandleFromIDB(IDB_AGENT_KEY),
    getHandleFromIDB(IDB_SUPERVISOR_KEY),
    getHandleFromIDB(IDB_TEMPLATES_KEY),
    getHandleFromIDB(IDB_LAST_WORKBOOK_KEY)
  ]);

  if (savedWbName && typeof savedWbName === 'string') {
    lastSavedWorkbookName = savedWbName;
  }
  if (savedSupHandle) {
    cachedSupervisorDirHandle = savedSupHandle;
  }
  if (savedTplHandle) {
    cachedTemplatesDirHandle = savedTplHandle;
  }
  if (savedAgentHandle) {
    cachedDirHandle = savedAgentHandle;
    setReconnectButtonText(`Reconnect (${savedAgentHandle.name})`);
    btnReconnect.style.display = 'inline-flex';
  }

  updateAgentDirUI();
  updateSupervisorDirUI();
  updateTemplatesDirUI();

  if (savedAgentHandle && autoReconnectEnabled) {
    const connectedImmediately = await attemptReconnectFolders({
      requestPermissionIfNeeded: false,
      silent: true
    });
    if (!connectedImmediately) {
      // Browser requires a user gesture after full restart before requestPermission() can fire.
      autoReconnectGesturePending = true;
      setReconnectButtonText(`Auto-Reconnect (${savedAgentHandle.name})`);
      setStatus('Click anywhere to auto-reconnect folder', 'warning');

      const onFirstUserGesture = async (evt) => {
        if (!autoReconnectGesturePending || !autoReconnectEnabled || parentDirHandle) {
          document.removeEventListener('click', onFirstUserGesture);
          return;
        }
        // Ignore if user is explicitly unchecking the auto-reconnect toggle, changing theme/font, or picking a new folder
        if (
          evt.target.closest('#chkAutoReconnect') ||
          evt.target.closest('#chkAutoReconnectBar') ||
          evt.target.closest('#btnLinkFile') ||
          evt.target.closest('#menuSelectAgentFolder') ||
          evt.target.closest('#btnResetAgentDir') ||
          evt.target.closest('#headerThemeGroup') ||
          evt.target.closest('#menuThemeGroup') ||
          evt.target.closest('.menu-fontsize-control')
        ) {
          return;
        }
        document.removeEventListener('click', onFirstUserGesture);
        autoReconnectGesturePending = false;
        await attemptReconnectFolders({ requestPermissionIfNeeded: true, silent: true });
      };
      document.addEventListener('click', onFirstUserGesture);
    }
  }
}

// --- 1. DIRECTORY SELECTION, SUPERVISOR EXPORT FOLDER, TEMPLATES FOLDER & WORKBOOK DISCOVERY ---
async function promptSelectAgentFolder() {
  closeToolsMenu();
  try {
    if (typeof window.showDirectoryPicker !== 'function') {
      return showModalAlert({
        title: 'Browser Not Supported',
        message: 'Your browser does not support showDirectoryPicker. Please use Chrome or Edge over localhost or HTTPS.'
      });
    }
    const dirHandle = await window.showDirectoryPicker({ mode: 'readwrite' });
    parentDirHandle = dirHandle;
    cachedDirHandle = dirHandle;
    autoReconnectGesturePending = false;
    await saveHandleToIDB(IDB_AGENT_KEY, dirHandle);
    btnReconnect.style.display = 'none';
    btnLinkFile.style.display = 'none';
    btnNewWorkbook.disabled = false;
    updateAgentDirUI();
    await scanDirectoryAndLoadWorkbooks();
  } catch (err) {
    if (err.name !== 'AbortError') {
      await showModalAlert({
        title: 'Could Not Open Agent Folder',
        message: err.message
      });
    }
  }
}

btnLinkFile.addEventListener('click', promptSelectAgentFolder);
menuSelectAgentFolder.addEventListener('click', promptSelectAgentFolder);

if (btnResetAgentDir) {
  btnResetAgentDir.addEventListener('click', async () => {
    closeToolsMenu();
    parentDirHandle = null;
    cachedDirHandle = null;
    workingFileHandle = null;
    workbook = null;
    availableWorkbooks = [];
    lastSavedWorkbookName = null;
    autoReconnectGesturePending = false;
    await deleteHandleFromIDB(IDB_AGENT_KEY);
    await deleteHandleFromIDB(IDB_LAST_WORKBOOK_KEY);
    workbookPickerWrap.style.display = 'none';
    btnReconnect.style.display = 'none';
    btnLinkFile.style.display = 'inline-flex';
    btnNewEntry.style.display = 'none';
    btnRollover.style.display = 'none';
    if (btnRolloverPrevSheet) btnRolloverPrevSheet.disabled = true;
    selectedForDelete.clear();
    if (batchActionDock) batchActionDock.style.display = 'none';
    if (bottomEditorDock) bottomEditorDock.style.display = 'none';
    sheetTabsBar.style.display = 'none';
    schemaBanner.style.display = 'none';
    setStatus('No Folder Linked', '');
    updateAgentDirUI();
    listBody.innerHTML = '<tr><td colspan="12" class="empty-table-cell">Click "Select Agent Folder" to connect your tracker directory.</td></tr>';
    showToast('Cleared saved Tracker Workbooks folder.', 'info');
  });
}

btnReconnect.addEventListener('click', async () => {
  if (!cachedDirHandle) return;
  const ok = await attemptReconnectFolders({ requestPermissionIfNeeded: true, silent: false });
  if (!ok) {
    await showModalAlert({
      title: 'Permission Required',
      message: 'Read/Write permission was not granted for the folder.'
    });
  }
});

btnSupervisorDir.addEventListener('click', async () => {
  closeToolsMenu();
  try {
    if (typeof window.showDirectoryPicker !== 'function') {
      return showModalAlert({
        title: 'Browser Not Supported',
        message: 'Your browser does not support showDirectoryPicker. Please use Chrome or Edge over localhost or HTTPS.'
      });
    }

    if (cachedSupervisorDirHandle && !supervisorDirHandle) {
      let perm = await cachedSupervisorDirHandle.queryPermission({ mode: 'readwrite' });
      if (perm !== 'granted') {
        perm = await cachedSupervisorDirHandle.requestPermission({ mode: 'readwrite' });
      }
      if (perm === 'granted') {
        supervisorDirHandle = cachedSupervisorDirHandle;
        updateSupervisorDirUI();
        showToast(`Reconnected Supervisor export folder: "${supervisorDirHandle.name}"`, 'success');
        if (workbook && workingFileHandle) {
          const buf = await serializeWorkbookBuffer();
          await attemptMirrorWrite(buf, false);
        }
        return;
      }
    }

    const supDir = await window.showDirectoryPicker({ mode: 'readwrite' });
    supervisorDirHandle = supDir;
    cachedSupervisorDirHandle = supDir;
    await saveHandleToIDB(IDB_SUPERVISOR_KEY, supDir);
    updateSupervisorDirUI();
    showToast(`Supervisor (VIEW ONLY) export folder set to "${supDir.name}"`, 'success');

    if (workbook && workingFileHandle) {
      const buf = await serializeWorkbookBuffer();
      await attemptMirrorWrite(buf, false);
    }
  } catch (err) {
    if (err.name !== 'AbortError') {
      await showModalAlert({
        title: 'Supervisor Folder Error',
        message: 'Could not select Supervisor export folder: ' + err.message
      });
    }
  }
});

btnResetSupervisorDir.addEventListener('click', async () => {
  closeToolsMenu();
  supervisorDirHandle = null;
  cachedSupervisorDirHandle = null;
  await deleteHandleFromIDB(IDB_SUPERVISOR_KEY);
  updateSupervisorDirUI();
  showToast('Supervisor (VIEW ONLY) export reset to Agent workbook folder.', 'info');
});

if (btnTemplatesDir) {
  btnTemplatesDir.addEventListener('click', async () => {
    closeToolsMenu();
    try {
      if (typeof window.showDirectoryPicker !== 'function') {
        return showModalAlert({
          title: 'Browser Not Supported',
          message: 'Your browser does not support showDirectoryPicker. Please use Chrome or Edge over localhost or HTTPS.'
        });
      }

      if (cachedTemplatesDirHandle && !templatesDirHandle) {
        let perm = await cachedTemplatesDirHandle.queryPermission({ mode: 'read' });
        if (perm !== 'granted') {
          perm = await cachedTemplatesDirHandle.requestPermission({ mode: 'read' });
        }
        if (perm === 'granted') {
          templatesDirHandle = cachedTemplatesDirHandle;
          await scanTemplatesDirectory();
          showToast(`Reconnected Templates folder: "${templatesDirHandle.name}" (${availableTemplates.length} template(s))`, 'success');
          return;
        }
      }

      const tplDir = await window.showDirectoryPicker({ mode: 'read' });
      templatesDirHandle = tplDir;
      cachedTemplatesDirHandle = tplDir;
      await saveHandleToIDB(IDB_TEMPLATES_KEY, tplDir);
      await scanTemplatesDirectory();
      showToast(`Templates folder set to "${tplDir.name}" (${availableTemplates.length} .xlsx template(s) found)`, 'success');
    } catch (err) {
      if (err.name !== 'AbortError') {
        await showModalAlert({
          title: 'Templates Folder Error',
          message: 'Could not select Templates folder: ' + err.message
        });
      }
    }
  });
}

if (btnResetTemplatesDir) {
  btnResetTemplatesDir.addEventListener('click', async () => {
    closeToolsMenu();
    templatesDirHandle = null;
    cachedTemplatesDirHandle = null;
    availableTemplates = [];
    await deleteHandleFromIDB(IDB_TEMPLATES_KEY);
    updateTemplatesDirUI();
    showToast('Templates folder reset to built-in default schema.', 'info');
  });
}

function extractMonthYearWeight(fileName, fallbackLastModified) {
  const months = {
    january: 0, jan: 0,
    february: 1, feb: 1,
    march: 2, mar: 2,
    april: 3, apr: 3,
    may: 4,
    june: 5, jun: 5,
    july: 6, jul: 6,
    august: 7, aug: 7,
    september: 8, sep: 8, sept: 8,
    october: 9, oct: 9,
    november: 10, nov: 10,
    december: 11, dec: 11
  };
  const regex = /\b(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)[\s\-_]+(20\d{2})\b/i;
  const match = fileName.match(regex);
  if (match) {
    const mIdx = months[match[1].toLowerCase()];
    const year = parseInt(match[2], 10);
    if (mIdx !== undefined && !isNaN(year)) {
      return new Date(year, mIdx, 28).getTime();
    }
  }
  return fallbackLastModified;
}

async function scanDirectoryAndLoadWorkbooks(preserveFileName = null) {
  if (!parentDirHandle) return;
  setStatus(`Scanning ${parentDirHandle.name}...`, '');
  btnNewWorkbook.disabled = false;

  availableWorkbooks = [];
  for await (const entry of parentDirHandle.values()) {
    if (entry.kind !== 'file') continue;
    const name = entry.name;
    const lower = name.toLowerCase();

    if (
      lower.endsWith('.xlsx') &&
      !lower.endsWith('(view only).xlsx') &&
      !name.startsWith('~$')
    ) {
      try {
        const fileObj = await entry.getFile();
        const sortWeight = extractMonthYearWeight(name, fileObj.lastModified);
        availableWorkbooks.push({
          name,
          handle: entry,
          lastModified: fileObj.lastModified,
          sortWeight
        });
      } catch (e) {
        console.warn('Skipping unreadable file:', name, e);
      }
    }
  }

  if (availableWorkbooks.length === 0) {
    workbookPickerWrap.style.display = 'none';
    workingFileHandle = null;
    workbook = null;
    sheetTabsBar.style.display = 'none';
    schemaBanner.style.display = 'none';
    setStatus(`No working .xlsx in ${parentDirHandle.name}`, 'warning');
    listBody.innerHTML = `
      <tr>
        <td colspan="12" style="text-align:center; color:var(--text-muted); padding: 36px;">
          <div style="margin-bottom:12px;">
            No working <code>.xlsx</code> files found in folder <strong>${escapeHtml(parentDirHandle.name)}</strong>.
          </div>
          <button type="button" id="btnEmptyStateBootstrap" class="btn-primary">
            ✨ Bootstrap New Tracker Workbook (.xlsx)
          </button>
        </td>
      </tr>
    `;
    const emptyBtn = document.getElementById('btnEmptyStateBootstrap');
    if (emptyBtn) {
      emptyBtn.addEventListener('click', promptCreateNewWorkbook);
    }
    return;
  }

  availableWorkbooks.sort((a, b) => {
    if (b.sortWeight !== a.sortWeight) return b.sortWeight - a.sortWeight;
    return b.lastModified - a.lastModified;
  });

  workbookSelect.innerHTML = '';
  availableWorkbooks.forEach((wb, idx) => {
    const opt = document.createElement('option');
    opt.value = String(idx);
    opt.textContent = wb.name;
    workbookSelect.appendChild(opt);
  });

  workbookPickerWrap.style.display = availableWorkbooks.length >= 1 ? 'flex' : 'none';

  let selectedIdx = 0;
  if (preserveFileName) {
    const foundIdx = availableWorkbooks.findIndex(w => w.name === preserveFileName);
    if (foundIdx !== -1) selectedIdx = foundIdx;
  }

  workbookSelect.value = String(selectedIdx);
  workingFileHandle = availableWorkbooks[selectedIdx].handle;
  lastSavedWorkbookName = workingFileHandle.name;
  await saveHandleToIDB(IDB_LAST_WORKBOOK_KEY, workingFileHandle.name);
  updateFilePathsBarUI();

  await loadWorkbookFromDisk();
}

workbookSelect.addEventListener('change', async () => {
  const idx = Number(workbookSelect.value);
  if (availableWorkbooks[idx]) {
    commitActiveScratchInlineEdit();
    if (hasAnyPendingScratchMutations()) {
      await flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: true });
    }
    resetScratchWorkbookRuntimeState();
    activeWorkspaceTab = DAILY_TRACKER_WS_TAB;
    workingFileHandle = availableWorkbooks[idx].handle;
    lastSavedWorkbookName = workingFileHandle.name;
    await saveHandleToIDB(IDB_LAST_WORKBOOK_KEY, workingFileHandle.name);
    clearTransientFiltersOnNavigation({ resetScopeToCurrent: true });
    dockEditingTarget = null;
    clearMirrorRetryState();
    setDeleteMode(false);
    updateFilePathsBarUI();
    try {
      await loadWorkbookFromDisk();
    } catch (err) {
      console.error('Error loading selected workbook:', err);
      await showModalAlert({
        title: 'Could Not Load Workbook',
        message: err.message
      });
    }
  }
});

// --- 2. READ WORKBOOK, PARSE REFERENCE SHEETS, CHECK SCHEMA & INDEX ALL DAILY TABS ---
const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32Bytes(u8) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < u8.length; i++) {
    crc = CRC32_TABLE[(crc ^ u8[i]) & 0xFF] ^ (crc >>> 8);
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

async function inflateRawBytes(compBytes) {
  const ds = new DecompressionStream('deflate-raw');
  const writer = ds.writable.getWriter();
  writer.write(compBytes);
  writer.close();
  const ab = await new Response(ds.readable).arrayBuffer();
  return new Uint8Array(ab);
}

async function deflateRawBytes(rawBytes) {
  const cs = new CompressionStream('deflate-raw');
  const writer = cs.writable.getWriter();
  writer.write(rawBytes);
  writer.close();
  const ab = await new Response(cs.readable).arrayBuffer();
  return new Uint8Array(ab);
}

async function unpackZipArchive(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const len = bytes.length;

  let eocdOffset = -1;
  for (let i = len - 22; i >= Math.max(0, len - 65557); i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      eocdOffset = i;
      break;
    }
  }
  if (eocdOffset === -1) {
    throw new Error('Invalid XLSX archive: End of Central Directory not found.');
  }

  const totalEntries = view.getUint16(eocdOffset + 10, true);
  const cdOffset = view.getUint32(eocdOffset + 16, true);
  const entries = new Map();
  const dec = new TextDecoder('utf-8');

  let ptr = cdOffset;
  for (let e = 0; e < totalEntries; e++) {
    if (ptr + 46 > len || view.getUint32(ptr, true) !== 0x02014b50) break;
    const compMethod = view.getUint16(ptr + 10, true);
    const compSize = view.getUint32(ptr + 20, true);
    const nameLen = view.getUint16(ptr + 28, true);
    const extraLen = view.getUint16(ptr + 30, true);
    const commentLen = view.getUint16(ptr + 32, true);
    const localHeaderOffset = view.getUint32(ptr + 42, true);

    const name = dec.decode(bytes.subarray(ptr + 46, ptr + 46 + nameLen)).replace(/\\/g, '/');
    const localNameLen = view.getUint16(localHeaderOffset + 26, true);
    const localExtraLen = view.getUint16(localHeaderOffset + 28, true);
    const dataStart = localHeaderOffset + 30 + localNameLen + localExtraLen;
    const compData = bytes.subarray(dataStart, dataStart + compSize);

    let rawData;
    if (compMethod === 0) {
      rawData = compData.slice();
    } else if (compMethod === 8) {
      rawData = await inflateRawBytes(compData);
    } else {
      rawData = compData.slice();
    }
    entries.set(name, rawData);
    ptr += 46 + nameLen + extraLen + commentLen;
  }

  return entries;
}

async function packZipArchive(entriesMap) {
  const enc = new TextEncoder();
  const localParts = [];
  const cdParts = [];
  let localOffset = 0;
  let totalCdBytes = 0;

  for (const [name, rawData] of entriesMap.entries()) {
    const nameBytes = enc.encode(name);
    const u8Raw = rawData instanceof Uint8Array ? rawData : enc.encode(String(rawData));
    const crc = crc32Bytes(u8Raw);
    const compData = await deflateRawBytes(u8Raw);
    const compSize = compData.length;
    const uncompSize = u8Raw.length;

    const lfh = new Uint8Array(30 + nameBytes.length);
    const lfhView = new DataView(lfh.buffer);
    lfhView.setUint32(0, 0x04034b50, true);
    lfhView.setUint16(4, 20, true);
    lfhView.setUint16(6, 0, true);
    lfhView.setUint16(8, 8, true); // deflate
    lfhView.setUint16(10, 0, true);
    lfhView.setUint16(12, 0x21, true);
    lfhView.setUint32(14, crc, true);
    lfhView.setUint32(18, compSize, true);
    lfhView.setUint32(22, uncompSize, true);
    lfhView.setUint16(26, nameBytes.length, true);
    lfhView.setUint16(28, 0, true);
    lfh.set(nameBytes, 30);

    localParts.push(lfh, compData);

    const cdfh = new Uint8Array(46 + nameBytes.length);
    const cdView = new DataView(cdfh.buffer);
    cdView.setUint32(0, 0x02014b50, true);
    cdView.setUint16(4, 20, true);
    cdView.setUint16(6, 20, true);
    cdView.setUint16(8, 0, true);
    cdView.setUint16(10, 8, true);
    cdView.setUint16(12, 0, true);
    cdView.setUint16(14, 0x21, true);
    cdView.setUint32(16, crc, true);
    cdView.setUint32(20, compSize, true);
    cdView.setUint32(24, uncompSize, true);
    cdView.setUint16(28, nameBytes.length, true);
    cdView.setUint16(30, 0, true);
    cdView.setUint16(32, 0, true);
    cdView.setUint16(34, 0, true);
    cdView.setUint16(36, 0, true);
    cdView.setUint32(38, 0, true);
    cdView.setUint32(42, localOffset, true);
    cdfh.set(nameBytes, 46);

    cdParts.push(cdfh);
    totalCdBytes += cdfh.length;
    localOffset += lfh.length + compData.length;
  }

  const eocd = new Uint8Array(22);
  const eocdView = new DataView(eocd.buffer);
  const count = cdParts.length;
  eocdView.setUint32(0, 0x06054b50, true);
  eocdView.setUint16(4, 0, true);
  eocdView.setUint16(6, 0, true);
  eocdView.setUint16(8, count, true);
  eocdView.setUint16(10, count, true);
  eocdView.setUint32(12, totalCdBytes, true);
  eocdView.setUint32(16, localOffset, true);
  eocdView.setUint16(20, 0, true);

  const out = new Uint8Array(localOffset + totalCdBytes + 22);
  let pos = 0;
  for (const part of localParts) {
    out.set(part, pos);
    pos += part.length;
  }
  for (const part of cdParts) {
    out.set(part, pos);
    pos += part.length;
  }
  out.set(eocd, pos);
  return out.buffer;
}

const STANDARD_DYNAMIC_ARRAY_METADATA_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray">' +
  '<metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes>' +
  '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>' +
  '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>';

const WORKBOOK_CALC_FEATURES_EXT =
  '<ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures">' +
  '<xcalcf:calcFeatures>' +
  '<xcalcf:feature name="microsoft.com:RD"/>' +
  '<xcalcf:feature name="microsoft.com:Single"/>' +
  '<xcalcf:feature name="microsoft.com:FV"/>' +
  '<xcalcf:feature name="microsoft.com:CNMTM"/>' +
  '<xcalcf:feature name="microsoft.com:LET_WF"/>' +
  '<xcalcf:feature name="microsoft.com:LAMBDA_WF"/>' +
  '<xcalcf:feature name="microsoft.com:ARRAYTEXT_WF"/>' +
  '</xcalcf:calcFeatures></ext>';

const DAILY_SHEET_DATA_VALIDATIONS_EXTLST =
  '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">' +
  '<x14:dataValidations count="2" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">' +
  '<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1">' +
  '<x14:formula1><xm:f>\'Validation List\'!$E$2:$E$100</xm:f></x14:formula1><xm:sqref>I2:I101</xm:sqref>' +
  '</x14:dataValidation>' +
  '<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1">' +
  '<x14:formula1><xm:f>\'Validation List\'!$A$2:$A$50</xm:f></x14:formula1><xm:sqref>F2:F108</xm:sqref>' +
  '</x14:dataValidation>' +
  '</x14:dataValidations></ext></extLst>';

const DEFAULT_EXTERNAL_LINK_RELS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="LKQ-Bidder%20Support%20shared%20lists.xlsx" TargetMode="External"/>' +
  '</Relationships>';

function parseWorksheetPathsFromZipEntries(entries) {
  const dec = new TextDecoder('utf-8');
  const wbBytes = entries.get('xl/workbook.xml');
  const relsBytes = entries.get('xl/_rels/workbook.xml.rels');
  const map = new Map(); // sheetName -> zipPath
  if (!wbBytes || !relsBytes) return map;

  const wbXml = dec.decode(wbBytes);
  const relsXml = dec.decode(relsBytes);

  const rIdToTarget = new Map();
  const relRegex = /<Relationship\b[^>]*>/g;
  let m;
  while ((m = relRegex.exec(relsXml)) !== null) {
    const tag = m[0];
    const idMatch = tag.match(/\bId="([^"]+)"/);
    const targetMatch = tag.match(/\bTarget="([^"]+)"/);
    if (idMatch && targetMatch) {
      let target = targetMatch[1].replace(/^\/+/, '');
      if (!target.startsWith('xl/')) target = 'xl/' + target;
      rIdToTarget.set(idMatch[1], target);
    }
  }

  const sheetRegex = /<sheet\b[^>]*>/g;
  while ((m = sheetRegex.exec(wbXml)) !== null) {
    const tag = m[0];
    const nameMatch = tag.match(/\bname="([^"]+)"/);
    const rIdMatch = tag.match(/\br:id="([^"]+)"/);
    if (nameMatch && rIdMatch) {
      const sName = nameMatch[1]
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'");
      const targetPath = rIdToTarget.get(rIdMatch[1]);
      if (targetPath && entries.has(targetPath)) {
        map.set(sName, targetPath);
      }
    }
  }
  return map;
}

async function captureOriginalArchiveState(arrayBuffer) {
  try {
    const entries = await unpackZipArchive(arrayBuffer);
    const dec = new TextDecoder('utf-8');
    const externalLinkFiles = new Map();
    for (const [k, v] of entries.entries()) {
      if (k.startsWith('xl/externalLinks/')) {
        externalLinkFiles.set(k, v);
      }
    }
    const metadataXmlBytes = entries.get('xl/metadata.xml') || null;
    const sheetPaths = parseWorksheetPathsFromZipEntries(entries);
    const arrayFormulasBySheet = new Map(); // lowerSheetName -> Map(cellRef -> { fTag, cm, t })
    let templateExtLstXml = null;

    for (const [sheetName, zipPath] of sheetPaths.entries()) {
      const lower = sheetName.trim().toLowerCase();
      const xml = dec.decode(entries.get(zipPath));

      if (lower === 'template' || isTemplatePlaceholderTab(lower) || DATE_TAB_REGEX.test(sheetName.trim())) {
        const extMatch = xml.match(/<extLst>[\s\S]*?<\/extLst>/);
        if (extMatch && extMatch[0].includes('dataValidations') && !templateExtLstXml) {
          templateExtLstXml = extMatch[0].replace(/\s+xr:uid="[^"]*"/g, '');
        }
      }

      const cellMap = new Map();
      const cellRegex = /<c\b([^>]*\br="([A-Z]+\d+)"[^>]*)>(<f\b[^>]*>[\s\S]*?<\/f>)(<v[^>]*>[\s\S]*?<\/v>|<v\/>)?<\/c>/g;
      let cm;
      while ((cm = cellRegex.exec(xml)) !== null) {
        const cAttrs = cm[1];
        const cellRef = cm[2];
        const fTag = cm[3];
        if (/\bt="array"/.test(fTag) || /\bcm="/.test(cAttrs) || /_xlfn\._xlws\.FILTER|_xlfn\.LET/.test(fTag)) {
          const tAttr = (cAttrs.match(/\bt="([^"]+)"/) || [])[1] || null;
          cellMap.set(cellRef, { fTag, tAttr });
        }
      }
      if (cellMap.size > 0) {
        arrayFormulasBySheet.set(lower, cellMap);
      }
    }

    return {
      externalLinkFiles,
      metadataXmlBytes,
      arrayFormulasBySheet,
      templateExtLstXml: templateExtLstXml || DAILY_SHEET_DATA_VALIDATIONS_EXTLST
    };
  } catch (err) {
    console.warn('Could not capture original archive metadata:', err);
    return null;
  }
}

function parseSharedStringsFromXml(ssXml) {
  if (!ssXml) return [];
  const list = [];
  const siRegex = /<si\b[^>]*>([\s\S]*?)<\/si>/g;
  let m;
  while ((m = siRegex.exec(ssXml)) !== null) {
    const inner = m[1];
    const tParts = [];
    const tRegex = /<t\b[^>]*>([\s\S]*?)<\/t>/g;
    let tm;
    while ((tm = tRegex.exec(inner)) !== null) {
      tParts.push(tm[1]);
    }
    list.push(tParts.join(''));
  }
  return list;
}

function synthesizeExternalLinkXmlFromSheets(validationXml, masterXml) {
  const extractRowsForExternalSheet = (sheetXml) => {
    if (!sheetXml) return '';
    const rowsOut = [];
    const rowRegex = /<row\b[^>]*\br="(\d+)"[^>]*>([\s\S]*?)<\/row>/g;
    let rm;
    while ((rm = rowRegex.exec(sheetXml)) !== null) {
      const rNum = rm[1];
      const rInner = rm[2];
      const cellsOut = [];
      const cRegex = /<c\b([^>]*\br="([A-Z]+\d+)"[^>]*)>(?:<f\b[^>]*>[\s\S]*?<\/f>)?(?:<v([^>]*)>([\s\S]*?)<\/v>|<v\/>)?<\/c>/g;
      let cm;
      while ((cm = cRegex.exec(rInner)) !== null) {
        const cAttrs = cm[1];
        const cRef = cm[2];
        const vAttrs = cm[3] || '';
        const vVal = cm[4];
        if (vVal === undefined || vVal === '') continue;
        const isStr = /\bt="str"/.test(cAttrs);
        cellsOut.push(`<cell r="${cRef}"${isStr ? ' t="str"' : ''}><v${vAttrs}>${vVal}</v></cell>`);
      }
      if (cellsOut.length > 0) {
        rowsOut.push(`<row r="${rNum}">${cellsOut.join('')}</row>`);
      }
    }
    return rowsOut.join('');
  };

  const valRows = extractRowsForExternalSheet(validationXml);
  const masterRows = extractRowsForExternalSheet(masterXml);

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1">' +
    '<sheetNames><sheetName val="Validation List"/><sheetName val="Master Account List"/></sheetNames>' +
    '<sheetDataSet>' +
    `<sheetData sheetId="0">${valRows}</sheetData>` +
    `<sheetData sheetId="1">${masterRows}</sheetData>` +
    '</sheetDataSet>' +
    '</externalBook></externalLink>'
  );
}

/**
 * Post-processes the XLSX archive produced by ExcelJS so Desktop Excel never flags
 * the workbook as corrupted or strips formulas:
 * 1. Restores dynamic array metadata (`xl/metadata.xml` + `cm="1"` on array cells) and
 *    converts ghost spill cells on `Master Accounts` & `Validation List` from `t="s"` back to `t="str"`.
 * 2. Restores `xl/externalLinks/*` and `<externalReferences>` when `[1]` external links exist.
 * 3. Ensures `Column C` XLOOKUP formula cells on `Template` and daily sheets have `t="str"` and `<v/>`.
 * 4. Restores `<extLst>` cross-sheet Data Validation dropdowns (`Requested Action` & `Status`) on daily sheets.
 */
async function repairSavedXlsxArchive(excelJsBuffer, archiveState) {
  try {
    const entries = await unpackZipArchive(excelJsBuffer);
    const dec = new TextDecoder('utf-8');
    const enc = new TextEncoder();

    const ssBytes = entries.get('xl/sharedStrings.xml');
    const sharedStrings = ssBytes ? parseSharedStringsFromXml(dec.decode(ssBytes)) : [];
    const sheetPaths = parseWorksheetPathsFromZipEntries(entries);

    let hasAnyDynamicArrayFormula = false;
    let hasAnyExternalRefFormula = false;
    let hasValidationSheetInWb = false;
    let repairedValidationXml = '';
    let repairedMasterXml = '';

    for (const [sName] of sheetPaths.entries()) {
      const low = sName.trim().toLowerCase();
      if (low === 'validation list' || low === 'validation lists') {
        hasValidationSheetInWb = true;
      }
    }

    for (const [sheetName, zipPath] of sheetPaths.entries()) {
      const lower = sheetName.trim().toLowerCase();
      let xml = dec.decode(entries.get(zipPath));

      const isDynamicArrayRefSheet =
        lower === 'master accounts' ||
        lower === 'master account' ||
        lower === 'validation list' ||
        lower === 'validation lists' ||
        /\bt="array"|_xlfn\._xlws\.FILTER|_xlfn\.LET/.test(xml);

      if (isDynamicArrayRefSheet) {
        const origCellMap = archiveState && archiveState.arrayFormulasBySheet
          ? archiveState.arrayFormulasBySheet.get(lower)
          : null;

        // 1. Restore cm="1" and exact <f t="array" ref="..."> on dynamic array anchor cells
        xml = xml.replace(
          /<c\b([^>]*\br="([A-Z]+\d+)"[^>]*)>(<f\b[^>]*>[\s\S]*?<\/f>)(<v[^>]*>([\s\S]*?)<\/v>|<v\/>)?<\/c>/g,
          (full, cAttrs, cellRef, fTag, vFull, vInner) => {
            const origInfo = origCellMap ? origCellMap.get(cellRef) : null;
            const isArrayF = Boolean(
              origInfo ||
              /\bt="array"/.test(fTag) ||
              /_xlfn\._xlws\.FILTER|_xlfn\.LET/.test(fTag)
            );
            if (!isArrayF) return full;

            hasAnyDynamicArrayFormula = true;
            const finalFTag = origInfo ? origInfo.fTag : fTag;
            if (/\[1\]/.test(finalFTag)) {
              hasAnyExternalRefFormula = true;
            }

            let attrs = cAttrs;
            if (!/\bcm="/.test(attrs)) {
              attrs += ' cm="1"';
            }
            const isNumVal = vInner !== undefined && vInner !== '' && !isNaN(Number(vInner));
            if (origInfo && origInfo.tAttr && !/\bt="/.test(attrs)) {
              attrs += ` t="${origInfo.tAttr}"`;
            } else if (!isNumVal && !/\bt="/.test(attrs)) {
              attrs += ' t="str"';
            }
            return `<c${attrs}>${finalFTag}${vFull || '<v/>'}</c>`;
          }
        );

        // 2. Convert spilled string cells in rows 2+ from t="s" (shared string) back to t="str" (formula string result)
        xml = xml.replace(
          /<row\b([^>]*\br="(\d+)"[^>]*)>([\s\S]*?)<\/row>/g,
          (rowFull, rowAttrs, rNumStr, rowInner) => {
            const rNum = parseInt(rNumStr, 10);
            if (rNum < 2) return rowFull;
            const fixedInner = rowInner.replace(
              /<c\b([^>]*?)\bt="s"([^>]*)><v>(\d+)<\/v><\/c>/g,
              (cFull, a1, a2, idxStr) => {
                const idx = parseInt(idxStr, 10);
                const resolved = sharedStrings[idx] !== undefined ? sharedStrings[idx] : '';
                const needsSpace = resolved.startsWith(' ') || resolved.endsWith(' ');
                return `<c${a1}t="str"${a2}><v${needsSpace ? ' xml:space="preserve"' : ''}>${resolved}</v></c>`;
              }
            );
            return `<row${rowAttrs}>${fixedInner}</row>`;
          }
        );

        if (lower === 'validation list' || lower === 'validation lists') {
          repairedValidationXml = xml;
        } else if (lower === 'master accounts' || lower === 'master account') {
          repairedMasterXml = xml;
        }
      }

      // 3. For Template & Daily Sheets: ensure Column C XLOOKUP formulas have t="str" and <v/>, and restore Data Validations <extLst>
      const isDailyOrTemplate =
        lower === 'template' ||
        isTemplatePlaceholderTab(lower) ||
        DATE_TAB_REGEX.test(sheetName.trim());

      if (isDailyOrTemplate) {
        xml = xml.replace(
          /<c\b([^>]*\br="C(\d+)"[^>]*)>(<f\b[^>]*>[\s\S]*?<\/f>)(<v[^>]*>[\s\S]*?<\/v>|<v\/>)?<\/c>/g,
          (full, cAttrs, rNumStr, fTag, vFull) => {
            let cleanAttrs = cAttrs.replace(/\s+cm="[^"]*"/g, '');
            if (/\bt="[^"]*"/.test(cleanAttrs)) {
              cleanAttrs = cleanAttrs.replace(/\bt="[^"]*"/, 't="str"');
            } else {
              cleanAttrs += ' t="str"';
            }
            const cleanFTag = fTag.replace(/^<f\b[^>]*>/, '<f>');
            let cleanV = vFull || '<v/>';
            if (cleanV === '<v></v>') cleanV = '<v/>';
            return `<c${cleanAttrs}>${cleanFTag}${cleanV}</c>`;
          }
        );

        if (hasValidationSheetInWb && !xml.includes('dataValidations')) {
          const extLstToInject = (archiveState && archiveState.templateExtLstXml)
            ? archiveState.templateExtLstXml
            : DAILY_SHEET_DATA_VALIDATIONS_EXTLST;
          if (!xml.includes('<extLst>')) {
            xml = xml.replace(/<\/worksheet>\s*$/, `${extLstToInject}</worksheet>`);
          }
        }
      }

      entries.set(zipPath, enc.encode(xml));
    }

    // 4. Remove stale calcChain.xml if present so Excel cleanly rebuilds the calculation chain
    entries.delete('xl/calcChain.xml');

    // 5. Restore or synthesize xl/externalLinks/* and xl/metadata.xml when needed
    let ctXml = dec.decode(entries.get('[Content_Types].xml'));
    let relsXml = dec.decode(entries.get('xl/_rels/workbook.xml.rels'));
    let wbXml = dec.decode(entries.get('xl/workbook.xml'));

    ctXml = ctXml.replace(/<Override[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g, '');
    relsXml = relsXml.replace(/<Relationship[^>]*Target="calcChain\.xml"[^>]*\/>/g, '');

    let maxRId = 0;
    const rIdMatches = relsXml.match(/\bId="rId(\d+)"/g) || [];
    for (const rm of rIdMatches) {
      const n = parseInt(rm.replace(/\D+/g, ''), 10);
      if (n > maxRId) maxRId = n;
    }

    if (hasAnyExternalRefFormula) {
      if (archiveState && archiveState.externalLinkFiles && archiveState.externalLinkFiles.size > 0) {
        for (const [extPath, extBytes] of archiveState.externalLinkFiles.entries()) {
          entries.set(extPath, extBytes);
        }
      } else {
        entries.set(
          'xl/externalLinks/externalLink1.xml',
          enc.encode(synthesizeExternalLinkXmlFromSheets(repairedValidationXml, repairedMasterXml))
        );
        entries.set(
          'xl/externalLinks/_rels/externalLink1.xml.rels',
          enc.encode(DEFAULT_EXTERNAL_LINK_RELS_XML)
        );
      }

      if (!ctXml.includes('/xl/externalLinks/externalLink1.xml')) {
        ctXml = ctXml.replace(
          /<\/Types>\s*$/,
          '<Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/></Types>'
        );
      }

      let extRId = null;
      const existingExtRel = relsXml.match(/<Relationship\b[^>]*\bId="(rId\d+)"[^>]*externalLinks\/externalLink1\.xml[^>]*\/>/);
      if (existingExtRel) {
        extRId = existingExtRel[1];
      } else {
        extRId = `rId${++maxRId}`;
        relsXml = relsXml.replace(
          /<\/Relationships>\s*$/,
          `<Relationship Id="${extRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>`
        );
      }

      if (!wbXml.includes('<externalReferences')) {
        wbXml = wbXml.replace(
          /<\/sheets>/,
          `</sheets><externalReferences><externalReference r:id="${extRId}"/></externalReferences>`
        );
      }
    }

    if (hasAnyDynamicArrayFormula) {
      const metaBytes = (archiveState && archiveState.metadataXmlBytes)
        ? archiveState.metadataXmlBytes
        : enc.encode(STANDARD_DYNAMIC_ARRAY_METADATA_XML);
      entries.set('xl/metadata.xml', metaBytes);

      if (!ctXml.includes('/xl/metadata.xml')) {
        ctXml = ctXml.replace(
          /<\/Types>\s*$/,
          '<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/></Types>'
        );
      }

      if (!relsXml.includes('Target="metadata.xml"')) {
        const metaRId = `rId${++maxRId}`;
        relsXml = relsXml.replace(
          /<\/Relationships>\s*$/,
          `<Relationship Id="${metaRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/></Relationships>`
        );
      }
    }

    if (!wbXml.includes('xcalcf:calcFeatures')) {
      if (wbXml.includes('</extLst>')) {
        wbXml = wbXml.replace(/<\/extLst>/, `${WORKBOOK_CALC_FEATURES_EXT}</extLst>`);
      } else {
        wbXml = wbXml.replace(/<\/workbook>\s*$/, `<extLst>${WORKBOOK_CALC_FEATURES_EXT}</extLst></workbook>`);
      }
    }

    entries.set('[Content_Types].xml', enc.encode(ctXml));
    entries.set('xl/_rels/workbook.xml.rels', enc.encode(relsXml));
    entries.set('xl/workbook.xml', enc.encode(wbXml));

    return await packZipArchive(entries);
  } catch (err) {
    console.warn('OpenXML post-processor warning (falling back to ExcelJS buffer):', err);
    return excelJsBuffer;
  }
}

/**
 * Neutralizes docProps/core.xml and docProps/app.xml in the in-memory ZIP headers
 * before passing to ExcelJS. ExcelJS's CoreXform / AppXform parsers throw
 * `Unexpected xml node in parseOpen: {"name":"lastModifiedBy",...}` when reading
 * workbooks whose docProps/core.xml uses a default XML namespace without `cp:` prefixes.
 */
function sanitizeXlsxBufferForExcelJS(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer.slice(0));
  const len = bytes.length;

  for (let i = 0; i <= len - 30; i++) {
    if (bytes[i] !== 0x50 || bytes[i + 1] !== 0x4b) continue;

    const b2 = bytes[i + 2];
    const b3 = bytes[i + 3];
    let nameLen = 0;
    let nameOffset = 0;

    if (b2 === 0x03 && b3 === 0x04) {
      nameLen = bytes[i + 26] | (bytes[i + 27] << 8);
      nameOffset = i + 30;
    } else if (b2 === 0x01 && b3 === 0x02 && i + 46 <= len) {
      nameLen = bytes[i + 28] | (bytes[i + 29] << 8);
      nameOffset = i + 46;
    } else {
      continue;
    }

    if (nameLen <= 0 || nameLen > 256 || nameOffset + nameLen > len) continue;

    let rawName = '';
    for (let k = 0; k < nameLen; k++) {
      rawName += String.fromCharCode(bytes[nameOffset + k]);
    }
    const normName = rawName.replace(/\\/g, '/').replace(/^\/+/, '').toLowerCase();

    if (normName === 'docprops/core.xml') {
      const replacement = 'skip.xml';
      const targetStart = nameOffset + nameLen - replacement.length;
      for (let k = 0; k < replacement.length; k++) {
        bytes[targetStart + k] = replacement.charCodeAt(k);
      }
    } else if (normName === 'docprops/app.xml') {
      const replacement = 'skp.xml';
      const targetStart = nameOffset + nameLen - replacement.length;
      for (let k = 0; k < replacement.length; k++) {
        bytes[targetStart + k] = replacement.charCodeAt(k);
      }
    }
  }

  return bytes.buffer;
}

function getMasterAccountsWorksheet(wb = workbook) {
  if (!wb || typeof wb.eachSheet !== 'function') return null;
  let found = null;
  wb.eachSheet((ws) => {
    const lower = String(ws.name || '').trim().toLowerCase();
    if (lower === 'master accounts' || lower === 'master account') {
      found = ws;
    }
  });
  return found;
}

function getMockMasterAccountsList() {
  if (typeof window !== 'undefined' && Array.isArray(window.MOCK_MASTER_ACCOUNTS)) {
    return window.MOCK_MASTER_ACCOUNTS.map(item => ({
      broker: String(item.broker || '').trim(),
      master: String(item.master || '').trim().toUpperCase(),
      bidder: String(item.bidder || '').trim().toUpperCase(),
      yard: String(item.yard || '').trim().toUpperCase()
    })).filter(m => m.broker || m.master);
  }
  return [];
}

function updateDevMockAccountsUI() {
  if (chkDevMockAccounts) {
    chkDevMockAccounts.checked = Boolean(useDevMockAccounts);
  }
  if (devMockStateBadge) {
    devMockStateBadge.style.display = useDevMockAccounts ? 'inline-block' : 'none';
    devMockStateBadge.textContent = isLocalhostDevEnvironment() ? 'LOCALHOST MOCK' : 'DEV MOCK';
  }
}

async function loadWorkbookFromDisk(preserveSelectedSheet = null, options = {}) {
  if (!workingFileHandle) return;
  commitActiveScratchInlineEdit();
  if (hasAnyPendingScratchMutations() && !options.skipFlushPendingScratch && workbook) {
    await flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: true });
  }
  if (!options.preserveStatus && !pendingMirrorBuffer) {
    setStatus('Reading...', '');
  }

  const file = await workingFileHandle.getFile();
  const arrayBuffer = await file.arrayBuffer();
  originalArchiveState = await captureOriginalArchiveState(arrayBuffer);
  const sanitizedBuffer = sanitizeXlsxBufferForExcelJS(arrayBuffer);

  const loadedWorkbook = new ExcelJS.Workbook();
  await loadedWorkbook.xlsx.load(sanitizedBuffer);
  workbook = loadedWorkbook;
  reapplyPendingScratchEditsToWorkbook(workbook);

  allRows = [];
  dateSheets = [];
  scratchSheets = [];
  masterAccountsMap = [];
  hasMasterAccountsSheet = false;
  validationActions = [];
  validationStatuses = [];

  let hasTemplate = false;
  let hasValidation = false;
  let placeholderDateSheet = null;
  const dailySheetObjs = [];

  // Safety check for optional Master Accounts sheet (removed from new monthly workbook bootstrap)
  const masterAccountsWs = getMasterAccountsWorksheet(workbook);
  hasMasterAccountsSheet = Boolean(masterAccountsWs);

  if (useDevMockAccounts) {
    // Dev / Localhost Mock Mode: bypass Excel 'Master Accounts' sheet lookup and populate dummy test records
    masterAccountsMap = getMockMasterAccountsList();
  } else if (masterAccountsWs) {
    parseMasterAccountsSheet(masterAccountsWs);
  }

  // Pass 1: Parse Reference Sheets, collect daily MM-DD-YYYY tabs, and discover custom Scratch Sheets
  workbook.eachSheet((worksheet) => {
    const rawName = worksheet.name.trim();
    const lowerName = rawName.toLowerCase();

    if (DATE_TAB_REGEX.test(rawName)) {
      dateSheets.push(rawName);
      dailySheetObjs.push({ worksheet, rawName });
    } else if (isTemplatePlaceholderTab(rawName)) {
      placeholderDateSheet = worksheet;
    } else if (lowerName === 'template') {
      hasTemplate = true;
    } else if (lowerName === 'validation lists' || lowerName === 'validation list') {
      hasValidation = true;
      parseValidationSheet(worksheet);
    } else if (isCustomScratchSheetName(rawName, worksheet.state)) {
      const bounds = measureScratchWorksheetBounds(worksheet);
      scratchSheets.push({
        name: worksheet.name,
        nonEmptyRows: bounds.nonEmptyRows,
        maxRow: bounds.maxRow,
        maxCol: bounds.maxCol
      });
    }
  });

  // Pass 2: Sort Daily MM-DD-YYYY Tabs oldest -> newest before indexing so un-keyed cross-day rows inherit prior Tracking IDs
  dailySheetObjs.sort((a, b) => parseDateTab(a.rawName) - parseDateTab(b.rawName));
  dailySheetObjs.forEach(({ worksheet, rawName }) => {
    indexWorksheetRows(worksheet, rawName);
  });

  // Enrich masterAccountsMap from existing daily rows when not running isolated Dev Mock Mode
  if (!useDevMockAccounts) {
    const seenAccountKeys = new Set(
      masterAccountsMap.map(m => `${m.broker.toLowerCase()}::${m.master.toLowerCase()}`)
    );
    allRows.forEach(r => {
      if (!r.broker && !r.master) return;
      const k = `${r.broker.toLowerCase()}::${r.master.toLowerCase()}`;
      if (!seenAccountKeys.has(k)) {
        seenAccountKeys.add(k);
        masterAccountsMap.push({
          broker: r.broker,
          master: r.master.toUpperCase(),
          bidder: r.bidder ? r.bidder.toUpperCase() : '',
          yard: r.yard ? r.yard.toUpperCase() : ''
        });
      }
    });
  }

  populateReferenceControls();
  updateDevMockAccountsUI();

  // Sort date tabs newest first for UI dropdown
  dateSheets.sort((a, b) => parseDateTab(b) - parseDateTab(a));

  populateSheetDropdown(preserveSelectedSheet);
  renderSheetTabsBar();
  updateSheetStepperState();

  sheetSelect.disabled = dateSheets.length === 0;
  scopeSelect.disabled = dateSheets.length === 0;
  searchInput.disabled = dateSheets.length === 0;

  btnNewEntry.style.display = 'inline-flex';
  btnRollover.style.display = 'inline-flex';
  btnRepairSchema.disabled = false;
  if (menuNewScratchSheet) {
    menuNewScratchSheet.disabled = false;
  }
  btnAddBlankTab.disabled = false;
  if (btnRolloverPrevSheet) {
    btnRolloverPrevSheet.disabled = dateSheets.length === 0;
  }
  btnRenameTab.disabled = dateSheets.length === 0;
  btnDeleteTab.disabled = dateSheets.length === 0;
  btnDeleteMode.disabled = allRows.length === 0;
  if (bottomEditorDock) {
    bottomEditorDock.style.display = 'block';
    syncBottomDockStateAfterLoad();
  }
  restoreOpenScratchTabsForWorkbook();
  updateFilePathsBarUI();

  // Check if workbook is a fresh template with "Rename to Current Date" or missing reference tabs (Master Accounts is optional)
  const missingParts = [];
  if (dateSheets.length === 0 && !placeholderDateSheet) missingParts.push('Daily MM-DD-YYYY Tab');
  if (!hasTemplate) missingParts.push('template');
  if (!hasValidation) missingParts.push('validation lists');

  if (dateSheets.length === 0 && placeholderDateSheet && missingParts.length === 0) {
    const todayStr = getTodayFormatted();
    schemaBanner.style.display = 'flex';
    schemaBannerText.innerHTML = `
      <strong>📅 New Tracker Template Ready:</strong> Found initial tab <code>${escapeHtml(placeholderDateSheet.name)}</code>.
      Click <strong>Set Today's Date (${escapeHtml(todayStr)})</strong> to rename it to today's tab with all formulas &amp; dropdowns intact.
    `;
    btnBannerBootstrap.textContent = `📅 Set Today's Date (${todayStr})`;
  } else if (missingParts.length > 0) {
    schemaBanner.style.display = 'flex';
    schemaBannerText.innerHTML = `
      <strong>✨ Workbook Schema Incomplete:</strong> Missing <code>${escapeHtml(missingParts.join(', '))}</code>.
      Click Bootstrap to initialize all missing standard tabs and headers non-destructively.
    `;
    btnBannerBootstrap.textContent = '🛠️ Bootstrap Schema';
  } else {
    schemaBanner.style.display = 'none';
    btnBannerBootstrap.textContent = '🛠️ Bootstrap Schema';
  }

  if (!options.preserveStatus && !pendingMirrorBuffer) {
    setStatus(`Synced: ${file.name}`, 'ready');
  }
  renderList();
  syncWorkspaceTabsUI();
}


function cellText(row, colIndex) {
  const val = row.getCell(colIndex).value;
  if (val === null || val === undefined) return '';
  if (typeof val === 'object') {
    if (val instanceof Date) return val.toISOString().split('T')[0];
    if (Array.isArray(val.richText)) return val.richText.map(t => t.text || '').join('').trim();
    if (val.text !== undefined && val.text !== null && typeof val.text !== 'object') {
      const s = String(val.text).trim();
      return s === '[object Object]' ? '' : s;
    }
    if (val.formula !== undefined || val.sharedFormula !== undefined) {
      if (val.result === undefined || val.result === null || typeof val.result === 'object') {
        return '';
      }
      const s = String(val.result).trim();
      return s === '[object Object]' ? '' : s;
    }
    if (val.result !== undefined && val.result !== null && typeof val.result !== 'object') {
      const s = String(val.result).trim();
      return s === '[object Object]' ? '' : s;
    }
    return '';
  }
  const str = String(val).trim();
  return str === '[object Object]' ? '' : str;
}

function getWorksheetColumnMap(worksheet) {
  const row1 = worksheet.getRow(1);
  const h9 = cellText(row1, 9).toLowerCase();
  const h10 = cellText(row1, 10).toLowerCase();
  const h11 = cellText(row1, 11).toLowerCase();

  const isTemplateLayout =
    h9.includes('status') ||
    h10.includes('action taken') ||
    h11.includes('sent email');

  if (isTemplateLayout) {
    return {
      isTemplateLayout: true,
      statusCol: 9,     // Col I: Status
      notesCol: 10,     // Col J: Action Taken/Notes
      emailSubjCol: 11, // Col K: Sent Email Subject (remains visible)
      trackingCol: 12   // Col L: Hidden Tracking ID (UUIDv4)
    };
  }

  return {
    isTemplateLayout: false,
    notesCol: 9,        // Col I: Action Notes
    statusCol: 10,      // Col J: Request Status
    emailSubjCol: null,
    trackingCol: 11     // Col K: Hidden Tracking ID (UUIDv4)
  };
}

function buildTemplateBrokerFormula(rowNumber, brokerVal, masterVal) {
  let fallbackLiteral = '""';
  const cleanBroker = String(brokerVal || '').trim();
  const cleanMaster = String(masterVal || '').trim();

  if (cleanBroker) {
    const matchedInMaster = cleanMaster && masterAccountsMap.some(
      m => m.master.toLowerCase() === cleanMaster.toLowerCase()
    );
    if (!matchedInMaster) {
      fallbackLiteral = `"${cleanBroker.replace(/"/g, '""')}"`;
    }
  }
  return `IF($D${rowNumber}="", ${fallbackLiteral}, IF(_xlfn.XLOOKUP($D${rowNumber}, 'Master Accounts'!A:A, 'Master Accounts'!B:B, "")="", ${fallbackLiteral}, _xlfn.XLOOKUP($D${rowNumber}, 'Master Accounts'!A:A, 'Master Accounts'!B:B, "")))`;
}

function resolveBrokerFromMaster(masterVal, currentBrokerVal) {
  const cleanBroker = String(currentBrokerVal || '').trim();
  if (cleanBroker) return cleanBroker;
  const cleanMaster = String(masterVal || '').trim().toLowerCase();
  if (!cleanMaster) return '';
  const match = masterAccountsMap.find(m => m.master.toLowerCase() === cleanMaster);
  return match ? match.broker : '';
}

function setBrokerCellWithFormula(row, rowNumber, brokerVal, masterVal, colMap) {
  const resolvedBroker = resolveBrokerFromMaster(masterVal, brokerVal);
  const ownerWb = (row && row.worksheet && row.worksheet.workbook) ? row.worksheet.workbook : workbook;
  const wbHasMasterSheet = Boolean(getMasterAccountsWorksheet(ownerWb));
  if (colMap && colMap.isTemplateLayout && wbHasMasterSheet) {
    row.getCell(3).value = {
      formula: buildTemplateBrokerFormula(rowNumber, resolvedBroker, masterVal),
      result: resolvedBroker || ''
    };
  } else {
    row.getCell(3).value = resolvedBroker;
  }
}

// --- 3. WORKBOOK BOOTSTRAPPING, SCHEMA REPAIR & TAB MANAGEMENT ---
function applyStandardSheetFormatting(ws) {
  const widths = [30, 16, 22, 14, 14, 22, 10, 20, 36, 16, 38];
  const headerRow = ws.getRow(1);
  headerRow.height = 22;

  STANDARD_HEADERS.forEach((title, idx) => {
    const colIdx = idx + 1;
    ws.getColumn(colIdx).width = widths[idx] || 16;
    const cell = headerRow.getCell(colIdx);
    cell.value = title;
    cell.font = { bold: true, color: { argb: 'FFF8FAFC' }, size: 11 };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1E293B' }
    };
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
  });

  headerRow.commit();
  ws.getColumn(11).hidden = true;
}

function bootstrapWorkbookSchemaInMemory(wb) {
  const todayStr = getTodayFormatted();
  let templateWs = null;
  let validationWs = null;
  let placeholderDateWs = null;
  let hasAnyDateTab = false;

  wb.eachSheet((ws) => {
    const raw = ws.name.trim();
    const lower = raw.toLowerCase();
    if (DATE_TAB_REGEX.test(raw)) hasAnyDateTab = true;
    else if (isTemplatePlaceholderTab(raw)) placeholderDateWs = ws;
    else if (lower === 'template') templateWs = ws;
    else if (lower === 'validation lists' || lower === 'validation list') validationWs = ws;
  });

  // 1. Ensure template sheet exists (preserve existing Template columns if already present)
  if (!templateWs) {
    templateWs = wb.addWorksheet('template');
    applyStandardSheetFormatting(templateWs);
  } else {
    const tplMap = getWorksheetColumnMap(templateWs);
    templateWs.getColumn(tplMap.trackingCol).hidden = true;
    if (tplMap.emailSubjCol) {
      templateWs.getColumn(tplMap.emailSubjCol).hidden = false;
    }
  }

  // 2. Ensure validation lists sheet exists
  if (!validationWs) {
    validationWs = wb.addWorksheet('validation lists');
    validationWs.getColumn(1).width = 26;
    validationWs.getColumn(2).width = 22;
    const hRow = validationWs.getRow(1);
    hRow.getCell(1).value = 'Requested Actions';
    hRow.getCell(2).value = 'Status Options';
    hRow.font = { bold: true };
    hRow.commit();

    const maxLen = Math.max(DEFAULT_ACTIONS.length, DEFAULT_STATUSES.length);
    for (let i = 0; i < maxLen; i++) {
      const r = validationWs.getRow(i + 2);
      if (DEFAULT_ACTIONS[i]) r.getCell(1).value = DEFAULT_ACTIONS[i];
      if (DEFAULT_STATUSES[i]) r.getCell(2).value = DEFAULT_STATUSES[i];
      r.commit();
    }
  }

  // Note: Master Accounts sheet is intentionally excluded when creating/bootstrapping monthly workbooks.

  // 3. Ensure at least one MM-DD-YYYY daily tab exists:
  //    If the workbook has "Rename to Current Date", rename it in-place to today's date so all 74 template rows & formulas stay intact!
  let createdDateTab = null;
  if (!hasAnyDateTab) {
    createdDateTab = todayStr;
    if (placeholderDateWs) {
      placeholderDateWs.name = todayStr;
      const colMap = getWorksheetColumnMap(placeholderDateWs);
      placeholderDateWs.getColumn(colMap.trackingCol).hidden = true;
      if (colMap.emailSubjCol) {
        placeholderDateWs.getColumn(colMap.emailSubjCol).hidden = false;
      }
    } else {
      const dailyWs = wb.addWorksheet(todayStr);
      cloneTemplateStructure(templateWs, dailyWs);
      placeWorksheetFirstInWorkbook(wb, dailyWs);
    }
  }

  // 4. Remove default empty "Sheet1" if it is completely blank and other sheets exist
  const sheet1 = wb.getWorksheet('Sheet1');
  if (sheet1 && wb.worksheets.length > 1) {
    let sheet1HasData = false;
    sheet1.eachRow({ includeEmpty: false }, (row) => {
      for (let c = 1; c <= 11; c++) {
        if (cellText(row, c) !== '') sheet1HasData = true;
      }
    });
    if (!sheet1HasData) {
      wb.removeWorksheet(sheet1.id);
    }
  }

  return createdDateTab;
}

async function promptCreateNewWorkbook() {
  closeToolsMenu();
  if (!parentDirHandle) {
    return showModalAlert({
      title: 'Select Agent Folder First',
      message: 'Please connect an Agent Folder before creating a new workbook.'
    });
  }

  if (templatesDirHandle) {
    await scanTemplatesDirectory();
  }

  const now = new Date();
  const monthName = now.toLocaleString('en-US', { month: 'long' });
  const defaultFileName = `Agent Bidder Support Tracker - ${monthName} ${now.getFullYear()}.xlsx`;

  let selectOptions = null;
  if (availableTemplates.length > 0) {
    selectOptions = [
      ...availableTemplates.map((t, idx) => ({
        value: `tpl::${idx}`,
        label: `📋 From Templates Folder: ${t.name}`
      })),
      {
        value: 'builtin',
        label: '✨ Built-in Standard Schema (template, validation lists & today\'s tab — no Master Accounts)'
      }
    ];
  }

  const promptResult = await showModalPrompt({
    title: '✨ Create New Monthly Workbook',
    badge: parentDirHandle.name,
    message: availableTemplates.length > 0
      ? `Choose a template from "${templatesDirHandle.name}" or the built-in schema, and enter a filename for the new monthly tracker (.xlsx):`
      : 'Enter a filename for the new Excel tracker workbook (.xlsx). It will be initialized with template, validation lists, and today\'s daily tab (Master Accounts excluded).',
    selectLabel: 'Workbook Blueprint / Template Source',
    selectOptions,
    inputLabel: 'Workbook Filename (.xlsx)',
    defaultValue: defaultFileName,
    confirmText: '✨ Create Workbook',
    validate: (val) => {
      let checkName = val.trim();
      if (!checkName.toLowerCase().endsWith('.xlsx')) checkName += '.xlsx';
      if (checkName.toLowerCase().endsWith('(view only).xlsx') || checkName.startsWith('~$')) {
        return 'Filename cannot start with "~$" or end with "(VIEW ONLY).xlsx".';
      }
      return null;
    }
  });
  if (!promptResult) return;

  const rawName = typeof promptResult === 'object' ? promptResult.value : promptResult;
  const chosenSource = typeof promptResult === 'object' ? promptResult.selectValue : 'builtin';

  let fileName = rawName.trim();
  if (!fileName.toLowerCase().endsWith('.xlsx')) {
    fileName += '.xlsx';
  }

  try {
    let exists = false;
    try {
      await parentDirHandle.getFileHandle(fileName, { create: false });
      exists = true;
    } catch (_) {
      exists = false;
    }

    if (exists) {
      const overwrite = await showModalConfirm({
        title: '⚠️ Workbook Already Exists',
        badge: fileName,
        message: `"${fileName}" already exists in "${parentDirHandle.name}".\n\nDo you want to overwrite it with a fresh workbook?`,
        confirmText: 'Overwrite Workbook',
        variant: 'danger'
      });
      if (!overwrite) return;
    }

    const newWb = new ExcelJS.Workbook();
    if (chosenSource && chosenSource.startsWith('tpl::')) {
      const tplIdx = parseInt(chosenSource.replace('tpl::', ''), 10);
      const tplEntry = availableTemplates[tplIdx];
      if (tplEntry && tplEntry.handle) {
        const tplFile = await tplEntry.handle.getFile();
        const tplBuffer = await tplFile.arrayBuffer();
        originalArchiveState = await captureOriginalArchiveState(tplBuffer);
        const sanitizedTpl = sanitizeXlsxBufferForExcelJS(tplBuffer);
        await newWb.xlsx.load(sanitizedTpl);
      } else {
        originalArchiveState = null;
      }
    } else {
      originalArchiveState = null;
    }

    bootstrapWorkbookSchemaInMemory(newWb);

    const newFileHandle = await parentDirHandle.getFileHandle(fileName, { create: true });
    workingFileHandle = newFileHandle;
    workbook = newWb;

    const saved = await saveWorkbookToDisk();
    if (saved) {
      showToast(`Created new monthly workbook: "${fileName}"`, 'success');
      await scanDirectoryAndLoadWorkbooks(fileName);
    }
  } catch (err) {
    console.error('Error creating new workbook:', err);
    await showModalAlert({
      title: 'Could Not Create Workbook',
      message: err.message
    });
  }
}

async function handleRepairCurrentWorkbookSchema() {
  closeToolsMenu();
  if (!workbook || !workingFileHandle) return;

  await loadWorkbookFromDisk(sheetSelect.value);
  const createdTab = bootstrapWorkbookSchemaInMemory(workbook);

  // Ensure any existing MM-DD-YYYY tabs have headers on Row 1 and proper tracking column visibility
  dateSheets.forEach(tabName => {
    const ws = workbook.getWorksheet(tabName);
    if (!ws) return;
    const row1 = ws.getRow(1);
    const hasHeader = [1, 2, 3].some(c => cellText(row1, c) !== '');
    if (!hasHeader) {
      const templateSheet = findTemplateSheet();
      if (templateSheet) {
        cloneTemplateStructure(templateSheet, ws);
      } else {
        applyStandardSheetFormatting(ws);
      }
    } else {
      const colMap = getWorksheetColumnMap(ws);
      ws.getColumn(colMap.trackingCol).hidden = true;
      if (colMap.emailSubjCol) {
        ws.getColumn(colMap.emailSubjCol).hidden = false;
      }
    }
  });

  const targetSelectTab = createdTab || sheetSelect.value;
  const saved = await saveWorkbookToDisk();
  if (saved) {
    await loadWorkbookFromDisk(targetSelectTab, {
      preserveStatus: saved === 'mirror-locked'
    });
    showToast(
      createdTab
        ? `Initialized daily tab "${createdTab}" with all formulas & validations intact.`
        : 'Workbook schema verified.',
      'success'
    );
  }
}

btnNewWorkbook.addEventListener('click', promptCreateNewWorkbook);
btnRepairSchema.addEventListener('click', handleRepairCurrentWorkbookSchema);
btnBannerBootstrap.addEventListener('click', handleRepairCurrentWorkbookSchema);

// --- SHEET CREATION, ROLLOVER FROM PREVIOUS SHEET, RENAMING & DELETION ---
btnAddBlankTab.addEventListener('click', async () => {
  closeToolsMenu();
  if (!workbook) return;

  const latestSheet = dateSheets.length > 0 ? dateSheets[0] : getTodayFormatted();
  const suggestedDate = dateSheets.includes(getTodayFormatted())
    ? addDaysToDateTab(latestSheet, 1)
    : getTodayFormatted();

  const hasPriorSheets = dateSheets.length > 0;
  const selectOptions = hasPriorSheets
    ? [
        {
          value: 'rollover_prev',
          label: `↪ Rollover unfinished requests from previous sheet (${latestSheet})`
        },
        {
          value: 'blank',
          label: '📄 Create blank daily tab only (headers & template formulas)'
        }
      ]
    : null;

  const rawInput = await showModalPrompt({
    title: '➕ New Daily Tab',
    badge: 'MM-DD-YYYY',
    message: hasPriorSheets
      ? `Enter the date for the new daily tab and choose whether to roll over unfinished items from the previous sheet or start blank:`
      : 'Enter the date for the new blank daily worksheet tab (cloned from your template headers):',
    selectLabel: 'Initial Tab Contents',
    selectOptions,
    defaultSelectValue: hasPriorSheets ? 'rollover_prev' : '',
    inputLabel: 'Daily Tab Date (MM-DD-YYYY)',
    defaultValue: suggestedDate,
    placeholder: 'MM-DD-YYYY',
    confirmText: '➕ Create Tab',
    validate: (val) => {
      if (!DATE_TAB_REGEX.test(val)) {
        return 'Please use exact MM-DD-YYYY format (e.g., 11-06-2026).';
      }
      if (workbook.getWorksheet(val)) {
        return `Worksheet tab "${val}" already exists in this workbook.`;
      }
      return null;
    }
  });
  if (!rawInput) return;
  const newTabName = (typeof rawInput === 'object' ? rawInput.value : rawInput).trim();
  const creationMode = (typeof rawInput === 'object' && rawInput.selectValue) ? rawInput.selectValue : 'blank';

  await loadWorkbookFromDisk(sheetSelect.value);
  if (workbook.getWorksheet(newTabName)) {
    return showModalAlert({
      title: 'Tab Already Exists',
      message: `Worksheet tab "${newTabName}" already exists!`
    });
  }

  let placeholderWs = null;
  workbook.eachSheet(ws => {
    if (isTemplatePlaceholderTab(ws.name)) placeholderWs = ws;
  });

  let newSheet = null;
  if (placeholderWs && dateSheets.length === 0) {
    placeholderWs.name = newTabName;
    newSheet = placeholderWs;
  } else {
    const templateSheet = findTemplateSheet();
    newSheet = workbook.addWorksheet(newTabName);
    placeWorksheetFirst(newSheet);

    if (templateSheet) {
      cloneTemplateStructure(templateSheet, newSheet);
    } else {
      applyStandardSheetFormatting(newSheet);
    }
  }

  let rolledCount = 0;
  let prevSourceName = null;
  if (creationMode === 'rollover_prev' && dateSheets.length > 0 && newSheet) {
    const plan = computeRolloverAnalysis(newTabName, false, 'previous');
    prevSourceName = plan.immediatePrevSheet || null;
    if (plan.valid && plan.toRollOver.length > 0) {
      const targetColMap = getWorksheetColumnMap(newSheet);
      newSheet.getColumn(targetColMap.trackingCol).hidden = true;
      if (targetColMap.emailSubjCol) {
        newSheet.getColumn(targetColMap.emailSubjCol).hidden = false;
      }
      let writeRowIdx = getNextEmptyRow(newSheet);
      for (const item of plan.toRollOver) {
        const rNum = writeRowIdx++;
        const r = newSheet.getRow(rNum);
        writeItemToWorksheetRow(r, rNum, item, targetColMap);
        unifyHistoricalGroupTrackingIds(item._entries, item.key);
        rolledCount++;
      }
    }
  }

  const saved = await saveWorkbookToDisk();
  if (saved) {
    activeWorkspaceTab = DAILY_TRACKER_WS_TAB;
    clearTransientFiltersOnNavigation({ resetScopeToCurrent: true });
    await loadWorkbookFromDisk(newTabName, {
      preserveStatus: saved === 'mirror-locked'
    });
    if (creationMode === 'rollover_prev' && prevSourceName) {
      showToast(
        `Created "${newTabName}" and rolled over ${rolledCount} unfinished request(s) from ${prevSourceName}.`,
        'success'
      );
    } else {
      showToast(`Created blank daily tab "${newTabName}".`, 'success');
    }
  }
});

if (btnRolloverPrevSheet) {
  btnRolloverPrevSheet.addEventListener('click', async () => {
    closeToolsMenu();
    if (!workbook) return;
    try {
      const activeSheet = sheetSelect.value;
      await loadWorkbookFromDisk(activeSheet);
      const preferredTarget = (activeSheet && DATE_TAB_REGEX.test(activeSheet)) ? activeSheet : null;
      openRolloverModal({ preferredSourceMode: 'previous', preferredTargetDate: preferredTarget });
    } catch (err) {
      console.error('Error opening rollover from previous sheet:', err);
      await showModalAlert({
        title: 'Rollover Read Error',
        message: 'Could not read workbook for rollover: ' + err.message,
        variant: 'danger'
      });
    }
  });
}

btnRenameTab.addEventListener('click', async () => {
  closeToolsMenu();
  const currentTab = sheetSelect.value;
  if (!workbook || !currentTab || !DATE_TAB_REGEX.test(currentTab)) {
    return showModalAlert({
      title: 'No Daily Tab Selected',
      message: 'Select a valid daily MM-DD-YYYY tab to rename.'
    });
  }

  const rawNewName = await showModalPrompt({
    title: '✏️ Rename Current Day Tab',
    badge: currentTab,
    message: `Enter a new date name (MM-DD-YYYY) for worksheet tab "${currentTab}":`,
    inputLabel: 'New Tab Name (MM-DD-YYYY)',
    defaultValue: currentTab,
    placeholder: 'MM-DD-YYYY',
    confirmText: '✏️ Rename Tab',
    validate: (val) => {
      if (!DATE_TAB_REGEX.test(val)) {
        return 'New tab name must use exact MM-DD-YYYY format.';
      }
      if (val !== currentTab && workbook.getWorksheet(val)) {
        return `A worksheet named "${val}" already exists in this workbook.`;
      }
      return null;
    }
  });
  if (!rawNewName) return;
  const newTabName = (typeof rawNewName === 'object' ? rawNewName.value : rawNewName).trim();

  if (newTabName === currentTab) return;

  await loadWorkbookFromDisk(currentTab);
  if (workbook.getWorksheet(newTabName)) {
    return showModalAlert({
      title: 'Tab Already Exists',
      message: `A worksheet named "${newTabName}" already exists in this workbook.`
    });
  }

  const ws = workbook.getWorksheet(currentTab);
  if (!ws) {
    return showModalAlert({
      title: 'Worksheet Not Found',
      message: `Worksheet "${currentTab}" was not found.`
    });
  }

  ws.name = newTabName;
  const saved = await saveWorkbookToDisk();
  if (saved) {
    activeWorkspaceTab = DAILY_TRACKER_WS_TAB;
    clearTransientFiltersOnNavigation({ resetScopeToCurrent: true });
    await loadWorkbookFromDisk(newTabName, {
      preserveStatus: saved === 'mirror-locked'
    });
    showToast(`Renamed tab "${currentTab}" → "${newTabName}".`, 'success');
  }
});

btnDeleteTab.addEventListener('click', async () => {
  closeToolsMenu();
  const currentTab = sheetSelect.value;
  if (!workbook || !currentTab || !DATE_TAB_REGEX.test(currentTab)) {
    return showModalAlert({
      title: 'No Daily Tab Selected',
      message: 'Select a valid daily MM-DD-YYYY tab to delete.'
    });
  }

  const rowsOnTab = allRows.filter(r => r.sheetName === currentTab);
  const confirmed = await showModalConfirm({
    title: `🗑️ Delete Daily Tab "${currentTab}"?`,
    badge: `${rowsOnTab.length} request(s)`,
    message:
      `Are you sure you want to permanently delete worksheet tab "${currentTab}"?\n\n` +
      `This tab currently contains ${rowsOnTab.length} request(s).\n` +
      `Reference tabs (template, validation lists) will not be affected.`,
    confirmText: '🗑️ Delete Tab',
    variant: 'danger'
  });
  if (!confirmed) return;

  await loadWorkbookFromDisk(currentTab);
  const ws = workbook.getWorksheet(currentTab);
  if (!ws) {
    return showModalAlert({
      title: 'Worksheet Not Found',
      message: `Worksheet "${currentTab}" was not found.`
    });
  }

  if (workbook.worksheets.length <= 1) {
    const fallbackTemplate = workbook.addWorksheet('template');
    applyStandardSheetFormatting(fallbackTemplate);
  }

  workbook.removeWorksheet(ws.id);

  const saved = await saveWorkbookToDisk();
  if (saved) {
    activeWorkspaceTab = DAILY_TRACKER_WS_TAB;
    clearTransientFiltersOnNavigation({ resetScopeToCurrent: true });
    await loadWorkbookFromDisk(null, {
      preserveStatus: saved === 'mirror-locked'
    });
    showToast(`Deleted worksheet tab "${currentTab}".`, 'success');
  }
});

// --- 4. DYNAMIC REFERENCE SHEET PARSING (validation lists & Master Accounts) ---
function parseValidationSheet(worksheet) {
  const actionsSet = new Set();
  const statusesColB = new Set();
  const statusesColE = new Set();

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const actionVal = cellText(row, 1);
    const statusValB = cellText(row, 2);
    const statusValE = cellText(row, 5);

    if (actionVal) actionsSet.add(actionVal);
    if (statusValB) statusesColB.add(statusValB);
    if (statusValE) statusesColE.add(statusValE);
  });

  validationActions = Array.from(actionsSet);
  validationStatuses = statusesColE.size > 0
    ? Array.from(statusesColE)
    : Array.from(statusesColB);
}

function parseMasterAccountsSheet(worksheet) {
  const seen = new Set();
  const row1 = worksheet.getRow(1);
  const h1 = cellText(row1, 1).toLowerCase();
  const h2 = cellText(row1, 2).toLowerCase();
  const h3 = cellText(row1, 3).toLowerCase();

  // In "Tracker Template - Make a Copy.xlsx", Col A is "Master Account" and Col B is "Account Name"
  const isMasterFirst =
    h1.includes('master') ||
    h2.includes('account name') ||
    (h2.includes('broker') && !h1.includes('broker'));
  const hasBidderInCol3 = h3.includes('bidder');

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    let broker = '';
    let master = '';
    let bidder = '';

    if (isMasterFirst) {
      master = cellText(row, 1).toUpperCase();
      broker = cellText(row, 2);
      bidder = hasBidderInCol3 ? cellText(row, 3).toUpperCase() : '';
    } else {
      broker = cellText(row, 1);
      master = cellText(row, 2).toUpperCase();
      bidder = cellText(row, 3).toUpperCase();
    }

    if (!broker && !master) return;
    const dedupeKey = `${broker.toLowerCase()}::${master.toLowerCase()}`;
    if (seen.has(dedupeKey)) return;
    seen.add(dedupeKey);

    masterAccountsMap.push({ broker, master, bidder });
  });
}

function populateSelectWithStatuses(selectEl) {
  if (!selectEl) return;
  const statusesToUse = validationStatuses.length > 0 ? validationStatuses : DEFAULT_STATUSES;
  const currentSelected = selectEl.value;
  selectEl.innerHTML = '';
  statusesToUse.forEach(st => {
    const opt = document.createElement('option');
    opt.value = st;
    opt.textContent = st;
    selectEl.appendChild(opt);
  });
  if (currentSelected && statusesToUse.includes(currentSelected)) {
    selectEl.value = currentSelected;
  }
}

function updateAutoFillFieldHints() {
  const hasAutoFillData = Array.isArray(masterAccountsMap) && masterAccountsMap.length > 0;
  const modeHint = useDevMockAccounts
    ? `Mock Auto-Fill active (${masterAccountsMap.length} test accounts)`
    : hasAutoFillData
      ? `Auto-fills from ${hasMasterAccountsSheet ? 'Master Accounts' : 'workbook history'} (or reverse)`
      : 'Enter manually (Master Accounts sheet not present)';

  if (fMaster) {
    fMaster.placeholder = hasAutoFillData ? 'Auto-fills from Broker (or reverse)' : 'Enter Master # manually';
    fMaster.title = modeHint;
  }
  if (dockMaster) {
    dockMaster.placeholder = hasAutoFillData ? 'Master # (Auto-fills)' : 'Master #';
    dockMaster.title = modeHint;
  }
  if (fBroker) {
    fBroker.placeholder = hasAutoFillData ? 'Select or type broker...' : 'Type broker name...';
    fBroker.title = modeHint;
  }
  if (dockBroker) {
    dockBroker.placeholder = hasAutoFillData ? 'Select or type broker...' : 'Type broker name...';
    dockBroker.title = modeHint;
  }
}

function populateReferenceControls() {
  const actionsToUse = validationActions.length > 0 ? validationActions : DEFAULT_ACTIONS;
  actionListEl.innerHTML = '';
  actionsToUse.forEach(action => {
    const opt = document.createElement('option');
    opt.value = action;
    actionListEl.appendChild(opt);
  });

  populateSelectWithStatuses(fStatus);
  populateSelectWithStatuses(dockStatus);
  populateSelectWithStatuses(batchStatusSelect);

  brokerListEl.innerHTML = '';
  const safeAccounts = Array.isArray(masterAccountsMap) ? masterAccountsMap : [];
  const uniqueBrokers = Array.from(new Set(safeAccounts.map(m => m.broker).filter(Boolean)));
  uniqueBrokers.forEach(broker => {
    const opt = document.createElement('option');
    opt.value = broker;
    brokerListEl.appendChild(opt);
  });

  updateAutoFillFieldHints();
}

function refreshAccountAutoFillSource() {
  masterAccountsMap = [];
  if (useDevMockAccounts) {
    masterAccountsMap = getMockMasterAccountsList();
  } else if (workbook) {
    const masterWs = getMasterAccountsWorksheet(workbook);
    hasMasterAccountsSheet = Boolean(masterWs);
    if (masterWs) {
      parseMasterAccountsSheet(masterWs);
    }
    const seenAccountKeys = new Set(
      masterAccountsMap.map(m => `${m.broker.toLowerCase()}::${m.master.toLowerCase()}`)
    );
    allRows.forEach(r => {
      if (!r.broker && !r.master) return;
      const k = `${r.broker.toLowerCase()}::${r.master.toLowerCase()}`;
      if (!seenAccountKeys.has(k)) {
        seenAccountKeys.add(k);
        masterAccountsMap.push({
          broker: r.broker,
          master: r.master.toUpperCase(),
          bidder: r.bidder ? r.bidder.toUpperCase() : '',
          yard: r.yard ? r.yard.toUpperCase() : ''
        });
      }
    });
  }
  populateReferenceControls();
  updateDevMockAccountsUI();
}

if (chkDevMockAccounts) {
  chkDevMockAccounts.checked = Boolean(useDevMockAccounts);
  chkDevMockAccounts.addEventListener('change', (e) => {
    useDevMockAccounts = Boolean(e.target.checked);
    localStorage.setItem(DEV_MOCK_STORAGE_KEY, useDevMockAccounts ? 'true' : 'false');
    refreshAccountAutoFillSource();
    showToast(
      useDevMockAccounts
        ? `Dev Mock Master Accounts enabled (${masterAccountsMap.length} dummy accounts loaded).`
        : (hasMasterAccountsSheet
            ? 'Switched auto-fill back to workbook Master Accounts sheet.'
            : 'Mock data disabled — graceful manual entry / workbook history mode active.'),
      'info'
    );
  });
}

// Initialize Dev Mock Master Accounts immediately on startup when enabled (e.g., localhost testing)
if (useDevMockAccounts) {
  masterAccountsMap = getMockMasterAccountsList();
}
populateReferenceControls();
updateDevMockAccountsUI();

function applyBrokerAutoFillToPair(brokerInputEl, masterInputEl, bidderInputEl, yardInputEl = null) {
  if (!brokerInputEl || !Array.isArray(masterAccountsMap) || masterAccountsMap.length === 0) return;
  const typedBroker = brokerInputEl.value.trim().toLowerCase();
  if (!typedBroker) return;

  const match = masterAccountsMap.find(m => m.broker && m.broker.toLowerCase() === typedBroker);
  if (match) {
    if (masterInputEl && !masterInputEl.value.trim() && match.master) {
      masterInputEl.value = match.master;
    }
    if (bidderInputEl && !bidderInputEl.value.trim() && match.bidder) {
      bidderInputEl.value = match.bidder;
    }
    if (yardInputEl && !yardInputEl.value.trim() && match.yard) {
      yardInputEl.value = match.yard;
    }
  }
}

function applyMasterReverseAutoFillToPair(masterInputEl, brokerInputEl, bidderInputEl, yardInputEl = null) {
  if (!masterInputEl || !Array.isArray(masterAccountsMap) || masterAccountsMap.length === 0) return;
  const typedMaster = masterInputEl.value.trim().toLowerCase();
  if (!typedMaster) return;

  const match = masterAccountsMap.find(m => m.master && m.master.toLowerCase() === typedMaster);
  if (match) {
    if (brokerInputEl && !brokerInputEl.value.trim() && match.broker) {
      brokerInputEl.value = match.broker;
    }
    if (bidderInputEl && !bidderInputEl.value.trim() && match.bidder) {
      bidderInputEl.value = match.bidder;
    }
    if (yardInputEl && !yardInputEl.value.trim() && match.yard) {
      yardInputEl.value = match.yard;
    }
  }
}

fBroker.addEventListener('input', () => applyBrokerAutoFillToPair(fBroker, fMaster, fBidder, fYard));
fBroker.addEventListener('change', () => applyBrokerAutoFillToPair(fBroker, fMaster, fBidder, fYard));
fMaster.addEventListener('input', () => applyMasterReverseAutoFillToPair(fMaster, fBroker, fBidder, fYard));
fMaster.addEventListener('change', () => applyMasterReverseAutoFillToPair(fMaster, fBroker, fBidder, fYard));

if (dockBroker && dockMaster && dockBidder) {
  dockBroker.addEventListener('input', () => applyBrokerAutoFillToPair(dockBroker, dockMaster, dockBidder, dockYard));
  dockBroker.addEventListener('change', () => applyBrokerAutoFillToPair(dockBroker, dockMaster, dockBidder, dockYard));
  dockMaster.addEventListener('input', () => applyMasterReverseAutoFillToPair(dockMaster, dockBroker, dockBidder, dockYard));
  dockMaster.addEventListener('change', () => applyMasterReverseAutoFillToPair(dockMaster, dockBroker, dockBidder, dockYard));
}

function getNextEmptyRow(worksheet) {
  const colMap = getWorksheetColumnMap(worksheet);
  let lastRow = 1;
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const subject = cellText(row, 1);
    const ticket = cellText(row, 2);
    const master = cellText(row, 4);
    const broker = resolveBrokerFromMaster(master, cellText(row, 3));
    const bidder = cellText(row, 5);
    const action = cellText(row, 6);
    const yard = cellText(row, 7);
    const lots = cellText(row, 8);
    const status = cellText(row, colMap.statusCol);
    const notes = cellText(row, colMap.notesCol);
    const hasData = [subject, ticket, broker, master, bidder, action, yard, lots, notes, status]
      .some(val => val !== '');
    if (hasData && rowNumber > lastRow) lastRow = rowNumber;
  });
  return lastRow + 1;
}

function findTemplateSheet() {
  let found = null;
  workbook.eachSheet((ws) => {
    if (ws.name.trim().toLowerCase() === 'template') {
      found = ws;
    }
  });
  if (!found && dateSheets.length > 0) {
    found = workbook.getWorksheet(dateSheets[0]);
  }
  return found;
}

function placeWorksheetFirstInWorkbook(wb, targetSheet) {
  try {
    let order = 2;
    wb.eachSheet((ws) => {
      if (ws.id === targetSheet.id) {
        ws.orderNo = 1;
      } else {
        ws.orderNo = order++;
      }
    });
  } catch (e) {
    console.warn('Could not reorder worksheet tabs:', e);
  }
}

function placeWorksheetFirst(targetSheet) {
  if (workbook) placeWorksheetFirstInWorkbook(workbook, targetSheet);
}

function cloneTemplateStructure(templateSheet, newSheet) {
  if (Array.isArray(templateSheet.views) && templateSheet.views.length > 0) {
    newSheet.views = JSON.parse(JSON.stringify(templateSheet.views));
  }
  if (templateSheet.pageSetup) {
    newSheet.pageSetup = JSON.parse(JSON.stringify(templateSheet.pageSetup));
  }

  if (Array.isArray(templateSheet.columns)) {
    templateSheet.columns.forEach((col, idx) => {
      if (!col) return;
      const targetCol = newSheet.getColumn(idx + 1);
      targetCol.width = col.width || 15;
      targetCol.hidden = col.hidden || false;
      if (col.style) {
        targetCol.style = JSON.parse(JSON.stringify(col.style));
      }
    });
  }

  const headerRow = templateSheet.getRow(1);
  const newHeader = newSheet.getRow(1);
  if (headerRow.height) newHeader.height = headerRow.height;

  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const targetCell = newHeader.getCell(colNumber);
    targetCell.value = cell.value;
    if (cell.style) {
      targetCell.style = JSON.parse(JSON.stringify(cell.style));
    }
  });
  newHeader.commit();

  const colMap = getWorksheetColumnMap(newSheet);
  newSheet.getColumn(colMap.trackingCol).hidden = true;
  if (colMap.emailSubjCol) {
    newSheet.getColumn(colMap.emailSubjCol).hidden = false;
  }

  // If cloning from the full 74-row Tracker Template, preserve rows 2..74 cell styles and Column C XLOOKUP formulas
  if (colMap.isTemplateLayout) {
    const maxTemplateRow = Math.max(templateSheet.rowCount || 74, 74);
    for (let r = 2; r <= maxTemplateRow; r++) {
      const srcRow = templateSheet.getRow(r);
      const dstRow = newSheet.getRow(r);
      if (srcRow.height) dstRow.height = srcRow.height;
      srcRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        if (colNumber <= 12 && cell.style) {
          dstRow.getCell(colNumber).style = JSON.parse(JSON.stringify(cell.style));
        }
      });
      setBrokerCellWithFormula(dstRow, r, '', '', colMap);
      dstRow.commit();
    }
  }
}

function writeItemToWorksheetRow(row, rowNumber, item, colMap) {
  const masterUpper = item.master ? item.master.toUpperCase() : '';
  row.getCell(1).value = item.subject;
  row.getCell(2).value = item.ticket ? item.ticket.toUpperCase() : '';
  setBrokerCellWithFormula(row, rowNumber, item.broker, masterUpper, colMap);
  row.getCell(4).value = masterUpper;
  row.getCell(5).value = item.bidder ? item.bidder.toUpperCase() : '';
  row.getCell(6).value = item.action;
  row.getCell(7).value = item.yard ? item.yard.toUpperCase() : '';
  row.getCell(8).value = item.lots;
  row.getCell(colMap.notesCol).value = item.notes;
  row.getCell(colMap.statusCol).value = item.status;
  row.getCell(colMap.trackingCol).value = item.key;
  row.commit();
}

function prepareWorkbookFormulasBeforeSave(wb) {
  if (!wb) return;
  wb.eachSheet((ws) => {
    const rawName = ws.name.trim();
    const lower = rawName.toLowerCase();
    const isDailyOrTemplate =
      lower === 'template' ||
      isTemplatePlaceholderTab(rawName) ||
      DATE_TAB_REGEX.test(rawName);
    if (!isDailyOrTemplate) return;

    const colMap = getWorksheetColumnMap(ws);
    ws.getColumn(colMap.trackingCol).hidden = true;
    if (colMap.emailSubjCol) {
      ws.getColumn(colMap.emailSubjCol).hidden = false;
    }

    if (!colMap.isTemplateLayout) return;

    const maxR = Math.max(ws.rowCount || 74, 74);
    for (let r = 2; r <= maxR; r++) {
      const row = ws.getRow(r);
      const masterVal = cellText(row, 4).toUpperCase();
      const brokerVal = cellText(row, 3);
      setBrokerCellWithFormula(row, r, brokerVal, masterVal, colMap);
    }
  });
}

async function serializeWorkbookBuffer() {
  prepareWorkbookFormulasBeforeSave(workbook);
  const rawBuffer = await workbook.xlsx.writeBuffer();
  return await repairSavedXlsxArchive(rawBuffer, originalArchiveState);
}

// --- 8. DUAL-WRITE MIRRORING + AUTO-RETRY & MANUAL RETRY ---
function clearMirrorRetryState() {
  pendingMirrorBuffer = null;
  mirrorRetryCount = 0;
  mirrorCountdownSec = 0;
  if (mirrorCountdownInterval) {
    clearInterval(mirrorCountdownInterval);
    mirrorCountdownInterval = null;
  }
  if (btnRetryMirror) {
    btnRetryMirror.style.display = 'none';
  }
}

function scheduleMirrorAutoRetry() {
  if (mirrorCountdownInterval) {
    clearInterval(mirrorCountdownInterval);
  }
  btnRetryMirror.style.display = 'inline-block';

  if (mirrorRetryCount >= MIRROR_MAX_AUTO_RETRIES) {
    btnRetryMirror.textContent = '🔄 Retry Mirror Sync';
    setStatus('⚠️ Saved Working File (Mirror locked — click Retry when free)', 'warning');
    return;
  }

  mirrorCountdownSec = MIRROR_RETRY_INTERVAL_SEC;
  btnRetryMirror.textContent = `🔄 Retry Mirror (${mirrorCountdownSec}s)`;

  mirrorCountdownInterval = setInterval(async () => {
    mirrorCountdownSec--;
    if (mirrorCountdownSec > 0) {
      btnRetryMirror.textContent = `🔄 Retry Mirror (${mirrorCountdownSec}s)`;
    } else {
      clearInterval(mirrorCountdownInterval);
      mirrorCountdownInterval = null;
      mirrorRetryCount++;
      await retryMirrorSyncNow(true);
    }
  }, 1000);
}

async function attemptMirrorWrite(buffer, isRetry = false) {
  const targetMirrorDir = supervisorDirHandle || parentDirHandle;
  if (!targetMirrorDir || !workingFileHandle) return true;

  const mirrorFileName = workingFileHandle.name.replace(/\.xlsx$/i, '') + ' (VIEW ONLY).xlsx';
  try {
    const mirrorHandle = await targetMirrorDir.getFileHandle(mirrorFileName, { create: true });
    const mirrorWritable = await mirrorHandle.createWritable();
    await mirrorWritable.write(buffer);
    await mirrorWritable.close();

    const wasRetrying = Boolean(pendingMirrorBuffer) || isRetry;
    clearMirrorRetryState();

    const mirrorLocationNote = supervisorDirHandle ? ` (Mirrored to ${supervisorDirHandle.name})` : '';
    setStatus(`Saved & Mirrored: ${workingFileHandle.name}${mirrorLocationNote}`, 'ready');
    updateFilePathsBarUI();
    if (wasRetrying) {
      showToast(`✅ Supervisor mirror "${mirrorFileName}" synced!`, 'success');
    }
    return true;
  } catch (mirrorErr) {
    console.warn(`Supervisor mirror file "${mirrorFileName}" is locked or busy:`, mirrorErr);
    pendingMirrorBuffer = buffer;
    setStatus('⚠️ Saved Working File (Mirror in use by Supervisor)', 'warning');
    scheduleMirrorAutoRetry();
    if (!isRetry) {
      showToast('⚠️ Saved Working File (Mirror in use by Supervisor — auto-retrying every 15s)', 'warning');
    }
    return false;
  }
}

async function retryMirrorSyncNow(isAuto = false) {
  if (!workingFileHandle || !workbook) return;
  btnRetryMirror.textContent = '⏳ Retrying Mirror...';

  let bufferToWrite = pendingMirrorBuffer;
  if (!bufferToWrite) {
    try {
      bufferToWrite = await serializeWorkbookBuffer();
    } catch (e) {
      console.error('Could not serialize buffer for mirror retry:', e);
      return;
    }
  }

  const ok = await attemptMirrorWrite(bufferToWrite, true);
  if (!ok && !isAuto) {
    showToast('⚠️ Supervisor (VIEW ONLY) file is still locked in Excel. Will keep retrying.', 'warning');
  }
}

btnRetryMirror.addEventListener('click', () => {
  mirrorRetryCount = 0; // Reset auto-retry counter on manual click
  retryMirrorSyncNow(false);
});

async function saveWorkbookToDisk() {
  if (!workingFileHandle) return false;

  setStatus('Saving to Excel...', '');
  let buffer;
  try {
    buffer = await serializeWorkbookBuffer();
  } catch (err) {
    console.error('Failed to serialize workbook buffer:', err);
    setStatus('Serialization Error', 'locked');
    await showModalAlert({
      title: 'Excel Serialization Error',
      message: 'Failed to generate Excel buffer: ' + err.message,
      variant: 'danger'
    });
    return false;
  }

  // 1. Primary Working File Write
  try {
    const writable = await workingFileHandle.createWritable();
    await writable.write(buffer);
    await writable.close();
  } catch (err) {
    console.error('Primary file write error:', err);
    setStatus('File Locked! Close Excel & Retry', 'locked');
    await showModalAlert({
      title: '🔒 Save Blocked by Windows (File Locked)',
      badge: workingFileHandle.name,
      message:
        `The working workbook ("${workingFileHandle.name}") is currently open and locked in Desktop Excel.\n\n` +
        'Please close the file in Desktop Excel and click Save again.',
      variant: 'danger'
    });
    return false;
  }

  // 2. Dual-Write Mirroring to (VIEW ONLY).xlsx with Auto-Retry & Manual Retry fallback
  const mirrorOk = await attemptMirrorWrite(buffer, false);
  return mirrorOk ? true : 'mirror-locked';
}

