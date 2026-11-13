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

function getProfileIDBKeys(profileId = activeProfileId) {
  if (!profileId || profileId === DEFAULT_TRACKER_PROFILE_ID) {
    return {
      agentKey: IDB_AGENT_KEY,
      supervisorKey: IDB_SUPERVISOR_KEY,
      templatesKey: IDB_TEMPLATES_KEY,
      lastWbKey: IDB_LAST_WORKBOOK_KEY
    };
  }
  const prefix = `profile::${profileId}::`;
  return {
    agentKey: `${prefix}agentFolderHandle`,
    supervisorKey: `${prefix}supervisorFolderHandle`,
    templatesKey: `${prefix}templatesFolderHandle`,
    lastWbKey: `${prefix}lastActiveWorkbookName`
  };
}

function getStagedPrimarySaveIDBKey(workbookName, profileId = activeProfileId) {
  return `stagedPrimarySave::${profileId || DEFAULT_TRACKER_PROFILE_ID}::${String(workbookName || '').trim()}`;
}

function getActiveTrackerProfile() {
  let found = trackerProfiles.find(p => p.id === activeProfileId);
  if (!found) {
    found = trackerProfiles[0] || { id: DEFAULT_TRACKER_PROFILE_ID, name: 'Daily Tracker', mirrorEnabled: true };
    activeProfileId = found.id;
  }
  return found;
}

function isActiveProfileMirrorEnabled() {
  const prof = getActiveTrackerProfile();
  return prof ? prof.mirrorEnabled !== false : true;
}

function persistTrackerProfiles() {
  try {
    localStorage.setItem(TRACKER_PROFILES_STORAGE_KEY, JSON.stringify(trackerProfiles));
    localStorage.setItem(ACTIVE_PROFILE_STORAGE_KEY, activeProfileId);
  } catch (err) {
    console.warn('Could not persist tracker profiles:', err);
  }
}

function syncMirrorCopyCheckboxUI() {
  if (chkEnableMirrorCopy) {
    chkEnableMirrorCopy.checked = isActiveProfileMirrorEnabled();
  }
}

if (chkEnableMirrorCopy) {
  syncMirrorCopyCheckboxUI();
  chkEnableMirrorCopy.addEventListener('change', async (e) => {
    const prof = getActiveTrackerProfile();
    if (!prof) return;
    prof.mirrorEnabled = Boolean(e.target.checked);
    persistTrackerProfiles();
    if (!prof.mirrorEnabled) {
      clearMirrorRetryState();
    }
    updateFilePathsBarUI();
    updateStatusBarTelemetry();
    if (prof.mirrorEnabled && workbook && workingFileHandle) {
      try {
        const buf = await serializeWorkbookBuffer();
        await attemptMirrorWrite(buf, false);
      } catch (_) { /* ignore */ }
    } else {
      showToast(
        prof.mirrorEnabled
          ? `Enabled (VIEW ONLY).xlsx mirroring for "${prof.name}".`
          : `Disabled (VIEW ONLY).xlsx mirroring for "${prof.name}" (saves working file only).`,
        'info'
      );
    }
  });
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

  const mirrorEnabled = isActiveProfileMirrorEnabled();

  if (pathViewOnlyText) {
    if (!mirrorEnabled) {
      pathViewOnlyText.textContent = 'Disabled for this tracker tab';
    } else if ((supervisorDirHandle || parentDirHandle) && workingFileHandle) {
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

  const canOpenViewOnly = Boolean(mirrorEnabled && (parentDirHandle || supervisorDirHandle) && workingFileHandle);
  if (btnOpenViewOnly) {
    btnOpenViewOnly.style.display = canOpenViewOnly ? 'inline-block' : 'none';
  }
  if (menuOpenViewOnly) {
    menuOpenViewOnly.disabled = !canOpenViewOnly;
  }
  syncMirrorCopyCheckboxUI();
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
  syncMirrorCopyCheckboxUI();
  if (typeof syncWorkspaceTabsUI === 'function') {
    syncWorkspaceTabsUI();
  }
  if (!('indexedDB' in window)) return;
  const idbKeys = getProfileIDBKeys(activeProfileId);
  const [savedAgentHandle, savedSupHandle, savedTplHandle, savedWbName] = await Promise.all([
    getHandleFromIDB(idbKeys.agentKey),
    getHandleFromIDB(idbKeys.supervisorKey),
    getHandleFromIDB(idbKeys.templatesKey),
    getHandleFromIDB(idbKeys.lastWbKey)
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
          evt.target.closest('.menu-fontsize-control') ||
          evt.target.closest('#trackerProfilesBar') ||
          evt.target.closest('#workspaceTabsBar')
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
    const idbKeys = getProfileIDBKeys(activeProfileId);
    await saveHandleToIDB(idbKeys.agentKey, dirHandle);
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

function resetActiveTrackerUIDisconnectedState() {
  workbookPickerWrap.style.display = 'none';
  btnReconnect.style.display = cachedDirHandle ? 'inline-flex' : 'none';
  btnLinkFile.style.display = 'inline-flex';
  btnNewEntry.style.display = 'none';
  btnRollover.style.display = 'none';
  btnNewWorkbook.disabled = !parentDirHandle;
  btnRepairSchema.disabled = true;
  if (btnOpenSchemaMapper) btnOpenSchemaMapper.disabled = true;
  if (menuNewScratchSheet) menuNewScratchSheet.disabled = true;
  btnAddBlankTab.disabled = true;
  if (btnRolloverPrevSheet) btnRolloverPrevSheet.disabled = true;
  btnRenameTab.disabled = true;
  btnDeleteTab.disabled = true;
  btnDeleteMode.disabled = true;
  selectedForDelete.clear();
  if (batchActionDock) batchActionDock.style.display = 'none';
  if (bottomEditorDock) bottomEditorDock.style.display = 'none';
  sheetTabsBar.style.display = 'none';
  schemaBanner.style.display = 'none';
  activeSchemaConfig = {
    isStandardSchema: true,
    mode: 'auto',
    columns: [],
    primaryCol: 1,
    ticketCol: 2,
    secondaryCol: 3,
    statusCol: 10,
    notesCol: 9,
    trackingCol: 11
  };
  if (typeof renderTableHeader === 'function') renderTableHeader();
  if (typeof syncDynamicEditorFormsUI === 'function') syncDynamicEditorFormsUI();
  setStatus('No Folder Linked', '');
  updateAgentDirUI();
  updateSupervisorDirUI();
  updateTemplatesDirUI();
  const prof = getActiveTrackerProfile();
  const label = prof ? escapeHtml(prof.name) : 'Tracker';
  listBody.innerHTML = `<tr><td colspan="12" class="empty-table-cell">Click <strong>Select Agent Folder</strong> to connect the workbook directory for <strong>${label}</strong>.</td></tr>`;
  if (typeof syncWorkspaceTabsUI === 'function') syncWorkspaceTabsUI();
}

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
    const idbKeys = getProfileIDBKeys(activeProfileId);
    await deleteHandleFromIDB(idbKeys.agentKey);
    await deleteHandleFromIDB(idbKeys.lastWbKey);
    trackerProfileRuntimeCache.delete(activeProfileId);
    resetActiveTrackerUIDisconnectedState();
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
    await saveHandleToIDB(getProfileIDBKeys(activeProfileId).supervisorKey, supDir);
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
  await deleteHandleFromIDB(getProfileIDBKeys(activeProfileId).supervisorKey);
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
      await saveHandleToIDB(getProfileIDBKeys(activeProfileId).templatesKey, tplDir);
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
    await deleteHandleFromIDB(getProfileIDBKeys(activeProfileId).templatesKey);
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
  await saveHandleToIDB(getProfileIDBKeys(activeProfileId).lastWbKey, workingFileHandle.name);
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
    await saveHandleToIDB(getProfileIDBKeys(activeProfileId).lastWbKey, workingFileHandle.name);
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

function expandValidationSqrefRanges(xmlFragment) {
  if (!xmlFragment) return xmlFragment;
  const expandRangeToken = (token) => {
    const m = token.match(/^(\$?[A-Z]+)\$?2(?::(\$?[A-Z]+)\$?(\d+))?$/i);
    if (!m) return token;
    const col1 = m[1].replace(/\$/g, '');
    const col2 = (m[2] || col1).replace(/\$/g, '');
    if (col1.toUpperCase() !== col2.toUpperCase()) return token;
    const endRow = m[3] ? parseInt(m[3], 10) : 2;
    if (endRow >= 100) return token;
    return `${col1}2:${col1}200`;
  };
  return xmlFragment
    .replace(/<xm:sqref>([^<]+)<\/xm:sqref>/g, (_, inner) => {
      const updated = inner.trim().split(/\s+/).map(expandRangeToken).join(' ');
      return `<xm:sqref>${updated}</xm:sqref>`;
    })
    .replace(/\bsqref="([^"]+)"/g, (_, inner) => {
      const updated = inner.trim().split(/\s+/).map(expandRangeToken).join(' ');
      return `sqref="${updated}"`;
    });
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
    let capturedExtLstXml = null;
    let bestExtScore = -1;
    let capturedStdDataValidationsXml = null;
    let bestStdScore = -1;

    for (const [sheetName, zipPath] of sheetPaths.entries()) {
      const trimmedName = sheetName.trim();
      const lower = trimmedName.toLowerCase();
      const xml = dec.decode(entries.get(zipPath));
      const isDateSheet = DATE_TAB_REGEX.test(trimmedName);
      const dateWeight = isDateSheet ? parseDateTab(trimmedName) : 0;

      if (lower === 'template' || isTemplatePlaceholderTab(lower) || isDateSheet) {
        const extMatch = xml.match(/<extLst>[\s\S]*?<\/extLst>/);
        if (extMatch && extMatch[0].includes('dataValidations')) {
          const cleanedExt = expandValidationSqrefRanges(extMatch[0].replace(/\s+xr:uid="[^"]*"/g, ''));
          const dvCount = (cleanedExt.match(/<x14:dataValidation\b/g) || []).length;
          const score = dvCount * 1e15 + dateWeight;
          if (score > bestExtScore) {
            bestExtScore = score;
            capturedExtLstXml = cleanedExt;
          }
        }

        const stdMatch = xml.match(/<dataValidations\b[^>]*>[\s\S]*?<\/dataValidations>/);
        if (stdMatch) {
          const cleanedStd = expandValidationSqrefRanges(stdMatch[0].replace(/\s+xr:uid="[^"]*"/g, ''));
          const stdCount = (cleanedStd.match(/<dataValidation\b/g) || []).length;
          const stdScore = stdCount * 1e15 + dateWeight;
          if (stdScore > bestStdScore) {
            bestStdScore = stdScore;
            capturedStdDataValidationsXml = cleanedStd;
          }
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
      capturedExtLstXml,
      capturedStdDataValidationsXml,
      templateExtLstXml: capturedExtLstXml || DAILY_SHEET_DATA_VALIDATIONS_EXTLST
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

      if (/<f\b[^>]*>[\s\S]*?\[1\]/.test(xml)) {
        hasAnyExternalRefFormula = true;
      }

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

        if ( archiveState && archiveState.capturedStdDataValidationsXml ) {
          const targetStdXml = archiveState.capturedStdDataValidationsXml;
          const targetStdCount = (targetStdXml.match(/<dataValidation\b/g) || []).length;
          const existingStdMatch = xml.match(/<dataValidations\b[^>]*>[\s\S]*?<\/dataValidations>/);
          if (!existingStdMatch) {
            if (/(<hyperlinks|<printOptions|<pageMargins|<pageSetup|<headerFooter|<drawing|<extLst|<\/worksheet>)/.test(xml)) {
              xml = xml.replace(
                /(<hyperlinks|<printOptions|<pageMargins|<pageSetup|<headerFooter|<drawing|<extLst|<\/worksheet>)/,
                `${targetStdXml}$1`
              );
            }
          } else {
            const existingStdCount = (existingStdMatch[0].match(/<dataValidation\b/g) || []).length;
            if (existingStdCount < targetStdCount) {
              xml = xml.replace(/<dataValidations\b[^>]*>[\s\S]*?<\/dataValidations>/, targetStdXml);
            }
          }
        }

        const customSynthesizedExtLst = (!activeSchemaConfig.isStandardSchema && hasValidationSheetInWb && typeof buildCustomSchemaDataValidationsExtLst === 'function')
          ? buildCustomSchemaDataValidationsExtLst()
          : null;

        let desiredExtLst = (archiveState && archiveState.capturedExtLstXml)
          ? archiveState.capturedExtLstXml
          : (customSynthesizedExtLst || (activeSchemaConfig.isStandardSchema && hasValidationSheetInWb ? DAILY_SHEET_DATA_VALIDATIONS_EXTLST : null));

        // If custom schema has a status column validation on the validation sheet that wasn't in an older capturedExtLstXml, merge or prefer the fuller one
        if (
          desiredExtLst &&
          customSynthesizedExtLst &&
          desiredExtLst !== customSynthesizedExtLst &&
          activeSchemaConfig &&
          activeSchemaConfig.statusCol
        ) {
          const stLetter = colNumberToExcelLetters(activeSchemaConfig.statusCol);
          const sqrefRe = new RegExp(`\\b${stLetter}2:`, 'i');
          if (!sqrefRe.test(desiredExtLst) && !(archiveState && archiveState.capturedStdDataValidationsXml && sqrefRe.test(archiveState.capturedStdDataValidationsXml))) {
            desiredExtLst = customSynthesizedExtLst;
          }
        }

        if (desiredExtLst) {
          const desiredExtCount = (desiredExtLst.match(/<x14:dataValidation\b/g) || []).length;
          if (!xml.includes('dataValidations')) {
            if (!xml.includes('<extLst>')) {
              xml = xml.replace(/<\/worksheet>\s*$/, `${desiredExtLst}</worksheet>`);
            }
          } else if (xml.includes('<x14:dataValidations')) {
            const currentExtCount = (xml.match(/<x14:dataValidation\b/g) || []).length;
            if (currentExtCount < desiredExtCount) {
              xml = xml.replace(/<extLst>[\s\S]*?<\/extLst>/, desiredExtLst);
            }
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

function getWorkbookSeriesKey(fileName) {
  return String(fileName || '')
    .replace(/\.xlsx$/i, '')
    .replace(/\s*\(view only\)\s*$/i, '')
    .replace(/[\s\-_]+(january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)[\s\-_]+20\d{2}\b/ig, '')
    .trim()
    .toLowerCase();
}

function getSavedSchemaOverrideForWorkbook(fileName, profileId = activeProfileId) {
  if (!fileName || !schemaOverridesStore) return null;
  const exactKey = `${profileId}::file::${String(fileName).trim().toLowerCase()}`;
  if (schemaOverridesStore[exactKey]) return schemaOverridesStore[exactKey];
  const series = getWorkbookSeriesKey(fileName);
  if (series) {
    const seriesKey = `${profileId}::series::${series}`;
    if (schemaOverridesStore[seriesKey]) return schemaOverridesStore[seriesKey];
  }
  return null;
}

function saveSchemaOverrideForWorkbook(fileName, overrideObj, profileId = activeProfileId) {
  if (!fileName) return;
  const exactKey = `${profileId}::file::${String(fileName).trim().toLowerCase()}`;
  const series = getWorkbookSeriesKey(fileName);
  const seriesKey = series ? `${profileId}::series::${series}` : null;

  if (!overrideObj) {
    delete schemaOverridesStore[exactKey];
    if (seriesKey) delete schemaOverridesStore[seriesKey];
  } else {
    schemaOverridesStore[exactKey] = overrideObj;
    if (seriesKey) schemaOverridesStore[seriesKey] = overrideObj;
  }
  try {
    localStorage.setItem(SCHEMA_OVERRIDES_STORAGE_KEY, JSON.stringify(schemaOverridesStore));
  } catch (e) {
    console.warn('Could not persist schema overrides:', e);
  }
}

function extractDetectedRow1HeadersFromSheet(ws) {
  if (!ws) return { columns: [], existingTrackingCol: null, maxCol: 0 };
  const row1 = ws.getRow(1);
  const scanLimit = Math.min(Math.max(row1.cellCount || 0, ws.columnCount || 0, 14), 60);
  let maxNonEmptyCol = 0;
  for (let c = 1; c <= scanLimit; c++) {
    if (cellText(row1, c) !== '') {
      maxNonEmptyCol = c;
    }
  }
  if (maxNonEmptyCol === 0) {
    return { columns: [], existingTrackingCol: null, maxCol: 0 };
  }

  const maxRowScan = Math.min(ws.rowCount || 1, 18);
  const colHasShiftedUuidNextToIt = (colIdx) => {
    let nextHasUuid = false;
    let curHasNonUuid = false;
    for (let r = 2; r <= maxRowScan; r++) {
      const row = ws.getRow(r);
      const curVal = cellText(row, colIdx);
      const nextVal = cellText(row, colIdx + 1);
      if (nextVal && UUID_REGEX.test(nextVal)) nextHasUuid = true;
      if (curVal && !UUID_REGEX.test(curVal)) curHasNonUuid = true;
    }
    return nextHasUuid && curHasNonUuid;
  };

  let existingTrackingCol = null;
  const rawCols = [];
  for (let c = 1; c <= maxNonEmptyCol; c++) {
    const header = cellText(row1, c);
    const lower = header.toLowerCase();
    if (/^(tracking\s*id|tracking_id|uuid|guid|row\s*guid)$/i.test(lower)) {
      // Handle mid-month Status column added where Tracking ID used to be (while UUIDs moved to c + 1)
      if (c === maxNonEmptyCol && colHasShiftedUuidNextToIt(c)) {
        existingTrackingCol = c + 1;
        const colLetter = colNumberToExcelLetters(c);
        const inferredLabel = 'Status';
        const dropdownOpts = findValidationOptionsForHeader(inferredLabel);
        rawCols.push({
          colIndex: c,
          fieldKey: `col_${c}`,
          colLetter,
          header: inferredLabel,
          dropdownOptions: dropdownOpts
        });
      } else {
        existingTrackingCol = c;
      }
      continue;
    }
    const colLetter = colNumberToExcelLetters(c);
    const label = header || `Column ${colLetter}`;
    const dropdownOpts = findValidationOptionsForHeader(label);
    rawCols.push({
      colIndex: c,
      fieldKey: `col_${c}`,
      colLetter,
      header: label,
      dropdownOptions: dropdownOpts
    });
  }

  return {
    columns: rawCols,
    existingTrackingCol,
    maxCol: maxNonEmptyCol
  };
}

function findValidationOptionsForHeader(headerLabel) {
  if (!headerLabel || !validationListsByHeader || validationListsByHeader.size === 0) return [];
  const norm = String(headerLabel).trim().toLowerCase();
  if (!norm) return [];
  if (validationListsByHeader.has(norm)) {
    return (validationListsByHeader.get(norm) || []).filter(v => !UUID_REGEX.test(String(v).trim()));
  }
  // Singular/plural & partial match (e.g. "Requested Action" <-> "Requested Actions", "Status" <-> "Status Options")
  for (const [vHeader, opts] of validationListsByHeader.entries()) {
    if (
      vHeader === `${norm}s` ||
      `${vHeader}s` === norm ||
      (norm.length >= 4 && vHeader.startsWith(norm)) ||
      (vHeader.length >= 4 && norm.startsWith(vHeader))
    ) {
      return (opts || []).filter(v => !UUID_REGEX.test(String(v).trim()));
    }
  }
  return [];
}

function excelLettersToColNumber(letters) {
  const clean = String(letters || '').trim().toUpperCase();
  if (!clean) return 0;
  let num = 0;
  for (let i = 0; i < clean.length; i++) {
    const code = clean.charCodeAt(i);
    if (code < 65 || code > 90) return 0;
    num = num * 26 + (code - 64);
  }
  return num;
}

function resolveFormulaListValuesFromWorkbook(wb, rawFormula) {
  if (!wb || !rawFormula) return [];
  const formula = String(rawFormula).trim().replace(/^=/, '').trim();
  if (!formula) return [];

  // 1. Inline quoted list: "Pending,In Progress,Complete"
  if (formula.startsWith('"') && formula.endsWith('"') && formula.length >= 2) {
    return formula
      .slice(1, -1)
      .split(',')
      .map(s => s.trim())
      .filter(s => s && !UUID_REGEX.test(s));
  }

  // 2. Worksheet range reference: 'Validation List'!$A$2:$A$50 or validation lists!$B$2:$B$20
  const rangeMatch = formula.match(/^(?:'([^']+)'|([^!]+))!\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/);
  if (!rangeMatch) return [];

  const targetSheetName = (rangeMatch[1] || rangeMatch[2] || '').replace(/''/g, "'").trim().toLowerCase();
  const colStart = excelLettersToColNumber(rangeMatch[3]);
  const rowStart = Math.max(1, parseInt(rangeMatch[4] || '2', 10) || 2);
  const rowEnd = Math.min(500, Math.max(rowStart, parseInt(rangeMatch[6] || String(rowStart + 100), 10) || (rowStart + 100)));
  if (!colStart) return [];

  let refSheet = null;
  wb.eachSheet((ws) => {
    if (ws.name.trim().toLowerCase() === targetSheetName) {
      refSheet = ws;
    }
  });
  if (!refSheet) return [];

  const values = [];
  const seen = new Set();
  const maxScan = Math.min(Math.max(refSheet.rowCount || 0, rowEnd), 500);
  for (let r = rowStart; r <= maxScan; r++) {
    const val = cellText(refSheet.getRow(r), colStart);
    if (!val || UUID_REGEX.test(val)) continue;
    // Skip Row 1 if the range was $A$1:$A$50 and Row 1 is a column header like "Status Options"
    if (r === 1 && /\b(options?|list|actions?|status(?:es)?)\b/i.test(val) && !KNOWN_STATUS_VALUES.has(val.toLowerCase())) {
      continue;
    }
    const key = val.toLowerCase();
    if (!seen.has(key)) {
      seen.add(key);
      values.push(val);
    }
  }
  return values;
}

function extractDataValidationOptionsByCol(wb, archiveState) {
  const byCol = new Map();
  if (!wb) return byCol;

  const recordOptions = (colNum, opts) => {
    if (!colNum || !Array.isArray(opts) || opts.length === 0) return;
    const existing = byCol.get(colNum) || [];
    if (opts.length > existing.length) {
      byCol.set(colNum, opts);
    }
  };

  // 1. Inspect ExcelJS in-memory worksheet dataValidations.model across sheets
  wb.eachSheet((ws) => {
    const raw = ws.name.trim();
    const isDailyOrTpl =
      raw.toLowerCase() === 'template' ||
      isTemplatePlaceholderTab(raw) ||
      DATE_TAB_REGEX.test(raw);
    if (!isDailyOrTpl) return;
    const model = ws.dataValidations && ws.dataValidations.model;
    if (!model || typeof model !== 'object') return;
    Object.entries(model).forEach(([addr, rule]) => {
      if (!rule || rule.type !== 'list' || !Array.isArray(rule.formulae) || !rule.formulae[0]) return;
      const colMatch = String(addr).match(/^\$?([A-Za-z]{1,3})\$?\d+/);
      if (!colMatch) return;
      const colNum = excelLettersToColNumber(colMatch[1]);
      const opts = resolveFormulaListValuesFromWorkbook(wb, rule.formulae[0]);
      recordOptions(colNum, opts);
    });
  });

  // 2. Inspect captured OpenXML <extLst> (<x14:dataValidations>) and <dataValidations>
  const extXml = archiveState && archiveState.capturedExtLstXml ? archiveState.capturedExtLstXml : '';
  if (extXml) {
    const x14Blocks = extXml.match(/<x14:dataValidation\b[\s\S]*?<\/x14:dataValidation>/g) || [];
    x14Blocks.forEach(block => {
      if (!/\btype="list"/i.test(block)) return;
      const fMatch = block.match(/<x14:formula1>\s*<xm:f>([\s\S]*?)<\/xm:f>\s*<\/x14:formula1>/);
      const sqMatch = block.match(/<xm:sqref>([^<]+)<\/xm:sqref>/);
      if (!fMatch || !sqMatch) return;
      const opts = resolveFormulaListValuesFromWorkbook(wb, fMatch[1]);
      if (opts.length === 0) return;
      String(sqMatch[1]).trim().split(/\s+/).forEach(tok => {
        const colMatch = tok.match(/^\$?([A-Za-z]{1,3})\$?\d+/);
        if (colMatch) recordOptions(excelLettersToColNumber(colMatch[1]), opts);
      });
    });
  }

  const stdXml = archiveState && archiveState.capturedStdDataValidationsXml ? archiveState.capturedStdDataValidationsXml : '';
  if (stdXml) {
    const dvBlocks = stdXml.match(/<dataValidation\b[\s\S]*?<\/dataValidation>/g) || [];
    dvBlocks.forEach(block => {
      if (!/\btype="list"/i.test(block)) return;
      const sqMatch = block.match(/\bsqref="([^"]+)"/i);
      const fMatch = block.match(/<formula1>([\s\S]*?)<\/formula1>/i);
      if (!sqMatch || !fMatch) return;
      const rawFormula = fMatch[1].replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
      const opts = resolveFormulaListValuesFromWorkbook(wb, rawFormula);
      if (opts.length === 0) return;
      String(sqMatch[1]).trim().split(/\s+/).forEach(tok => {
        const colMatch = tok.match(/^\$?([A-Za-z]{1,3})\$?\d+/);
        if (colMatch) recordOptions(excelLettersToColNumber(colMatch[1]), opts);
      });
    });
  }

  return byCol;
}

function isBrokerStandardHeaderRow(ws) {
  if (!ws) return true;
  const row1 = ws.getRow(1);
  let anyHeader = false;
  for (let c = 1; c <= 12; c++) {
    if (cellText(row1, c) !== '') {
      anyHeader = true;
      break;
    }
  }
  if (!anyHeader) return true;

  let maxCol = 0;
  for (let c = 1; c <= 40; c++) {
    if (cellText(row1, c) !== '') maxCol = c;
  }
  if (maxCol > 12) return false;

  const h1 = cellText(row1, 1).toLowerCase();
  const h2 = cellText(row1, 2).toLowerCase();
  const h3 = cellText(row1, 3).toLowerCase();
  const h4 = cellText(row1, 4).toLowerCase();
  const h5 = cellText(row1, 5).toLowerCase();
  const h6 = cellText(row1, 6).toLowerCase();
  const h7 = cellText(row1, 7).toLowerCase();
  const h8 = cellText(row1, 8).toLowerCase();

  const standardMatches = [
    h2.includes('ticket'),
    h3.includes('broker') || h3.includes('account'),
    h4.includes('master'),
    h5.includes('bidder'),
    h6.includes('action'),
    h7.includes('yard'),
    h8.includes('lot')
  ].filter(Boolean).length;

  return h1.includes('subject') && standardMatches >= 2;
}

function findPrimarySchemaHeaderWorksheet(wb) {
  if (!wb) return null;
  let tplWs = null;
  let placeholderWs = null;
  const dateTabList = [];
  let firstCustomWs = null;

  wb.eachSheet((ws) => {
    const raw = ws.name.trim();
    const lower = raw.toLowerCase();
    if (lower === 'template') tplWs = ws;
    else if (isTemplatePlaceholderTab(raw)) placeholderWs = ws;
    else if (DATE_TAB_REGEX.test(raw)) dateTabList.push(ws);
    else if (!isProtectedReferenceSheetName(raw) && !firstCustomWs) firstCustomWs = ws;
  });

  if (dateTabList.length > 0) {
    // Sort newest first, then pick the sheet with the most visible non-tracking columns
    // so mid-month added columns (e.g. Status in Col J) are never lost to an older or template-cloned tab
    dateTabList.sort((a, b) => parseDateTab(b.name.trim()) - parseDateTab(a.name.trim()));
    let bestDaily = null;
    let bestColCount = 0;
    for (const ws of dateTabList) {
      const det = extractDetectedRow1HeadersFromSheet(ws);
      if (det.columns.length > bestColCount) {
        bestColCount = det.columns.length;
        bestDaily = ws;
      }
    }
    if (bestDaily) return bestDaily;
  }
  if (placeholderWs) return placeholderWs;
  if (tplWs) return tplWs;
  return firstCustomWs;
}

function resolveWorkbookSchemaConfig(wb = workbook, fileName = '') {
  const refWs = findPrimarySchemaHeaderWorksheet(wb);
  const detected = extractDetectedRow1HeadersFromSheet(refWs);
  const autoIsStandard = isBrokerStandardHeaderRow(refWs);
  const savedOverride = getSavedSchemaOverrideForWorkbook(fileName, activeProfileId);
  const mode = (savedOverride && ['auto', 'custom', 'standard'].includes(savedOverride.mode))
    ? savedOverride.mode
    : 'auto';

  let useStandard = mode === 'standard'
    ? true
    : (mode === 'custom' ? false : autoIsStandard);

  // Cannot use custom mode if Row 1 has zero headers at all
  if (!useStandard && detected.columns.length === 0) {
    useStandard = true;
  }

  if (useStandard) {
    const stdMap = refWs ? getStandardWorksheetColumnMapOnly(refWs) : {
      isTemplateLayout: false,
      statusCol: 10,
      notesCol: 9,
      emailSubjCol: null,
      trackingCol: 11
    };
    activeSchemaConfig = {
      isStandardSchema: true,
      mode,
      autoDetectedStandard: autoIsStandard,
      columns: detected.columns,
      primaryCol: 1,
      ticketCol: 2,
      secondaryCol: 3,
      statusCol: stdMap.statusCol,
      notesCol: stdMap.notesCol,
      trackingCol: stdMap.trackingCol
    };
    return activeSchemaConfig;
  }

  const cols = detected.columns;
  const dvOptsByCol = extractDataValidationOptionsByCol(wb, originalArchiveState);
  cols.forEach(col => {
    if ((!col.dropdownOptions || col.dropdownOptions.length === 0) && dvOptsByCol.has(col.colIndex)) {
      col.dropdownOptions = dvOptsByCol.get(col.colIndex);
    }
  });

  const validColSet = new Set(cols.map(c => c.colIndex));
  const pickValidOverride = (val) => {
    if (val === 0 || val === '0' || val === 'none') return null;
    const n = Number(val);
    return Number.isFinite(n) && validColSet.has(n) ? n : undefined;
  };

  // Auto-infer semantic roles from column headers
  const findColByRegex = (regex, excludeSet = new Set()) => {
    const hit = cols.find(c => !excludeSet.has(c.colIndex) && regex.test(c.header));
    return hit ? hit.colIndex : null;
  };

  const inferredStatusCol =
    findColByRegex(/\b(status|state|stage|disposition|outcome)\b/i) ||
    findColByRegex(/status|state/i);
  const statusCol = savedOverride && pickValidOverride(savedOverride.statusCol) !== undefined
    ? pickValidOverride(savedOverride.statusCol)
    : inferredStatusCol;

  const usedSet = new Set(statusCol ? [statusCol] : []);

  const inferredNotesCol =
    findColByRegex(/\b(notes?|comments?|remarks?|updates?|timeline|action taken)\b/i, usedSet) ||
    findColByRegex(/note|comment|detail/i, usedSet);
  const notesCol = savedOverride && pickValidOverride(savedOverride.notesCol) !== undefined
    ? pickValidOverride(savedOverride.notesCol)
    : inferredNotesCol;
  if (notesCol) usedSet.add(notesCol);

  const inferredTicketCol = findColByRegex(/\b(tickets?|case\s*#?|incident|ref\s*#?|reference)\b/i, usedSet);
  const ticketCol = savedOverride && pickValidOverride(savedOverride.ticketCol) !== undefined
    ? pickValidOverride(savedOverride.ticketCol)
    : inferredTicketCol;
  if (ticketCol) usedSet.add(ticketCol);

  const inferredPrimaryCol =
    findColByRegex(/\b(subject|title|summary|request|vin|lot|item|issue|topic|account|customer|name)\b/i, usedSet) ||
    (cols.find(c => !usedSet.has(c.colIndex)) || cols[0]).colIndex;
  const primaryCol = (savedOverride && pickValidOverride(savedOverride.primaryCol)) || inferredPrimaryCol;
  if (primaryCol) usedSet.add(primaryCol);

  const inferredSecondaryCol =
    findColByRegex(/\b(account|broker|dealer|customer|buyer|assigned|owner|specialist|agent|dept|department|state|master)\b/i, usedSet) ||
    (cols.find(c => !usedSet.has(c.colIndex)) ? cols.find(c => !usedSet.has(c.colIndex)).colIndex : null);
  const secondaryCol = savedOverride && pickValidOverride(savedOverride.secondaryCol) !== undefined
    ? pickValidOverride(savedOverride.secondaryCol)
    : inferredSecondaryCol;

  const maxVisibleCol = cols.reduce((m, c) => Math.max(m, c.colIndex), 0);
  // Ensure trackingCol is always strictly past all visible columns (e.g. Col K = 11 when Status is in Col J = 10)
  const trackingCol = Math.max(detected.existingTrackingCol || 0, maxVisibleCol + 1);

  // Sync Status column dropdown options with validationStatuses
  if (statusCol) {
    const stColObj = cols.find(c => c.colIndex === statusCol);
    if (stColObj) {
      if (stColObj.dropdownOptions && stColObj.dropdownOptions.length > 0 && validationStatuses.length === 0) {
        validationStatuses = stColObj.dropdownOptions.filter(s => s && !UUID_REGEX.test(String(s).trim()));
      } else if ((!stColObj.dropdownOptions || stColObj.dropdownOptions.length === 0) && validationStatuses.length > 0) {
        stColObj.dropdownOptions = validationStatuses.filter(s => s && !UUID_REGEX.test(String(s).trim()));
      }
    }
  }

  activeSchemaConfig = {
    isStandardSchema: false,
    mode,
    autoDetectedStandard: autoIsStandard,
    columns: cols,
    primaryCol,
    ticketCol: ticketCol || null,
    secondaryCol: secondaryCol || null,
    statusCol: statusCol || null,
    notesCol: notesCol || null,
    trackingCol
  };
  return activeSchemaConfig;
}

function expandDataValidationModelRanges(modelObj) {
  if (!modelObj || typeof modelObj !== 'object') return {};
  const expanded = {};
  Object.entries(modelObj).forEach(([addr, rule]) => {
    if (!rule) return;
    const rangeMatch = String(addr).match(/^(\$?[A-Za-z]{1,3}\$?2):(\$?[A-Za-z]{1,3}\$?)(\d+)$/);
    if (rangeMatch && rangeMatch[1].replace(/\$/g, '').toUpperCase().replace(/\d+/g, '') === rangeMatch[2].replace(/\$/g, '').toUpperCase()) {
      const endRow = Math.max(parseInt(rangeMatch[3], 10) || 2, 500);
      expanded[`${rangeMatch[1]}:${rangeMatch[2]}${endRow}`] = JSON.parse(JSON.stringify(rule));
    } else {
      expanded[addr] = JSON.parse(JSON.stringify(rule));
    }
  });
  return expanded;
}

function syncWorkbookSheetsWithActiveSchema(wb = workbook) {
  if (!wb || !activeSchemaConfig || activeSchemaConfig.isStandardSchema || !Array.isArray(activeSchemaConfig.columns) || activeSchemaConfig.columns.length === 0) {
    return;
  }
  const donorWs = findPrimarySchemaHeaderWorksheet(wb);
  const donorRow1 = donorWs ? donorWs.getRow(1) : null;
  let bestDvModel = null;
  let bestDvCount = 0;

  wb.eachSheet((ws) => {
    const raw = ws.name.trim();
    if (raw.toLowerCase() !== 'template' && !DATE_TAB_REGEX.test(raw)) return;
    const m = ws.dataValidations && ws.dataValidations.model;
    const count = m && typeof m === 'object' ? Object.keys(m).length : 0;
    if (count > bestDvCount) {
      bestDvCount = count;
      bestDvModel = m;
    }
  });

  // If no ExcelJS dataValidations.model exists yet for statusCol, synthesize one from validationSheetMeta
  const statusColIdx = activeSchemaConfig.statusCol;
  const synthesizedDvModel = bestDvModel ? expandDataValidationModelRanges(bestDvModel) : {};
  if (statusColIdx && validationSheetMeta && validationSheetMeta.statusRangeFormula) {
    const colLet = colNumberToExcelLetters(statusColIdx);
    const hasStatusRule = Object.keys(synthesizedDvModel).some(k =>
      k.replace(/\$/g, '').toUpperCase().startsWith(`${colLet}2`)
    );
    if (!hasStatusRule) {
      synthesizedDvModel[`${colLet}2:${colLet}500`] = {
        type: 'list',
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formulae: [validationSheetMeta.statusRangeFormula]
      };
    }
  }
  const hasAnyDvRules = Object.keys(synthesizedDvModel).length > 0;

  wb.eachSheet((ws) => {
    const raw = ws.name.trim();
    const isDailyOrTpl =
      raw.toLowerCase() === 'template' ||
      isTemplatePlaceholderTab(raw) ||
      DATE_TAB_REGEX.test(raw);
    if (!isDailyOrTpl) return;

    const row1 = ws.getRow(1);
    let modifiedHeader = false;
    activeSchemaConfig.columns.forEach((col) => {
      const idx = col.colIndex;
      ws.getColumn(idx).hidden = false;
      const curHdr = cellText(row1, idx);
      if (!curHdr || /^(tracking\s*id|tracking_id|uuid|guid|row\s*guid)$/i.test(curHdr)) {
        const cell = row1.getCell(idx);
        cell.value = col.header;
        const styleDonorCell = (donorRow1 && donorRow1.getCell(idx).style)
          ? donorRow1.getCell(idx)
          : (idx > 1 ? row1.getCell(idx - 1) : null);
        if (styleDonorCell && styleDonorCell.style) {
          cell.style = JSON.parse(JSON.stringify(styleDonorCell.style));
        }
        if (donorWs && donorWs.getColumn(idx) && donorWs.getColumn(idx).width) {
          ws.getColumn(idx).width = donorWs.getColumn(idx).width;
        } else if (!ws.getColumn(idx).width || ws.getColumn(idx).width < 12) {
          ws.getColumn(idx).width = 18;
        }
        modifiedHeader = true;
      }
    });

    const trkCol = activeSchemaConfig.trackingCol;
    if (trkCol) {
      const trkHdr = cellText(row1, trkCol);
      if (!trkHdr) {
        row1.getCell(trkCol).value = 'Tracking ID';
        modifiedHeader = true;
      }
      ws.getColumn(trkCol).hidden = true;
    }
    if (modifiedHeader) {
      row1.commit();
    }

    if (hasAnyDvRules && ws.dataValidations) {
      const curModel = (ws.dataValidations.model && typeof ws.dataValidations.model === 'object')
        ? ws.dataValidations.model
        : {};
      ws.dataValidations.model = {
        ...expandDataValidationModelRanges(curModel),
        ...JSON.parse(JSON.stringify(synthesizedDvModel))
      };
    }
  });
}

async function loadWorkbookFromDisk(preserveSelectedSheet = null, options = {}) {
  if (!workingFileHandle) return;
  commitActiveScratchInlineEdit();
  if (hasAnyPendingScratchMutations() && !options.skipFlushPendingScratch && workbook) {
    await flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: true });
  }
  if (!options.preserveStatus && !pendingMirrorBuffer && !pendingPrimarySaveState) {
    setStatus('Reading...', '');
  }

  let file = await workingFileHandle.getFile();
  let arrayBuffer = await file.arrayBuffer();

  // Check IndexedDB Fail-Safe Staging Cache for unsaved primary .xlsx edits (e.g., from when Desktop Excel locked the file)
  if (!options.skipStagedRecoveryCheck) {
    const stagedKey = getStagedPrimarySaveIDBKey(workingFileHandle.name, activeProfileId);
    const stagedRecord = await getHandleFromIDB(stagedKey);
    if (stagedRecord && stagedRecord.buffer && stagedRecord.savedAt) {
      const stagedDateStr = new Date(stagedRecord.savedAt).toLocaleString();
      const diskDateStr = file.lastModified ? new Date(file.lastModified).toLocaleString() : 'Unknown';
      const recoverNow = await showModalConfirm({
        title: '🛡️ Recover Staged Fail-Safe Edits?',
        badge: workingFileHandle.name,
        message:
          `TrackerUI found unsaved edits for "${workingFileHandle.name}" staged in browser fail-safe storage (from ${stagedDateStr} while the file was locked in Excel; current disk timestamp: ${diskDateStr}).\n\n` +
          `Would you like to write and restore those staged edits now?`,
        confirmText: '💾 Restore & Save Staged Edits',
        cancelText: '🗑️ Discard Staged Cache'
      });

      if (recoverNow) {
        try {
          const writable = await workingFileHandle.createWritable();
          await writable.write(stagedRecord.buffer);
          await writable.close();
          await deleteHandleFromIDB(stagedKey);
          clearPrimarySaveRetryUI();
          await attemptMirrorWrite(stagedRecord.buffer, false);
          file = await workingFileHandle.getFile();
          arrayBuffer = await file.arrayBuffer();
          showToast(`Recovered and saved staged edits to "${workingFileHandle.name}".`, 'success');
        } catch (lockErr) {
          // File is still locked in Excel; load the staged buffer into memory and keep Retry Primary Save active!
          arrayBuffer = stagedRecord.buffer.slice ? stagedRecord.buffer.slice(0) : stagedRecord.buffer;
          pendingPrimarySaveState = {
            profileId: activeProfileId,
            workbookName: workingFileHandle.name,
            savedAt: stagedRecord.savedAt,
            buffer: stagedRecord.buffer
          };
          showPrimarySaveRetryUI();
          showToast('Loaded your staged edits in memory (file is still open in Excel — click Retry Primary Save when closed).', 'warning');
        }
      } else {
        await deleteHandleFromIDB(stagedKey);
        clearPrimarySaveRetryUI();
      }
    }
  }

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
  validationListsByHeader = new Map();
  validationSheetMeta = null;
  actionCategoryMap = new Map();

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

  // Resolve Standard vs Dynamic Custom Department Schema before indexing daily rows
  resolveWorkbookSchemaConfig(workbook, workingFileHandle.name);
  syncWorkbookSheetsWithActiveSchema(workbook);

  // Pass 2: Sort Daily MM-DD-YYYY Tabs oldest -> newest before indexing so un-keyed cross-day rows inherit prior Tracking IDs
  dailySheetObjs.sort((a, b) => parseDateTab(a.rawName) - parseDateTab(b.rawName));
  dailySheetObjs.forEach(({ worksheet, rawName }) => {
    indexWorksheetRows(worksheet, rawName);
  });

  // Clean validationStatuses of any accidental UUIDs and fall back to DEFAULT_STATUSES without scraping arbitrary past sheet cells
  validationStatuses = validationStatuses.filter(s => s && !UUID_REGEX.test(String(s).trim()));
  if (!activeSchemaConfig.isStandardSchema && validationStatuses.length === 0) {
    const stColObj = activeSchemaConfig.statusCol
      ? activeSchemaConfig.columns.find(c => c.colIndex === activeSchemaConfig.statusCol)
      : null;
    if (stColObj && Array.isArray(stColObj.dropdownOptions) && stColObj.dropdownOptions.length > 0) {
      validationStatuses = stColObj.dropdownOptions.filter(s => s && !UUID_REGEX.test(String(s).trim()));
    } else {
      validationStatuses = [...DEFAULT_STATUSES];
    }
  }

  // Enrich masterAccountsMap from existing daily rows when standard schema is active and not running isolated Dev Mock Mode
  if (activeSchemaConfig.isStandardSchema && !useDevMockAccounts) {
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

  if (typeof renderTableHeader === 'function') {
    renderTableHeader();
  }
  populateReferenceControls();
  if (typeof syncDynamicEditorFormsUI === 'function') {
    syncDynamicEditorFormsUI();
  }
  updateDevMockAccountsUI();

  // Sort date tabs newest first for UI dropdown
  dateSheets.sort((a, b) => parseDateTab(b) - parseDateTab(a));

  populateSheetDropdown(preserveSelectedSheet);
  renderSheetTabsBar();
  updateSheetStepperState();

  sheetSelect.disabled = dateSheets.length === 0;
  scopeSelect.disabled = dateSheets.length === 0;
  searchInput.disabled = dateSheets.length === 0;
  if (searchInput) {
    if (!activeSchemaConfig.isStandardSchema && activeSchemaConfig.columns.length > 0) {
      const sampleHeaders = activeSchemaConfig.columns.slice(0, 5).map(c => c.header).join(', ');
      searchInput.placeholder = `Search across ALL daily tabs (${sampleHeaders}, or Tracking ID)...`;
    } else {
      searchInput.placeholder = 'Search across ALL daily tabs (Subject, Ticket #, Broker, Master #, Bidder #, Lot #, or Tracking ID)...';
    }
  }

  btnNewEntry.style.display = 'inline-flex';
  btnRollover.style.display = 'inline-flex';
  btnRepairSchema.disabled = false;
  if (btnOpenSchemaMapper) {
    btnOpenSchemaMapper.disabled = false;
  }
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

  // Check if workbook is a fresh template with "Rename to Current Date" or missing reference tabs
  const missingParts = [];
  if (dateSheets.length === 0 && !placeholderDateSheet) missingParts.push('Daily MM-DD-YYYY Tab');
  if (activeSchemaConfig.isStandardSchema) {
    if (!hasTemplate) missingParts.push('template');
    if (!hasValidation) missingParts.push('validation lists');
  }

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
    schemaBannerText.innerHTML = !activeSchemaConfig.isStandardSchema
      ? `<strong>📅 Custom Department Workbook Ready:</strong> Detected ${activeSchemaConfig.columns.length} custom columns, but no <code>MM-DD-YYYY</code> daily tab yet. Click below to initialize today's daily tab preserving your custom headers.`
      : `
        <strong>✨ Workbook Schema Incomplete:</strong> Missing <code>${escapeHtml(missingParts.join(', '))}</code>.
        Click Bootstrap to initialize all missing standard tabs and headers non-destructively.
      `;
    btnBannerBootstrap.textContent = !activeSchemaConfig.isStandardSchema
      ? `📅 Initialize Today's Tab (${getTodayFormatted()})`
      : '🛠️ Bootstrap Schema';
  } else {
    schemaBanner.style.display = 'none';
    btnBannerBootstrap.textContent = '🛠️ Bootstrap Schema';
  }

  if (!options.preserveStatus && !pendingMirrorBuffer && !pendingPrimarySaveState) {
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

function getStandardWorksheetColumnMapOnly(worksheet) {
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
      isStandardSchema: true,
      isCustomSchema: false,
      isTemplateLayout: true,
      statusCol: 9,     // Col I: Status
      notesCol: 10,     // Col J: Action Taken/Notes
      emailSubjCol: 11, // Col K: Sent Email Subject (remains visible)
      trackingCol: 12   // Col L: Hidden Tracking ID (UUIDv4)
    };
  }

  return {
    isStandardSchema: true,
    isCustomSchema: false,
    isTemplateLayout: false,
    notesCol: 9,        // Col I: Action Notes
    statusCol: 10,      // Col J: Request Status
    emailSubjCol: null,
    trackingCol: 11     // Col K: Hidden Tracking ID (UUIDv4)
  };
}

function getWorksheetColumnMap(worksheet) {
  const ownerWb = worksheet && worksheet.workbook ? worksheet.workbook : workbook;
  if (ownerWb === workbook && activeSchemaConfig && !activeSchemaConfig.isStandardSchema) {
    return {
      isStandardSchema: false,
      isCustomSchema: true,
      isTemplateLayout: false,
      columns: activeSchemaConfig.columns,
      primaryCol: activeSchemaConfig.primaryCol,
      ticketCol: activeSchemaConfig.ticketCol,
      secondaryCol: activeSchemaConfig.secondaryCol,
      statusCol: activeSchemaConfig.statusCol,
      notesCol: activeSchemaConfig.notesCol,
      emailSubjCol: null,
      trackingCol: activeSchemaConfig.trackingCol
    };
  }
  return getStandardWorksheetColumnMapOnly(worksheet);
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

function applyCustomHeadersSheetFormatting(ws, customHeadersList) {
  const cleanHeaders = Array.isArray(customHeadersList)
    ? customHeadersList.map(h => String(h || '').trim()).filter(Boolean)
    : [];
  const headersWithTracking = [...cleanHeaders];
  if (!headersWithTracking.some(h => /^(tracking\s*id|tracking_id|uuid|guid)$/i.test(h))) {
    headersWithTracking.push('Tracking ID');
  }

  const headerRow = ws.getRow(1);
  headerRow.height = 22;

  headersWithTracking.forEach((title, idx) => {
    const colIdx = idx + 1;
    const isTracking = idx === headersWithTracking.length - 1 && /^(tracking\s*id|tracking_id|uuid|guid)$/i.test(title);
    ws.getColumn(colIdx).width = isTracking ? 38 : Math.max(16, Math.min(34, title.length + 8));
    const cell = headerRow.getCell(colIdx);
    cell.value = title;
    cell.font = { bold: true, color: { argb: 'FFF8FAFC' }, size: 11 };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1E293B' }
    };
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
    if (isTracking) {
      ws.getColumn(colIdx).hidden = true;
    }
  });

  headerRow.commit();
}

function bootstrapWorkbookSchemaInMemory(wb, customHeadersList = null) {
  const todayStr = getTodayFormatted();
  let templateWs = null;
  let validationWs = null;
  let placeholderDateWs = null;
  let hasAnyDateTab = false;
  let customHeaderSourceWs = null;

  wb.eachSheet((ws) => {
    const raw = ws.name.trim();
    const lower = raw.toLowerCase();
    if (DATE_TAB_REGEX.test(raw)) hasAnyDateTab = true;
    else if (isTemplatePlaceholderTab(raw)) placeholderDateWs = ws;
    else if (lower === 'template') templateWs = ws;
    else if (lower === 'validation lists' || lower === 'validation list') validationWs = ws;
    else if (!isProtectedReferenceSheetName(raw) && !customHeaderSourceWs) {
      const det = extractDetectedRow1HeadersFromSheet(ws);
      if (det.columns.length >= 2 && !isBrokerStandardHeaderRow(ws)) {
        customHeaderSourceWs = ws;
      }
    }
  });

  // 1. Ensure template sheet exists (preserve existing or custom department columns if present)
  if (!templateWs) {
    templateWs = wb.addWorksheet('template');
    if (Array.isArray(customHeadersList) && customHeadersList.length > 0) {
      applyCustomHeadersSheetFormatting(templateWs, customHeadersList);
    } else if (customHeaderSourceWs) {
      cloneTemplateStructure(customHeaderSourceWs, templateWs);
      const det = extractDetectedRow1HeadersFromSheet(templateWs);
      const trkCol = det.existingTrackingCol || (det.maxCol + 1);
      if (!det.existingTrackingCol) {
        templateWs.getRow(1).getCell(trkCol).value = 'Tracking ID';
        templateWs.getRow(1).commit();
      }
      templateWs.getColumn(trkCol).hidden = true;
    } else {
      applyStandardSheetFormatting(templateWs);
    }
  } else {
    const det = extractDetectedRow1HeadersFromSheet(templateWs);
    if (!isBrokerStandardHeaderRow(templateWs) && det.columns.length > 0) {
      const trkCol = det.existingTrackingCol || (det.maxCol + 1);
      if (!det.existingTrackingCol) {
        templateWs.getRow(1).getCell(trkCol).value = 'Tracking ID';
        templateWs.getRow(1).commit();
      }
      templateWs.getColumn(trkCol).hidden = true;
    } else {
      const tplMap = getStandardWorksheetColumnMapOnly(templateWs);
      templateWs.getColumn(tplMap.trackingCol).hidden = true;
      if (tplMap.emailSubjCol) {
        templateWs.getColumn(tplMap.emailSubjCol).hidden = false;
      }
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
      const det = extractDetectedRow1HeadersFromSheet(placeholderDateWs);
      if (!isBrokerStandardHeaderRow(placeholderDateWs) && det.columns.length > 0) {
        const trkCol = det.existingTrackingCol || (det.maxCol + 1);
        placeholderDateWs.getColumn(trkCol).hidden = true;
      } else {
        const colMap = getStandardWorksheetColumnMapOnly(placeholderDateWs);
        placeholderDateWs.getColumn(colMap.trackingCol).hidden = true;
        if (colMap.emailSubjCol) {
          placeholderDateWs.getColumn(colMap.emailSubjCol).hidden = false;
        }
      }
    } else if (customHeaderSourceWs && customHeaderSourceWs.name.toLowerCase() === 'sheet1') {
      // Rename custom-header Sheet1 directly to today's daily tab so any existing rows on Sheet1 are preserved!
      customHeaderSourceWs.name = todayStr;
      const det = extractDetectedRow1HeadersFromSheet(customHeaderSourceWs);
      const trkCol = det.existingTrackingCol || (det.maxCol + 1);
      if (!det.existingTrackingCol) {
        customHeaderSourceWs.getRow(1).getCell(trkCol).value = 'Tracking ID';
        customHeaderSourceWs.getRow(1).commit();
      }
      customHeaderSourceWs.getColumn(trkCol).hidden = true;
      placeWorksheetFirstInWorkbook(wb, customHeaderSourceWs);
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
      for (let c = 1; c <= 20; c++) {
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
  const activeProf = getActiveTrackerProfile();
  const isDefaultProf = !activeProf || activeProf.id === DEFAULT_TRACKER_PROFILE_ID;
  const defaultFileName = isDefaultProf
    ? `Agent Bidder Support Tracker - ${monthName} ${now.getFullYear()}.xlsx`
    : `${activeProf.name} - ${monthName} ${now.getFullYear()}.xlsx`;

  const selectOptions = [
    ...availableTemplates.map((t, idx) => ({
      value: `tpl::${idx}`,
      label: `📋 From Templates Folder: ${t.name}`
    })),
    {
      value: 'builtin',
      label: '✨ Built-in Broker Support Schema (11 standard columns — no Master Accounts)'
    },
    {
      value: 'custom_cols',
      label: '🧩 Custom Department Schema (Specify your own Row-1 column headers...)'
    }
  ];

  const promptResult = await showModalPrompt({
    title: '✨ Create New Monthly Workbook',
    badge: parentDirHandle.name,
    message: availableTemplates.length > 0
      ? `Choose a template from "${templatesDirHandle.name}", the standard schema, or custom department columns, and enter a filename (.xlsx):`
      : 'Choose the standard schema or custom department columns, and enter a filename for the new monthly tracker (.xlsx):',
    selectLabel: 'Workbook Blueprint / Schema Source',
    selectOptions,
    defaultSelectValue: availableTemplates.length > 0 ? 'tpl::0' : (isDefaultProf ? 'builtin' : 'custom_cols'),
    inputLabel: 'Workbook Filename (.xlsx)',
    defaultValue: defaultFileName,
    confirmText: '✨ Continue',
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

  let customHeadersList = null;
  if (chosenSource === 'custom_cols') {
    const rawColsPrompt = await showModalPrompt({
      title: '🧩 Define Custom Department Columns',
      badge: 'Row-1 Headers',
      message:
        'Enter your department tracker\'s column headers separated by commas.\n' +
        '(A hidden immutable UUIDv4 "Tracking ID" column will be added automatically at the end.)',
      inputLabel: 'Comma-Separated Column Headers',
      defaultValue: 'Subject / Item, Ticket #, Account #, Department Action, Notes, Status',
      confirmText: '✨ Create Custom Workbook',
      validate: (val) => {
        const parts = String(val || '').split(',').map(s => s.trim()).filter(Boolean);
        if (parts.length < 2) {
          return 'Please enter at least 2 column headers separated by commas.';
        }
        return null;
      }
    });
    if (!rawColsPrompt) return;
    const rawColsStr = typeof rawColsPrompt === 'object' ? rawColsPrompt.value : rawColsPrompt;
    customHeadersList = String(rawColsStr).split(',').map(s => s.trim()).filter(Boolean);
  }

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

    bootstrapWorkbookSchemaInMemory(newWb, customHeadersList);
    if (customHeadersList) {
      saveSchemaOverrideForWorkbook(fileName, { mode: 'custom' }, activeProfileId);
    }

    const newFileHandle = await parentDirHandle.getFileHandle(fileName, { create: true });
    workingFileHandle = newFileHandle;
    workbook = newWb;
    resolveWorkbookSchemaConfig(workbook, fileName);

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
      } else if (!activeSchemaConfig.isStandardSchema && activeSchemaConfig.columns.length > 0) {
        applyCustomHeadersSheetFormatting(ws, activeSchemaConfig.columns.map(c => c.header));
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
    } else if (!activeSchemaConfig.isStandardSchema && activeSchemaConfig.columns.length > 0) {
      applyCustomHeadersSheetFormatting(newSheet, activeSchemaConfig.columns.map(c => c.header));
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
    if (!activeSchemaConfig.isStandardSchema && activeSchemaConfig.columns.length > 0) {
      applyCustomHeadersSheetFormatting(fallbackTemplate, activeSchemaConfig.columns.map(c => c.header));
    } else {
      applyStandardSheetFormatting(fallbackTemplate);
    }
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
let validationSheetMeta = null;

function parseValidationSheet(worksheet) {
  const actionsSet = new Set();
  const statusesColB = new Set();
  const statusesColE = new Set();
  let statusHeaderMatchedCol = null;

  const row1 = worksheet.getRow(1);
  const headerByCol = new Map();
  const setsByCol = new Map();
  const maxRowByCol = new Map();
  let categoryColIndex = null;

  const maxHeaderScan = Math.min(Math.max(worksheet.columnCount || 12, 12), 30);
  let row1LooksLikeDataValues = false;
  for (let c = 1; c <= maxHeaderScan; c++) {
    const rawHeader = cellText(row1, c);
    if (!rawHeader || UUID_REGEX.test(rawHeader)) continue;
    const lowerHeader = rawHeader.toLowerCase();
    // Detect if Row 1 is actually a status option value (e.g. "Pending", "In Progress") rather than a column header
    if (KNOWN_STATUS_VALUES.has(lowerHeader) && !/\b(options?|list|status(?:es)?)\b/i.test(lowerHeader)) {
      row1LooksLikeDataValues = true;
    }
    headerByCol.set(c, lowerHeader);
    setsByCol.set(c, new Set());
    maxRowByCol.set(c, 1);
    if (!statusHeaderMatchedCol && /\b(status|statuses|state|stage|disposition|outcome)\b/i.test(rawHeader)) {
      statusHeaderMatchedCol = c;
    }
    if (
      c >= 2 &&
      c !== 5 &&
      /\b(category|parent|group|bucket|pipeline|department|domain)\b/i.test(rawHeader) &&
      !categoryColIndex
    ) {
      categoryColIndex = c;
    }
  }

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1 && !row1LooksLikeDataValues) return;
    const actionVal = cellText(row, 1);
    const statusValB = cellText(row, 2);
    const statusValE = cellText(row, 5);

    if (actionVal && !UUID_REGEX.test(actionVal)) actionsSet.add(actionVal);
    if (statusValB && !UUID_REGEX.test(statusValB)) statusesColB.add(statusValB);
    if (statusValE && !UUID_REGEX.test(statusValE)) statusesColE.add(statusValE);

    if (actionVal && !UUID_REGEX.test(actionVal) && categoryColIndex) {
      const catVal = cellText(row, categoryColIndex);
      if (catVal && !UUID_REGEX.test(catVal)) {
        actionCategoryMap.set(actionVal.toLowerCase(), catVal);
      }
    }

    for (let colIdx = 1; colIdx <= maxHeaderScan; colIdx++) {
      const cellVal = cellText(row, colIdx);
      if (cellVal && !UUID_REGEX.test(cellVal)) {
        if (!setsByCol.has(colIdx)) setsByCol.set(colIdx, new Set());
        setsByCol.get(colIdx).add(cellVal);
        maxRowByCol.set(colIdx, Math.max(maxRowByCol.get(colIdx) || 1, rowNumber));
      }
    }
  });

  headerByCol.forEach((lowerHeader, colIdx) => {
    const vals = Array.from(setsByCol.get(colIdx) || []);
    if (vals.length > 0) {
      validationListsByHeader.set(lowerHeader, vals);
    }
  });

  validationActions = Array.from(actionsSet);

  let chosenStatusColIdx = null;
  let chosenStatuses = [];

  if (statusHeaderMatchedCol && (setsByCol.get(statusHeaderMatchedCol) || new Set()).size > 0) {
    chosenStatusColIdx = statusHeaderMatchedCol;
    chosenStatuses = Array.from(setsByCol.get(statusHeaderMatchedCol));
  } else if (statusesColE.size > 0) {
    chosenStatusColIdx = 5;
    chosenStatuses = Array.from(statusesColE);
  } else if (statusesColB.size > 0) {
    chosenStatusColIdx = 2;
    chosenStatuses = Array.from(statusesColB);
  } else {
    // Check if any column on the validation sheet contains known status values (e.g. Col A on a single-list validation tab)
    for (const [colIdx, valSet] of setsByCol.entries()) {
      const arr = Array.from(valSet);
      const knownHits = arr.filter(v => KNOWN_STATUS_VALUES.has(v.toLowerCase())).length;
      if (knownHits >= 1 && arr.length > 0) {
        chosenStatusColIdx = colIdx;
        chosenStatuses = arr;
        break;
      }
    }
  }

  validationStatuses = chosenStatuses.filter(s => s && !UUID_REGEX.test(String(s).trim()));

  const startRow = row1LooksLikeDataValues ? 1 : 2;
  let statusRangeFormula = null;
  if (chosenStatusColIdx) {
    const colLet = colNumberToExcelLetters(chosenStatusColIdx);
    const endRow = Math.max(maxRowByCol.get(chosenStatusColIdx) || 10, chosenStatuses.length + startRow + 2);
    const escapedSheetName = worksheet.name.replace(/'/g, "''");
    statusRangeFormula = `'${escapedSheetName}'!$${colLet}$${startRow}:$${colLet}$${endRow}`;
  }

  validationSheetMeta = {
    sheetName: worksheet.name,
    startRow,
    statusColIdx: chosenStatusColIdx,
    statusRangeFormula
  };
}

function buildCustomSchemaDataValidationsExtLst() {
  if (!activeSchemaConfig || activeSchemaConfig.isStandardSchema || !activeSchemaConfig.statusCol) {
    return null;
  }
  const statusColLet = colNumberToExcelLetters(activeSchemaConfig.statusCol);
  let formulaInner = '';
  if (validationSheetMeta && validationSheetMeta.statusRangeFormula) {
    formulaInner = validationSheetMeta.statusRangeFormula;
  } else {
    const stColObj = activeSchemaConfig.columns.find(c => c.colIndex === activeSchemaConfig.statusCol);
    const opts = (stColObj && Array.isArray(stColObj.dropdownOptions) && stColObj.dropdownOptions.length > 0)
      ? stColObj.dropdownOptions
      : validationStatuses;
    const cleanOpts = (opts || []).filter(s => s && !UUID_REGEX.test(String(s).trim()));
    if (cleanOpts.length === 0) return null;
    const joined = cleanOpts.join(',');
    if (joined.length > 250) return null;
    formulaInner = `"${joined.replace(/"/g, '""')}"`;
  }

  const escapedFormula = formulaInner
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

  return `<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:dataValidations count="1" xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"><x14:formula1><xm:f>${escapedFormula}</xm:f></x14:formula1><xm:sqref>${statusColLet}2:${statusColLet}500</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>`;
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
    if (colMap.isCustomSchema) {
      const hasCustomData =
        (Array.isArray(colMap.columns) && colMap.columns.some(c => cellText(row, c.colIndex) !== '')) ||
        (colMap.notesCol && cellText(row, colMap.notesCol) !== '') ||
        (colMap.statusCol && cellText(row, colMap.statusCol) !== '');
      if (hasCustomData && rowNumber > lastRow) lastRow = rowNumber;
      return;
    }
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
  let tplWs = null;
  workbook.eachSheet((ws) => {
    if (ws.name.trim().toLowerCase() === 'template') {
      tplWs = ws;
    }
  });
  const primaryDailyWs = findPrimarySchemaHeaderWorksheet(workbook);

  // For custom schemas or when a daily tab has more columns than an outdated template sheet (e.g. Status added mid-month),
  // prefer the richest daily sheet so new tabs inherit all current columns, widths, and data validations
  if (primaryDailyWs && primaryDailyWs !== tplWs) {
    if (activeSchemaConfig && !activeSchemaConfig.isStandardSchema) {
      return primaryDailyWs;
    }
    if (tplWs) {
      const tplCols = extractDetectedRow1HeadersFromSheet(tplWs).columns.length;
      const dailyCols = extractDetectedRow1HeadersFromSheet(primaryDailyWs).columns.length;
      if (dailyCols > tplCols) {
        return primaryDailyWs;
      }
    }
  }

  if (tplWs) return tplWs;
  if (primaryDailyWs) return primaryDailyWs;
  if (dateSheets.length > 0) {
    return workbook.getWorksheet(dateSheets[0]);
  }
  return null;
}

/**
 * Orders all MM-DD-YYYY daily worksheet tabs chronologically from left to right
 * (oldest date on the left -> newest date on the right) while keeping all non-date
 * sheets (Scratch, Master Accounts, validation lists, template) to their right.
 */
function orderWorkbookSheetsChronologically(wb = workbook) {
  if (!wb) return;
  try {
    const dateTabs = [];
    const otherTabs = [];
    wb.eachSheet((ws) => {
      const raw = ws.name.trim();
      if (DATE_TAB_REGEX.test(raw)) {
        dateTabs.push(ws);
      } else {
        otherTabs.push(ws);
      }
    });

    dateTabs.sort((a, b) => parseDateTab(a.name.trim()) - parseDateTab(b.name.trim()));
    const ordered = [...dateTabs, ...otherTabs];
    ordered.forEach((ws, idx) => {
      ws.orderNo = idx + 1;
    });
  } catch (e) {
    console.warn('Could not reorder worksheet tabs chronologically:', e);
  }
}

function placeWorksheetFirstInWorkbook(wb, _targetSheet) {
  orderWorkbookSheetsChronologically(wb);
}

function placeWorksheetFirst(_targetSheet) {
  if (workbook) orderWorkbookSheetsChronologically(workbook);
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

  // Copy ExcelJS dataValidations.model so newly created daily tabs inherit dropdown lists
  if (templateSheet.dataValidations && templateSheet.dataValidations.model && newSheet.dataValidations) {
    newSheet.dataValidations.model = expandDataValidationModelRanges(templateSheet.dataValidations.model);
  }

  const colMap = getWorksheetColumnMap(newSheet);
  if (colMap.isCustomSchema && Array.isArray(colMap.columns)) {
    colMap.columns.forEach(c => {
      newSheet.getColumn(c.colIndex).hidden = false;
      if (!cellText(newHeader, c.colIndex)) {
        newHeader.getCell(c.colIndex).value = c.header;
      }
    });
    if (colMap.trackingCol && !cellText(newHeader, colMap.trackingCol)) {
      newHeader.getCell(colMap.trackingCol).value = 'Tracking ID';
    }
    newHeader.commit();
  }
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
  if (colMap && colMap.isCustomSchema) {
    const customMap = (item && item.customValues && typeof item.customValues === 'object')
      ? item.customValues
      : {};
    const cols = Array.isArray(colMap.columns) ? colMap.columns : [];
    cols.forEach(col => {
      const idx = col.colIndex;
      const key = `col_${idx}`;
      let val = customMap[key] !== undefined ? customMap[key] : '';
      if (val && UUID_REGEX.test(String(val).trim())) {
        val = '';
      }
      if (!val) {
        if (idx === colMap.primaryCol && item.subject) val = item.subject;
        else if (idx === colMap.ticketCol && item.ticket) val = item.ticket;
        else if (idx === colMap.secondaryCol && item.broker) val = item.broker;
        else if (idx === colMap.notesCol && item.notes) val = item.notes;
        else if (idx === colMap.statusCol && item.status && !UUID_REGEX.test(String(item.status).trim())) val = item.status;
      }
      if (idx === colMap.ticketCol && val) {
        val = String(val).toUpperCase();
      }
      row.getCell(idx).value = val !== undefined && val !== null ? String(val) : '';
    });
    if (colMap.notesCol && item.notes !== undefined && customMap[`col_${colMap.notesCol}`] === undefined) {
      row.getCell(colMap.notesCol).value = item.notes || '';
    }
    if (colMap.statusCol && item.status !== undefined && customMap[`col_${colMap.statusCol}`] === undefined) {
      const cleanSt = UUID_REGEX.test(String(item.status || '').trim()) ? getDefaultStatus() : (item.status || '');
      row.getCell(colMap.statusCol).value = cleanSt;
    }
    row.getCell(colMap.trackingCol).value = item.key;
    row.commit();
    return;
  }

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
  orderWorkbookSheetsChronologically(wb);
  if (wb === workbook) {
    syncWorkbookSheetsWithActiveSchema(wb);
  }
  wb.eachSheet((ws) => {
    const rawName = ws.name.trim();
    const lower = rawName.toLowerCase();
    const isDailyOrTemplate =
      lower === 'template' ||
      isTemplatePlaceholderTab(rawName) ||
      DATE_TAB_REGEX.test(rawName);
    if (!isDailyOrTemplate) return;

    const colMap = getWorksheetColumnMap(ws);
    if (colMap.isCustomSchema && Array.isArray(colMap.columns)) {
      colMap.columns.forEach(c => {
        ws.getColumn(c.colIndex).hidden = false;
      });
    }
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

// --- 8. DUAL-WRITE MIRRORING, PRIMARY FAIL-SAFE CACHE & RETRY ---
function showPrimarySaveRetryUI() {
  if (btnRetryPrimarySave) {
    btnRetryPrimarySave.style.display = 'inline-block';
    btnRetryPrimarySave.textContent = '💾 Retry Primary Save';
  }
  setStatus('⚠️ Staged in Cache (Excel Locked — Close & Retry)', 'locked');
}

function clearPrimarySaveRetryUI() {
  pendingPrimarySaveState = null;
  if (btnRetryPrimarySave) {
    btnRetryPrimarySave.style.display = 'none';
  }
}

async function retryStagedPrimarySaveNow() {
  if (!workingFileHandle) return;
  if (btnRetryPrimarySave) {
    btnRetryPrimarySave.textContent = '⏳ Saving...';
  }

  let bufferToWrite = null;
  try {
    if (workbook) {
      bufferToWrite = await serializeWorkbookBuffer();
    } else if (pendingPrimarySaveState && pendingPrimarySaveState.buffer) {
      bufferToWrite = pendingPrimarySaveState.buffer;
    }
  } catch (err) {
    console.error('Failed to serialize workbook during primary retry:', err);
    showPrimarySaveRetryUI();
    return;
  }

  if (!bufferToWrite) {
    clearPrimarySaveRetryUI();
    return;
  }

  try {
    const writable = await workingFileHandle.createWritable();
    await writable.write(bufferToWrite);
    await writable.close();

    const stagedKey = getStagedPrimarySaveIDBKey(workingFileHandle.name, activeProfileId);
    await deleteHandleFromIDB(stagedKey);
    clearPrimarySaveRetryUI();

    const mirrorOk = await attemptMirrorWrite(bufferToWrite, false);
    await loadWorkbookFromDisk(sheetSelect ? sheetSelect.value : null, {
      skipStagedRecoveryCheck: true,
      preserveStatus: !mirrorOk
    });
    showToast(`✅ Saved "${workingFileHandle.name}" & cleared staged fail-safe cache!`, 'success');
  } catch (err) {
    console.warn('Primary file is still locked during retry:', err);
    const stagedRecord = {
      profileId: activeProfileId,
      workbookName: workingFileHandle.name,
      savedAt: Date.now(),
      buffer: bufferToWrite
    };
    pendingPrimarySaveState = stagedRecord;
    await saveHandleToIDB(getStagedPrimarySaveIDBKey(workingFileHandle.name, activeProfileId), stagedRecord);
    showPrimarySaveRetryUI();
    showToast('⚠️ Working file is still open in Desktop Excel. Close Excel and click Retry Primary Save.', 'warning');
  }
}

if (btnRetryPrimarySave) {
  btnRetryPrimarySave.addEventListener('click', () => {
    retryStagedPrimarySaveNow();
  });
}

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
  if (!isActiveProfileMirrorEnabled()) {
    clearMirrorRetryState();
    if (workingFileHandle) {
      setStatus(`Saved: ${workingFileHandle.name} (Mirror Off)`, 'ready');
    }
    updateFilePathsBarUI();
    return true;
  }

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

  // 1. Primary Working File Write (with IndexedDB Fail-Safe Staging when Desktop Excel locks the file)
  try {
    const writable = await workingFileHandle.createWritable();
    await writable.write(buffer);
    await writable.close();

    const stagedKey = getStagedPrimarySaveIDBKey(workingFileHandle.name, activeProfileId);
    await deleteHandleFromIDB(stagedKey);
    clearPrimarySaveRetryUI();
  } catch (err) {
    console.error('Primary file write error:', err);
    const stagedRecord = {
      profileId: activeProfileId,
      workbookName: workingFileHandle.name,
      savedAt: Date.now(),
      buffer
    };
    pendingPrimarySaveState = stagedRecord;
    await saveHandleToIDB(getStagedPrimarySaveIDBKey(workingFileHandle.name, activeProfileId), stagedRecord);
    showPrimarySaveRetryUI();

    await showModalAlert({
      title: '🔒 File Locked in Excel — Edits Safely Staged!',
      badge: workingFileHandle.name,
      message:
        `The working workbook ("${workingFileHandle.name}") is currently open and locked in Desktop Excel.\n\n` +
        `✅ No work was lost: your changes are kept in memory and safely staged in TrackerUI's local fail-safe cache.\n` +
        `Close the file in Desktop Excel and click "💾 Retry Primary Save" in the bottom status bar.`,
      variant: 'danger'
    });
    return false;
  }

  // 2. Dual-Write Mirroring to (VIEW ONLY).xlsx with Auto-Retry & Manual Retry fallback
  const mirrorOk = await attemptMirrorWrite(buffer, false);
  return mirrorOk ? true : 'mirror-locked';
}

// ============================================================================
// 9. MULTI-DEPARTMENT TRACKER WORKSPACE PROFILES (TABS, FOLDERS & SNAPSHOTS)
// ============================================================================
function captureActiveProfileSnapshot() {
  return {
    parentDirHandle,
    cachedDirHandle,
    supervisorDirHandle,
    cachedSupervisorDirHandle,
    templatesDirHandle,
    cachedTemplatesDirHandle,
    availableTemplates: [...availableTemplates],
    lastSavedWorkbookName,
    availableWorkbooks: [...availableWorkbooks],
    workingFileHandle,
    workbook,
    originalArchiveState,
    allRows: [...allRows],
    dateSheets: [...dateSheets],
    scratchSheets: [...scratchSheets],
    openScratchTabs: [...openScratchTabs],
    activeWorkspaceTab,
    selectedSheetValue: sheetSelect ? sheetSelect.value : '',
    scopeValue: scopeSelect ? scopeSelect.value : 'current',
    searchQuery: searchInput ? searchInput.value : '',
    activeSchemaConfig: JSON.parse(JSON.stringify(activeSchemaConfig)),
    masterAccountsMap: [...masterAccountsMap],
    hasMasterAccountsSheet,
    validationActions: [...validationActions],
    validationStatuses: [...validationStatuses],
    validationListsByHeader: new Map(validationListsByHeader),
    actionCategoryMap: new Map(actionCategoryMap),
    pendingMirrorBuffer,
    pendingPrimarySaveState,
    dockEditingTarget: dockEditingTarget ? { ...dockEditingTarget } : null
  };
}

function applyProfileSnapshotToGlobals(snap) {
  parentDirHandle = snap.parentDirHandle || null;
  cachedDirHandle = snap.cachedDirHandle || null;
  supervisorDirHandle = snap.supervisorDirHandle || null;
  cachedSupervisorDirHandle = snap.cachedSupervisorDirHandle || null;
  templatesDirHandle = snap.templatesDirHandle || null;
  cachedTemplatesDirHandle = snap.cachedTemplatesDirHandle || null;
  availableTemplates = Array.isArray(snap.availableTemplates) ? [...snap.availableTemplates] : [];
  lastSavedWorkbookName = snap.lastSavedWorkbookName || null;
  availableWorkbooks = Array.isArray(snap.availableWorkbooks) ? [...snap.availableWorkbooks] : [];
  workingFileHandle = snap.workingFileHandle || null;
  workbook = snap.workbook || null;
  originalArchiveState = snap.originalArchiveState || null;
  allRows = Array.isArray(snap.allRows) ? [...snap.allRows] : [];
  dateSheets = Array.isArray(snap.dateSheets) ? [...snap.dateSheets] : [];
  scratchSheets = Array.isArray(snap.scratchSheets) ? [...snap.scratchSheets] : [];
  openScratchTabs = Array.isArray(snap.openScratchTabs) ? [...snap.openScratchTabs] : [];
  activeWorkspaceTab = DAILY_TRACKER_WS_TAB;
  activeSchemaConfig = snap.activeSchemaConfig
    ? JSON.parse(JSON.stringify(snap.activeSchemaConfig))
    : {
        isStandardSchema: true,
        mode: 'auto',
        columns: [],
        primaryCol: 1,
        ticketCol: 2,
        secondaryCol: 3,
        statusCol: 10,
        notesCol: 9,
        trackingCol: 11
      };
  masterAccountsMap = Array.isArray(snap.masterAccountsMap) ? [...snap.masterAccountsMap] : [];
  hasMasterAccountsSheet = Boolean(snap.hasMasterAccountsSheet);
  validationActions = Array.isArray(snap.validationActions) ? [...snap.validationActions] : [];
  validationStatuses = Array.isArray(snap.validationStatuses) ? [...snap.validationStatuses] : [];
  validationListsByHeader = snap.validationListsByHeader instanceof Map
    ? new Map(snap.validationListsByHeader)
    : new Map();
  actionCategoryMap = snap.actionCategoryMap instanceof Map
    ? new Map(snap.actionCategoryMap)
    : new Map();
  pendingMirrorBuffer = snap.pendingMirrorBuffer || null;
  pendingPrimarySaveState = snap.pendingPrimarySaveState || null;
  dockEditingTarget = snap.dockEditingTarget ? { ...snap.dockEditingTarget } : null;

  // Rebuild workbookSelect UI
  if (workbookSelect) {
    workbookSelect.innerHTML = '';
    availableWorkbooks.forEach((wbItem, idx) => {
      const opt = document.createElement('option');
      opt.value = String(idx);
      opt.textContent = wbItem.name;
      if (workingFileHandle && wbItem.name === workingFileHandle.name) {
        opt.selected = true;
      }
      workbookSelect.appendChild(opt);
    });
    workbookPickerWrap.style.display = availableWorkbooks.length > 0 ? 'flex' : 'none';
  }

  btnReconnect.style.display = (!parentDirHandle && cachedDirHandle) ? 'inline-flex' : 'none';
  btnLinkFile.style.display = parentDirHandle ? 'none' : 'inline-flex';
  btnNewWorkbook.disabled = !parentDirHandle;

  updateAgentDirUI();
  updateSupervisorDirUI();
  updateTemplatesDirUI();

  if (workbook && workingFileHandle) {
    if (typeof renderTableHeader === 'function') renderTableHeader();
    populateReferenceControls();
    if (typeof syncDynamicEditorFormsUI === 'function') syncDynamicEditorFormsUI();
    populateSheetDropdown(snap.selectedSheetValue || null);
    if (scopeSelect) scopeSelect.value = snap.scopeValue || 'current';
    if (searchInput) {
      searchInput.value = snap.searchQuery || '';
      searchInput.disabled = dateSheets.length === 0;
      if (!activeSchemaConfig.isStandardSchema && activeSchemaConfig.columns.length > 0) {
        const sampleHeaders = activeSchemaConfig.columns.slice(0, 5).map(c => c.header).join(', ');
        searchInput.placeholder = `Search across ALL daily tabs (${sampleHeaders}, or Tracking ID)...`;
      } else {
        searchInput.placeholder = 'Search across ALL daily tabs (Subject, Ticket #, Broker, Master #, Bidder #, Lot #, or Tracking ID)...';
      }
    }
    sheetSelect.disabled = dateSheets.length === 0;
    scopeSelect.disabled = dateSheets.length === 0;
    btnNewEntry.style.display = 'inline-flex';
    btnRollover.style.display = 'inline-flex';
    btnRepairSchema.disabled = false;
    if (btnOpenSchemaMapper) btnOpenSchemaMapper.disabled = false;
    if (menuNewScratchSheet) menuNewScratchSheet.disabled = false;
    btnAddBlankTab.disabled = false;
    if (btnRolloverPrevSheet) btnRolloverPrevSheet.disabled = dateSheets.length === 0;
    btnRenameTab.disabled = dateSheets.length === 0;
    btnDeleteTab.disabled = dateSheets.length === 0;
    btnDeleteMode.disabled = allRows.length === 0;
    schemaBanner.style.display = 'none';

    renderSheetTabsBar();
    updateSheetStepperState();
    if (bottomEditorDock) {
      bottomEditorDock.style.display = 'block';
      syncBottomDockStateAfterLoad();
    }
    if (pendingPrimarySaveState) {
      showPrimarySaveRetryUI();
    } else if (pendingMirrorBuffer) {
      scheduleMirrorAutoRetry();
    } else {
      clearPrimarySaveRetryUI();
      clearMirrorRetryState();
      setStatus(`Synced: ${workingFileHandle.name}`, 'ready');
    }
    renderList();
    if (typeof setActiveWorkspaceTab === 'function') {
      setActiveWorkspaceTab(DAILY_TRACKER_WS_TAB);
    }
  } else {
    resetActiveTrackerUIDisconnectedState();
  }
}

async function switchTrackerProfile(targetProfileId) {
  const targetProf = trackerProfiles.find(p => p.id === targetProfileId);
  if (!targetProf) return;

  if (targetProfileId === activeProfileId) {
    if (activeWorkspaceTab !== DAILY_TRACKER_WS_TAB && typeof setActiveWorkspaceTab === 'function') {
      setActiveWorkspaceTab(DAILY_TRACKER_WS_TAB);
    }
    return;
  }

  commitActiveScratchInlineEdit();
  if (hasAnyPendingScratchMutations() && workbook) {
    await flushPendingScratchEditsToDisk({ reloadAfter: false, silentToast: true });
  }

  trackerProfileRuntimeCache.set(activeProfileId, captureActiveProfileSnapshot());

  if (mirrorCountdownInterval) {
    clearInterval(mirrorCountdownInterval);
    mirrorCountdownInterval = null;
  }
  selectedForDelete.clear();
  activeKeyFilter = null;
  activeFieldFilter = null;
  editingTarget = null;
  dockEditingTarget = null;

  activeProfileId = targetProfileId;
  persistTrackerProfiles();
  syncMirrorCopyCheckboxUI();

  if (trackerProfileRuntimeCache.has(targetProfileId)) {
    const cachedSnap = trackerProfileRuntimeCache.get(targetProfileId);
    applyProfileSnapshotToGlobals(cachedSnap);
    if (typeof syncWorkspaceTabsUI === 'function') syncWorkspaceTabsUI();
    return;
  }

  // Fresh profile activation in this session: clear runtime globals and load IDB handles
  parentDirHandle = null;
  cachedDirHandle = null;
  supervisorDirHandle = null;
  cachedSupervisorDirHandle = null;
  templatesDirHandle = null;
  cachedTemplatesDirHandle = null;
  availableTemplates = [];
  lastSavedWorkbookName = null;
  availableWorkbooks = [];
  workingFileHandle = null;
  workbook = null;
  originalArchiveState = null;
  allRows = [];
  dateSheets = [];
  scratchSheets = [];
  openScratchTabs = [];
  activeWorkspaceTab = DAILY_TRACKER_WS_TAB;
  pendingMirrorBuffer = null;
  pendingPrimarySaveState = null;
  clearPrimarySaveRetryUI();
  clearMirrorRetryState();

  if (typeof setActiveWorkspaceTab === 'function') {
    setActiveWorkspaceTab(DAILY_TRACKER_WS_TAB);
  }
  resetActiveTrackerUIDisconnectedState();

  if ('indexedDB' in window) {
    const idbKeys = getProfileIDBKeys(activeProfileId);
    const [savedAgent, savedSup, savedTpl, savedWb] = await Promise.all([
      getHandleFromIDB(idbKeys.agentKey),
      getHandleFromIDB(idbKeys.supervisorKey),
      getHandleFromIDB(idbKeys.templatesKey),
      getHandleFromIDB(idbKeys.lastWbKey)
    ]);
    if (savedWb && typeof savedWb === 'string') lastSavedWorkbookName = savedWb;
    if (savedSup) cachedSupervisorDirHandle = savedSup;
    if (savedTpl) cachedTemplatesDirHandle = savedTpl;
    if (savedAgent) {
      cachedDirHandle = savedAgent;
      setReconnectButtonText(`Reconnect (${savedAgent.name})`);
      btnReconnect.style.display = 'inline-flex';
      updateAgentDirUI();
      updateSupervisorDirUI();
      updateTemplatesDirUI();
      if (autoReconnectEnabled) {
        await attemptReconnectFolders({ requestPermissionIfNeeded: true, silent: true });
      }
    }
  }
  if (typeof syncWorkspaceTabsUI === 'function') syncWorkspaceTabsUI();
}

async function createTrackerProfileInteractive() {
  closeToolsMenu();
  const suggestedNum = trackerProfiles.length + 1;
  const rawInput = await showModalPrompt({
    title: '＋ Add Department Tracker Tab',
    badge: 'Multi-Tracker',
    message:
      'Create an independent Tracker Workspace tab for another department or workflow. ' +
      'Each Tracker Tab links to its own folder and automatically adapts to standard or custom Row-1 Excel headers.',
    selectLabel: 'Supervisor (VIEW ONLY) Mirroring',
    selectOptions: [
      { value: 'mirror_on', label: '✅ Enable (VIEW ONLY).xlsx Mirror Copy (Default)' },
      { value: 'mirror_off', label: '📄 Working File Only (Disable Mirror Copy)' }
    ],
    defaultSelectValue: 'mirror_on',
    inputLabel: 'Department / Tracker Tab Name',
    defaultValue: `Dept ${suggestedNum} Tracker`,
    placeholder: 'e.g., Titles Dept Tracker, Escalations, Arbitration',
    confirmText: '＋ Add Tracker Tab',
    validate: (val) => {
      if (val.length > 40) return 'Please keep the tab name under 40 characters.';
      if (trackerProfiles.some(p => p.name.toLowerCase() === val.toLowerCase())) {
        return `A tracker tab named "${val}" already exists.`;
      }
      return null;
    }
  });
  if (!rawInput) return;

  const tabName = (typeof rawInput === 'object' ? rawInput.value : rawInput).trim();
  const mirrorChoice = (typeof rawInput === 'object' && rawInput.selectValue) ? rawInput.selectValue : 'mirror_on';
  const newId = `dept_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 6)}`;

  trackerProfiles.push({
    id: newId,
    name: tabName,
    mirrorEnabled: mirrorChoice !== 'mirror_off'
  });
  persistTrackerProfiles();

  await switchTrackerProfile(newId);
  showToast(`Added "${tabName}" workspace tab. Select its Agent Folder to load workbooks.`, 'success');
}

async function renameActiveTrackerProfileInteractive(targetProfileId = activeProfileId) {
  closeToolsMenu();
  const prof = trackerProfiles.find(p => p.id === targetProfileId);
  if (!prof) return;

  const rawNewName = await showModalPrompt({
    title: '✏️ Rename Tracker Tab',
    badge: prof.name,
    message: `Enter a display name for the "${prof.name}" workspace tab:`,
    inputLabel: 'Tracker Tab Name',
    defaultValue: prof.name,
    placeholder: 'e.g., Broker Support, Titles Dept',
    confirmText: 'Save Name',
    validate: (val) => {
      if (val.length > 40) return 'Please keep the tab name under 40 characters.';
      if (trackerProfiles.some(p => p.id !== prof.id && p.name.toLowerCase() === val.toLowerCase())) {
        return `Another tracker tab is already named "${val}".`;
      }
      return null;
    }
  });
  if (!rawNewName) return;
  const cleanName = (typeof rawNewName === 'object' ? rawNewName.value : rawNewName).trim();
  if (!cleanName || cleanName === prof.name) return;

  prof.name = cleanName;
  persistTrackerProfiles();
  if (typeof syncWorkspaceTabsUI === 'function') syncWorkspaceTabsUI();
  updateStatusBarTelemetry();
  showToast(`Renamed tracker tab to "${cleanName}".`, 'success');
}

async function removeTrackerProfileInteractive(targetProfileId) {
  if (!targetProfileId || targetProfileId === DEFAULT_TRACKER_PROFILE_ID) {
    return showModalAlert({
      title: 'Primary Tracker Tab Protected',
      message: 'The primary Daily Tracker tab cannot be removed (you can rename it or change its folder anytime).'
    });
  }
  const prof = trackerProfiles.find(p => p.id === targetProfileId);
  if (!prof) return;

  const confirmed = await showModalConfirm({
    title: `Remove "${prof.name}" Tab?`,
    badge: 'Workspace Tab Only',
    message:
      `Remove the "${prof.name}" workspace tab from TrackerUI?\n\n` +
      `Your Excel (.xlsx) files on disk will NOT be deleted or modified.`,
    confirmText: 'Remove Tab',
    variant: 'danger'
  });
  if (!confirmed) return;

  if (activeProfileId === targetProfileId) {
    await switchTrackerProfile(DEFAULT_TRACKER_PROFILE_ID);
  }

  const idbKeys = getProfileIDBKeys(targetProfileId);
  await Promise.all([
    deleteHandleFromIDB(idbKeys.agentKey),
    deleteHandleFromIDB(idbKeys.supervisorKey),
    deleteHandleFromIDB(idbKeys.templatesKey),
    deleteHandleFromIDB(idbKeys.lastWbKey)
  ]);

  trackerProfiles = trackerProfiles.filter(p => p.id !== targetProfileId);
  trackerProfileRuntimeCache.delete(targetProfileId);
  persistTrackerProfiles();
  if (typeof syncWorkspaceTabsUI === 'function') syncWorkspaceTabsUI();
  updateStatusBarTelemetry();
  showToast(`Removed "${prof.name}" workspace tab.`, 'info');
}

if (btnAddTrackerProfile) {
  btnAddTrackerProfile.addEventListener('click', createTrackerProfileInteractive);
}
if (menuAddTrackerProfile) {
  menuAddTrackerProfile.addEventListener('click', createTrackerProfileInteractive);
}
if (menuRenameTrackerProfile) {
  menuRenameTrackerProfile.addEventListener('click', () => renameActiveTrackerProfileInteractive(activeProfileId));
}

// ============================================================================
// 10. DYNAMIC COLUMN ROLE MAPPER MODAL (#schemaMapperModal)
// ============================================================================
let schemaMapperDetectedState = null;

function buildRoleSelectOptionsHtml(cols, selectedColIdx, allowNone = true) {
  let html = allowNone ? '<option value="">— None / Not Used —</option>' : '';
  cols.forEach(c => {
    const sel = Number(selectedColIdx) === c.colIndex ? ' selected' : '';
    const letter = colNumberToExcelLetters(c.colIndex);
    html += `<option value="${c.colIndex}"${sel}>Col ${letter} (${c.colIndex}): ${escapeHtml(c.header)}</option>`;
  });
  return html;
}

function refreshSchemaMapperPreviewUI() {
  if (!schemaMapperDetectedState || !schemaDetectedColsPreview) return;
  const modeVal = schemaModeSelect ? schemaModeSelect.value : 'auto';
  const effectiveStandard =
    modeVal === 'standard' ||
    (modeVal === 'auto' && schemaMapperDetectedState.autoIsStandard);

  [
    schemaPrimaryColSelect,
    schemaStatusColSelect,
    schemaNotesColSelect,
    schemaTicketColSelect,
    schemaSecondaryColSelect
  ].forEach(sel => {
    if (sel) sel.disabled = effectiveStandard;
  });

  if (schemaMapperModeBadge) {
    if (effectiveStandard) {
      schemaMapperModeBadge.textContent = modeVal === 'auto' ? 'Auto: Standard 11-Col' : 'Standard 11-Col';
      schemaMapperModeBadge.className = 'badge badge-complete';
    } else {
      schemaMapperModeBadge.textContent = modeVal === 'auto' ? 'Auto: Custom Dept Schema' : 'Custom Dept Schema';
      schemaMapperModeBadge.className = 'badge badge-progress';
    }
  }

  const pCol = Number(schemaPrimaryColSelect?.value) || null;
  const stCol = Number(schemaStatusColSelect?.value) || null;
  const nCol = Number(schemaNotesColSelect?.value) || null;
  const tCol = Number(schemaTicketColSelect?.value) || null;
  const secCol = Number(schemaSecondaryColSelect?.value) || null;

  const cols = schemaMapperDetectedState.columns;
  if (cols.length === 0) {
    schemaDetectedColsPreview.innerHTML = '<div class="control-label">No Row-1 headers detected on active sheets.</div>';
    return;
  }

  schemaDetectedColsPreview.innerHTML = cols.map(c => {
    const letter = colNumberToExcelLetters(c.colIndex);
    const roles = [];
    if (!effectiveStandard) {
      if (c.colIndex === pCol) roles.push('Primary');
      if (c.colIndex === stCol) roles.push('Status');
      if (c.colIndex === nCol) roles.push('Notes');
      if (c.colIndex === tCol) roles.push('Ticket #');
      if (c.colIndex === secCol) roles.push('Secondary');
    }
    const valOpts = findValidationOptionsForHeader(c.header);
    const roleTag = roles.length > 0
      ? `<span class="schema-col-role-tag">${escapeHtml(roles.join(' + '))}</span>`
      : '';
    const valTag = valOpts.length > 0
      ? `<span class="schema-col-role-tag" title="Bound to validation lists dropdown">▾ ${valOpts.length} opts</span>`
      : '';
    return `
      <span class="schema-col-chip">
        <span class="col-letter">Col ${letter}</span>
        <span>${escapeHtml(c.header)}</span>
        ${roleTag}${valTag}
      </span>
    `;
  }).join('');
}

function openSchemaMapperModal() {
  closeToolsMenu();
  if (!workbook || !workingFileHandle || !schemaMapperModal) return;

  const headerWs = findPrimarySchemaHeaderWorksheet(workbook);
  const detected = headerWs ? extractDetectedRow1HeadersFromSheet(headerWs) : { columns: [], maxCol: 0, existingTrackingCol: null };
  const autoIsStandard = headerWs ? isBrokerStandardHeaderRow(headerWs) : true;
  const cols = detected.columns.length > 0
    ? detected.columns
    : STANDARD_HEADERS.slice(0, 10).map((h, i) => ({ colIndex: i + 1, header: h }));

  const savedOverride = getSavedSchemaOverrideForWorkbook(workingFileHandle.name, activeProfileId);
  const currentMode = (savedOverride && savedOverride.mode) ? savedOverride.mode : (activeSchemaConfig.mode || 'auto');

  schemaMapperDetectedState = {
    columns: cols,
    autoIsStandard,
    trackingCol: detected.existingTrackingCol || (Math.max(...cols.map(c => c.colIndex), 0) + 1)
  };

  if (schemaMapperSubtitle) {
    schemaMapperSubtitle.textContent = `Workbook: ${workingFileHandle.name} • Source Sheet: ${headerWs ? headerWs.name : 'Default'}`;
  }
  if (schemaModeSelect) {
    schemaModeSelect.value = ['auto', 'custom', 'standard'].includes(currentMode) ? currentMode : 'auto';
  }

  const defaultPrimary = activeSchemaConfig.primaryCol || cols[0].colIndex;
  const defaultStatus = activeSchemaConfig.isStandardSchema ? '' : (activeSchemaConfig.statusCol || '');
  const defaultNotes = activeSchemaConfig.isStandardSchema ? '' : (activeSchemaConfig.notesCol || '');
  const defaultTicket = activeSchemaConfig.isStandardSchema ? '' : (activeSchemaConfig.ticketCol || '');
  const defaultSecondary = activeSchemaConfig.isStandardSchema ? '' : (activeSchemaConfig.secondaryCol || '');

  // If currently standard, pre-populate sensible inferred defaults in case user switches dropdown to 'custom'
  const findInfer = (re) => {
    const hit = cols.find(c => re.test(c.header));
    return hit ? hit.colIndex : '';
  };

  if (schemaPrimaryColSelect) {
    schemaPrimaryColSelect.innerHTML = buildRoleSelectOptionsHtml(cols, defaultPrimary || findInfer(/subject|title|request|vin|lot|item|account/i) || cols[0].colIndex, false);
  }
  if (schemaStatusColSelect) {
    schemaStatusColSelect.innerHTML = buildRoleSelectOptionsHtml(cols, defaultStatus !== '' ? defaultStatus : findInfer(/status|state|stage|disposition/i), true);
  }
  if (schemaNotesColSelect) {
    schemaNotesColSelect.innerHTML = buildRoleSelectOptionsHtml(cols, defaultNotes !== '' ? defaultNotes : findInfer(/note|comment|remark|action taken|update/i), true);
  }
  if (schemaTicketColSelect) {
    schemaTicketColSelect.innerHTML = buildRoleSelectOptionsHtml(cols, defaultTicket !== '' ? defaultTicket : findInfer(/ticket|case|incident|ref/i), true);
  }
  if (schemaSecondaryColSelect) {
    schemaSecondaryColSelect.innerHTML = buildRoleSelectOptionsHtml(cols, defaultSecondary !== '' ? defaultSecondary : findInfer(/broker|account|dealer|owner|assigned|dept/i), true);
  }
  if (schemaTrackingColInfo) {
    const trkIdx = activeSchemaConfig.isStandardSchema ? 11 : schemaMapperDetectedState.trackingCol;
    schemaTrackingColInfo.textContent = `Col ${colNumberToExcelLetters(trkIdx)} (${trkIdx}) — Hidden UUIDv4 Tracking ID`;
  }

  refreshSchemaMapperPreviewUI();
  schemaMapperModal.classList.add('open');
}

if (btnOpenSchemaMapper) {
  btnOpenSchemaMapper.addEventListener('click', openSchemaMapperModal);
}
if (schemaModeSelect) {
  schemaModeSelect.addEventListener('change', refreshSchemaMapperPreviewUI);
}
[
  schemaPrimaryColSelect,
  schemaStatusColSelect,
  schemaNotesColSelect,
  schemaTicketColSelect,
  schemaSecondaryColSelect
].forEach(el => {
  if (el) el.addEventListener('change', refreshSchemaMapperPreviewUI);
});

if (btnCancelSchemaMapper) {
  btnCancelSchemaMapper.addEventListener('click', () => {
    schemaMapperModal.classList.remove('open');
  });
}
if (schemaMapperModal) {
  schemaMapperModal.addEventListener('click', (e) => {
    if (e.target === schemaMapperModal) schemaMapperModal.classList.remove('open');
  });
}

if (btnResetSchemaMapper) {
  btnResetSchemaMapper.addEventListener('click', async () => {
    if (!workingFileHandle) return;
    saveSchemaOverrideForWorkbook(workingFileHandle.name, null, activeProfileId);
    schemaMapperModal.classList.remove('open');
    await loadWorkbookFromDisk(sheetSelect ? sheetSelect.value : null, { skipStagedRecoveryCheck: true });
    showToast('Reset column mapping to Auto-Detect.', 'info');
  });
}

if (schemaMapperForm) {
  schemaMapperForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    if (!workingFileHandle) return;
    const mode = schemaModeSelect ? schemaModeSelect.value : 'auto';
    const parseOptCol = (sel) => {
      if (!sel || sel.value === '') return null;
      const n = parseInt(sel.value, 10);
      return Number.isFinite(n) && n >= 1 ? n : null;
    };
    const override = {
      mode,
      primaryCol: parseOptCol(schemaPrimaryColSelect) || 1,
      statusCol: parseOptCol(schemaStatusColSelect),
      notesCol: parseOptCol(schemaNotesColSelect),
      ticketCol: parseOptCol(schemaTicketColSelect),
      secondaryCol: parseOptCol(schemaSecondaryColSelect)
    };
    saveSchemaOverrideForWorkbook(workingFileHandle.name, override, activeProfileId);
    schemaMapperModal.classList.remove('open');
    await loadWorkbookFromDisk(sheetSelect ? sheetSelect.value : null, { skipStagedRecoveryCheck: true });
    showToast(`Saved column role mapping for "${workingFileHandle.name}".`, 'success');
  });
}

// ============================================================================
// 11. BACKGROUND WEB WORKER PARSER & SUPERVISOR MULTI-WORKBOOK MASTER ROLLUP
// ============================================================================
let rollupWorkerInstance = null;
let rollupWorkerFailed = false;
let rollupWorkerReqSeq = 0;
const rollupWorkerPendingMap = new Map();

function discoverExcelJsScriptUrl() {
  if (typeof document !== 'undefined' && document.scripts) {
    for (const s of Array.from(document.scripts)) {
      if (s.src && /exceljs/i.test(s.src)) {
        return s.src;
      }
    }
  }
  return 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js';
}

function getOrCreateRollupWorker() {
  if (rollupWorkerFailed || typeof Worker === 'undefined' || typeof Blob === 'undefined') {
    return null;
  }
  if (rollupWorkerInstance) return rollupWorkerInstance;

  const excelJsUrl = discoverExcelJsScriptUrl();
  const workerSource = `
    let excelReady = false;
    try {
      importScripts(${JSON.stringify(excelJsUrl)});
      excelReady = Boolean(self.ExcelJS);
    } catch (err) {
      excelReady = false;
    }

    const DATE_RE = /^\\d{1,2}-\\d{1,2}-\\d{2,4}$/;
    function parseDateMs(name) {
      const p = String(name || '').trim().split('-').map(Number);
      if (p.length !== 3) return 0;
      let [m, d, y] = p;
      if (y < 100) y += 2000;
      return new Date(y, m - 1, d).getTime();
    }
    function cellStr(row, c) {
      const v = row.getCell(c).value;
      if (v === null || v === undefined) return '';
      if (typeof v === 'object') {
        if (v instanceof Date) return v.toISOString().split('T')[0];
        if (Array.isArray(v.richText)) return v.richText.map(t => t.text || '').join('').trim();
        if (v.text !== undefined && v.text !== null && typeof v.text !== 'object') return String(v.text).trim();
        if (v.result !== undefined && v.result !== null && typeof v.result !== 'object') return String(v.result).trim();
        return '';
      }
      return String(v).trim();
    }

    self.onmessage = async (evt) => {
      const { id, fileName, buffer } = evt.data || {};
      if (!excelReady || !self.ExcelJS) {
        self.postMessage({ id, ok: false, error: 'ExcelJS unavailable in Worker' });
        return;
      }
      try {
        const wb = new self.ExcelJS.Workbook();
        await wb.xlsx.load(buffer);
        const categoryPairs = [];
        const dailySheets = [];
        let templateWs = null;

        wb.eachSheet((ws) => {
          const raw = ws.name.trim();
          const low = raw.toLowerCase();
          if (DATE_RE.test(raw)) {
            dailySheets.push({ ws, name: raw, ms: parseDateMs(raw) });
          } else if (low === 'template') {
            templateWs = ws;
          } else if (low === 'validation lists' || low === 'validation list') {
            const r1 = ws.getRow(1);
            let catCol = null;
            for (let c = 2; c <= 16; c++) {
              const h = cellStr(r1, c);
              if (c !== 5 && /\\b(category|parent|group|bucket|pipeline|department|domain)\\b/i.test(h)) {
                catCol = c;
                break;
              }
            }
            if (catCol) {
              ws.eachRow({ includeEmpty: false }, (r, rn) => {
                if (rn === 1) return;
                const act = cellStr(r, 1);
                const cat = cellStr(r, catCol);
                if (act && cat) categoryPairs.push([act.toLowerCase(), cat]);
              });
            }
          }
        });

        dailySheets.sort((a, b) => a.ms - b.ms);
        const headerWs = dailySheets.length > 0 ? dailySheets[dailySheets.length - 1].ws : templateWs;
        let isStandard = true;
        let pCol = 1, tCol = 2, secCol = 3, actCol = 6, nCol = 9, stCol = 10, trkCol = 11;
        let maxScanCol = 11;

        if (headerWs) {
          const r1 = headerWs.getRow(1);
          const h1 = cellStr(r1, 1).toLowerCase();
          const h3 = cellStr(r1, 3).toLowerCase();
          const h4 = cellStr(r1, 4).toLowerCase();
          const h6 = cellStr(r1, 6).toLowerCase();
          let stdHits = 0;
          if (h1.includes('subject')) stdHits++;
          if (h3.includes('broker') || h3.includes('account')) stdHits++;
          if (h4.includes('master')) stdHits++;
          if (h6.includes('action')) stdHits++;
          isStandard = stdHits >= 3;

          if (isStandard) {
            const h9 = cellStr(r1, 9).toLowerCase();
            const h10 = cellStr(r1, 10).toLowerCase();
            const h11 = cellStr(r1, 11).toLowerCase();
            if (h9.includes('status') || h10.includes('action taken') || h11.includes('sent email')) {
              stCol = 9; nCol = 10; trkCol = 12; maxScanCol = 12;
            }
          } else {
            const cols = [];
            let foundTrk = null;
            for (let c = 1; c <= 40; c++) {
              const h = cellStr(r1, c);
              if (!h) continue;
              if (/^(tracking\\s*id|tracking_id|uuid|guid)$/i.test(h)) {
                foundTrk = c;
                continue;
              }
              cols.push({ c, h });
            }
            if (cols.length > 0) {
              const pick = (re) => { const f = cols.find(x => re.test(x.h)); return f ? f.c : null; };
              stCol = pick(/status|state|stage|disposition/i);
              nCol = pick(/note|comment|remark|action taken|update/i);
              tCol = pick(/ticket|case|incident|ref/i);
              actCol = pick(/action|type|category|reason|issue|subtype|request type/i);
              pCol = pick(/subject|title|summary|request|vin|lot|item/i) || cols[0].c;
              secCol = pick(/broker|account|dealer|customer|owner|assigned|dept/i) || (cols[1] ? cols[1].c : null);
              maxScanCol = Math.max(...cols.map(x => x.c));
              trkCol = foundTrk || (maxScanCol + 1);
            }
          }
        }

        const rows = [];
        for (const { ws, name, ms } of dailySheets) {
          ws.eachRow({ includeEmpty: false }, (r, rn) => {
            if (rn === 1) return;
            let hasAny = false;
            for (let c = 1; c <= maxScanCol; c++) {
              if (c !== trkCol && cellStr(r, c) !== '') { hasAny = true; break; }
            }
            if (!hasAny) return;

            const subj = pCol ? cellStr(r, pCol) : '';
            const tkt = tCol ? cellStr(r, tCol).toUpperCase() : '';
            let acct = '';
            if (isStandard) {
              const brk = cellStr(r, 3);
              const mst = cellStr(r, 4).toUpperCase();
              acct = brk && mst ? (brk + ' (' + mst + ')') : (brk || mst);
            } else if (secCol) {
              acct = cellStr(r, secCol);
            }
            const act = actCol ? cellStr(r, actCol) : '';
            const nts = nCol ? cellStr(r, nCol) : '';
            const st = (stCol ? cellStr(r, stCol) : '') || 'Pending';
            const key = (trkCol ? cellStr(r, trkCol).toLowerCase() : '') || '';

            rows.push({
              workbookName: fileName,
              sheetName: name,
              sheetMs: ms,
              rowNumber: rn,
              subject: subj || '(Untitled)',
              ticket: tkt,
              account: acct,
              action: act || (isStandard ? '' : 'Custom Dept Entry'),
              notes: nts,
              status: st,
              key,
              isStandard
            });
          });
        }

        self.postMessage({ id, ok: true, rows, categoryPairs });
      } catch (e) {
        self.postMessage({ id, ok: false, error: e && e.message ? e.message : String(e) });
      }
    };
  `;

  try {
    const blob = new Blob([workerSource], { type: 'application/javascript' });
    const url = URL.createObjectURL(blob);
    const worker = new Worker(url);
    worker.onmessage = (evt) => {
      const data = evt.data || {};
      const pending = rollupWorkerPendingMap.get(data.id);
      if (!pending) return;
      rollupWorkerPendingMap.delete(data.id);
      if (data.ok) {
        pending.resolve({ rows: data.rows || [], categoryPairs: data.categoryPairs || [], usedWorker: true });
      } else {
        pending.reject(new Error(data.error || 'Worker parse error'));
      }
    };
    worker.onerror = (err) => {
      rollupWorkerFailed = true;
      rollupWorkerPendingMap.forEach(p => p.reject(err));
      rollupWorkerPendingMap.clear();
      try { worker.terminate(); } catch (_) { /* ignore */ }
      rollupWorkerInstance = null;
    };
    rollupWorkerInstance = worker;
    return worker;
  } catch (err) {
    rollupWorkerFailed = true;
    return null;
  }
}

async function parseWorkbookBufferMainThreadFallback(sanitizedBuffer, fileName) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(sanitizedBuffer);
  const categoryPairs = [];
  const dailySheets = [];
  let templateWs = null;

  wb.eachSheet((ws) => {
    const raw = ws.name.trim();
    const low = raw.toLowerCase();
    if (DATE_TAB_REGEX.test(raw)) {
      dailySheets.push({ ws, name: raw, ms: parseDateTab(raw) });
    } else if (low === 'template') {
      templateWs = ws;
    } else if (low === 'validation lists' || low === 'validation list') {
      const r1 = ws.getRow(1);
      let catCol = null;
      for (let c = 2; c <= 16; c++) {
        const h = cellText(r1, c);
        if (c !== 5 && /\b(category|parent|group|bucket|pipeline|department|domain)\b/i.test(h)) {
          catCol = c;
          break;
        }
      }
      if (catCol) {
        ws.eachRow({ includeEmpty: false }, (r, rn) => {
          if (rn === 1) return;
          const act = cellText(r, 1);
          const cat = cellText(r, catCol);
          if (act && cat) categoryPairs.push([act.toLowerCase(), cat]);
        });
      }
    }
  });

  dailySheets.sort((a, b) => a.ms - b.ms);
  const headerWs = dailySheets.length > 0 ? dailySheets[dailySheets.length - 1].ws : templateWs;
  const isStandard = headerWs ? isBrokerStandardHeaderRow(headerWs) : true;

  let pCol = 1, tCol = 2, secCol = 3, actCol = 6, nCol = 9, stCol = 10, trkCol = 11, maxScanCol = 11;
  if (headerWs && isStandard) {
    const stdMap = getStandardWorksheetColumnMapOnly(headerWs);
    stCol = stdMap.statusCol;
    nCol = stdMap.notesCol;
    trkCol = stdMap.trackingCol;
    maxScanCol = trkCol;
  } else if (headerWs) {
    const det = extractDetectedRow1HeadersFromSheet(headerWs);
    if (det.columns.length > 0) {
      const pick = (re) => { const f = det.columns.find(x => re.test(x.header)); return f ? f.colIndex : null; };
      stCol = pick(/status|state|stage|disposition/i);
      nCol = pick(/note|comment|remark|action taken|update/i);
      tCol = pick(/ticket|case|incident|ref/i);
      actCol = pick(/action|type|category|reason|issue|subtype|request type/i);
      pCol = pick(/subject|title|summary|request|vin|lot|item/i) || det.columns[0].colIndex;
      secCol = pick(/broker|account|dealer|customer|owner|assigned|dept/i) || (det.columns[1] ? det.columns[1].colIndex : null);
      maxScanCol = det.maxCol;
      trkCol = det.existingTrackingCol || (det.maxCol + 1);
    }
  }

  const rows = [];
  for (const { ws, name, ms } of dailySheets) {
    ws.eachRow({ includeEmpty: false }, (r, rn) => {
      if (rn === 1) return;
      let hasAny = false;
      for (let c = 1; c <= maxScanCol; c++) {
        if (c !== trkCol && cellText(r, c) !== '') { hasAny = true; break; }
      }
      if (!hasAny) return;
      const subj = pCol ? cellText(r, pCol) : '';
      const tkt = tCol ? cellText(r, tCol).toUpperCase() : '';
      let acct = '';
      if (isStandard) {
        const brk = cellText(r, 3);
        const mst = cellText(r, 4).toUpperCase();
        acct = brk && mst ? `${brk} (${mst})` : (brk || mst);
      } else if (secCol) {
        acct = cellText(r, secCol);
      }
      const act = actCol ? cellText(r, actCol) : '';
      const nts = nCol ? cellText(r, nCol) : '';
      const st = (stCol ? cellText(r, stCol) : '') || 'Pending';
      const key = trkCol ? cellText(r, trkCol).toLowerCase() : '';
      rows.push({
        workbookName: fileName,
        sheetName: name,
        sheetMs: ms,
        rowNumber: rn,
        subject: subj || '(Untitled)',
        ticket: tkt,
        account: acct,
        action: act || (isStandard ? '' : 'Custom Dept Entry'),
        notes: nts,
        status: st,
        key,
        isStandard
      });
    });
  }

  return { rows, categoryPairs, usedWorker: false };
}

async function parseWorkbookBufferInBackground(rawBuffer, fileName) {
  const sanitized = sanitizeXlsxBufferForExcelJS(rawBuffer);
  const worker = getOrCreateRollupWorker();
  if (worker) {
    try {
      const reqId = ++rollupWorkerReqSeq;
      const workerCopy = sanitized.slice(0);
      return await new Promise((resolve, reject) => {
        rollupWorkerPendingMap.set(reqId, { resolve, reject });
        worker.postMessage({ id: reqId, fileName, buffer: workerCopy }, [workerCopy]);
      });
    } catch (workerErr) {
      rollupWorkerFailed = true;
      console.warn('Web Worker parse fallback to main thread:', workerErr);
    }
  }
  return await parseWorkbookBufferMainThreadFallback(sanitized, fileName);
}

// --- Supervisor Master Rollup State & Controller ---
let rollupCustomDirHandle = null;
let masterRollupState = {
  folderName: '',
  filesScanned: [],
  rawRows: [],
  dedupedRows: [],
  localCategoryMap: new Map(),
  usedWorker: false
};

function consolidateRollupRowsPerWorkbook(rawRows) {
  const groups = [];
  rawRows.forEach(row => {
    let matched = null;
    if (row.key) {
      matched = groups.find(g => g.workbookName === row.workbookName && g.key && g.key === row.key);
    }
    if (!matched && row.ticket) {
      const tokens = splitNormalizedTicketTokens(row.ticket);
      if (tokens.length > 0) {
        matched = groups.find(g =>
          g.workbookName === row.workbookName &&
          g.ticketTokens.some(t => tokens.includes(t))
        );
      }
    }
    if (!matched) {
      const normSubj = row.subject.trim().toLowerCase();
      if (normSubj && normSubj !== '(untitled)') {
        matched = groups.find(g =>
          g.workbookName === row.workbookName &&
          !g.key &&
          !row.key &&
          g.normSubj === normSubj &&
          g.account.toLowerCase() === row.account.toLowerCase()
        );
      }
    }

    if (matched) {
      const ticketMerge = mergeCommaDelimitedTickets(matched.latest.ticket, row.ticket);
      matched.ticketTokens = ticketMerge.allTokens;
      if (row.sheetMs >= matched.latest.sheetMs) {
        matched.latest = {
          ...row,
          ticket: ticketMerge.merged,
          action: row.action || matched.latest.action,
          account: row.account || matched.latest.account,
          notes: row.notes || matched.latest.notes
        };
      } else {
        matched.latest.ticket = ticketMerge.merged;
      }
      if (row.key && !matched.key) matched.key = row.key;
      matched.occurrences++;
    } else {
      const tTokens = splitNormalizedTicketTokens(row.ticket);
      groups.push({
        workbookName: row.workbookName,
        key: row.key || '',
        normSubj: row.subject.trim().toLowerCase(),
        account: row.account || '',
        ticketTokens: tTokens,
        occurrences: 1,
        latest: { ...row, ticket: tTokens.join(', ') || row.ticket }
      });
    }
  });

  return groups.map(g => ({
    ...g.latest,
    occurrences: g.occurrences
  }));
}

function resolveRollupRowParentCategory(row) {
  const clean = String(row.action || '').trim();
  if (clean) {
    const low = clean.toLowerCase();
    if (masterRollupState.localCategoryMap.has(low)) {
      return masterRollupState.localCategoryMap.get(low);
    }
  }
  return resolveActionParentCategory(clean);
}

async function scanAndAggregateRollupFolder() {
  const targetDir = rollupCustomDirHandle || supervisorDirHandle || parentDirHandle;
  if (!targetDir) {
    if (masterRollupSubtitle) {
      masterRollupSubtitle.textContent = 'No folder connected yet — click "Select Folder..." to choose a Supervisor or Team folder.';
    }
    if (masterRollupSummaryGrid) masterRollupSummaryGrid.innerHTML = '';
    if (masterRollupCategoryBar) masterRollupCategoryBar.innerHTML = '';
    if (masterRollupBody) {
      masterRollupBody.innerHTML = '<tr><td colspan="9" class="empty-table-cell">Click <strong>Select Folder...</strong> in the top-right to scan team tracker workbooks.</td></tr>';
    }
    if (btnExportMasterRollupXlsx) btnExportMasterRollupXlsx.disabled = true;
    return;
  }

  const mode = masterRollupSourceMode ? masterRollupSourceMode.value : 'auto';
  if (masterRollupSubtitle) {
    masterRollupSubtitle.textContent = `Scanning "${targetDir.name}"...`;
  }
  if (masterRollupBody) {
    masterRollupBody.innerHTML = '<tr><td colspan="9" class="empty-table-cell">⏳ Parsing workbooks in background...</td></tr>';
  }

  try {
    const candidates = [];
    for await (const entry of targetDir.values()) {
      if (entry.kind !== 'file') continue;
      const name = entry.name;
      const lower = name.toLowerCase();
      if (!lower.endsWith('.xlsx')) continue;
      if (name.startsWith('~$')) continue;
      if (/^master_tracker_rollup_/i.test(name)) continue;
      const isMirror = lower.endsWith('(view only).xlsx');
      const baseKey = lower.replace(/\s*\(view only\)\.xlsx$/i, '.xlsx');
      candidates.push({ name, handle: entry, isMirror, baseKey });
    }

    let selectedEntries = [];
    if (mode === 'mirrors') {
      selectedEntries = candidates.filter(c => c.isMirror);
    } else if (mode === 'working') {
      selectedEntries = candidates.filter(c => !c.isMirror);
    } else {
      // Auto: prefer (VIEW ONLY).xlsx mirror per base workbook when present, else working file
      const byBase = new Map();
      candidates.forEach(c => {
        const existing = byBase.get(c.baseKey);
        if (!existing || (c.isMirror && !existing.isMirror)) {
          byBase.set(c.baseKey, c);
        }
      });
      selectedEntries = Array.from(byBase.values());
    }

    selectedEntries.sort((a, b) => a.name.localeCompare(b.name));

    const allParsedRows = [];
    const localCatMap = new Map();
    let anyWorkerUsed = false;

    for (const item of selectedEntries) {
      try {
        const file = await item.handle.getFile();
        const buf = await file.arrayBuffer();
        const displayWbName = item.name.replace(/\s*\(VIEW ONLY\)\.xlsx$/i, '.xlsx');
        const res = await parseWorkbookBufferInBackground(buf, displayWbName);
        if (res.usedWorker) anyWorkerUsed = true;
        (res.categoryPairs || []).forEach(([actLow, cat]) => localCatMap.set(actLow, cat));
        (res.rows || []).forEach(r => allParsedRows.push(r));
      } catch (fileErr) {
        console.warn(`Skipped unreadable workbook "${item.name}" in Master Rollup:`, fileErr);
      }
    }

    masterRollupState = {
      folderName: targetDir.name,
      filesScanned: selectedEntries.map(e => e.name.replace(/\s*\(VIEW ONLY\)\.xlsx$/i, '.xlsx')),
      rawRows: allParsedRows.map(r => ({ ...r, parentCategory: '' })),
      dedupedRows: [],
      localCategoryMap: localCatMap,
      usedWorker: anyWorkerUsed
    };

    masterRollupState.rawRows.forEach(r => {
      r.parentCategory = resolveRollupRowParentCategory(r);
    });
    masterRollupState.dedupedRows = consolidateRollupRowsPerWorkbook(masterRollupState.rawRows).map(r => ({
      ...r,
      parentCategory: resolveRollupRowParentCategory(r)
    }));

    if (masterRollupWorkerBadge) {
      masterRollupWorkerBadge.textContent = anyWorkerUsed ? '⚡ Web Worker Parser' : 'Local Async Parser';
    }
    if (masterRollupSubtitle) {
      masterRollupSubtitle.textContent = `Folder: ${targetDir.name} • Scanned ${selectedEntries.length} workbook(s)`;
    }

    populateMasterRollupFilterDropdowns();
    renderMasterRollupUI();
  } catch (err) {
    console.error('Master Rollup scan error:', err);
    if (masterRollupBody) {
      masterRollupBody.innerHTML = `<tr><td colspan="9" class="empty-table-cell">Could not scan folder: ${escapeHtml(err.message)}</td></tr>`;
    }
  }
}

function populateMasterRollupFilterDropdowns() {
  const uniqueWbs = Array.from(new Set(masterRollupState.filesScanned)).sort((a, b) => a.localeCompare(b));
  if (masterRollupWorkbookFilter) {
    const prev = masterRollupWorkbookFilter.value;
    masterRollupWorkbookFilter.innerHTML = '<option value="__ALL__">All Workbooks</option>' +
      uniqueWbs.map(w => `<option value="${escapeHtml(w)}">${escapeHtml(w)}</option>`).join('');
    if (uniqueWbs.includes(prev)) masterRollupWorkbookFilter.value = prev;
  }

  const baseList = (masterRollupDedupCheck && masterRollupDedupCheck.checked)
    ? masterRollupState.dedupedRows
    : masterRollupState.rawRows;
  const uniqueCats = Array.from(new Set(baseList.map(r => r.parentCategory))).sort((a, b) => a.localeCompare(b));
  if (masterRollupCategoryFilter) {
    const prevCat = masterRollupCategoryFilter.value;
    masterRollupCategoryFilter.innerHTML = '<option value="__ALL__">All Parent Categories</option>' +
      uniqueCats.map(c => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');
    if (uniqueCats.includes(prevCat)) masterRollupCategoryFilter.value = prevCat;
  }
}

function getFilteredMasterRollupRows() {
  const sourceRows = (masterRollupDedupCheck && masterRollupDedupCheck.checked)
    ? masterRollupState.dedupedRows
    : masterRollupState.rawRows;

  const wbFilter = masterRollupWorkbookFilter ? masterRollupWorkbookFilter.value : '__ALL__';
  const catFilter = masterRollupCategoryFilter ? masterRollupCategoryFilter.value : '__ALL__';
  const stFilter = masterRollupStatusFilter ? masterRollupStatusFilter.value : '__ALL__';
  const q = (masterRollupSearchInput ? masterRollupSearchInput.value : '').trim().toLowerCase();

  return sourceRows
    .filter(r => {
      if (wbFilter !== '__ALL__' && r.workbookName !== wbFilter) return false;
      if (catFilter !== '__ALL__' && r.parentCategory !== catFilter) return false;
      if (stFilter === '__OPEN__' && isCompletedStatus(r.status)) return false;
      if (stFilter === '__COMPLETE__' && !isCompletedStatus(r.status)) return false;
      if (q) {
        const hay = `${r.workbookName} ${r.sheetName} ${r.subject} ${r.ticket} ${r.account} ${r.action} ${r.parentCategory} ${r.notes} ${r.status}`.toLowerCase();
        if (!hay.includes(q)) return false;
      }
      return true;
    })
    .sort((a, b) => b.sheetMs - a.sheetMs || a.workbookName.localeCompare(b.workbookName));
}

function renderMasterRollupUI() {
  const baseRows = (masterRollupDedupCheck && masterRollupDedupCheck.checked)
    ? masterRollupState.dedupedRows
    : masterRollupState.rawRows;
  const filtered = getFilteredMasterRollupRows();

  const totalCount = baseRows.length;
  const completedCount = baseRows.filter(r => isCompletedStatus(r.status)).length;
  const openCount = totalCount - completedCount;
  const escalatedCount = baseRows.filter(r => /escalat|overdue|waiting/i.test(r.status || '')).length;
  const pct = totalCount > 0 ? Math.round((completedCount / totalCount) * 100) : 0;

  if (masterRollupSummaryGrid) {
    masterRollupSummaryGrid.innerHTML = `
      <div class="rollup-kpi-card">
        <span class="rollup-kpi-label">Workbooks Scanned</span>
        <span class="rollup-kpi-value">${masterRollupState.filesScanned.length}</span>
      </div>
      <div class="rollup-kpi-card">
        <span class="rollup-kpi-label">${masterRollupDedupCheck?.checked ? 'Unique Requests' : 'Total Daily Rows'}</span>
        <span class="rollup-kpi-value">${totalCount}</span>
      </div>
      <div class="rollup-kpi-card">
        <span class="rollup-kpi-label">Open / Unfinished</span>
        <span class="rollup-kpi-value kpi-warn">${openCount}</span>
      </div>
      <div class="rollup-kpi-card">
        <span class="rollup-kpi-label">Escalated / Waiting</span>
        <span class="rollup-kpi-value">${escalatedCount}</span>
      </div>
      <div class="rollup-kpi-card">
        <span class="rollup-kpi-label">Completed (${pct}%)</span>
        <span class="rollup-kpi-value kpi-good">${completedCount}</span>
      </div>
    `;
  }

  // Category Breakdown Pills
  if (masterRollupCategoryBar) {
    const countsByCat = new Map();
    baseRows.forEach(r => {
      countsByCat.set(r.parentCategory, (countsByCat.get(r.parentCategory) || 0) + 1);
    });
    const activeCat = masterRollupCategoryFilter ? masterRollupCategoryFilter.value : '__ALL__';
    const sortedCats = Array.from(countsByCat.entries()).sort((a, b) => b[1] - a[1]);
    masterRollupCategoryBar.innerHTML = sortedCats.length === 0
      ? ''
      : [
          `<button type="button" class="rollup-cat-pill ${activeCat === '__ALL__' ? 'active' : ''}" data-cat="__ALL__">All Categories <strong>(${totalCount})</strong></button>`,
          ...sortedCats.map(([cat, cnt]) => `
            <button type="button" class="rollup-cat-pill ${activeCat === cat ? 'active' : ''}" data-cat="${escapeHtml(cat)}">
              ${escapeHtml(cat)} <strong>(${cnt})</strong>
            </button>
          `)
        ].join('');
  }

  if (masterRollupStatsText) {
    masterRollupStatsText.textContent = `${filtered.length} of ${totalCount} shown`;
  }
  if (btnExportMasterRollupXlsx) {
    btnExportMasterRollupXlsx.disabled = filtered.length === 0;
  }

  if (!masterRollupBody) return;
  if (filtered.length === 0) {
    masterRollupBody.innerHTML = '<tr><td colspan="9" class="empty-table-cell">No aggregated requests match the current filter.</td></tr>';
    return;
  }

  const maxRender = 500;
  const slice = filtered.slice(0, maxRender);
  masterRollupBody.innerHTML = slice.map(r => {
    const statusBadgeClass = isCompletedStatus(r.status)
      ? 'badge-complete'
      : (/escalat|overdue/i.test(r.status) ? 'badge-escalated' : (/progress|waiting/i.test(r.status) ? 'badge-progress' : 'badge-pending'));
    const occNote = r.occurrences && r.occurrences > 1 ? ` <span class="control-label" title="Rolled across ${r.occurrences} daily tabs">(${r.occurrences}d)</span>` : '';
    return `
      <tr>
        <td><strong>${escapeHtml(r.workbookName.replace(/\.xlsx$/i, ''))}</strong></td>
        <td><span class="key-chip" style="cursor:default;">${escapeHtml(r.sheetName)}</span>${occNote}</td>
        <td>${escapeHtml(r.subject)}</td>
        <td>${r.ticket ? `<code>${escapeHtml(r.ticket)}</code>` : '—'}</td>
        <td>${escapeHtml(r.account || '—')}</td>
        <td>${escapeHtml(r.action || '—')}</td>
        <td><span class="badge badge-progress">${escapeHtml(r.parentCategory)}</span></td>
        <td class="notes-cell" title="${escapeHtml(r.notes)}">${escapeHtml(r.notes ? r.notes.slice(0, 140) : '—')}</td>
        <td><span class="badge ${statusBadgeClass}">${escapeHtml(r.status)}</span></td>
      </tr>
    `;
  }).join('');
}

async function exportMasterRollupToXlsx() {
  const rows = getFilteredMasterRollupRows();
  if (rows.length === 0) return;

  try {
    const outWb = new ExcelJS.Workbook();
    const ws = outWb.addWorksheet('Master Rollup');
    const headers = [
      'Workbook',
      'Latest Daily Tab',
      'Days Tracked',
      'Subject / Primary',
      'Ticket Number',
      'Account / Subtitle',
      'Requested Action / Subtype',
      'Parent Category',
      'Action Notes',
      'Status',
      'Tracking ID'
    ];
    const widths = [26, 16, 14, 34, 20, 24, 24, 24, 42, 16, 38];
    const r1 = ws.getRow(1);
    r1.height = 22;
    headers.forEach((h, i) => {
      ws.getColumn(i + 1).width = widths[i];
      const cell = r1.getCell(i + 1);
      cell.value = h;
      cell.font = { bold: true, color: { argb: 'FFF8FAFC' }, size: 11 };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
    });
    r1.commit();

    rows.forEach((r, idx) => {
      const row = ws.getRow(idx + 2);
      row.getCell(1).value = r.workbookName;
      row.getCell(2).value = r.sheetName;
      row.getCell(3).value = r.occurrences || 1;
      row.getCell(4).value = r.subject;
      row.getCell(5).value = r.ticket;
      row.getCell(6).value = r.account;
      row.getCell(7).value = r.action;
      row.getCell(8).value = r.parentCategory;
      row.getCell(9).value = r.notes;
      row.getCell(10).value = r.status;
      row.getCell(11).value = r.key;
      row.commit();
    });

    // Sheet 2: Category & Status Summary
    const sumWs = outWb.addWorksheet('Category Summary');
    sumWs.getColumn(1).width = 28;
    sumWs.getColumn(2).width = 16;
    sumWs.getColumn(3).width = 16;
    sumWs.getColumn(4).width = 16;
    const sr1 = sumWs.getRow(1);
    ['Parent Category', 'Total Requests', 'Open', 'Completed'].forEach((h, i) => {
      const c = sr1.getCell(i + 1);
      c.value = h;
      c.font = { bold: true, color: { argb: 'FFF8FAFC' }, size: 11 };
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
    });
    sr1.commit();

    const byCat = new Map();
    rows.forEach(r => {
      const entry = byCat.get(r.parentCategory) || { total: 0, open: 0, done: 0 };
      entry.total++;
      if (isCompletedStatus(r.status)) entry.done++;
      else entry.open++;
      byCat.set(r.parentCategory, entry);
    });
    Array.from(byCat.entries())
      .sort((a, b) => b[1].total - a[1].total)
      .forEach(([cat, stats], idx) => {
        const sr = sumWs.getRow(idx + 2);
        sr.getCell(1).value = cat;
        sr.getCell(2).value = stats.total;
        sr.getCell(3).value = stats.open;
        sr.getCell(4).value = stats.done;
        sr.commit();
      });

    const buf = await outWb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Master_Tracker_Rollup_${getTodayFormatted()}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
    showToast(`Exported ${rows.length} aggregated rows to Master_Tracker_Rollup_${getTodayFormatted()}.xlsx`, 'success');
  } catch (err) {
    console.error('Export Master Rollup error:', err);
    showModalAlert({ title: 'Export Error', message: err.message, variant: 'danger' });
  }
}

async function openMasterRollupModal() {
  closeToolsMenu();
  if (!masterRollupModal) return;
  masterRollupModal.classList.add('open');
  await scanAndAggregateRollupFolder();
}

if (menuOpenMasterRollup) {
  menuOpenMasterRollup.addEventListener('click', openMasterRollupModal);
}
if (btnPickMasterRollupFolder) {
  btnPickMasterRollupFolder.addEventListener('click', async () => {
    try {
      if (typeof window.showDirectoryPicker !== 'function') return;
      rollupCustomDirHandle = await window.showDirectoryPicker({ mode: 'read' });
      await scanAndAggregateRollupFolder();
    } catch (err) {
      if (err.name !== 'AbortError') {
        showModalAlert({ title: 'Folder Selection Error', message: err.message });
      }
    }
  });
}
if (btnRefreshMasterRollup) {
  btnRefreshMasterRollup.addEventListener('click', scanAndAggregateRollupFolder);
}
if (btnExportMasterRollupXlsx) {
  btnExportMasterRollupXlsx.addEventListener('click', exportMasterRollupToXlsx);
}
if (btnCloseMasterRollup) {
  btnCloseMasterRollup.addEventListener('click', () => masterRollupModal.classList.remove('open'));
}
if (masterRollupModal) {
  masterRollupModal.addEventListener('click', (e) => {
    if (e.target === masterRollupModal) masterRollupModal.classList.remove('open');
  });
}
if (masterRollupSourceMode) {
  masterRollupSourceMode.addEventListener('change', scanAndAggregateRollupFolder);
}
if (masterRollupDedupCheck) {
  masterRollupDedupCheck.addEventListener('change', () => {
    populateMasterRollupFilterDropdowns();
    renderMasterRollupUI();
  });
}
[masterRollupWorkbookFilter, masterRollupCategoryFilter, masterRollupStatusFilter].forEach(el => {
  if (el) el.addEventListener('change', renderMasterRollupUI);
});
if (masterRollupSearchInput) {
  masterRollupSearchInput.addEventListener('input', renderMasterRollupUI);
}
if (masterRollupCategoryBar) {
  masterRollupCategoryBar.addEventListener('click', (e) => {
    const pill = e.target.closest('.rollup-cat-pill');
    if (!pill || !masterRollupCategoryFilter) return;
    masterRollupCategoryFilter.value = pill.dataset.cat || '__ALL__';
    renderMasterRollupUI();
  });
}

// ============================================================================
// 12. PORTABLE WORKSPACE CONFIG (.JSON) EXPORT & IMPORT
// ============================================================================
function exportWorkspaceConfigJson() {
  closeToolsMenu();
  const payload = {
    schema: 'TrackerUI_WorkspaceConfig_v1',
    appVersion: APP_VERSION,
    exportedAt: new Date().toISOString(),
    preferences: {
      themePreference,
      fontScalePercent,
      autoReconnectEnabled
    },
    trackerProfiles,
    schemaOverrides: schemaOverridesStore,
    formLayouts: formLayoutsStore,
    actionCategoryOverrides: actionCategoryOverridesStore,
    referenceDefaults: {
      statuses: validationStatuses.length > 0 ? validationStatuses : DEFAULT_STATUSES,
      actions: validationActions.length > 0 ? validationActions : DEFAULT_ACTIONS
    }
  };

  const jsonStr = JSON.stringify(payload, null, 2);
  const blob = new Blob([jsonStr], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `trackerui-config-${getTodayFormatted()}.json`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
  showToast('Exported portable workspace configuration (.json).', 'success');
}

async function importWorkspaceConfigJsonFile(file) {
  if (!file) return;
  try {
    const rawText = await file.text();
    const parsed = JSON.parse(rawText);
    if (!parsed || typeof parsed !== 'object') {
      throw new Error('Invalid JSON configuration object.');
    }

    if (parsed.preferences && typeof parsed.preferences === 'object') {
      if (['auto', 'light', 'dark'].includes(parsed.preferences.themePreference)) {
        setThemePreference(parsed.preferences.themePreference, { persist: true, announce: false });
      }
      if (Number.isFinite(Number(parsed.preferences.fontScalePercent))) {
        applyFontScaleToDOM(Number(parsed.preferences.fontScalePercent), { persist: true });
      }
      if (typeof parsed.preferences.autoReconnectEnabled === 'boolean') {
        setAutoReconnectPreference(parsed.preferences.autoReconnectEnabled);
      }
    }

    if (Array.isArray(parsed.trackerProfiles) && parsed.trackerProfiles.length > 0) {
      const mergedProfiles = [...trackerProfiles];
      parsed.trackerProfiles.forEach(inc => {
        if (!inc || typeof inc.id !== 'string' || !inc.id.trim()) return;
        const cleanId = inc.id.trim();
        const cleanName = String(inc.name || 'Tracker').trim() || 'Tracker';
        const existing = mergedProfiles.find(p => p.id === cleanId);
        if (existing) {
          existing.name = cleanName;
          existing.mirrorEnabled = inc.mirrorEnabled !== false;
        } else {
          mergedProfiles.push({
            id: cleanId,
            name: cleanName,
            mirrorEnabled: inc.mirrorEnabled !== false
          });
        }
      });
      trackerProfiles = mergedProfiles;
      persistTrackerProfiles();
    }

    if (parsed.schemaOverrides && typeof parsed.schemaOverrides === 'object' && !Array.isArray(parsed.schemaOverrides)) {
      schemaOverridesStore = { ...schemaOverridesStore, ...parsed.schemaOverrides };
      localStorage.setItem(SCHEMA_OVERRIDES_STORAGE_KEY, JSON.stringify(schemaOverridesStore));
    }

    if (parsed.formLayouts && typeof parsed.formLayouts === 'object' && !Array.isArray(parsed.formLayouts)) {
      formLayoutsStore = { ...formLayoutsStore, ...parsed.formLayouts };
      saveFormLayoutsStore();
    }

    if (parsed.actionCategoryOverrides && typeof parsed.actionCategoryOverrides === 'object' && !Array.isArray(parsed.actionCategoryOverrides)) {
      actionCategoryOverridesStore = { ...actionCategoryOverridesStore, ...parsed.actionCategoryOverrides };
      localStorage.setItem(ACTION_CATEGORY_OVERRIDES_STORAGE_KEY, JSON.stringify(actionCategoryOverridesStore));
    }

    syncMirrorCopyCheckboxUI();
    if (workbook && workingFileHandle) {
      await loadWorkbookFromDisk(sheetSelect ? sheetSelect.value : null, { skipStagedRecoveryCheck: true });
    } else if (typeof syncWorkspaceTabsUI === 'function') {
      syncWorkspaceTabsUI();
    }
    showToast(`Imported workspace configuration from "${file.name}".`, 'success');
  } catch (err) {
    console.error('Config JSON import error:', err);
    await showModalAlert({
      title: 'Could Not Import Config (.json)',
      message: err.message,
      variant: 'danger'
    });
  }
}

if (btnExportConfigJson) {
  btnExportConfigJson.addEventListener('click', exportWorkspaceConfigJson);
}
if (btnImportConfigJson && configJsonFileInput) {
  btnImportConfigJson.addEventListener('click', () => {
    closeToolsMenu();
    configJsonFileInput.value = '';
    configJsonFileInput.click();
  });
  configJsonFileInput.addEventListener('change', (e) => {
    const file = e.target.files && e.target.files[0];
    if (file) importWorkspaceConfigJsonFile(file);
  });
}


