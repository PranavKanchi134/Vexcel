// Popup toggle logic

const enableToggle = document.getElementById('enableToggle');
const statusText = document.getElementById('statusText');
const optCmdN = document.getElementById('optCmdN');
const optCmdT = document.getElementById('optCmdT');
const debugModeToggle = document.getElementById('debugModeToggle');
const perfSummary = document.getElementById('perfSummary');
const syncButton = document.getElementById('syncButton');
const syncStatus = document.getElementById('syncStatus');

// Load current state
chrome.runtime.sendMessage({ type: 'GET_STATE' }, (settings) => {
  if (chrome.runtime.lastError) {
    statusText.textContent = 'Error loading state';
    return;
  }
  applyState(settings);
  loadPerfSummary(settings);
});

enableToggle.addEventListener('change', () => {
  chrome.runtime.sendMessage({ type: 'TOGGLE_ENABLED' }, (settings) => {
    if (chrome.runtime.lastError) return;
    applyState(settings);
  });
});

optCmdN.addEventListener('change', () => updateOptIn());
optCmdT.addEventListener('change', () => updateOptIn());
debugModeToggle.addEventListener('change', () => {
  chrome.runtime.sendMessage({
    type: 'UPDATE_SETTINGS',
    settings: { debugMode: debugModeToggle.checked }
  }, (settings) => {
    if (chrome.runtime.lastError) return;
    applyState(settings);
    loadPerfSummary(settings);
  });
});
syncButton.addEventListener('click', () => syncActiveSheetFromClipboard());

function updateOptIn() {
  const optInShortcuts = {
    'cmd+n': optCmdN.checked,
    'cmd+t': optCmdT.checked
  };
  chrome.runtime.sendMessage({
    type: 'UPDATE_SETTINGS',
    settings: { optInShortcuts }
  });
}

function applyState(settings) {
  if (!settings) return;
  enableToggle.checked = settings.enabled;
  statusText.textContent = settings.enabled ? 'Shortcuts active on Google Sheets' : 'Shortcuts disabled';
  statusText.className = 'status ' + (settings.enabled ? 'active' : 'inactive');

  if (settings.optInShortcuts) {
    optCmdN.checked = settings.optInShortcuts['cmd+n'] || false;
    optCmdT.checked = settings.optInShortcuts['cmd+t'] || false;
  }
  debugModeToggle.checked = !!settings.debugMode;
}

function loadPerfSummary(settings) {
  if (!settings || !settings.debugMode) {
    perfSummary.textContent = 'Enable Debug Mode to see rolling command timings.';
    return;
  }

  chrome.runtime.sendMessage({ type: 'GET_DEBUG_STATS' }, (summary) => {
    if (chrome.runtime.lastError || !summary || !summary.commands) {
      perfSummary.textContent = 'No command timings yet.';
      return;
    }

    const important = ['fillDown', 'fillRight', 'pasteValues', 'findReplace', 'formatCycle', 'fontColorCycle', 'autoColorSelection'];
    const parts = [];
    for (const key of important) {
      const stat = summary.commands[key];
      if (!stat) continue;
      parts.push(`${key}: ${stat.avgMs}ms avg`);
    }
    perfSummary.textContent = parts.length ? parts.join(' | ') : 'No command timings yet.';
  });
}

async function syncActiveSheetFromClipboard() {
  setSyncStatus('Reading clipboard path...', '');
  syncButton.disabled = true;

  try {
    const filePath = (await navigator.clipboard.readText()).trim();
    if (!filePath) {
      throw new Error('Clipboard is empty. Copy an .xlsx filepath first.');
    }
    if (!/\.(xlsx|xlsm)$/i.test(filePath)) {
      throw new Error('Clipboard does not look like an .xlsx or .xlsm filepath.');
    }

    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab || !tab.id || !tab.url || !tab.url.includes('docs.google.com/spreadsheets/')) {
      throw new Error('Open the target Google Sheet first, then try Sync again.');
    }

    const context = await sendMessageToTab(tab.id, { type: 'GET_SHEET_CONTEXT' });
    if (!context || !context.ok) {
      throw new Error((context && context.reason) || 'Could not detect the active Google Sheets tab.');
    }

    setSyncStatus(`Syncing ${basename(filePath)} into ${context.sheetTitle}...`, '');

    const response = await fetch('http://127.0.0.1:8765/sync', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filePath,
        spreadsheetId: context.spreadsheetId,
        sheetTitle: context.sheetTitle
      })
    });

    const payload = await response.json().catch(() => ({}));
    if (!response.ok || !payload.ok) {
      throw new Error(payload.error || payload.reason || `Sync failed (${response.status})`);
    }

    setSyncStatus(
      `Synced ${basename(filePath)} to ${payload.sheetTitle} (${payload.rowCount} rows x ${payload.columnCount} cols).`,
      'success'
    );
  } catch (error) {
    if (error && /Failed to fetch/i.test(error.message || '')) {
      setSyncStatus('Sync server is not running. Start ./scripts/gsheets-sync-server and try again.', 'error');
    } else {
      setSyncStatus(error.message || 'Sync failed.', 'error');
    }
  } finally {
    syncButton.disabled = false;
  }
}

function sendMessageToTab(tabId, message) {
  return new Promise((resolve, reject) => {
    chrome.tabs.sendMessage(tabId, message, (response) => {
      if (chrome.runtime.lastError) {
        reject(chrome.runtime.lastError);
      } else {
        resolve(response);
      }
    });
  });
}

function basename(filePath) {
  return filePath.split(/[\\/]/).pop() || filePath;
}

function setSyncStatus(text, tone) {
  syncStatus.textContent = text;
  syncStatus.className = tone ? `sync-status ${tone}` : 'sync-status';
}
