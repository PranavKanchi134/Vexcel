// Popup toggle logic

const enableToggle = document.getElementById('enableToggle');
const statusText = document.getElementById('statusText');
const optCmdN = document.getElementById('optCmdN');
const optCmdT = document.getElementById('optCmdT');
const debugModeToggle = document.getElementById('debugModeToggle');
const perfSummary = document.getElementById('perfSummary');

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
