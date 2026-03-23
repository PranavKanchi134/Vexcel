// Vexcel Service Worker - manages extension state and icon toggling

const SETTINGS_KEY = 'vexcel_settings';
let debugStats = { commands: {}, updatedAt: 0 };

const DEFAULT_SETTINGS = {
  enabled: true,
  debugMode: false,
  optInShortcuts: {
    'cmd+n': false, // New sheet - off by default (overrides Chrome new window)
    'cmd+t': false  // Table/alternating colors - off by default (overrides Chrome new tab)
  }
};

// Initialize state on install
chrome.runtime.onInstalled.addListener(async () => {
  const data = await chrome.storage.local.get(SETTINGS_KEY);
  if (!data[SETTINGS_KEY]) {
    await chrome.storage.local.set({ [SETTINGS_KEY]: DEFAULT_SETTINGS });
  }
  await updateIcon((data[SETTINGS_KEY] || DEFAULT_SETTINGS).enabled);
});

// Handle messages from popup and content scripts
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === 'GET_STATE') {
    chrome.storage.local.get(SETTINGS_KEY).then(data => {
      sendResponse(data[SETTINGS_KEY] || DEFAULT_SETTINGS);
    });
    return true; // async response
  }

  if (message.type === 'TOGGLE_ENABLED') {
    handleToggle().then(newState => sendResponse(newState));
    return true;
  }

  if (message.type === 'UPDATE_SETTINGS') {
    handleUpdateSettings(message.settings).then(settings => sendResponse(settings));
    return true;
  }

  if (message.type === 'PERF_EVENT') {
    debugStats = message.summary || debugStats;
    sendResponse({ ok: true });
    return false;
  }

  if (message.type === 'GET_DEBUG_STATS') {
    sendResponse(debugStats);
    return false;
  }
});

async function handleToggle() {
  const data = await chrome.storage.local.get(SETTINGS_KEY);
  const settings = data[SETTINGS_KEY] || DEFAULT_SETTINGS;
  settings.enabled = !settings.enabled;
  await chrome.storage.local.set({ [SETTINGS_KEY]: settings });
  await updateIcon(settings.enabled);
  await broadcastState(settings);
  return settings;
}

async function handleUpdateSettings(newSettings) {
  const data = await chrome.storage.local.get(SETTINGS_KEY);
  const existing = data[SETTINGS_KEY] || DEFAULT_SETTINGS;
  // Deep merge optInShortcuts to avoid losing keys
  const settings = {
    ...existing,
    ...newSettings,
    optInShortcuts: {
      ...(existing.optInShortcuts || {}),
      ...(newSettings.optInShortcuts || {})
    }
  };
  await chrome.storage.local.set({ [SETTINGS_KEY]: settings });
  await broadcastState(settings);
  return settings;
}

async function updateIcon(enabled) {
  const suffix = enabled ? '' : '-inactive';
  await chrome.action.setIcon({
    path: {
      16: `icons/icon-16${suffix}.png`,
      48: `icons/icon-48${suffix}.png`,
      128: `icons/icon-128${suffix}.png`
    }
  });
  await chrome.action.setBadgeText({ text: enabled ? 'ON' : 'OFF' });
  await chrome.action.setBadgeBackgroundColor({
    color: enabled ? '#1a73e8' : '#999999'
  });
}

async function broadcastState(settings) {
  const tabs = await chrome.tabs.query({ url: 'https://docs.google.com/spreadsheets/*' });
  for (const tab of tabs) {
    try {
      await chrome.tabs.sendMessage(tab.id, {
        type: 'STATE_UPDATE',
        settings
      });
    } catch (e) {
      // Tab may not have content script loaded yet
    }
  }
}

// Update icon when a Google Sheets tab becomes active
chrome.tabs.onActivated.addListener(async () => {
  const data = await chrome.storage.local.get(SETTINGS_KEY);
  const settings = data[SETTINGS_KEY] || DEFAULT_SETTINGS;
  await updateIcon(settings.enabled);
});
