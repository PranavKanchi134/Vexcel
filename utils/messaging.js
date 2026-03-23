// Chrome runtime message helpers

const VexcelMessaging = (() => {
  /**
   * Send a message to the background service worker and return a promise.
   */
  function sendToBackground(message) {
    return new Promise((resolve, reject) => {
      chrome.runtime.sendMessage(message, (response) => {
        if (chrome.runtime.lastError) {
          reject(chrome.runtime.lastError);
        } else {
          resolve(response);
        }
      });
    });
  }

  /**
   * Get the current extension settings.
   */
  function getState() {
    return sendToBackground({ type: 'GET_STATE' });
  }

  function getDebugStats() {
    return sendToBackground({ type: 'GET_DEBUG_STATS' });
  }

  /**
   * Listen for state updates from the background.
   */
  function onStateUpdate(callback) {
    chrome.runtime.onMessage.addListener((message) => {
      if (message.type === 'STATE_UPDATE') {
        callback(message.settings);
      }
    });
  }

  return { sendToBackground, getState, getDebugStats, onStateUpdate };
})();
