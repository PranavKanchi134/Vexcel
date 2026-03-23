// Vexcel Content Script - Core event interception pipeline
// Injected into Google Sheets pages (all frames) at document_start

(() => {
  let enabled = true;
  let debugMode = false;
  let optInShortcuts = {};
  let cachedIframes = [];
  let iframeObserver = null;

  const isTopFrame = (window === window.top);

  // Initialize the extension
  function init() {
    // Register all shortcuts
    VexcelShortcutRegistry.initialize();

    // Only initialize overlays in the top frame
    if (isTopFrame) {
      VexcelAcceleratorOverlay.init();
      VexcelTraceArrows.init();
      observeIframes();
    }

    // Attach capture-phase keydown listener BEFORE Sheets registers its own
    document.addEventListener('keydown', handleKeyDown, true);
    document.addEventListener('keyup', handleKeyUp, true);

    // ── Cross-frame accelerator bridging ──────────────────────
    // The Sheets grid lives in an iframe, so Option tap and accelerator
    // key events must be forwarded to the top frame where the overlay lives.
    if (isTopFrame) {
      // Top frame: listen for forwarded events from iframes
      window.addEventListener('message', (msg) => {
        if (!msg.data || msg.data.source !== 'vexcel') return;
        if (msg.data.type === 'OPTION_TAP') {
          VexcelAcceleratorOverlay.activate();
          broadcastAccelState(true);
        } else if (msg.data.type === 'ACCEL_KEYDOWN') {
          // Create a synthetic event for the overlay to process
          VexcelAcceleratorOverlay.handleKeyDown(
            new KeyboardEvent('keydown', {
              key: msg.data.key, code: msg.data.code,
              altKey: msg.data.altKey, metaKey: msg.data.metaKey,
              ctrlKey: msg.data.ctrlKey, shiftKey: msg.data.shiftKey,
              bubbles: true, cancelable: true
            })
          );
          // Broadcast updated state back to iframes
          broadcastAccelState(VexcelAcceleratorOverlay.isActive());
        }
      });

      // Listen for click-away to dismiss accelerator
      document.addEventListener('mousedown', (e) => {
        if (VexcelAcceleratorOverlay.isActive()) {
          const host = document.getElementById('vexcel-accelerator-host');
          if (host && host.contains(e.target)) return;
          VexcelAcceleratorOverlay.deactivate();
          broadcastAccelState(false);
        }
      }, false);
    } else {
      // Iframe: listen for accelerator state broadcasts from top frame
      window.addEventListener('message', (msg) => {
        if (!msg.data || msg.data.source !== 'vexcel') return;
        if (msg.data.type === 'ACCEL_STATE') {
          accelActiveInTop = msg.data.active;
        }
      });
    }

    // Get initial state (retry on failure since service worker may not be ready)
    function fetchState(retries = 3) {
      VexcelMessaging.getState().then(settings => {
        if (settings) {
          enabled = settings.enabled;
          debugMode = !!settings.debugMode;
          optInShortcuts = settings.optInShortcuts || {};
        } else {
          enabled = true; // Default to enabled once we get a response
        }
      }).catch(() => {
        if (retries > 0) {
          setTimeout(() => fetchState(retries - 1), 500);
        } else {
          enabled = true; // Default to enabled after all retries exhausted
        }
      });
    }
    fetchState();

    // Listen for state updates from popup/background
    VexcelMessaging.onStateUpdate(settings => {
      enabled = settings.enabled;
      debugMode = !!settings.debugMode;
      optInShortcuts = settings.optInShortcuts || {};
    });

    console.log('[Vexcel] Content script initialized');
  }

  // ── Cross-frame accelerator state ───────────────────────────
  // Tracks whether the accelerator overlay is active in the top frame.
  // Used by iframes to know when to forward key events.
  let accelActiveInTop = false;

  function broadcastAccelState(active) {
    accelActiveInTop = active;
    try {
      const iframes = cachedIframes.length
        ? cachedIframes
        : window.top.document.querySelectorAll('iframe');
      for (const iframe of iframes) {
        try {
          iframe.contentWindow.postMessage(
            { source: 'vexcel', type: 'ACCEL_STATE', active },
            '*'
          );
        } catch (e) { /* cross-origin iframe */ }
      }
    } catch (e) {}
  }

  function observeIframes() {
    try {
      const doc = window.top.document;
      const refresh = () => {
        cachedIframes = Array.from(doc.querySelectorAll('iframe')).filter(iframe => iframe.isConnected);
      };
      refresh();
      if (iframeObserver) iframeObserver.disconnect();
      iframeObserver = new MutationObserver(refresh);
      iframeObserver.observe(doc.documentElement || doc.body, { childList: true, subtree: true });
    } catch (err) {}
  }

  // ── Option tap tracking for iframes ───────────────────────
  // Iframes can't call VexcelAcceleratorOverlay (it lives in the top frame's
  // JS context), so we track the Option key locally and postMessage to top.
  let iframeOptionDown = null;
  let iframeOtherKey = false;

  /**
   * Main keydown handler - capture phase.
   */
  function handleKeyDown(e) {
    if (!enabled) return;

    // ── Accelerator handling ──────────────────────────────────
    if (isTopFrame) {
      // Top frame: handle directly
      if (VexcelAcceleratorOverlay.handleKeyDown(e)) {
        broadcastAccelState(VexcelAcceleratorOverlay.isActive());
        return;
      }
    } else {
      // Iframe: if accelerator is active in top frame, forward key events there
      if (accelActiveInTop) {
        e.preventDefault();
        e.stopPropagation();
        e.stopImmediatePropagation();
        try {
          window.top.postMessage({
            source: 'vexcel', type: 'ACCEL_KEYDOWN',
            key: e.key, code: e.code,
            altKey: e.altKey, metaKey: e.metaKey,
            ctrlKey: e.ctrlKey, shiftKey: e.shiftKey
          }, '*');
        } catch (err) {}
        return;
      }

      // Iframe: track Option key for tap detection
      if (e.key === 'Alt' && !e.metaKey && !e.ctrlKey && !e.shiftKey) {
        iframeOptionDown = e.timeStamp;
        iframeOtherKey = false;
        // Block the Alt keydown so Google Sheets doesn't activate its menu bar
        e.preventDefault();
        e.stopPropagation();
        e.stopImmediatePropagation();
        return;
      }
      if (iframeOptionDown) iframeOtherKey = true;
    }

    // ── Normal shortcut handling ──────────────────────────────

    // Skip modifier-only events
    if (VexcelKeyUtils.isModifierOnly(e)) return;

    // Normalize the event to a combo string
    const combo = VexcelKeyUtils.normalizeEvent(e);
    if (!combo) return;

    // Look up the shortcut FIRST, then check context
    const shortcut = VexcelShortcutRegistry.lookup(combo);
    if (!shortcut) return;

    // Check if we should intercept in the current context
    if (!VexcelDom.shouldIntercept()) {
      console.warn(`[Vexcel] Shortcut "${combo}" blocked by shouldIntercept — dialog/menu/input open`);
      return;
    }

    // Check opt-in shortcuts
    if (shortcut.optIn && !optInShortcuts[combo]) return;

    // Check pass-through shortcuts (just let them propagate)
    if (shortcut.passThrough) return;

    // Block the native browser action (e.g. Cmd+R = reload, Cmd+D = bookmark)
    e.preventDefault();
    e.stopPropagation();
    e.stopImmediatePropagation();

    // Execute the action.
    // NOTE: Keyboard events only fire in the FOCUSED frame (they don't cross
    // frame boundaries), so there's no risk of double-firing between the
    // top frame and an iframe. We must execute from whichever frame has focus,
    // since that's the only one that receives the event.
    const commandId = shortcut.commandId || combo;
    VexcelPerf.start(commandId, { combo, label: shortcut.label });
    VexcelPerf.mark(commandId, 'captured');
    VexcelShortcutActions.execute(shortcut, e).then(outcome => {
      if (outcome.strategy) VexcelPerf.mark(commandId, 'strategy', { strategy: outcome.strategy });
      const perfRecord = VexcelPerf.finish(commandId, outcome) || {};
      if (outcome.ok) {
        if (debugMode) {
          console.log(`[Vexcel] Executed: ${shortcut.label} (${combo}) via ${outcome.strategy || 'action'} in ${perfRecord.durationMs || 0}ms`);
          if (perfRecord.slow) {
            console.warn(`[Vexcel] Slow command: ${commandId} took ${perfRecord.durationMs}ms`);
          }
        } else {
          console.log(`[Vexcel] Executed: ${shortcut.label} (${combo})`);
        }
        showToast(outcome.message || shortcut.label, 'success');
      } else {
        console.warn(`[Vexcel] Failed: ${shortcut.label} (${combo})`, outcome.reason || '');
        showToast(outcome.message || `${shortcut.label} failed`, 'error');
      }
    });
  }

  /**
   * Keyup handler for accelerator Option tap detection.
   * Works in both top frame (direct) and iframes (via postMessage).
   */
  function handleKeyUp(e) {
    if (!enabled) return;

    if (isTopFrame) {
      const handled = VexcelAcceleratorOverlay.handleKeyUp(e);
      if (handled) {
        e.preventDefault();
        e.stopPropagation();
        e.stopImmediatePropagation();
        broadcastAccelState(VexcelAcceleratorOverlay.isActive());
      }
    } else {
      // Iframe: detect Option tap and forward to top frame
      if (e.key === 'Alt' && iframeOptionDown && !iframeOtherKey) {
        const dur = e.timeStamp - iframeOptionDown;
        iframeOptionDown = null;
        if (dur < 500) {
          e.preventDefault();
          e.stopPropagation();
          e.stopImmediatePropagation();
          try {
            window.top.postMessage({ source: 'vexcel', type: 'OPTION_TAP' }, '*');
            accelActiveInTop = true; // Optimistically set — top frame will correct if needed
          } catch (err) {}
        }
      }
      if (e.key === 'Alt') iframeOptionDown = null;
    }
  }

  // ── Toast notification ───────────────────────────────────
  // Always render in the top frame so it's visible regardless of which frame
  // triggered the shortcut. Uses try/catch since iframe→top access can fail
  // if frames are cross-origin (shouldn't happen with Google Sheets, but safe).
  let toastEl = null;
  let toastTimer = null;

  function showToast(label, tone = 'success') {
    try {
      const topDocument = window.top.document;
      if (!topDocument.body) return;

      if (!toastEl) {
        // Check if another frame already created the toast element
        toastEl = topDocument.getElementById('vexcel-toast');
      }
      if (!toastEl) {
        toastEl = topDocument.createElement('div');
        toastEl.id = 'vexcel-toast';
        toastEl.style.cssText =
          'position:fixed;bottom:24px;left:50%;transform:translateX(-50%);' +
          'background:#1c1c2e;color:#ccc;border:1px solid #383860;' +
          'border-radius:6px;padding:6px 14px;font:500 12px -apple-system,sans-serif;' +
          'z-index:2147483645;pointer-events:none;opacity:0;' +
          'transition:opacity .15s;box-shadow:0 4px 12px rgba(0,0,0,.5);';
        topDocument.body.appendChild(toastEl);
      }
      if (tone === 'error') {
        toastEl.style.background = '#3b1f27';
        toastEl.style.borderColor = '#8d3b50';
        toastEl.style.color = '#ffd7df';
      } else {
        toastEl.style.background = '#1c1c2e';
        toastEl.style.borderColor = '#383860';
        toastEl.style.color = '#ccc';
      }
      toastEl.textContent = label;
      toastEl.style.opacity = '1';
      clearTimeout(toastTimer);
      toastTimer = setTimeout(() => { toastEl.style.opacity = '0'; }, 1200);
    } catch (e) {
      // Cross-origin — can't access top frame, skip toast silently
    }
  }

  // Initialize when the script loads
  init();
})();
