// Cell operations shortcuts: Fill Down, Fill Right, Paste Special, etc.

const VexcelCellOperations = (() => {

  /**
   * Dispatch a Ctrl+key combo to the active element.
   * Google Sheets on Mac uses Ctrl (not Cmd) for many native shortcuts.
   * Cmd+D/R are intercepted by Chrome, but Ctrl+D/R reach Sheets.
   */
  function dispatchCtrlKey(keyChar, keyCode) {
    // Try dispatching to all possible targets
    const targets = [];
    // 1. Current document's active element
    if (document.activeElement && document.activeElement !== document.body) {
      targets.push(document.activeElement);
    }
    // 2. Top frame's active element
    try {
      const topDoc = window.top.document;
      if (topDoc.activeElement && topDoc.activeElement !== topDoc.body) {
        targets.push(topDoc.activeElement);
      }
    } catch (e) {}
    // 3. Cell editor elements
    const cellInputs = document.querySelectorAll('.cell-input, [contenteditable="true"]');
    for (const ci of cellInputs) {
      const r = ci.getBoundingClientRect();
      if (r.width > 0 && r.height > 0) targets.push(ci);
    }
    // 4. Fallback to document body
    if (targets.length === 0) targets.push(document.body);

    const opts = {
      key: keyChar, code: `Key${keyChar.toUpperCase()}`, keyCode: keyCode,
      ctrlKey: true, metaKey: false,
      bubbles: true, cancelable: true
    };
    for (const target of targets) {
      target.dispatchEvent(new KeyboardEvent('keydown', opts));
      target.dispatchEvent(new KeyboardEvent('keyup', { ...opts, cancelable: false }));
    }

    return targets.length > 0;
  }

  async function executeFillAction(commandId, label, keyChar, keyCode) {
    const primary = commandId === 'fillDown'
      ? VexcelFastActions.fillDown
      : VexcelFastActions.fillRight;
    const command = await primary();
    if (command.ok) return command;

    if (commandId === 'fillRight') {
      return {
        ok: false,
        message: `${label} failed`,
        reason: command.reason || 'no verified fill-right path succeeded'
      };
    }

    console.warn(`[Vexcel] ${label}: fast paths failed, trying synthetic Ctrl+${keyChar.toUpperCase()}`);
    return dispatchCtrlKey(keyChar, keyCode)
      ? { ok: true, strategy: 'syntheticCtrl', verified: false, message: `${label} (best effort)` }
      : { ok: false, message: `${label} failed`, reason: 'synthetic fallback had no target' };
  }

  const shortcuts = {
    // Fill Down - Cmd+D (overrides Chrome bookmark)
    'cmd+d': {
      label: 'Fill Down',
      commandId: 'fillDown',
      action: () => executeFillAction('fillDown', 'Fill Down', 'd', 68),
      overridesChrome: true
    },

    // Fill Right - Cmd+R (overrides Chrome reload)
    'cmd+r': {
      label: 'Fill Right',
      commandId: 'fillRight',
      action: () => executeFillAction('fillRight', 'Fill Right', 'r', 82),
      overridesChrome: true
    },

    // Paste Special - Cmd+Option+V (opens Paste Special submenu)
    'cmd+option+v': {
      label: 'Paste Special',
      action: () => VexcelMenuNavigator.clickMenuPath(['Edit', 'Paste special'])
    },

    // Paste Values Only - Cmd+Shift+V (common Excel shortcut)
    'cmd+shift+v': {
      label: 'Paste Values',
      commandId: 'pasteValues',
      action: () => VexcelFastActions.pasteValues(),
      overridesChrome: true
    },

    // Find & Replace - Cmd+H (overrides Chrome hide)
    'cmd+h': {
      label: 'Find & Replace',
      commandId: 'findReplace',
      action: () => VexcelFastActions.findReplace(),
      overridesChrome: true
    },

    // New Sheet - Cmd+N (opt-in, overrides Chrome new window)
    'cmd+n': {
      label: 'New Sheet',
      action: () => VexcelMenuNavigator.clickMenuPath(['Insert', 'Sheet']),
      overridesChrome: true,
      optIn: true
    },

    // Create Table / Alternating colors - Cmd+T (opt-in, overrides Chrome new tab)
    'cmd+t': {
      label: 'Alternating Colors',
      action: () => VexcelMenuNavigator.clickMenuPath(['Format', 'Alternating colors']),
      overridesChrome: true,
      optIn: true
    }
  };

  return { shortcuts };
})();
