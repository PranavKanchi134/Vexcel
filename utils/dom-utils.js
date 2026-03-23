// Centralized Google Sheets DOM selectors and utilities

const VexcelDom = (() => {
  // Selectors for key Google Sheets UI elements
  const SELECTORS = {
    // Top-level menu bar items
    menuBar: '#docs-menubar',
    menuFile: '#docs-file-menu',
    menuEdit: '#docs-edit-menu',
    menuView: '#docs-view-menu',
    menuInsert: '#docs-insert-menu',
    menuFormat: '#docs-format-menu',
    menuData: '#docs-data-menu',
    menuTools: '#docs-tools-menu',
    menuExtensions: '#docs-extensions-menu',
    menuHelp: '#docs-help-menu',

    // Name box (cell reference / Go To) — lives in top frame
    nameBox: '#t-name-box input, .jfk-textinput[aria-label="Name Box"], input[aria-label="Name Box"]',

    // Formula bar — lives in top frame
    formulaBar: '#t-formula-bar-input, .cell-input.formula-bar-input',

    // Cell editor (the active cell input area) — may be in iframe
    cellEditor: '.cell-input, #waffle-cell-editor, [role="textbox"].cell-input',

    // Sheet tab bar — lives in top frame
    sheetTabBar: '.docs-sheet-tab-bar',
    sheetTabs: '.docs-sheet-tab',
    activeSheetTab: '.docs-sheet-active-tab',

    // Toolbar — lives in top frame
    toolbar: '#toolbar',

    // Modal dialogs
    modalDialog: '.modal-dialog',
    dialogOverlay: '.modal-dialog-bg',

    // Tool finder (Help > Search menus)
    toolFinder: '.docs-tool-finder-input input',

    // Menu items (generic)
    menuItem: '[role="menuitem"]',
    menuItemCheckbox: '[role="menuitemcheckbox"]',

    // Grid/spreadsheet area
    grid: '.waffle-content-container, .grid-container',
  };

  /**
   * Get the top-level document (menus, toolbar, name box live here).
   * Falls back to current document if cross-origin.
   */
  function topDoc() {
    try { return window.top.document; }
    catch (e) { return document; }
  }

  /**
   * Query an element using our centralized selectors.
   * @param {string} selectorKey - Key from SELECTORS
   * @param {Document|Element} root - Root to search in (default: current document)
   */
  function query(selectorKey, root = document) {
    const selector = SELECTORS[selectorKey];
    if (!selector) return null;
    return root.querySelector(selector);
  }

  /**
   * Query all matching elements.
   */
  function queryAll(selectorKey, root = document) {
    const selector = SELECTORS[selectorKey];
    if (!selector) return [];
    return Array.from(root.querySelectorAll(selector));
  }

  /**
   * Check if a modal dialog is currently open.
   * Checks both the current frame and the top frame.
   */
  function isDialogOpen() {
    // Check current document
    if (_checkDialogInDoc(document)) return true;

    // Also check top frame (dialogs render in the top frame)
    const top = topDoc();
    if (top !== document && _checkDialogInDoc(top)) return true;

    return false;
  }

  function _checkDialogInDoc(doc) {
    const dialog = doc.querySelector(SELECTORS.modalDialog);
    if (dialog && isElementVisible(dialog)) return true;

    const googDialog = doc.querySelector('[role="dialog"]');
    if (googDialog && isElementVisible(googDialog)) return true;

    // Check for open dropdown/context menus — but NOT inside the
    // accelerator overlay (which has its own shadow DOM)
    const openMenu = doc.querySelector('[role="menu"]');
    if (openMenu && isElementVisible(openMenu)) {
      // Don't count the accelerator overlay's elements
      const accHost = doc.getElementById('vexcel-accelerator-host');
      if (!accHost || !accHost.contains(openMenu)) return true;
    }

    return false;
  }

  /**
   * Reliable visibility check that works for fixed/absolute positioned elements.
   */
  function isElementVisible(el) {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    if (rect.width === 0 || rect.height === 0) return false;
    const style = getComputedStyle(el);
    if (style.display === 'none' || style.visibility === 'hidden') return false;
    if (parseFloat(style.opacity) < 0.1) return false;
    // Off-screen elements
    if (rect.bottom < 0 || rect.right < 0) return false;
    if (rect.top > (window.innerHeight || 9999)) return false;
    if (rect.left > (window.innerWidth || 9999)) return false;
    return true;
  }

  /**
   * Check if focus is in a text input (not the cell editor).
   * We don't want to intercept shortcuts when typing in dialogs, find bar, etc.
   */
  function isInTextInput() {
    const active = document.activeElement;
    if (!active) return false;

    const tag = active.tagName.toLowerCase();
    const activeType = (active.type || '').toLowerCase();
    const insideDialog = !!active.closest('[role="dialog"], .modal-dialog, .docs-dialog');
    const visible = isElementVisible(active);

    // Google Sheets frequently keeps focus on hidden/off-screen capture fields
    // while the grid is active. Those should NOT block shortcut interception.
    if (tag === 'input' || tag === 'textarea') {
      if (insideDialog) return true;
      if (!visible) return false;
      if (tag === 'input' && activeType === 'hidden') return false;
      return true;
    }

    // Block contenteditable ONLY when it's inside a dialog/modal.
    // The Google Sheets grid itself is contenteditable — we must NOT block it.
    if (active.isContentEditable) {
      if (insideDialog) return true;
    }

    return false;
  }

  /**
   * Check if we should intercept shortcuts in the current context.
   */
  function shouldIntercept() {
    if (isDialogOpen()) return false;
    if (isInTextInput()) return false;
    return true;
  }

  /**
   * Find an element by its aria-label (partial or full match).
   */
  function findByAriaLabel(label, root = null, exact = false) {
    const doc = root || topDoc();
    const escapedLabel = label.replace(/\\/g, '\\\\').replace(/"/g, '\\"');
    if (exact) {
      return doc.querySelector(`[aria-label="${escapedLabel}"]`);
    }
    // Try exact first, then contains
    let el = doc.querySelector(`[aria-label="${escapedLabel}"]`);
    if (!el) {
      const all = doc.querySelectorAll('[aria-label]');
      for (const candidate of all) {
        if (candidate.getAttribute('aria-label').includes(label)) {
          el = candidate;
          break;
        }
      }
    }
    return el;
  }

  /**
   * Get the name box element for focusing (Go To).
   * Name box lives in the top frame.
   */
  function getNameBox() {
    const top = topDoc();
    return top.querySelector(SELECTORS.nameBox);
  }

  /**
   * Get all sheet tabs.
   * Sheet tabs live in the top frame.
   */
  function getSheetTabs() {
    const top = topDoc();
    return Array.from(top.querySelectorAll(SELECTORS.sheetTabs));
  }

  /**
   * Get the currently active sheet tab.
   */
  function getActiveSheetTab() {
    const top = topDoc();
    return top.querySelector(SELECTORS.activeSheetTab);
  }

  return {
    SELECTORS,
    topDoc,
    query,
    queryAll,
    isDialogOpen,
    isInTextInput,
    shouldIntercept,
    findByAriaLabel,
    getNameBox,
    getSheetTabs,
    getActiveSheetTab
  };
})();
