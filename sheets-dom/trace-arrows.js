// Trace Precedents / Dependents for Google Sheets
// Parses formulas, shows referenced cells with values, keyboard navigable

const VexcelTraceArrows = (() => {

  let panelHost = null;
  let panelRoot = null;
  let contentEl = null;
  let visible = false;
  let selectedIdx = 0;
  let currentRefs = [];
  let keyHandler = null;

  // ── Init ──────────────────────────────────────────────────
  function init() {
    const existing = document.getElementById('vexcel-trace-host');
    if (existing) existing.remove();

    panelHost = document.createElement('div');
    panelHost.id = 'vexcel-trace-host';
    panelHost.style.cssText =
      'position:fixed;top:0;left:0;width:100%;height:100%;' +
      'z-index:2147483646;pointer-events:none;display:none;';

    const attach = () => {
      document.body.appendChild(panelHost);
      panelRoot = panelHost.attachShadow({ mode: 'open' });

      const style = document.createElement('style');
      style.textContent = CSS;
      panelRoot.appendChild(style);

      contentEl = document.createElement('div');
      contentEl.id = 'vt-root';
      panelRoot.appendChild(contentEl);
    };

    if (document.body) attach();
    else document.addEventListener('DOMContentLoaded', attach);
  }

  // ── Trace Precedents ──────────────────────────────────────
  function tracePrecedents() {
    clearArrows();

    const cellAddr = getCurrentCellAddress();
    const formula = getFormula();

    if (!formula || !formula.startsWith('=')) {
      showMessage('No formula in the selected cell.');
      return;
    }

    const refs = parseFormulaRefs(formula);
    if (refs.length === 0) {
      showMessage('No cell references found in formula.');
      return;
    }

    // Fetch cell values for each ref, then show the panel
    fetchCellValues(refs).then(refData => {
      showPrecedentsPanel(cellAddr, formula, refData);
    });
  }

  // ── Trace Dependents ──────────────────────────────────────
  function traceDependents() {
    clearArrows();

    const cellAddr = getCurrentCellAddress();
    if (!cellAddr) {
      showMessage('Cannot determine current cell address.');
      return;
    }

    showDependentsPanel(cellAddr);
  }

  // ── Clear ─────────────────────────────────────────────────
  function clearArrows() {
    visible = false;
    currentRefs = [];
    selectedIdx = 0;
    if (panelHost) panelHost.style.display = 'none';
    if (contentEl) contentEl.innerHTML = '';
    if (keyHandler) {
      document.removeEventListener('keydown', keyHandler, true);
      keyHandler = null;
    }
  }

  // ── Fetch cell values via the formula bar ──────────────────
  // Navigate to each ref, read its value from the formula bar, then return
  async function fetchCellValues(refs) {
    const topDoc = getTopDocument();
    const results = [];

    // Save current cell address to return to it later
    const originalAddr = getCurrentCellAddress();

    for (const ref of refs) {
      const cleanRef = ref.replace(/\$/g, '');
      // For ranges (A1:B5), just show the range notation
      if (cleanRef.includes(':')) {
        results.push({ ref, value: '(range)' });
        continue;
      }
      // For sheet-qualified refs, show as-is
      if (cleanRef.includes('!')) {
        results.push({ ref, value: '(other sheet)' });
        continue;
      }

      // Navigate to the cell and read its displayed value
      const val = await readCellValue(cleanRef, topDoc);
      results.push({ ref, value: val || '' });
    }

    // Return to original cell
    if (originalAddr) {
      await doNavigateToCell(originalAddr, topDoc);
      await sleep(100);
    }

    return results;
  }

  /**
   * Read a cell's displayed value by navigating to it via the Name Box.
   * Uses keyboard simulation (select all, type, enter) for reliable input.
   */
  async function readCellValue(cellAddr, topDoc) {
    const ok = await doNavigateToCell(cellAddr, topDoc);
    if (!ok) return '';

    await sleep(200);

    // Read the formula bar content (shows the cell's value/formula)
    return readFormulaBar(topDoc);
  }

  /**
   * Navigate to a cell by focusing the Name Box and typing into it.
   * Uses keyboard events (Ctrl+A to select all, then type each char) for reliability.
   */
  async function doNavigateToCell(cellAddr, topDoc) {
    const nameBox = findNameBox(topDoc);
    if (!nameBox) {
      console.warn('[Vexcel] Name Box not found');
      return false;
    }

    // Click the Name Box to focus it (more reliable than .focus())
    const rect = nameBox.getBoundingClientRect();
    const x = rect.left + rect.width / 2;
    const y = rect.top + rect.height / 2;
    const win = nameBox.ownerDocument.defaultView || window;
    const clickOpts = { bubbles: true, cancelable: true, view: win, clientX: x, clientY: y };
    nameBox.dispatchEvent(new MouseEvent('mousedown', clickOpts));
    nameBox.dispatchEvent(new MouseEvent('mouseup', clickOpts));
    nameBox.dispatchEvent(new MouseEvent('click', clickOpts));
    nameBox.focus();
    await sleep(50);

    // Select all existing text (Ctrl+A / Cmd+A)
    nameBox.dispatchEvent(new KeyboardEvent('keydown', {
      key: 'a', code: 'KeyA', keyCode: 65, ctrlKey: true, metaKey: true, bubbles: true
    }));
    nameBox.select();
    await sleep(30);

    // Clear and set value using both native setter and direct assignment
    const nativeSet = Object.getOwnPropertyDescriptor(
      HTMLInputElement.prototype, 'value'
    );
    if (nativeSet && nativeSet.set) {
      nativeSet.set.call(nameBox, cellAddr);
    } else {
      nameBox.value = cellAddr;
    }
    nameBox.dispatchEvent(new Event('input', { bubbles: true }));
    nameBox.dispatchEvent(new Event('change', { bubbles: true }));

    // Also type each character individually as backup
    for (const char of cellAddr) {
      nameBox.dispatchEvent(new KeyboardEvent('keydown', {
        key: char, code: `Key${char.toUpperCase()}`, keyCode: char.charCodeAt(0), bubbles: true
      }));
      nameBox.dispatchEvent(new KeyboardEvent('keypress', {
        key: char, code: `Key${char.toUpperCase()}`, charCode: char.charCodeAt(0), bubbles: true
      }));
      nameBox.dispatchEvent(new KeyboardEvent('keyup', {
        key: char, code: `Key${char.toUpperCase()}`, keyCode: char.charCodeAt(0), bubbles: true
      }));
    }
    await sleep(30);

    // Press Enter to navigate
    const enterOpts = { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true };
    nameBox.dispatchEvent(new KeyboardEvent('keydown', enterOpts));
    nameBox.dispatchEvent(new KeyboardEvent('keypress', enterOpts));
    nameBox.dispatchEvent(new KeyboardEvent('keyup', enterOpts));

    return true;
  }

  /**
   * Read the formula bar content.
   */
  function readFormulaBar(topDoc) {
    const selectors = [
      '#t-formula-bar-input',
      '.cell-input[aria-label="Formula Bar"]',
      '.formulabar-input',
    ];
    for (const sel of selectors) {
      const el = topDoc.querySelector(sel);
      if (el) {
        const text = (el.textContent || el.innerText || '').trim();
        if (text) return text;
      }
    }

    // Fallback: check any .cell-input in the formula bar area
    const inputs = topDoc.querySelectorAll('.cell-input');
    for (const el of inputs) {
      const rect = el.getBoundingClientRect();
      if (rect.top < 80 && rect.width > 100 && rect.height > 0) {
        return (el.textContent || el.innerText || '').trim();
      }
    }

    return '';
  }

  // ── Get formula from the formula bar ──────────────────────
  function getFormula() {
    const topDoc = getTopDocument();

    const selectors = [
      '#t-formula-bar-input',
      '.cell-input[aria-label="Formula Bar"]',
      '.formulabar-input',
      '.cell-input',
    ];

    for (const sel of selectors) {
      const els = topDoc.querySelectorAll(sel);
      for (const el of els) {
        if (sel === '.cell-input') {
          const rect = el.getBoundingClientRect();
          if (rect.top > 80 || rect.width < 100 || rect.height === 0) continue;
        }
        const text = el.textContent.trim();
        if (text) return text;
      }
    }

    const ariaEls = topDoc.querySelectorAll('[aria-label*="ormula" i]');
    for (const el of ariaEls) {
      const text = el.textContent.trim();
      if (text && text.startsWith('=')) return text;
    }

    return '';
  }

  // ── Get current cell address from Name Box ────────────────
  function getCurrentCellAddress() {
    const topDoc = getTopDocument();

    const selectors = [
      '#t-name-box input',
      '.jfk-textinput[aria-label="Name Box"]',
      'input[aria-label="Name Box"]',
      '#t-name-box',
    ];

    for (const sel of selectors) {
      const el = topDoc.querySelector(sel);
      if (!el) continue;
      const val = (el.value || el.textContent || '').trim();
      if (val && /^[A-Z]+\d+/i.test(val)) return val;
    }

    return '';
  }

  // ── Formula parser ────────────────────────────────────────
  function parseFormulaRefs(formula) {
    if (!formula || !formula.startsWith('=')) return [];

    const body = formula.slice(1);
    const cleaned = body.replace(/"[^"]*"/g, match => ' '.repeat(match.length));

    const refPattern =
      /(?:(?:'(?:[^']|'')*'|[A-Za-z_]\w*)!)?(\$?[A-Z]{1,3}\$?\d{1,7})(?::(\$?[A-Z]{1,3}\$?\d{1,7}))?/gi;

    const functionNames = new Set([
      'IF', 'OR', 'AND', 'NOT', 'TRUE', 'FALSE', 'PI', 'NA', 'NOW', 'ROW', 'ABS',
      'LOG', 'EXP', 'SIN', 'COS', 'TAN', 'LEN', 'MOD', 'MIN', 'MAX', 'SUM',
      'MID', 'DAY', 'DATE', 'CODE', 'CHAR', 'INT',
    ]);

    const refs = [];
    const seen = new Set();
    let match;

    while ((match = refPattern.exec(cleaned)) !== null) {
      const full = match[0];
      const cellPart = match[1];
      const afterIdx = match.index + full.length;
      if (afterIdx < cleaned.length && cleaned[afterIdx] === '(') continue;
      const upperCell = cellPart.replace(/\$/g, '').toUpperCase();
      if (functionNames.has(upperCell)) continue;
      const colMatch = upperCell.match(/^([A-Z]+)/);
      if (colMatch && colMatch[1].length > 3) continue;
      const key = full.toUpperCase();
      if (!seen.has(key)) {
        seen.add(key);
        refs.push(full);
      }
    }

    return refs;
  }

  // ── Navigate to a cell via the Name Box ───────────────────
  function navigateToCell(ref) {
    const topDoc = getTopDocument();
    // Use the reliable navigation helper (async, but we fire-and-forget here)
    doNavigateToCell(ref, topDoc);
    return true;
  }

  function findNameBox(doc) {
    const selectors = [
      '#t-name-box input',
      '.jfk-textinput[aria-label="Name Box"]',
      'input[aria-label="Name Box"]',
    ];
    for (const sel of selectors) {
      const el = doc.querySelector(sel);
      if (el) return el;
    }
    return null;
  }

  // ── Open Find & Replace to search for a cell reference ────
  async function searchInFormulas(cellAddr) {
    await VexcelMenuNavigator.clickMenuPath(['Edit', 'Find and replace']);
    await sleep(600);

    const topDoc = getTopDocument();
    const inputs = topDoc.querySelectorAll(
      '[role="dialog"] input[type="text"], [role="dialog"] input:not([type])'
    );
    for (const input of inputs) {
      const rect = input.getBoundingClientRect();
      if (rect.width > 50 && rect.height > 0) {
        input.focus();
        const nativeSet = Object.getOwnPropertyDescriptor(
          HTMLInputElement.prototype, 'value'
        );
        if (nativeSet && nativeSet.set) {
          nativeSet.set.call(input, cellAddr);
        } else {
          input.value = cellAddr;
        }
        input.dispatchEvent(new Event('input', { bubbles: true }));
        break;
      }
    }

    await sleep(200);
    const checkboxes = topDoc.querySelectorAll('[role="dialog"] [role="checkbox"], [role="dialog"] .goog-checkbox');
    for (const cb of checkboxes) {
      const label = cb.textContent || cb.getAttribute('aria-label') || '';
      if (label.toLowerCase().includes('formula')) {
        if (cb.getAttribute('aria-checked') !== 'true') cb.click();
        break;
      }
    }
    const labels = topDoc.querySelectorAll('[role="dialog"] label');
    for (const label of labels) {
      const text = (label.textContent || '').toLowerCase();
      if (text.includes('formula') || text.includes('within')) {
        const checkbox = label.querySelector('input[type="checkbox"]');
        if (checkbox && !checkbox.checked) checkbox.click();
        break;
      }
    }
  }

  // ── Keyboard handler for the panel ─────────────────────────
  function attachKeyboard(refData) {
    currentRefs = refData;
    selectedIdx = 0;

    keyHandler = (e) => {
      if (!visible) return;

      if (e.key === 'Escape') {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        clearArrows();
        return;
      }

      if (e.key === 'ArrowDown' || e.key === 'j') {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        selectedIdx = Math.min(selectedIdx + 1, currentRefs.length - 1);
        updateSelection();
        return;
      }

      if (e.key === 'ArrowUp' || e.key === 'k') {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        selectedIdx = Math.max(selectedIdx - 1, 0);
        updateSelection();
        return;
      }

      if (e.key === 'Enter') {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        if (currentRefs[selectedIdx]) {
          const ref = currentRefs[selectedIdx].ref.replace(/\$/g, '');
          clearArrows();
          navigateToCell(ref);
        }
        return;
      }

      // Number keys 1-9 to jump directly to a ref
      const num = parseInt(e.key);
      if (num >= 1 && num <= 9 && num <= currentRefs.length) {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        selectedIdx = num - 1;
        updateSelection();
        return;
      }
    };

    document.addEventListener('keydown', keyHandler, true);
  }

  function updateSelection() {
    if (!panelRoot) return;
    const rows = panelRoot.querySelectorAll('.vt-ref-row');
    rows.forEach((row, i) => {
      row.classList.toggle('vt-ref-selected', i === selectedIdx);
    });
  }

  // ── Precedents Panel UI ───────────────────────────────────
  function showPrecedentsPanel(cellAddr, formula, refData) {
    if (!contentEl) return;
    contentEl.innerHTML = '';
    panelHost.style.display = 'block';
    visible = true;

    const panel = mkEl('div', 'vt-panel');

    // Header
    const header = mkEl('div', 'vt-header');
    const title = mkEl('span', 'vt-title');
    title.textContent = `Precedents of ${cellAddr}`;
    header.appendChild(title);
    const hint = mkEl('span', 'vt-kbd-hint');
    hint.textContent = 'Up/Down Enter Esc';
    header.appendChild(hint);
    const closeBtn = mkEl('span', 'vt-close');
    closeBtn.textContent = '\u2715';
    closeBtn.addEventListener('click', clearArrows);
    header.appendChild(closeBtn);
    panel.appendChild(header);

    // Formula display
    const formulaEl = mkEl('div', 'vt-formula');
    formulaEl.textContent = formula;
    panel.appendChild(formulaEl);

    // Reference list
    const list = mkEl('div', 'vt-list');
    refData.forEach((item, i) => {
      const row = mkEl('div', 'vt-ref-row');
      if (i === 0) row.classList.add('vt-ref-selected');

      // Number badge for keyboard shortcut
      const numBadge = mkEl('span', 'vt-num');
      numBadge.textContent = String(i + 1);
      row.appendChild(numBadge);

      // Cell reference
      const refLabel = mkEl('span', 'vt-ref-label');
      refLabel.textContent = item.ref;
      row.appendChild(refLabel);

      // Cell value
      const valLabel = mkEl('span', 'vt-ref-value');
      valLabel.textContent = item.value || '(empty)';
      row.appendChild(valLabel);

      row.addEventListener('click', () => {
        const cleanRef = item.ref.replace(/\$/g, '');
        clearArrows();
        navigateToCell(cleanRef);
      });

      list.appendChild(row);
    });
    panel.appendChild(list);

    // Select All button
    if (refData.length > 1) {
      const selectAll = mkEl('div', 'vt-select-all');
      selectAll.textContent = 'Select All Precedents';
      selectAll.addEventListener('click', () => {
        const cleanRefs = refData.map(r => r.ref.replace(/\$/g, ''));
        clearArrows();
        navigateToCell(cleanRefs.join(','));
      });
      panel.appendChild(selectAll);
    }

    contentEl.appendChild(panel);

    // Attach keyboard navigation
    attachKeyboard(refData);
  }

  // ── Dependents Panel UI ───────────────────────────────────
  function showDependentsPanel(cellAddr) {
    if (!contentEl) return;
    contentEl.innerHTML = '';
    panelHost.style.display = 'block';
    visible = true;

    const panel = mkEl('div', 'vt-panel');

    const header = mkEl('div', 'vt-header');
    const title = mkEl('span', 'vt-title');
    title.textContent = `Dependents of ${cellAddr}`;
    header.appendChild(title);
    const closeBtn = mkEl('span', 'vt-close');
    closeBtn.textContent = '\u2715';
    closeBtn.addEventListener('click', clearArrows);
    header.appendChild(closeBtn);
    panel.appendChild(header);

    const info = mkEl('div', 'vt-info');
    info.textContent =
      'To find cells that depend on ' + cellAddr +
      ', we\'ll open Find & Replace to search within formulas.';
    panel.appendChild(info);

    const searchBtn = mkEl('div', 'vt-search-btn');
    searchBtn.textContent = 'Search Formulas for ' + cellAddr;
    searchBtn.addEventListener('click', () => {
      clearArrows();
      searchInFormulas(cellAddr);
    });
    panel.appendChild(searchBtn);

    // Keyboard: Enter to search, Esc to close
    keyHandler = (e) => {
      if (!visible) return;
      if (e.key === 'Escape') {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        clearArrows();
      } else if (e.key === 'Enter') {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        clearArrows();
        searchInFormulas(cellAddr);
      }
    };
    document.addEventListener('keydown', keyHandler, true);

    contentEl.appendChild(panel);
  }

  // ── Message (no formula, etc.) ────────────────────────────
  function showMessage(msg) {
    if (!contentEl) return;
    contentEl.innerHTML = '';
    panelHost.style.display = 'block';
    visible = true;

    const panel = mkEl('div', 'vt-panel vt-panel-sm');

    const header = mkEl('div', 'vt-header');
    const title = mkEl('span', 'vt-title');
    title.textContent = 'Trace';
    header.appendChild(title);
    const closeBtn = mkEl('span', 'vt-close');
    closeBtn.textContent = '\u2715';
    closeBtn.addEventListener('click', clearArrows);
    header.appendChild(closeBtn);
    panel.appendChild(header);

    const msgEl = mkEl('div', 'vt-info');
    msgEl.textContent = msg;
    panel.appendChild(msgEl);

    contentEl.appendChild(panel);

    // Esc to close
    keyHandler = (e) => {
      if (e.key === 'Escape') {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        clearArrows();
      }
    };
    document.addEventListener('keydown', keyHandler, true);

    setTimeout(() => { if (visible) clearArrows(); }, 4000);
  }

  // ── Helpers ───────────────────────────────────────────────
  function mkEl(tag, cls) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    return e;
  }

  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  function getTopDocument() {
    try { return window.top.document; }
    catch (e) { return document; }
  }

  // ── CSS ───────────────────────────────────────────────────
  const CSS = `
    :host { all: initial; }
    #vt-root { font-family: -apple-system, 'Segoe UI', sans-serif; }

    .vt-panel {
      position: fixed; top: 50px; right: 20px;
      background: #1c1c2e; border: 1px solid #383860;
      border-radius: 8px; padding: 0; width: 360px;
      box-shadow: 0 8px 32px rgba(0,0,0,.6);
      pointer-events: auto; overflow: hidden;
    }
    .vt-panel-sm { width: 260px; }

    .vt-header {
      display: flex; align-items: center; gap: 8px;
      padding: 10px 14px; background: #252540;
      border-bottom: 1px solid #383860;
    }
    .vt-title {
      font-size: 13px; font-weight: 600; color: #4a8fdb; flex: 1;
    }
    .vt-kbd-hint {
      font-size: 9px; color: #555; padding: 2px 5px;
      border: 1px solid #333; border-radius: 3px;
      font-family: 'SF Mono', Menlo, monospace;
    }
    .vt-close {
      font-size: 14px; color: #666; cursor: pointer;
      padding: 2px 6px; border-radius: 3px;
      transition: background .1s, color .1s;
    }
    .vt-close:hover { background: rgba(255,255,255,.1); color: #ccc; }

    .vt-formula {
      padding: 8px 14px; font-size: 11px; color: #999;
      background: #1a1a2a; border-bottom: 1px solid #2a2a40;
      font-family: 'SF Mono', Menlo, monospace;
      word-break: break-all; max-height: 60px; overflow-y: auto;
    }

    .vt-list { padding: 4px 6px; }

    .vt-ref-row {
      display: flex; align-items: center; gap: 8px;
      padding: 7px 10px; border-radius: 5px; cursor: pointer;
      transition: background .08s; border: 1px solid transparent;
    }
    .vt-ref-row:hover { background: rgba(74,143,219,.12); }
    .vt-ref-selected {
      background: rgba(74,143,219,.2) !important;
      border-color: #4a8fdb;
    }

    .vt-num {
      display: inline-flex; align-items: center; justify-content: center;
      width: 18px; height: 18px; border-radius: 3px;
      background: #2a2a45; border: 1px solid #444;
      color: #888; font-size: 10px; font-weight: 600;
      font-family: 'SF Mono', Menlo, monospace; flex-shrink: 0;
    }
    .vt-ref-selected .vt-num {
      background: #4a8fdb; border-color: #4a8fdb; color: #fff;
    }

    .vt-ref-label {
      font-size: 12px; color: #ddd; font-weight: 600;
      font-family: 'SF Mono', Menlo, monospace;
      min-width: 55px; flex-shrink: 0;
    }

    .vt-ref-value {
      font-size: 11px; color: #8a8aaa; flex: 1;
      font-family: 'SF Mono', Menlo, monospace;
      overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
      text-align: right;
    }

    .vt-select-all {
      margin: 4px 10px 10px; padding: 8px 0;
      text-align: center; font-size: 11px; font-weight: 600;
      color: #4a8fdb; cursor: pointer;
      border: 1px solid #4a8fdb; border-radius: 5px;
      transition: background .1s;
    }
    .vt-select-all:hover { background: rgba(74,143,219,.15); }

    .vt-info {
      padding: 12px 14px; font-size: 12px; color: #999;
      line-height: 1.5;
    }

    .vt-search-btn {
      margin: 4px 14px 14px; padding: 10px 0;
      text-align: center; font-size: 12px; font-weight: 600;
      color: #fff; background: #4a8fdb; cursor: pointer;
      border-radius: 5px; transition: background .1s;
    }
    .vt-search-btn:hover { background: #3a7fcb; }
  `;

  return { init, tracePrecedents, traceDependents, clearArrows };

})();
