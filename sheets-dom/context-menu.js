// Google Sheets context menu interaction for operations only available via right-click
// (e.g., hide/unhide/resize rows and columns)
//
// IMPORTANT: Google Sheets renders its grid on CANVAS. Row/column headers are NOT
// individual DOM elements. The `.row-headers-background` and `.column-headers-background`
// are full-size overlay divs. To "click a row header" we must dispatch mouse events at
// the correct PIXEL COORDINATES within those overlay divs — using the active cell's
// position to calculate the right row/column.

const VexcelContextMenu = (() => {

  function topDoc() {
    try { return window.top.document; }
    catch (e) { return document; }
  }

  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /** Get all searchable documents — iframes FIRST, then top doc. */
  function allDocs() {
    const docs = [];
    try {
      for (const iframe of topDoc().querySelectorAll('iframe')) {
        try { if (iframe.contentDocument) docs.push(iframe.contentDocument); }
        catch (e) { /* cross-origin */ }
      }
    } catch (e) {}
    docs.push(topDoc());
    return docs;
  }

  /** Find first visible element matching selector across all frames (iframes first). */
  function findVisible(selector) {
    for (const d of allDocs()) {
      for (const el of d.querySelectorAll(selector)) {
        const r = el.getBoundingClientRect();
        if (r.width > 0 && r.height > 0) return el;
      }
    }
    return null;
  }

  /**
   * Simulate a full click at exact pixel coordinates on a target element.
   */
  function clickAtCoords(el, x, y) {
    const win = el.ownerDocument.defaultView || window;
    const opts = { bubbles: true, cancelable: true, view: win, clientX: x, clientY: y };
    el.dispatchEvent(new PointerEvent('pointerdown', { ...opts, pointerId: 1 }));
    el.dispatchEvent(new MouseEvent('mousedown', opts));
    el.dispatchEvent(new PointerEvent('pointerup', { ...opts, pointerId: 1 }));
    el.dispatchEvent(new MouseEvent('mouseup', opts));
    el.dispatchEvent(new MouseEvent('click', opts));
  }

  /**
   * Simulate a right-click at exact pixel coordinates on a target element.
   */
  function rightClickAtCoords(el, x, y) {
    const win = el.ownerDocument.defaultView || window;
    const opts = { bubbles: true, cancelable: true, view: win, clientX: x, clientY: y, button: 2, buttons: 2 };
    el.dispatchEvent(new PointerEvent('pointerdown', { ...opts, pointerId: 1 }));
    el.dispatchEvent(new MouseEvent('mousedown', opts));
    el.dispatchEvent(new PointerEvent('pointerup', { ...opts, pointerId: 1 }));
    el.dispatchEvent(new MouseEvent('mouseup', opts));
    el.dispatchEvent(new MouseEvent('contextmenu', opts));
  }

  /**
   * Simulate a right-click (context menu) on a given element's center.
   */
  function simulateRightClick(el) {
    const rect = el.getBoundingClientRect();
    rightClickAtCoords(el, rect.left + rect.width / 2, rect.top + rect.height / 2);
  }

  /**
   * Find and click a context menu item by label.
   * Also logs all visible menu items on first attempt for debugging.
   */
  async function clickContextMenuItem(label, logItems = false) {
    const doc = topDoc();
    const labelLower = label.toLowerCase();

    for (let attempt = 0; attempt < 15; attempt++) {
      const candidates = doc.querySelectorAll(
        '[role="menuitem"], [role="menuitemcheckbox"], .goog-menuitem'
      );

      // Log all visible items on first attempt for debugging
      if (attempt === 0 && logItems) {
        const visible = [];
        for (const item of candidates) {
          const rect = item.getBoundingClientRect();
          if (rect.width === 0 || rect.height === 0) continue;
          visible.push((item.textContent || '').replace(/\s+/g, ' ').trim());
        }
        console.log(`[Vexcel] Context menu items (${visible.length}):`, visible.join(' | '));
      }

      for (const item of candidates) {
        const rect = item.getBoundingClientRect();
        if (rect.width === 0 || rect.height === 0) continue;
        const text = (item.textContent || '').replace(/\s+/g, ' ').trim().toLowerCase();
        const aria = (item.getAttribute('aria-label') || '').toLowerCase();
        if (text.includes(labelLower) || aria.includes(labelLower)) {
          clickAtCoords(item, rect.left + rect.width / 2, rect.top + rect.height / 2);
          return true;
        }
      }
      await sleep(80);
    }
    console.warn(`[Vexcel] Context menu item not found: "${label}"`);
    return false;
  }

  /**
   * Find the active cell element in the GRID (not the formula bar).
   * The formula bar also has class .cell-input but is ~2400px wide.
   * The actual grid cell is much narrower (typically < 500px).
   * We search iframes first since the grid lives in an iframe.
   */
  function findActiveCell() {
    // Strategy 1: Find .cell-input that's NOT the formula bar (filter by width)
    for (const d of allDocs()) {
      for (const el of d.querySelectorAll('.cell-input')) {
        const r = el.getBoundingClientRect();
        if (r.width > 0 && r.height > 0 && r.width < 1000) {
          return el;
        }
      }
    }

    // Strategy 2: Active cell border/highlight overlay
    const activeBorder = findVisible('.active-cell-border') ||
                         findVisible('.autofill-cover');
    if (activeBorder) return activeBorder;

    // Strategy 3: Any cell-input (even formula bar — better than nothing)
    const anyCellInput = findVisible('.cell-input');
    if (anyCellInput) return anyCellInput;

    // Strategy 4: Grid container
    return findVisible('.grid-container') ||
           findVisible('.waffle-content-container');
  }

  /**
   * Find the row headers background div.
   * This is a single div overlaying all row headers on the left side.
   */
  function findRowHeaderArea() {
    return findVisible('.row-headers-background');
  }

  /**
   * Find the column headers background div.
   * This is a single div overlaying all column headers at the top.
   */
  function findColumnHeaderArea() {
    return findVisible('.column-headers-background');
  }

  /**
   * Double-click on the bottom border of a row header to auto-fit row height.
   * This is the native Google Sheets gesture for auto-resize.
   * We target the bottom edge of the row at the active cell's Y position.
   */
  function dblClickAtCoords(el, x, y) {
    const win = el.ownerDocument.defaultView || window;
    const opts = { bubbles: true, cancelable: true, view: win, clientX: x, clientY: y };
    // Full sequence: mousedown, mouseup, click, mousedown, mouseup, click, dblclick
    el.dispatchEvent(new MouseEvent('mousedown', opts));
    el.dispatchEvent(new MouseEvent('mouseup', opts));
    el.dispatchEvent(new MouseEvent('click', opts));
    el.dispatchEvent(new MouseEvent('mousedown', { ...opts, detail: 2 }));
    el.dispatchEvent(new MouseEvent('mouseup', { ...opts, detail: 2 }));
    el.dispatchEvent(new MouseEvent('click', { ...opts, detail: 2 }));
    el.dispatchEvent(new MouseEvent('dblclick', { ...opts, detail: 2 }));
  }

  /**
   * Auto-resize rows to fit content.
   * Strategies ordered by reliability (tool finder > context menu > double-click).
   */
  async function autoResizeRows() {
    console.log('[Vexcel] Auto-resize rows');

    // Strategy 1: Tool finder → opens resize dialog → select "Fit to data"
    // This is the most reliable approach — no canvas coordinate guessing needed.
    console.log('[Vexcel] Strategy 1: Tool finder "Resize row"');
    let ok = await VexcelMenuNavigator.useToolFinder('Resize row');
    if (!ok) ok = await VexcelMenuNavigator.useToolFinder('Resize rows');
    if (ok) {
      await sleep(400);
      return clickFitToData();
    }

    // Strategy 2: Right-click on row header area → context menu → Resize row
    const activeCell = findActiveCell();
    const rowArea = findRowHeaderArea();

    if (activeCell && rowArea) {
      const cellRect = activeCell.getBoundingClientRect();
      const rowAreaRect = rowArea.getBoundingClientRect();
      const targetX = rowAreaRect.left + rowAreaRect.width / 2;

      let targetY = cellRect.top + cellRect.height / 2;
      if (targetY < rowAreaRect.top || targetY > rowAreaRect.bottom) {
        targetY = rowAreaRect.top + 30;
      }

      console.log('[Vexcel] Strategy 2: Right-click row header at', targetX, targetY);
      clickAtCoords(rowArea, targetX, targetY);
      await sleep(200);
      rightClickAtCoords(rowArea, targetX, targetY);
      await sleep(500);

      ok = await clickContextMenuItem('Resize row', true);
      if (!ok) ok = await clickContextMenuItem('Resize rows');
      if (ok) {
        await sleep(400);
        return clickFitToData();
      }
      closeContextMenu();

      // Strategy 3: Double-click on the bottom border of the row header (native auto-fit gesture)
      let borderY = cellRect.bottom;
      if (borderY < rowAreaRect.top || borderY > rowAreaRect.bottom) {
        borderY = targetY + 10;
      }
      console.log('[Vexcel] Strategy 3: Double-click row header border at', targetX, borderY);
      dblClickAtCoords(rowArea, targetX, borderY);
      await sleep(300);
      return true; // Optimistic — no way to verify double-click auto-fit worked
    }

    console.warn('[Vexcel] All row resize strategies failed');
    return false;
  }

  /**
   * Auto-resize columns to fit content.
   * Strategies ordered by reliability (tool finder > context menu > double-click).
   */
  async function autoResizeColumns() {
    console.log('[Vexcel] Auto-resize columns');

    // Strategy 1: Tool finder → opens resize dialog → select "Fit to data"
    console.log('[Vexcel] Strategy 1: Tool finder "Resize column"');
    let ok = await VexcelMenuNavigator.useToolFinder('Resize column');
    if (!ok) ok = await VexcelMenuNavigator.useToolFinder('Resize columns');
    if (ok) {
      await sleep(400);
      return clickFitToData();
    }

    // Strategy 2: Right-click on column header area → context menu → Resize column
    const activeCell = findActiveCell();
    const colArea = findColumnHeaderArea();

    if (activeCell && colArea) {
      const cellRect = activeCell.getBoundingClientRect();
      const colAreaRect = colArea.getBoundingClientRect();
      const targetY = colAreaRect.top + colAreaRect.height / 2;

      let targetX = cellRect.left + cellRect.width / 2;
      if (targetX < colAreaRect.left || targetX > colAreaRect.right) {
        targetX = colAreaRect.left + 80;
      }

      console.log('[Vexcel] Strategy 2: Right-click column header at', targetX, targetY);
      clickAtCoords(colArea, targetX, targetY);
      await sleep(200);
      rightClickAtCoords(colArea, targetX, targetY);
      await sleep(500);

      ok = await clickContextMenuItem('Resize column', true);
      if (!ok) ok = await clickContextMenuItem('Resize columns');
      if (ok) {
        await sleep(400);
        return clickFitToData();
      }
      closeContextMenu();

      // Strategy 3: Double-click on the right border of the column header (native auto-fit gesture)
      let borderX = cellRect.right;
      if (borderX < colAreaRect.left || borderX > colAreaRect.right) {
        borderX = targetX + 30;
      }
      console.log('[Vexcel] Strategy 3: Double-click column header border at', borderX, targetY);
      dblClickAtCoords(colArea, borderX, targetY);
      await sleep(300);
      return true; // Optimistic — no way to verify double-click auto-fit worked
    }

    console.warn('[Vexcel] All column resize strategies failed');
    return false;
  }

  /**
   * In a resize dialog, click "Fit to data" radio and then OK.
   * Waits for the dialog to appear, then selects the right option.
   */
  async function clickFitToData() {
    const doc = topDoc();

    // Wait for the dialog to appear (retry up to 10 times)
    let dialog = null;
    for (let attempt = 0; attempt < 10; attempt++) {
      const dialogs = doc.querySelectorAll('[role="dialog"], .modal-dialog, [role="alertdialog"]');
      for (const d of dialogs) {
        const r = d.getBoundingClientRect();
        if (r.width > 50 && r.height > 50) { dialog = d; break; }
      }
      if (dialog) break;
      await sleep(150);
    }

    const root = dialog || doc;
    console.log('[Vexcel] Looking for "Fit to data" in', dialog ? 'dialog' : 'document');

    // Log all visible elements in dialog for debugging
    if (dialog) {
      const debugEls = [];
      for (const el of dialog.querySelectorAll('*')) {
        const text = (el.textContent || '').trim();
        if (text && text.length < 50 && el.children.length < 3) {
          debugEls.push(`${el.tagName}.${el.className}: "${text}"`);
        }
      }
      console.log('[Vexcel] Dialog contents:', debugEls.join(' | '));
    }

    // Find and click "Fit to data" — search labels, spans, radios, and divs
    const allEls = root.querySelectorAll('label, span, div, [role="radio"], input[type="radio"]');
    let foundFit = false;
    for (const el of allEls) {
      const text = (el.textContent || '').trim().toLowerCase();
      if (!text.includes('fit to data')) continue;
      if (el.children.length > 5) continue; // skip large containers
      const r = el.getBoundingClientRect();
      if (r.width === 0 || r.height === 0) continue;

      console.log('[Vexcel] Clicking "Fit to data":', el.tagName, el.className);
      clickAtCoords(el, r.left + r.width / 2, r.top + r.height / 2);

      // Also activate any radio input inside or associated with this label
      const radio = el.querySelector('input[type="radio"]') ||
                    el.closest('label')?.querySelector('input[type="radio"]');
      if (radio) {
        radio.checked = true;
        radio.dispatchEvent(new Event('change', { bubbles: true }));
        radio.dispatchEvent(new Event('click', { bubbles: true }));
        const rr = radio.getBoundingClientRect();
        if (rr.width > 0) clickAtCoords(radio, rr.left + rr.width / 2, rr.top + rr.height / 2);
      }
      foundFit = true;
      await sleep(300);
      break;
    }

    if (!foundFit) {
      // Fallback: look for a "Fit to data" option using aria attributes
      const radios = root.querySelectorAll('[role="radio"], [type="radio"]');
      for (const radio of radios) {
        const label = radio.getAttribute('aria-label') || '';
        const parent = radio.closest('label, div');
        const parentText = parent ? (parent.textContent || '').toLowerCase() : '';
        if (label.toLowerCase().includes('fit') || parentText.includes('fit to data')) {
          console.log('[Vexcel] Clicking radio via aria/parent:', label || parentText);
          const r = radio.getBoundingClientRect();
          if (r.width > 0) clickAtCoords(radio, r.left + r.width / 2, r.top + r.height / 2);
          foundFit = true;
          await sleep(300);
          break;
        }
      }
    }

    if (!foundFit) {
      console.warn('[Vexcel] "Fit to data" not found in dialog');
    }

    // Click OK button
    const buttons = root.querySelectorAll('button, [role="button"], .goog-buttonset-default');
    for (const btn of buttons) {
      const text = (btn.textContent || '').trim().toLowerCase();
      if (text === 'ok' || text === 'apply' || text === 'done') {
        console.log('[Vexcel] Clicking OK button');
        const r = btn.getBoundingClientRect();
        clickAtCoords(btn, r.left + r.width / 2, r.top + r.height / 2);
        return true;
      }
    }

    // Fallback: press Enter to dismiss dialog
    console.log('[Vexcel] No OK button found, pressing Enter');
    const target = doc.activeElement || doc.body;
    target.dispatchEvent(new KeyboardEvent('keydown', {
      key: 'Enter', code: 'Enter', keyCode: 13, bubbles: true, cancelable: true
    }));
    return true;
  }

  function closeContextMenu() {
    const doc = topDoc();
    doc.dispatchEvent(new KeyboardEvent('keydown', {
      key: 'Escape', code: 'Escape', keyCode: 27, bubbles: true
    }));
  }

  /**
   * Right-click on the current selection and pick a context menu item.
   * For row/column hide/unhide operations.
   */
  async function rightClickAndSelect(menuLabel) {
    const activeCell = findActiveCell();
    if (!activeCell) {
      console.warn('[Vexcel] Cannot find active area for right-click');
      return false;
    }
    simulateRightClick(activeCell);
    await sleep(350);
    return clickContextMenuItem(menuLabel);
  }

  // Public API
  async function hideRows()      { return rightClickAndSelect('Hide row'); }
  async function hideColumns()   { return rightClickAndSelect('Hide column'); }
  async function unhideRows()    { return rightClickAndSelect('Unhide row'); }
  async function unhideColumns() { return rightClickAndSelect('Unhide column'); }
  async function resizeRows()    { return autoResizeRows(); }
  async function resizeColumns() { return autoResizeColumns(); }

  return { hideRows, hideColumns, unhideRows, unhideColumns, resizeRows, resizeColumns,
           autoResizeRows, autoResizeColumns, rightClickAndSelect, clickContextMenuItem };
})();
