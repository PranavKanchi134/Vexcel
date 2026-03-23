// Google Sheets toolbar button interaction

const VexcelToolbar = (() => {

  // Dispatches a full pointer/mouse event sequence.
  function simulateClick(el) {
    const rect = el.getBoundingClientRect();
    const x = rect.left + rect.width / 2;
    const y = rect.top  + rect.height / 2;
    const win = el.ownerDocument.defaultView || window;
    const common = { bubbles: true, cancelable: true, view: win, clientX: x, clientY: y };

    el.dispatchEvent(new PointerEvent('pointerdown', { ...common, pointerId: 1 }));
    el.dispatchEvent(new MouseEvent('mousedown', common));
    el.dispatchEvent(new PointerEvent('pointerup',   { ...common, pointerId: 1 }));
    el.dispatchEvent(new MouseEvent('mouseup',   common));
    el.dispatchEvent(new MouseEvent('click',     common));
  }

  function topDoc() {
    try { return window.top.document; }
    catch (e) { return document; }
  }

  /**
   * Click a toolbar button by its aria-label.
   * Tries simulateClick first, then focus+space as keyboard fallback.
   */
  function clickButton(label) {
    const button = findToolbarButton(label);
    if (button) {
      console.log(`[Vexcel] Clicking toolbar button: "${label}"`, button);
      simulateClick(button);
      return true;
    }
    console.warn(`[Vexcel] Toolbar button not found: "${label}"`);
    return false;
  }

  function toggleButton(label) {
    return clickButton(label);
  }

  /**
   * Find a toolbar button by aria-label.
   * Searches ALL elements with aria-labels in the toolbar area.
   */
  function findToolbarButton(label) {
    const doc = topDoc();
    const candidates = doc.querySelectorAll('[aria-label]');
    const labelLower = label.toLowerCase();
    const labelWords = labelLower.split(/\s+/);

    // Pass 1: exact aria-label match
    for (const el of candidates) {
      if (!inToolbarArea(el)) continue;
      if (el.getAttribute('aria-label') === label) return el;
    }

    // Pass 2: case-insensitive exact match
    for (const el of candidates) {
      if (!inToolbarArea(el)) continue;
      if (el.getAttribute('aria-label').toLowerCase() === labelLower) return el;
    }

    // Pass 3: aria-label contains our label OR our label contains the aria-label
    for (const el of candidates) {
      if (!inToolbarArea(el)) continue;
      const al = el.getAttribute('aria-label').toLowerCase();
      if (al.includes(labelLower)) return el;
      if (labelLower.includes(al) && al.length > 2) return el;
    }

    // Pass 4: all words in our label appear in the aria-label
    if (labelWords.length >= 2) {
      for (const el of candidates) {
        if (!inToolbarArea(el)) continue;
        const al = el.getAttribute('aria-label').toLowerCase();
        if (labelWords.every(w => al.includes(w))) return el;
      }
    }

    // Pass 5: try by id prefix (Google Sheets uses t-bold, t-italic, etc.)
    const idMap = {
      'bold': 't-bold', 'italic': 't-italic', 'underline': 't-underline',
      'strikethrough': 't-strikethrough',
    };
    if (idMap[labelLower]) {
      const el = doc.getElementById(idMap[labelLower]);
      if (el) return el;
    }

    // Pass 6: search by data-tooltip
    for (const el of candidates) {
      if (!inToolbarArea(el)) continue;
      const tooltip = (el.getAttribute('data-tooltip') || '').toLowerCase();
      if (tooltip.includes(labelLower)) return el;
      if (labelLower.includes(tooltip) && tooltip.length > 2) return el;
    }

    return null;
  }

  function inToolbarArea(el) {
    const r = el.getBoundingClientRect();
    return r.top < 200 && r.height > 4 && r.height < 64 && r.width > 4;
  }

  /**
   * Click a toolbar dropdown and select an item.
   * Handles both text-based menus AND icon-based panels (like borders).
   */
  async function selectFromDropdown(dropdownLabel, itemLabel) {
    const dropdown = findToolbarButton(dropdownLabel);
    if (!dropdown) {
      console.warn(`[Vexcel] Dropdown button not found: "${dropdownLabel}"`);
      return false;
    }

    console.log(`[Vexcel] Opening dropdown: "${dropdownLabel}"`);
    simulateClick(dropdown);
    await new Promise(resolve => setTimeout(resolve, 400));

    const doc = topDoc();
    const itemLabelLower = itemLabel.toLowerCase();

    // Helper: check if an element IS the dropdown button or inside it
    // (we must not click the button itself as a "dropdown item")
    function isDropdownButton(el) {
      return el === dropdown || dropdown.contains(el) || el.contains(dropdown);
    }

    // Helper: check if an element is in the toolbar area (not in the dropdown panel)
    function isInToolbar(el) {
      const r = el.getBoundingClientRect();
      return r.top < 80 && r.height < 64;
    }

    // Strategy 1: Standard menu items (text-based)
    const menuItems = doc.querySelectorAll(
      '[role="menuitem"], [role="menuitemcheckbox"], [role="option"], .goog-menuitem, [role="listbox"] [role="option"]'
    );
    for (const item of menuItems) {
      if (isDropdownButton(item)) continue;
      const rect = item.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0) continue;
      const text = (item.textContent || '').replace(/\s+/g, ' ').trim().toLowerCase();
      const aria = (item.getAttribute('aria-label') || '').toLowerCase();
      if (text === itemLabelLower || aria === itemLabelLower ||
          text.includes(itemLabelLower) || aria.includes(itemLabelLower)) {
        console.log(`[Vexcel] Selecting dropdown item (menu): "${itemLabel}"`, item);
        simulateClick(item);
        return true;
      }
    }

    // Strategy 2: Icon-based panels (borders, etc.)
    // Search elements with matching aria-label or data-tooltip
    // that are in the dropdown panel (below the toolbar), NOT the button itself
    const allAria = doc.querySelectorAll('[aria-label], [data-tooltip]');
    for (const el of allAria) {
      if (isDropdownButton(el) || isInToolbar(el)) continue;
      const rect = el.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0) continue;

      const aria = (el.getAttribute('aria-label') || '').toLowerCase();
      const tooltip = (el.getAttribute('data-tooltip') || '').toLowerCase();

      // Only use exact or starts-with matching — no reverse-includes
      for (const source of [aria, tooltip]) {
        if (!source) continue;
        if (source === itemLabelLower || source.includes(itemLabelLower)) {
          console.log(`[Vexcel] Selecting dropdown item (icon): "${itemLabel}" via "${source}"`, el);
          // For Closure Library palette items, the event listener is on
          // the parent goog-palette-cell (td), not the inner icon div.
          // Walk up to find the palette cell and click that.
          const target = el.closest('.goog-palette-cell, [role="gridcell"]') || el;
          const cr = target.getBoundingClientRect();
          const cx = cr.left + cr.width / 2;
          const cy = cr.top + cr.height / 2;
          const w = target.ownerDocument.defaultView || window;
          const evOpts = { bubbles: true, cancelable: true, view: w, clientX: cx, clientY: cy };
          target.dispatchEvent(new MouseEvent('mouseover', evOpts));
          target.dispatchEvent(new MouseEvent('mouseenter', evOpts));
          await new Promise(r => setTimeout(r, 30));
          target.dispatchEvent(new PointerEvent('pointerdown', { ...evOpts, pointerId: 1 }));
          target.dispatchEvent(new MouseEvent('mousedown', evOpts));
          target.dispatchEvent(new PointerEvent('pointerup', { ...evOpts, pointerId: 1 }));
          target.dispatchEvent(new MouseEvent('mouseup', evOpts));
          target.dispatchEvent(new MouseEvent('click', evOpts));
          target.dispatchEvent(new Event('action', { bubbles: true }));
          return true;
        }
      }
    }

    // Strategy 3: Closure Library palette cells (borders, etc.)
    // These use goog-palette-cell with role="gridcell"
    const paletteCells = doc.querySelectorAll(
      '.goog-palette-cell, [role="gridcell"]'
    );
    for (const cell of paletteCells) {
      if (isDropdownButton(cell) || isInToolbar(cell)) continue;
      const rect = cell.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0) continue;

      const aria = (cell.getAttribute('aria-label') || '').toLowerCase();
      const tooltip = (cell.getAttribute('data-tooltip') || '').toLowerCase();
      const title = (cell.getAttribute('title') || '').toLowerCase();
      // Also check inner elements for labels
      const innerEl = cell.querySelector('[aria-label], [data-tooltip], [title]');
      const innerAria = innerEl ? (innerEl.getAttribute('aria-label') || '').toLowerCase() : '';
      const innerTooltip = innerEl ? (innerEl.getAttribute('data-tooltip') || '').toLowerCase() : '';

      for (const source of [aria, tooltip, title, innerAria, innerTooltip]) {
        if (!source) continue;
        if (source === itemLabelLower || source.includes(itemLabelLower)) {
          console.log(`[Vexcel] Selecting palette cell: "${itemLabel}" via "${source}"`, cell);
          // Closure palette cells need special activation:
          // 1. Mouseover to highlight
          // 2. Mousedown + mouseup + click
          // 3. Also dispatch 'action' event
          const cr = cell.getBoundingClientRect();
          const cx = cr.left + cr.width / 2;
          const cy = cr.top + cr.height / 2;
          const w = cell.ownerDocument.defaultView || window;
          const evOpts = { bubbles: true, cancelable: true, view: w, clientX: cx, clientY: cy };
          cell.dispatchEvent(new MouseEvent('mouseover', evOpts));
          cell.dispatchEvent(new MouseEvent('mouseenter', evOpts));
          await new Promise(r => setTimeout(r, 50));
          cell.dispatchEvent(new PointerEvent('pointerdown', { ...evOpts, pointerId: 1 }));
          cell.dispatchEvent(new MouseEvent('mousedown', evOpts));
          cell.dispatchEvent(new PointerEvent('pointerup', { ...evOpts, pointerId: 1 }));
          cell.dispatchEvent(new MouseEvent('mouseup', evOpts));
          cell.dispatchEvent(new MouseEvent('click', evOpts));
          // Fallback: dispatch action event used by Closure Library
          cell.dispatchEvent(new Event('action', { bubbles: true }));
          return true;
        }
      }
    }

    // Strategy 4: Search by class patterns common in Google Sheets dropdowns
    const dropdownItems = doc.querySelectorAll(
      '.goog-toolbar-menu-button-dropdown-item, [class*="border"], [class*="Border"]'
    );
    for (const el of dropdownItems) {
      if (isDropdownButton(el) || isInToolbar(el)) continue;
      const rect = el.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0) continue;

      const aria = (el.getAttribute('aria-label') || '').toLowerCase();
      const tooltip = (el.getAttribute('data-tooltip') || '').toLowerCase();
      const title = (el.getAttribute('title') || '').toLowerCase();

      for (const source of [aria, tooltip, title]) {
        if (!source) continue;
        if (source === itemLabelLower || source.includes(itemLabelLower)) {
          console.log(`[Vexcel] Selecting dropdown item (class): "${itemLabel}" via "${source}"`, el);
          simulateClick(el);
          return true;
        }
      }
    }

    console.warn(`[Vexcel] Dropdown item not found: "${itemLabel}"`);
    // Debug: log everything visible with aria-labels in the dropdown area
    const debugItems = [];
    for (const el of allAria) {
      const rect = el.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0 || rect.top < 40) continue;
      const aria = el.getAttribute('aria-label') || '';
      const tooltip = el.getAttribute('data-tooltip') || '';
      const tag = el.tagName.toLowerCase();
      if (aria || tooltip) debugItems.push(`${tag}[aria="${aria}"][tooltip="${tooltip}"]`);
    }
    console.log('[Vexcel] Visible dropdown elements:\n' + debugItems.join('\n'));

    closeDropdown(doc);
    return false;
  }

  function closeDropdown(doc) {
    doc.dispatchEvent(new KeyboardEvent('keydown', {
      key: 'Escape', code: 'Escape', keyCode: 27, bubbles: true
    }));
    setTimeout(() => {
      doc.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, clientX: 1, clientY: 1 }));
    }, 50);
  }

  return { clickButton, toggleButton, findToolbarButton, selectFromDropdown };
})();
