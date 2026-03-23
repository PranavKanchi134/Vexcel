// Google Sheets color picker interaction
// Opens toolbar color dropdowns and enables keyboard navigation of the palette

const VexcelColorPicker = (() => {
  const lastColorState = {
    'Fill color': null,
    'Text color': null
  };
  const buttonCache = new Map();
  const swatchIndexByToolbar = new Map();

  function topDoc() {
    try { return window.top.document; }
    catch (e) { return document; }
  }

  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  function simulateClick(el) {
    const rect = el.getBoundingClientRect();
    const x = rect.left + rect.width / 2;
    const y = rect.top + rect.height / 2;
    const win = el.ownerDocument.defaultView || window;
    const common = { bubbles: true, cancelable: true, view: win, clientX: x, clientY: y };
    el.dispatchEvent(new PointerEvent('pointerdown', { ...common, pointerId: 1 }));
    el.dispatchEvent(new MouseEvent('mousedown', common));
    el.dispatchEvent(new PointerEvent('pointerup', { ...common, pointerId: 1 }));
    el.dispatchEvent(new MouseEvent('mouseup', common));
    el.dispatchEvent(new MouseEvent('click', common));
  }

  let keyHandler = null;
  let highlightEl = null;
  let currentSwatches = [];
  let currentIdx = 0;

  /**
   * Open the color dropdown and enable keyboard navigation.
   * @param {string} toolbarLabel - aria-label of the color button ("Fill color" or "Text color")
   */
  async function openColorPicker(toolbarLabel) {
    const doc = topDoc();

    // Find and click the dropdown arrow next to the color button
    const button = getToolbarButton(toolbarLabel);
    if (!button) {
      console.warn(`[Vexcel] Color button not found: "${toolbarLabel}"`);
      return false;
    }

    // Click the dropdown arrow (small arrow next to the color button)
    const arrow = button.querySelector('[class*="dropdown"], [class*="arrow"]') ||
                  button.nextElementSibling;

    if (arrow && arrow.getBoundingClientRect().width > 0) {
      simulateClick(arrow);
    } else {
      simulateClick(button);
    }
    await waitForPalette(doc);

    // Gather all color swatches from the dropdown palette
    collectSwatches(doc, toolbarLabel);

    if (currentSwatches.length === 0) {
      console.warn('[Vexcel] No color swatches found in dropdown');
      return false;
    }

    // Start with the first swatch selected
    currentIdx = 0;
    showHighlight();
    attachKeyboard(doc);

    console.log(`[Vexcel] Color picker opened with ${currentSwatches.length} swatches. Use arrows + Enter.`);
    return true;
  }

  /**
   * Collect all visible color swatches from the open dropdown popup.
   * Finds the color palette TABLE (not just any palette cell), to avoid
   * grabbing border palette cells or toolbar icons.
   */
  function collectSwatches(doc, toolbarLabel = 'active') {
    currentSwatches = [];

    // Find the color palette table — it's a .goog-palette-table inside
    // a container that has MANY cells (color palette: 60-80+, border palette: ~10).
    // The color dropdown is the one with the most palette cells.
    const paletteTables = doc.querySelectorAll('table.goog-palette-table');
    let bestTable = null;
    let bestCount = 0;

    for (const table of paletteTables) {
      const r = table.getBoundingClientRect();
      // Must be visible
      if (r.width === 0 || r.height === 0) continue;
      const cellCount = table.querySelectorAll('.goog-palette-cell').length;
      // The color palette has 60-80+ cells; border palette has ~10
      if (cellCount > bestCount) {
        bestCount = cellCount;
        bestTable = table;
      }
    }

    if (bestTable && bestCount > 15) {
      const cells = bestTable.querySelectorAll('.goog-palette-cell');
      for (const cell of cells) {
        const rect = cell.getBoundingClientRect();
        if (rect.width < 3 || rect.height < 3) continue;
        currentSwatches.push(cell);
      }
      buildSwatchIndex(toolbarLabel, currentSwatches);
      console.log(`[Vexcel] Color picker: found palette table with ${currentSwatches.length} cells`);
      return;
    }

    // Fallback: find ANY visible popup container below toolbar, get its palette cells
    const popups = doc.querySelectorAll(
      '.goog-popup, .goog-menu, .goog-colorpalette'
    );
    for (const p of popups) {
      const r = p.getBoundingClientRect();
      if (r.width < 50 || r.height < 50) continue;
      const cells = p.querySelectorAll('.goog-palette-cell');
      if (cells.length > 15) {
        for (const cell of cells) {
          const rect = cell.getBoundingClientRect();
          if (rect.width < 3 || rect.height < 3) continue;
          currentSwatches.push(cell);
        }
        buildSwatchIndex(toolbarLabel, currentSwatches);
        console.log(`[Vexcel] Color picker: found popup with ${currentSwatches.length} cells`);
        return;
      }
    }

    // Last resort: collect all palette cells, but only from visible ones
    // that are clearly part of a color grid (many cells close together)
    const allCells = doc.querySelectorAll('.goog-palette-cell');
    const cellsByTop = {};
    for (const cell of allCells) {
      const rect = cell.getBoundingClientRect();
      if (rect.width < 3 || rect.height < 3) continue;
      const rowKey = Math.round(rect.top / 5) * 5;
      if (!cellsByTop[rowKey]) cellsByTop[rowKey] = [];
      cellsByTop[rowKey].push(cell);
    }
    // Find the group of rows with the most cells (the color palette)
    // Color palette rows have 10 cells each; toolbar items have fewer
    const rows = Object.values(cellsByTop).filter(r => r.length >= 8);
    rows.sort((a, b) => a[0].getBoundingClientRect().top - b[0].getBoundingClientRect().top);
    for (const row of rows) {
      currentSwatches.push(...row);
    }
    buildSwatchIndex(toolbarLabel, currentSwatches);
    console.log(`[Vexcel] Color picker: found ${currentSwatches.length} cells via row grouping`);
  }

  /**
   * Show a highlight ring around the currently selected swatch.
   */
  function showHighlight() {
    if (!highlightEl) {
      highlightEl = document.createElement('div');
      highlightEl.style.cssText =
        'position:fixed;z-index:2147483647;pointer-events:none;' +
        'border:3px solid #1a73e8;border-radius:3px;' +
        'box-shadow:0 0 0 2px rgba(26,115,232,.4), 0 0 8px rgba(26,115,232,.3);' +
        'transition:top .08s,left .08s;';
      (topDoc().body || document.body).appendChild(highlightEl);
    }

    const swatch = currentSwatches[currentIdx];
    if (!swatch) return;

    const rect = swatch.getBoundingClientRect();
    highlightEl.style.display = 'block';
    highlightEl.style.top = (rect.top - 3) + 'px';
    highlightEl.style.left = (rect.left - 3) + 'px';
    highlightEl.style.width = (rect.width + 2) + 'px';
    highlightEl.style.height = (rect.height + 2) + 'px';
  }

  function removeHighlight() {
    if (highlightEl) {
      highlightEl.style.display = 'none';
      highlightEl.remove();
      highlightEl = null;
    }
  }

  /**
   * Attach keyboard navigation to the open color palette.
   * Arrow keys move, Enter selects, Escape closes.
   */
  function attachKeyboard(doc) {
    detachKeyboard();

    // Figure out grid dimensions from swatch positions
    const cols = getGridCols();

    keyHandler = (e) => {
      if (e.key === 'Escape') {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        cleanup(doc);
        return;
      }

      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        const swatch = currentSwatches[currentIdx];
        if (swatch) {
          removeHighlight();
          detachKeyboard();
          // Click with full event sequence for Closure palette
          const cr = swatch.getBoundingClientRect();
          const cx = cr.left + cr.width / 2;
          const cy = cr.top + cr.height / 2;
          const win = swatch.ownerDocument.defaultView || window;
          const evOpts = { bubbles: true, cancelable: true, view: win, clientX: cx, clientY: cy };
          swatch.dispatchEvent(new MouseEvent('mouseover', evOpts));
          swatch.dispatchEvent(new MouseEvent('mouseenter', evOpts));
          swatch.dispatchEvent(new PointerEvent('pointerdown', { ...evOpts, pointerId: 1 }));
          swatch.dispatchEvent(new MouseEvent('mousedown', evOpts));
          swatch.dispatchEvent(new PointerEvent('pointerup', { ...evOpts, pointerId: 1 }));
          swatch.dispatchEvent(new MouseEvent('mouseup', evOpts));
          swatch.dispatchEvent(new MouseEvent('click', evOpts));
          swatch.dispatchEvent(new Event('action', { bubbles: true }));
          currentSwatches = [];
        }
        return;
      }

      let newIdx = currentIdx;

      if (e.key === 'ArrowRight' || e.key === 'l') {
        newIdx = Math.min(currentIdx + 1, currentSwatches.length - 1);
      } else if (e.key === 'ArrowLeft' || e.key === 'h') {
        newIdx = Math.max(currentIdx - 1, 0);
      } else if (e.key === 'ArrowDown' || e.key === 'j') {
        newIdx = Math.min(currentIdx + cols, currentSwatches.length - 1);
      } else if (e.key === 'ArrowUp' || e.key === 'k') {
        newIdx = Math.max(currentIdx - cols, 0);
      } else {
        return; // Don't intercept other keys
      }

      e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
      currentIdx = newIdx;
      showHighlight();
    };

    document.addEventListener('keydown', keyHandler, true);

    // Also close if user clicks outside
    const clickAway = (e) => {
      // If clicking on a swatch, let it go through
      if (currentSwatches.includes(e.target)) return;
      setTimeout(() => {
        cleanup(doc);
        document.removeEventListener('mousedown', clickAway, true);
      }, 100);
    };
    document.addEventListener('mousedown', clickAway, true);
  }

  function detachKeyboard() {
    if (keyHandler) {
      document.removeEventListener('keydown', keyHandler, true);
      keyHandler = null;
    }
  }

  function cleanup(doc) {
    removeHighlight();
    detachKeyboard();
    currentSwatches = [];
    currentIdx = 0;
    // Close dropdown
    doc.dispatchEvent(new KeyboardEvent('keydown', {
      key: 'Escape', code: 'Escape', keyCode: 27, bubbles: true
    }));
  }

  /**
   * Determine the number of columns in the palette grid
   * by looking at how many swatches share the same row (similar top position).
   * Uses a tolerance of 5px to account for sub-pixel differences.
   */
  function getGridCols() {
    if (currentSwatches.length < 2) return 1;

    // Group swatches by their approximate top position
    const firstTop = Math.round(currentSwatches[0].getBoundingClientRect().top);
    let cols = 0;
    for (const s of currentSwatches) {
      const top = Math.round(s.getBoundingClientRect().top);
      if (Math.abs(top - firstTop) <= 5) {
        cols++;
      } else {
        break;
      }
    }

    // Sanity check: Google Sheets color palette is typically 10 columns wide
    if (cols < 2) cols = 10;
    console.log(`[Vexcel] Color grid detected ${cols} columns`);
    return cols;
  }

  // Legacy API for direct hex application (still used by some paths)
  async function applyColor(toolbarLabel, hex) {
    if (hex === 'picker') {
      return openColorPicker(toolbarLabel);
    }

    const doc = topDoc();
    const button = getToolbarButton(toolbarLabel);
    if (!button) return false;

    const arrow = button.querySelector('[class*="dropdown"], [class*="arrow"]') ||
                  button.nextElementSibling;
    if (arrow && arrow.getBoundingClientRect().width > 0) {
      simulateClick(arrow);
    } else {
      simulateClick(button);
    }
    await waitForPalette(doc);

    let ok;
    if (hex === 'reset') ok = clickResetColor(doc);
    else if (hex === 'custom') ok = clickCustomColor(doc);
    else ok = await clickColorSwatch(doc, hex);

    if (ok && hex !== 'reset' && hex !== 'custom') {
      lastColorState[toolbarLabel] = hex;
    }
    return ok;
  }

  function applyLastUsedColor(toolbarLabel) {
    const button = getToolbarButton(toolbarLabel);
    if (!button || !lastColorState[toolbarLabel]) return false;
    simulateClick(button);
    return true;
  }

  async function clickColorSwatch(doc, targetHex) {
    const targetRgb = hexToRgb(targetHex);
    if (!targetRgb) return false;
    const normalizedHex = normalizeHex(targetHex);
    const cachedSwatch = swatchIndexByToolbar.get('active')?.get(normalizedHex);
    if (cachedSwatch && cachedSwatch.isConnected) {
      simulateClick(cachedSwatch);
      return true;
    }

    const allElements = doc.querySelectorAll('.goog-palette-cell, *');
    let bestMatch = null;
    let bestDist = Infinity;

    for (const el of allElements) {
      const rect = el.getBoundingClientRect();
      if (rect.width < 8 || rect.width > 50 || rect.height < 8 || rect.height > 50) continue;
      if (rect.top < 40) continue;
      const bg = getComputedStyle(el).backgroundColor;
      if (!bg || bg === 'transparent' || bg === 'rgba(0, 0, 0, 0)') continue;
      const rgb = parseRgb(bg);
      if (!rgb) continue;
      const dist = colorDistance(rgb, targetRgb);
      if (dist < bestDist) { bestDist = dist; bestMatch = el; }
    }

    if (bestMatch && bestDist < 50) {
      simulateClick(bestMatch);
      const activeIndex = swatchIndexByToolbar.get('active') || new Map();
      activeIndex.set(normalizedHex, bestMatch);
      swatchIndexByToolbar.set('active', activeIndex);
      return true;
    }

    doc.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', keyCode: 27, bubbles: true }));
    return false;
  }

  function clickCustomColor(doc) {
    const candidates = doc.querySelectorAll('[role="menuitem"], button, [class*="custom"]');
    for (const el of candidates) {
      const rect = el.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0 || rect.top < 40) continue;
      const text = (el.textContent || '').toLowerCase().trim();
      if (text.includes('custom')) { simulateClick(el); return true; }
    }
    return false;
  }

  function clickResetColor(doc) {
    const candidates = doc.querySelectorAll('[role="menuitem"], button, [class*="reset"], [class*="none"]');
    for (const el of candidates) {
      const rect = el.getBoundingClientRect();
      if (rect.width === 0 || rect.height === 0 || rect.top < 40) continue;
      const text = (el.textContent || '').toLowerCase().trim();
      if (text.includes('reset') || text.includes('none') || text.includes('default')) {
        simulateClick(el); return true;
      }
    }
    return false;
  }

  function hexToRgb(hex) {
    hex = hex.replace('#', '');
    if (hex.length === 3) hex = hex[0]+hex[0]+hex[1]+hex[1]+hex[2]+hex[2];
    if (hex.length !== 6) return null;
    return { r: parseInt(hex.substr(0,2),16), g: parseInt(hex.substr(2,2),16), b: parseInt(hex.substr(4,2),16) };
  }

  function parseRgb(str) {
    const m = str.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
    return m ? { r: +m[1], g: +m[2], b: +m[3] } : null;
  }

  function colorDistance(a, b) {
    return Math.sqrt((a.r-b.r)**2 + (a.g-b.g)**2 + (a.b-b.b)**2);
  }

  function normalizeHex(hex) {
    const clean = hex.replace('#', '').toLowerCase();
    if (clean.length === 3) {
      return `#${clean[0]}${clean[0]}${clean[1]}${clean[1]}${clean[2]}${clean[2]}`;
    }
    return `#${clean}`;
  }

  function getToolbarButton(label) {
    const cached = buttonCache.get(label);
    if (cached && cached.isConnected) return cached;
    const button = VexcelToolbar.findToolbarButton(label);
    if (button) buttonCache.set(label, button);
    return button;
  }

  function buildSwatchIndex(toolbarLabel, swatches) {
    const index = new Map();
    for (const swatch of swatches) {
      const bg = getComputedStyle(swatch).backgroundColor;
      const rgb = parseRgb(bg);
      if (!rgb) continue;
      const hex = `#${rgb.r.toString(16).padStart(2, '0')}${rgb.g.toString(16).padStart(2, '0')}${rgb.b.toString(16).padStart(2, '0')}`;
      index.set(hex, swatch);
    }
    swatchIndexByToolbar.set(toolbarLabel, index);
    swatchIndexByToolbar.set('active', index);
  }

  async function waitForPalette(doc) {
    for (let attempt = 0; attempt < 8; attempt++) {
      const visiblePalette = doc.querySelector('table.goog-palette-table .goog-palette-cell, .goog-colorpalette .goog-palette-cell');
      if (visiblePalette) {
        const rect = visiblePalette.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0) return true;
      }
      await sleep(25);
    }
    return false;
  }

  return { applyColor, applyLastUsedColor, openColorPicker };
})();
