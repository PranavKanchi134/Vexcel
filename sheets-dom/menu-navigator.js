// Google Sheets menu navigation - find and click menu items by path

const VexcelMenuNavigator = (() => {
  const commandStrategyCache = new Map();
  const menuPathCache = new Map();
  let lastToolFinderQuery = '';
  let lastToolFinderInput = null;

  function now() {
    return (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
  }

  // Always use the top-level document for menu queries.
  // Menus live in the top frame; iframes only contain the grid.
  function topDoc() {
    try { return window.top.document; }
    catch (e) { return document; }
  }

  // Dispatches a full pointer / mouse event sequence so that Closure
  // Library UI widgets (menus, buttons) register the interaction.
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

  /**
   * Navigate a menu path and click the final item.
   * @param {string[]} path - Array of menu labels, e.g. ['Format', 'Number', 'Date']
   * @returns {Promise<boolean>} - Whether the action succeeded
   *
   * Smart fallback: If a middle item can't be found, tries collapsing
   * remaining path segments (e.g., ['Edit', 'Fill', 'Fill down'] will
   * try 'Fill' first, and if not found, try 'Fill down' directly).
   */
  async function clickMenuPath(path) {
    if (!path || path.length === 0) return false;

    try {
      const cachedTopMenu = getCachedTopMenu(path[0]);
      // Step 1: Click the top-level menu
      const topMenu = cachedTopMenu || findTopMenu(path[0]);
      if (!topMenu) {
        console.warn(`[Vexcel] Top menu not found: "${path[0]}"`);
        return false;
      }
      if (!cachedTopMenu) VexcelSelectorCache.set(`top-menu:${path[0]}`, topMenu);

      console.log(`[Vexcel] Clicking top menu: "${path[0]}"`, topMenu);
      simulateClick(topMenu);
      await waitForMenuSurface(topMenu.ownerDocument || topDoc(), 100);

      // Step 2: Navigate through each level of the path
      let i = 1;
      while (i < path.length) {
        const label = path[i];
        const isLast = (i === path.length - 1);
        let item = await findMenuItem(label);

        if (!item && !isLast) {
          // Fallback: try combining this segment with the next one
          // e.g., path = ['Edit', 'Fill', 'Fill down'] — if 'Fill' not found,
          // try 'Fill down' directly in the current menu
          const combined = path.slice(i + 1).join(' ');
          console.log(`[Vexcel] Submenu "${label}" not found, trying direct: "${combined}"`);
          item = await findMenuItem(combined);
          if (item) {
            // Found the final item directly — click it and done
            console.log(`[Vexcel] Found direct item: "${combined}"`, item);
            simulateClick(item);
            return true;
          }

          // Also try each remaining segment individually
          for (let j = i + 1; j < path.length; j++) {
            item = await findMenuItem(path[j]);
            if (item) {
              if (j === path.length - 1) {
                // This is the final item — click it
                console.log(`[Vexcel] Found skipped-to item: "${path[j]}"`, item);
                simulateClick(item);
                return true;
              } else {
                // Middle item found — hover and continue from here
                console.log(`[Vexcel] Found skipped-to submenu: "${path[j]}"`, item);
                hoverElement(item);
                await waitForMenuSurface(item.ownerDocument || topDoc(), 100);
                i = j + 1;
                break;
              }
            }
          }

          // If we still haven't found anything, fail
          if (!item) {
            console.warn(`[Vexcel] Menu item not found: "${label}" (path: ${path.join(' > ')})`);
            logVisibleMenuItems();
            closeMenus();
            return false;
          }
          continue; // We've already advanced i in the inner loop
        }

        if (!item) {
          console.warn(`[Vexcel] Menu item not found: "${label}" (path: ${path.join(' > ')})`);
          logVisibleMenuItems();
          closeMenus();
          return false;
        }

        console.log(`[Vexcel] Found menu item: "${label}"`, item);

        if (!isLast) {
          // Hover to open submenu, then wait for it to render
          hoverElement(item);
          await waitForMenuSurface(item.ownerDocument || topDoc(), 100);
        } else {
          // Final item — click it
          simulateClick(item);
        }

        i++;
      }

      return true;
    } catch (err) {
      console.error('[Vexcel] Menu navigation error:', err);
      closeMenus();
      return false;
    }
  }

  /**
   * Find a top-level menu button by label.
   * Tries ID selectors first, then aria-label, then text content.
   */
  function findTopMenu(label) {
    const cached = getCachedTopMenu(label);
    if (cached) return cached;

    const idMap = {
      'File':       '#docs-file-menu',
      'Edit':       '#docs-edit-menu',
      'View':       '#docs-view-menu',
      'Insert':     '#docs-insert-menu',
      'Format':     '#docs-format-menu',
      'Data':       '#docs-data-menu',
      'Tools':      '#docs-tools-menu',
      'Extensions': '#docs-extensions-menu',
      'Help':       '#docs-help-menu'
    };

    const doc = topDoc();

    // 1. Try known ID selector
    if (idMap[label]) {
      const el = doc.querySelector(idMap[label]);
      if (el) {
        VexcelSelectorCache.set(`top-menu:${label}`, el);
        return el;
      }
    }

    // 2. Try aria-label attribute (exact, then contains)
    const byAria = doc.querySelector(`[aria-label="${label}"]`);
    if (byAria) {
      VexcelSelectorCache.set(`top-menu:${label}`, byAria);
      return byAria;
    }

    // 3. Try matching menu bar items by text content
    const candidates = doc.querySelectorAll(
      '[role="menubar"] [role="menuitem"], [role="menubar"] *, .menu-button, [id*="docs-"][id*="-menu"]'
    );
    for (const el of candidates) {
      const text = getLabelText(el);
      if (text === label || text.toLowerCase() === label.toLowerCase()) {
        VexcelSelectorCache.set(`top-menu:${label}`, el);
        return el;
      }
    }

    return null;
  }

  /**
   * Find a visible menu item matching the given label.
   * Retries up to 15 times with 80ms gaps (total ~1200ms).
   */
  async function findMenuItem(label) {
    const doc = topDoc();
    for (let attempt = 0; attempt < 15; attempt++) {
      const roots = getVisibleMenuRoots(doc);
      const searchRoots = roots.length ? roots : [doc];
      for (const root of searchRoots) {
        const candidates = root.querySelectorAll(
          '[role="menuitem"], [role="menuitemcheckbox"], [role="menuitemradio"], [role="option"]'
        );

        for (const item of candidates) {
          if (!isVisible(item)) continue;
          if (matchesLabel(item, label)) return item;
        }

        const listItems = root.querySelectorAll('li[class*="goog-menuitem"], li[class*="menu-item"]');
        for (const item of listItems) {
          if (!isVisible(item)) continue;
          if (matchesLabel(item, label)) return item;
        }
      }

      await sleep(40);
    }
    return null;
  }

  /**
   * Check if a DOM element's label matches the target label.
   * Uses multiple strategies for maximum flexibility.
   *
   * Strategy order (most specific → least specific):
   * 1. Exact match on aria-label or full text
   * 2. Starts-with match
   * 3. Contains match (item text contains our label)
   * 4. Reverse contains (our label contains item text, min length 3)
   * 5. Child span text matching
   * 6. All-words match (every word in our label appears in item text)
   */
  function matchesLabel(el, label) {
    const labelLower = label.toLowerCase().trim();
    const labelWords = labelLower.split(/\s+/);

    // SKIP top-level menu bar buttons — they should never match as submenu items
    if (el.classList.contains('menu-button') || el.id?.startsWith('docs-') && el.id?.endsWith('-menu')) {
      return false;
    }

    // Get the primary label text from the menu item content span
    // (not the full textContent which may include shortcut hints, subtext)
    const ariaLabel = (el.getAttribute('aria-label') || '').trim().toLowerCase();
    const contentSpan = el.querySelector('.goog-menuitem-content');
    const primaryText = contentSpan
      ? (contentSpan.textContent || '').replace(/\s+/g, ' ').trim().toLowerCase()
      : (el.textContent || '').replace(/\s+/g, ' ').trim().toLowerCase();

    // Strategy 1: Exact match
    for (const text of [ariaLabel, primaryText]) {
      if (!text) continue;
      if (text === labelLower) return true;
    }

    // Strategy 2: Starts-with match
    for (const text of [ariaLabel, primaryText]) {
      if (!text) continue;
      if (text.startsWith(labelLower)) return true;
    }

    // Strategy 3: Contains (item text contains our search label)
    for (const text of [ariaLabel, primaryText]) {
      if (!text) continue;
      if (text.includes(labelLower)) return true;
    }

    // Strategy 4: All-words match (every word in our label appears in item text)
    if (labelWords.length >= 2) {
      if (primaryText && labelWords.every(w => primaryText.includes(w))) return true;
      if (ariaLabel && labelWords.every(w => ariaLabel.includes(w))) return true;
    }

    // Strategy 5: Check child spans individually
    const spans = el.querySelectorAll('span, div');
    for (const span of spans) {
      const spanText = (span.textContent || '').replace(/\s+/g, ' ').trim().toLowerCase();
      if (!spanText || spanText.length > 80) continue;
      if (spanText === labelLower) return true;
      if (spanText.startsWith(labelLower)) return true;
      if (spanText.includes(labelLower)) return true;
    }

    // Strategy 6: Reverse contains — ONLY when item text is substantial
    // (at least half the length of our label, and at least 5 chars)
    // This prevents "insert" matching "Insert Column left"
    for (const text of [ariaLabel, primaryText]) {
      if (!text) continue;
      if (text.length >= 5 && text.length >= labelLower.length * 0.5 && labelLower.includes(text)) {
        return true;
      }
    }

    return false;
  }

  function getLabelText(el) {
    return (el.textContent || '').replace(/\s+/g, ' ').trim();
  }

  /**
   * Hover over an element to trigger submenu opening.
   */
  function hoverElement(el) {
    const rect = el.getBoundingClientRect();
    const x = rect.left + rect.width / 2;
    const y = rect.top + rect.height / 2;
    const win = el.ownerDocument.defaultView || window;
    const opts = { bubbles: true, clientX: x, clientY: y, view: win };

    el.dispatchEvent(new PointerEvent('pointerenter', { ...opts, pointerId: 1 }));
    el.dispatchEvent(new MouseEvent('mouseenter', opts));
    el.dispatchEvent(new MouseEvent('mouseover', opts));
    el.dispatchEvent(new MouseEvent('mousemove', opts));

    // Simulate moving from the left edge to center (helps some menu implementations)
    el.dispatchEvent(new MouseEvent('mousemove', {
      ...opts, clientX: rect.left + 5, clientY: y
    }));
    el.dispatchEvent(new MouseEvent('mousemove', {
      ...opts, clientX: x, clientY: y
    }));

    if (el.focus) el.focus();
  }

  /**
   * Close any open menus.
   */
  function closeMenus() {
    const doc = topDoc();
    // Press Escape multiple times to close nested menus
    for (let i = 0; i < 3; i++) {
      doc.dispatchEvent(new KeyboardEvent('keydown', {
        key: 'Escape', code: 'Escape', keyCode: 27, bubbles: true
      }));
    }
    // Fallback: click outside
    setTimeout(() => {
      doc.body.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, clientX: 1, clientY: 1 }));
      doc.body.dispatchEvent(new MouseEvent('click',     { bubbles: true, clientX: 1, clientY: 1 }));
    }, 50);
  }

  /**
   * Debug helper: log all currently visible menu items.
   */
  function logVisibleMenuItems() {
    const doc = topDoc();
    const items = doc.querySelectorAll(
      '[role="menuitem"], [role="menuitemcheckbox"], [role="menuitemradio"], [role="option"]'
    );
    const visible = [];
    for (const item of items) {
      if (!isVisible(item)) continue;
      const text = (item.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 60);
      const aria = (item.getAttribute('aria-label') || '').slice(0, 60);
      visible.push(aria || text);
    }
    // Use join so it shows inline in console, not as collapsed Array(N)
    console.log('[Vexcel] Visible menu items (' + visible.length + '):\n' + visible.join('\n'));
  }

  function isVisible(el) {
    if (!el) return false;
    const rect = el.getBoundingClientRect();
    if (rect.width === 0 || rect.height === 0) return false;
    const style = getComputedStyle(el);
    return style.display !== 'none' && style.visibility !== 'hidden';
  }

  /**
   * Fallback: Use Google Sheets' built-in "Search the menus" tool finder.
   * This searches all menu commands by name and executes them.
   * @param {string} query - The command to search for (e.g., "Insert row above")
   * @returns {Promise<boolean>}
   */
  async function useToolFinder(query, allowlist = []) {
    const startAt = now();
    const doc = topDoc();
    const queryLower = normalizeSearchText(query);
    const queryWords = queryLower.split(/\s+/).filter(Boolean);
    const normalizedAllowlist = Array.isArray(allowlist) ? allowlist.map(normalizeSearchText) : [];

    let searchInput = findToolFinderInput(doc);
    if (!searchInput) {
      closeMenus();
      await sleep(24);
    }

    searchInput = searchInput || await openToolFinderInput(doc);
    if (!searchInput) {
      console.warn('[Vexcel] Tool finder search input not found');
      closeMenus();
      return false;
    }

    searchInput.focus();
    await typeIntoInput(searchInput, query, doc);
    lastToolFinderQuery = query;

    await sleep(60);

    // Find the search input's position so we only look at results BELOW it
    const allInputs = doc.querySelectorAll('input[type="text"]');
    let searchY = 80; // default
    for (const inp of allInputs) {
      const ir = inp.getBoundingClientRect();
      if (ir.width > 100 && ir.height > 0 && isVisible(inp)) {
        searchY = ir.bottom;
        break;
      }
    }
    console.log(`[Vexcel] Tool finder: searching for "${query}", results expected below Y=${searchY}`);

    const minScore = queryWords.length >= 2 ? 60 : 40;

    for (let attempt = 0; attempt < 6; attempt++) {
      const results = getToolFinderResults(doc, searchY);
      const best = findBestToolFinderResult(results, queryLower, queryWords, normalizedAllowlist);

      if (best && best.score >= minScore) {
        console.log(`[Vexcel] Tool finder: clicking best match "${best.label}" (score=${best.score})`);
        simulateClick(best.result);
        menuPathCache.set(`tool:${queryLower}`, best.label);
        return true;
      }

      await sleep(35);
    }

    const finalResults = getToolFinderResults(doc, searchY);
    const preview = finalResults
      .map(result => getToolFinderResultLabel(result))
      .filter(label => !normalizedAllowlist.length || matchesAllowlist(label, normalizedAllowlist))
      .filter(Boolean)
      .slice(0, 8);
    console.warn('[Vexcel] Tool finder: no good results for "' + query + '". Saw: ' + preview.join(' | '));
    closeMenus();
    return false;
  }

  async function openToolFinderInput(doc) {
    let searchInput = findToolFinderInput(doc);
    if (searchInput) return searchInput;

    triggerToolFinderShortcut(doc);
    for (let attempt = 0; attempt < 6; attempt++) {
      await sleep(25);
      searchInput = findToolFinderInput(doc);
      if (searchInput) return searchInput;
    }

    // Fallback: open from the Help menu if the shortcut path did not render.
    const helpMenu = findTopMenu('Help');
    if (!helpMenu) return null;
    simulateClick(helpMenu);
    for (let attempt = 0; attempt < 8; attempt++) {
      await sleep(30);
      searchInput = findToolFinderInput(doc);
      if (searchInput) return searchInput;
    }

    return null;
  }

  function findToolFinderInput(doc) {
    if (lastToolFinderInput && lastToolFinderInput.isConnected && isVisible(lastToolFinderInput)) {
      return lastToolFinderInput;
    }
    const input = doc.querySelector(
      '.docs-tool-finder-input input, ' +
      '[aria-label="Search the menus"], ' +
      '[aria-label*="Search the menus"] input, ' +
      'input[type="text"][aria-label*="menu"]'
    );
    lastToolFinderInput = input || null;
    return input;
  }

  function triggerToolFinderShortcut(doc) {
    const target = doc.activeElement || doc.body;
    const evt = {
      key: '/',
      code: 'Slash',
      keyCode: 191,
      which: 191,
      altKey: true,
      bubbles: true,
      cancelable: true
    };
    target.dispatchEvent(new KeyboardEvent('keydown', evt));
    target.dispatchEvent(new KeyboardEvent('keyup', evt));
  }

  async function runCommand(query, path, options = {}) {
    const prefer = options.prefer || 'toolFinder';
    const cacheKey = options.commandKey || `${query || ''}::${(path || []).join('>')}`;
    const cached = commandStrategyCache.get(cacheKey);
    const allowlist = options.allowlist || [];
    const startedAt = now();
    const strategies = prefer === 'menu'
      ? [
          { name: 'menu', run: () => clickMenuPath(path) },
          { name: 'toolFinder', run: () => useToolFinder(query, allowlist) }
        ]
      : [
          { name: 'toolFinder', run: () => useToolFinder(query, allowlist) },
          { name: 'menu', run: () => clickMenuPath(path) }
        ];

    if (cached) {
      strategies.sort((a, b) => {
        if (a.name === cached && b.name !== cached) return -1;
        if (b.name === cached && a.name !== cached) return 1;
        return 0;
      });
    }

    for (const strategy of strategies) {
      if (strategy.name === 'toolFinder' && !query) continue;
      if (strategy.name === 'menu' && (!path || path.length === 0)) continue;
      const ok = await strategy.run();
      if (ok) {
        commandStrategyCache.set(cacheKey, strategy.name);
        return {
          ok: true,
          strategy: strategy.name,
          verified: true,
          durationMs: Math.round(now() - startedAt)
        };
      }
    }

    return {
      ok: false,
      reason: `command not found for ${query || (path || []).join(' > ')}`,
      durationMs: Math.round(now() - startedAt)
    };
  }

  function getToolFinderResults(doc, searchY) {
    const results = doc.querySelectorAll(
      '[role="menuitem"], [role="option"], .goog-menuitem'
    );

    return Array.from(results).filter(result => {
      if (!isVisible(result)) return false;
      const rect = result.getBoundingClientRect();
      if (rect.top < searchY) return false;
      if (result.closest('[role="menubar"]')) return false;
      if (result.id?.startsWith('docs-') && result.id?.endsWith('-menu')) return false;
      return true;
    });
  }

  function findBestToolFinderResult(results, queryLower, queryWords, allowlist = []) {
    const allowedResults = allowlist.length
      ? results.filter(result => matchesAllowlist(getToolFinderResultLabel(result), allowlist))
      : results;
    const exact = allowedResults.find(result => getToolFinderResultLabel(result) === queryLower);
    if (exact) {
      return {
        result: exact,
        score: 100,
        label: getToolFinderResultLabel(exact)
      };
    }

    let best = null;

    for (const result of allowedResults) {
      const score = scoreToolFinderResult(result, queryLower, queryWords);
      if (!best || score > best.score) {
        best = {
          result,
          score,
          label: getToolFinderResultLabel(result)
        };
      }
    }

    return best;
  }

  function scoreToolFinderResult(result, queryLower, queryWords) {
    const sources = [
      normalizeSearchText(result.getAttribute('aria-label') || ''),
      normalizeSearchText(result.textContent || '')
    ].filter(Boolean);

    let best = 0;

    for (const source of sources) {
      if (source === queryLower) return 100;
      if (source.startsWith(queryLower)) best = Math.max(best, 90);
      if (source.includes(queryLower)) best = Math.max(best, 80);

      if (queryWords.length > 1) {
        const hasAllWords = queryWords.every(word => source.includes(word));
        if (hasAllWords) best = Math.max(best, 70);
      }

      const overlap = queryWords.filter(word => source.includes(word)).length;
      if (queryWords.length === 1 && overlap === 1) best = Math.max(best, 50);
      if (queryWords.length > 1 && overlap === queryWords.length) best = Math.max(best, 65);
    }

    return best;
  }

  function getToolFinderResultLabel(result) {
    return normalizeSearchText(
      result.getAttribute('aria-label') ||
      result.textContent ||
      ''
    );
  }

  function matchesAllowlist(label, allowlist) {
    const normalizedLabel = normalizeSearchText(label);
    if (!normalizedLabel) return false;

    return allowlist.some(entry => {
      const normalizedEntry = normalizeSearchText(entry);
      if (!normalizedEntry) return false;
      if (normalizedLabel === normalizedEntry) return true;
      if (normalizedLabel.startsWith(normalizedEntry)) return true;
      if (normalizedLabel.includes(normalizedEntry)) return true;

      const words = normalizedEntry.split(/\s+/).filter(Boolean);
      return words.length > 1 && words.every(word => normalizedLabel.includes(word));
    });
  }

  function normalizeSearchText(text) {
    return (text || '').replace(/\s+/g, ' ').trim().toLowerCase();
  }

  function getCachedTopMenu(label) {
    return VexcelSelectorCache.get(`top-menu:${label}`, () => null);
  }

  function getVisibleMenuRoots(doc) {
    return Array.from(doc.querySelectorAll('[role="menu"]')).filter(isVisible);
  }

  async function waitForMenuSurface(doc, timeoutMs = 100) {
    const startedAt = now();
    while (now() - startedAt < timeoutMs) {
      if (getVisibleMenuRoots(doc).length > 0) return true;
      await sleep(16);
    }
    return false;
  }

  /**
   * Type text into an input using multiple strategies for maximum compatibility.
   */
  async function typeIntoInput(input, text, doc) {
    input.focus();

    // Clear existing text first
    const nativeSet = Object.getOwnPropertyDescriptor(
      HTMLInputElement.prototype, 'value'
    );

    // Strategy 1: Set value + dispatch input/change events
    if (nativeSet && nativeSet.set) {
      nativeSet.set.call(input, text);
    } else {
      input.value = text;
    }
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new InputEvent('input', {
      bubbles: true, inputType: 'insertText', data: text
    }));
    input.dispatchEvent(new Event('change', { bubbles: true }));

    // Strategy 2: Type character by character with full key event sequence
    // This is what Google Sheets' Closure Library often listens for
    await sleep(50);
    // Clear and re-type if the value didn't stick
    if (input.value !== text) {
      if (nativeSet && nativeSet.set) {
        nativeSet.set.call(input, '');
      } else {
        input.value = '';
      }
      input.dispatchEvent(new Event('input', { bubbles: true }));

      for (const char of text) {
        input.dispatchEvent(new KeyboardEvent('keydown', {
          key: char, code: `Key${char.toUpperCase()}`, keyCode: char.charCodeAt(0),
          bubbles: true, cancelable: true
        }));
        input.dispatchEvent(new KeyboardEvent('keypress', {
          key: char, code: `Key${char.toUpperCase()}`, keyCode: char.charCodeAt(0),
          charCode: char.charCodeAt(0), bubbles: true
        }));
        // Append the character
        if (nativeSet && nativeSet.set) {
          nativeSet.set.call(input, input.value + char);
        } else {
          input.value += char;
        }
        input.dispatchEvent(new InputEvent('input', {
          bubbles: true, inputType: 'insertText', data: char
        }));
        input.dispatchEvent(new KeyboardEvent('keyup', {
          key: char, code: `Key${char.toUpperCase()}`, keyCode: char.charCodeAt(0),
          bubbles: true
        }));
        await sleep(20);
      }
    }

    console.log(`[Vexcel] Tool finder: typed "${input.value}" into search box`);
  }

  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  return { clickMenuPath, findTopMenu, findMenuItem, closeMenus, useToolFinder, runCommand };
})();
