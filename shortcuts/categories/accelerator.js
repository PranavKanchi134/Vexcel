// Accelerator key definitions — Excel 365 Windows ribbon KeyTips for Google Sheets
//
// Matches Excel for Windows ribbon Alt-key accelerators:
//   Alt+H = Home    Alt+N = Insert    Alt+M = Formulas
//   Alt+A = Data    Alt+W = View
//
// Tree structure:
//   mode: 'toolbar' — children render as badges on actual toolbar buttons
//   mode: 'list'    — children render as a compact key-hint list panel
//   toolbar: string — aria-label of the toolbar button this leaf activates
//   action: fn      — menu or custom action this leaf runs (also used as fallback for toolbar)
//   children: {}    — branch node; pressing this key goes deeper
//
// IMPORTANT: All menu paths use the ACTUAL Google Sheets menu structure, which differs
// significantly from Excel. Key differences:
//   - Insert row: Insert > Rows > Insert 1 row above
//   - Insert column: Insert > Column > Insert Column left  (singular "Column")
//   - Delete row/column: Edit > Delete > row/column
//   - Fill operations: Edit > Fill > Fill down (3-level path)
//   - Bold/Italic etc: Format > Text > Bold
//   - Sort: Data > Sort range (not "Sort sheet")
//   - Filter: Data > Filter (not "Create a filter")
//   - Hide/unhide/resize rows/cols: Context menu only (no top-level menu path)

const VexcelAcceleratorKeys = (() => {

  // Known Google Sheets toolbar button aria-labels
  const TB = {
    bold:        'Bold',
    italic:      'Italic',
    underline:   'Underline',
    strike:      'Strikethrough',
    fillColor:   'Fill color',
    fontColor:   'Text color',
    alignLeft:   'Left align',
    alignCenter: 'Center align',
    alignRight:  'Right align',
    alignTop:    'Top align',
    alignMiddle: 'Middle align',
    alignBottom: 'Bottom align',
    merge:       'Merge cells',
    wrap:        'Text wrapping',
    link:        'Insert link',
    borders:     'Borders',
    paintFormat: 'Paint format',
    filter:      'Create a filter',  // toolbar aria-label (may differ from menu text)
  };

  // Menu path helper — tries menu navigation, falls back to tool finder
  const menu = (...path) => async () => {
    const ok = await VexcelMenuNavigator.clickMenuPath(path);
    if (!ok) {
      // Use the last segment as the tool finder query
      const query = path[path.length - 1];
      console.log(`[Vexcel] Menu path failed, trying tool finder: "${query}"`);
      return VexcelMenuNavigator.useToolFinder(query);
    }
    return ok;
  };

  // Color picker: opens the real Google Sheets color palette with keyboard navigation
  // Arrow keys to move between swatches, Enter to select, Escape to cancel
  function colorPickerAction(toolbarLabel) {
    return () => VexcelColorPicker.openColorPicker(toolbarLabel);
  }

  // Border action: click the borders dropdown and select by position index
  // Google Sheets border palette is a grid with items in this order:
  //   0: All borders     1: Inner borders    2: Horizontal borders
  //   3: Vertical borders 4: Outer borders   5: Clear borders
  //   (second row if present)
  //   6: Top border      7: Bottom border    8: Left border
  //   9: Right border
  // We also try aria-label/data-tooltip matching first.
  const borderPositionMap = {
    'All borders':     0,
    'Inner borders':   1,
    'Horizontal borders': 2,
    'Vertical borders': 3,
    'Outside borders': 4,
    'Outer borders':   4,
    'Clear borders':   5,
    'No borders':      5,
    'Top border':      6,
    'Bottom border':   7,
    'Left border':     8,
    'Right border':    9,
  };

  function borderAction(itemLabel) {
    return async () => {
      const doc = (() => { try { return window.top.document; } catch(e) { return document; } })();

      // Click the borders dropdown button
      const button = VexcelToolbar.findToolbarButton(TB.borders);
      if (!button) {
        console.warn('[Vexcel] Borders button not found');
        return VexcelMenuNavigator.useToolFinder(itemLabel);
      }

      // Simulate click on the button
      const rect = button.getBoundingClientRect();
      const x = rect.left + rect.width / 2;
      const y = rect.top + rect.height / 2;
      const win = button.ownerDocument.defaultView || window;
      const common = { bubbles: true, cancelable: true, view: win, clientX: x, clientY: y };
      button.dispatchEvent(new PointerEvent('pointerdown', { ...common, pointerId: 1 }));
      button.dispatchEvent(new MouseEvent('mousedown', common));
      button.dispatchEvent(new PointerEvent('pointerup', { ...common, pointerId: 1 }));
      button.dispatchEvent(new MouseEvent('mouseup', common));
      button.dispatchEvent(new MouseEvent('click', common));

      await new Promise(r => setTimeout(r, 400));

      const itemLower = itemLabel.toLowerCase();

      // Strategy 1: Try to find by aria-label or data-tooltip (exact or contains)
      const candidates = doc.querySelectorAll('[aria-label], [data-tooltip], [title]');
      for (const el of candidates) {
        if (el === button || button.contains(el) || el.contains(button)) continue;
        const r = el.getBoundingClientRect();
        if (r.width === 0 || r.height === 0 || r.top < 80) continue;
        const aria = (el.getAttribute('aria-label') || '').toLowerCase();
        const tooltip = (el.getAttribute('data-tooltip') || '').toLowerCase();
        const title = (el.getAttribute('title') || '').toLowerCase();
        for (const s of [aria, tooltip, title]) {
          if (s && (s === itemLower || s.includes(itemLower))) {
            console.log(`[Vexcel] Border: clicking "${itemLabel}" via label "${s}"`);
            clickPaletteCell(el);
            await new Promise(r => setTimeout(r, 100));
            doc.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', keyCode: 27, bubbles: true, cancelable: true }));
            return true;
          }
        }
      }

      // Strategy 2: Select by position index in the palette grid
      const paletteCells = doc.querySelectorAll('.goog-palette-cell');
      const visibleCells = [];
      for (const cell of paletteCells) {
        const r = cell.getBoundingClientRect();
        if (r.width > 0 && r.height > 0 && r.top > 60) visibleCells.push(cell);
      }

      const idx = borderPositionMap[itemLabel];
      if (idx !== undefined && idx < visibleCells.length) {
        console.log(`[Vexcel] Border: clicking position ${idx} for "${itemLabel}"`);
        clickPaletteCell(visibleCells[idx]);
        await new Promise(r => setTimeout(r, 100));
        doc.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', keyCode: 27, bubbles: true, cancelable: true }));
        return true;
      }

      // Strategy 3: For "Border color" / "Border style", look for menu items below the palette
      if (itemLower.includes('color') || itemLower.includes('style')) {
        const menuItems = doc.querySelectorAll('[role="menuitem"], .goog-menuitem');
        for (const mi of menuItems) {
          const r = mi.getBoundingClientRect();
          if (r.width === 0 || r.height === 0 || r.top < 80) continue;
          const text = (mi.textContent || '').replace(/\s+/g, ' ').trim().toLowerCase();
          if (text.includes(itemLower)) {
            console.log(`[Vexcel] Border: clicking menu item "${text}"`);
            clickPaletteCell(mi);
            // Don't close for color/style - they open sub-menus
            return true;
          }
        }
      }

      console.warn(`[Vexcel] Border item not found: "${itemLabel}"`);
      // Close the dropdown
      doc.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', keyCode: 27, bubbles: true }));
      return VexcelMenuNavigator.useToolFinder(itemLabel);
    };
  }

  function clickPaletteCell(el) {
    const r = el.getBoundingClientRect();
    const cx = r.left + r.width / 2;
    const cy = r.top + r.height / 2;
    const win = el.ownerDocument.defaultView || window;
    const opts = { bubbles: true, cancelable: true, view: win, clientX: cx, clientY: cy };
    el.dispatchEvent(new MouseEvent('mouseover', opts));
    el.dispatchEvent(new MouseEvent('mouseenter', opts));
    el.dispatchEvent(new PointerEvent('pointerdown', { ...opts, pointerId: 1 }));
    el.dispatchEvent(new MouseEvent('mousedown', opts));
    el.dispatchEvent(new PointerEvent('pointerup', { ...opts, pointerId: 1 }));
    el.dispatchEvent(new MouseEvent('mouseup', opts));
    el.dispatchEvent(new MouseEvent('click', opts));
    el.dispatchEvent(new Event('action', { bubbles: true }));
  }

  // Toolbar click with menu fallback, then tool finder fallback
  function tbOrMenu(toolbarLabel, ...menuPath) {
    return {
      toolbar: toolbarLabel,
      action: async () => {
        const ok = await VexcelMenuNavigator.clickMenuPath(menuPath);
        if (!ok) {
          const query = menuPath[menuPath.length - 1];
          console.log(`[Vexcel] Menu fallback failed, trying tool finder: "${query}"`);
          return VexcelMenuNavigator.useToolFinder(query);
        }
        return ok;
      },
    };
  }

  const tree = {

    // ═══════════════════════════════════════════════════════════
    // H — Home (Excel: Alt+H)
    // ═══════════════════════════════════════════════════════════
    'h': {
      label: 'Home', hint: 'Format, Fill, Borders',
      mode: 'toolbar',
      children: {
        // ── Direct toolbar buttons (with menu fallbacks) ─────────
        '1': { label: 'Bold',          ...tbOrMenu(TB.bold,      'Format', 'Text', 'Bold') },
        '2': { label: 'Italic',        ...tbOrMenu(TB.italic,    'Format', 'Text', 'Italic') },
        '3': { label: 'Underline',     ...tbOrMenu(TB.underline, 'Format', 'Text', 'Underline') },
        '5': { label: 'Strikethrough', ...tbOrMenu(TB.strike,    'Format', 'Text', 'Strikethrough') },
        'p': { label: 'Paint Format',  toolbar: TB.paintFormat },
        'w': { label: 'Wrap Text',     ...tbOrMenu(TB.wrap,      'Format', 'Text wrapping', 'Wrap') },
        'k': { label: 'Insert Link',   ...tbOrMenu(TB.link,      'Insert', 'Link') },

        // ── Fill color (Excel: Alt+H, H and Alt+H, G) ─────────
        // Opens the real color palette with arrow-key navigation
        'h': { label: 'Fill Color',    action: colorPickerAction(TB.fillColor) },
        'j': { label: 'No Fill',      action: () => VexcelColorPicker.applyColor(TB.fillColor, 'reset') },

        // ── Clear formatting ───────────────────────────────────
        'e': { label: 'Clear Format', action: menu('Format', 'Clear formatting') },

        // ── Decimal places ───────────────────────────────────
        '9': { label: 'Increase Decimal', action: () => VexcelToolbar.clickButton('Increase decimal places') },
        '0': { label: 'Decrease Decimal', action: () => VexcelToolbar.clickButton('Decrease decimal places') },

        // ── V — Paste sub-menu (Excel: Alt+H, V) ────────────
        'v': {
          label: 'Paste', mode: 'list',
          children: {
            'p': { label: 'Paste',            action: () => {
              const doc = (() => { try { return window.top.document; } catch(e) { return document; } })();
              (doc.activeElement || doc.body).dispatchEvent(new KeyboardEvent('keydown', {
                key: 'v', code: 'KeyV', keyCode: 86, metaKey: true, bubbles: true, cancelable: true
              }));
            }},
            'v': { label: 'Values Only',      action: menu('Edit', 'Paste special', 'Values only') },
            'f': { label: 'Formula Only',     action: menu('Edit', 'Paste special', 'Formula only') },
            'o': { label: 'Format Only',      action: menu('Edit', 'Paste special', 'Format only') },
            't': { label: 'Transposed',       action: menu('Edit', 'Paste special', 'Transposed') },
            'w': { label: 'Column Widths',    action: menu('Edit', 'Paste special', 'Column widths only') },
            'b': { label: 'All Except Borders', action: menu('Edit', 'Paste special', 'All except borders') },
            's': { label: 'Paste Special...', action: menu('Edit', 'Paste special') },
          }
        },

        // ── B — Borders sub-menu (Excel: Alt+H, B) ────────────
        'b': {
          label: 'Borders', mode: 'list',
          children: {
            'a': { label: 'All Borders',      action: borderAction('All borders') },
            's': { label: 'Outside Borders',   action: borderAction('Outside borders') },
            'p': { label: 'Top Border',        action: borderAction('Top border') },
            'o': { label: 'Bottom Border',     action: borderAction('Bottom border') },
            'l': { label: 'Left Border',       action: borderAction('Left border') },
            'r': { label: 'Right Border',      action: borderAction('Right border') },
            'n': { label: 'No Borders',        action: borderAction('Clear borders') },
            'c': { label: 'Border Color',      action: borderAction('Border color') },
          }
        },

        // ── F — Font sub-menu ──────────────────────────────────
        'f': {
          label: 'Font', mode: 'list',
          children: {
            'c': { label: 'Font Color', action: colorPickerAction(TB.fontColor) },
            's': { label: 'Font Size',  action: menu('Format', 'Font size') },
          }
        },

        // ── L — Fill operations ────────────────────────────────
        'l': {
          label: 'Fill', mode: 'list',
          children: {
            'd': { label: 'Fill Down',  action: menu('Edit', 'Fill', 'Fill down') },
            'r': { label: 'Fill Right', action: menu('Edit', 'Fill', 'Fill right') },
            'u': { label: 'Fill Up',    action: menu('Edit', 'Fill', 'Fill up') },
            'l': { label: 'Fill Left',  action: menu('Edit', 'Fill', 'Fill left') },
          }
        },

        // ── A — Align sub-menu ─────────────────────────────────
        'a': {
          label: 'Align', mode: 'toolbar',
          children: {
            'l': { label: 'Align Left',    ...tbOrMenu(TB.alignLeft,    'Format', 'Align', 'Left') },
            'c': { label: 'Center',        ...tbOrMenu(TB.alignCenter,  'Format', 'Align', 'Center') },
            'r': { label: 'Align Right',   ...tbOrMenu(TB.alignRight,   'Format', 'Align', 'Right') },
            't': { label: 'Align Top',     ...tbOrMenu(TB.alignTop,     'Format', 'Align', 'Top') },
            'm': { label: 'Align Middle',  ...tbOrMenu(TB.alignMiddle,  'Format', 'Align', 'Middle') },
            'b': { label: 'Align Bottom',  ...tbOrMenu(TB.alignBottom,  'Format', 'Align', 'Bottom') },
          }
        },

        // ── M — Merge sub-menu ─────────────────────────────────
        'm': {
          label: 'Merge', mode: 'list',
          children: {
            'c': { label: 'Merge & Center',      action: menu('Format', 'Merge cells', 'Merge all') },
            'a': { label: 'Merge All',          action: menu('Format', 'Merge cells', 'Merge all') },
            'h': { label: 'Merge Horizontally',  action: menu('Format', 'Merge cells', 'Merge horizontally') },
            'v': { label: 'Merge Vertically',   action: menu('Format', 'Merge cells', 'Merge vertically') },
            'u': { label: 'Unmerge',            action: menu('Format', 'Merge cells', 'Unmerge') },
          }
        },

        // ── N — Number Format sub-menu ─────────────────────────
        'n': {
          label: 'Number Fmt', mode: 'list',
          children: {
            'g': { label: 'General (Auto)',   action: menu('Format', 'Number', 'Automatic') },
            'n': { label: 'Number',          action: menu('Format', 'Number', 'Number') },
            'c': { label: 'Currency',        action: menu('Format', 'Number', 'Currency') },
            'p': { label: 'Percent',         action: menu('Format', 'Number', 'Percent') },
            'd': { label: 'Date',            action: menu('Format', 'Number', 'Date') },
            't': { label: 'Time',            action: menu('Format', 'Number', 'Time') },
            's': { label: 'Scientific',      action: menu('Format', 'Number', 'Scientific') },
            'x': { label: 'Custom...',       action: menu('Format', 'Number', 'Custom number format') },
            'i': { label: 'Increase Decimal', action: () => VexcelToolbar.clickButton('Increase decimal places') },
            'e': { label: 'Decrease Decimal', action: () => VexcelToolbar.clickButton('Decrease decimal places') },
          }
        },

        // ── I — Insert cells/rows/cols ─────────────────────────
        'i': {
          label: 'Insert', mode: 'list',
          children: {
            'r': { label: 'Row Above',    action: menu('Insert', 'Rows', 'Insert 1 row above') },
            'b': { label: 'Row Below',    action: menu('Insert', 'Rows', 'Insert 1 row below') },
            'c': { label: 'Column Left',  action: menu('Insert', 'Column', 'Insert Column left') },
            'l': { label: 'Column Right', action: menu('Insert', 'Column', 'Insert Column right') },
            's': { label: 'New Sheet',    action: menu('Insert', 'Sheet') },
          }
        },

        // ── D — Delete ────────────────────────────────────────
        'd': {
          label: 'Delete', mode: 'list',
          children: {
            'r': { label: 'Delete Row',    action: menu('Edit', 'Delete', 'row') },
            'c': { label: 'Delete Column', action: menu('Edit', 'Delete', 'column') },
            's': { label: 'Delete Sheet',  action: menu('Edit', 'Delete', 'sheet') },
          }
        },

        // ── O — Row/Column sizing ─────────────────────────────
        'o': {
          label: 'Row/Col Size', mode: 'list',
          children: {
            'r': { label: 'Auto-fit Row',    action: () => VexcelContextMenu.autoResizeRows() },
            'c': { label: 'Auto-fit Column', action: () => VexcelContextMenu.autoResizeColumns() },
            'h': { label: 'Hide Row',        action: () => VexcelContextMenu.hideRows() },
            'j': { label: 'Hide Column',     action: () => VexcelContextMenu.hideColumns() },
            'u': { label: 'Unhide Rows',     action: () => VexcelContextMenu.unhideRows() },
            'l': { label: 'Unhide Columns',  action: () => VexcelContextMenu.unhideColumns() },
          }
        },

        // ── S — Sort & Filter ──────────────────────────────────
        's': {
          label: 'Sort/Filter', mode: 'list',
          children: {
            'a': { label: 'Sort Range',     action: menu('Data', 'Sort range') },
            'f': { label: 'Toggle Filter',   ...tbOrMenu(TB.filter, 'Data', 'Filter') },
          }
        },
      }
    },

    // ═══════════════════════════════════════════════════════════
    // N — Insert (Excel: Alt+N)
    // ═══════════════════════════════════════════════════════════
    'n': {
      label: 'Insert', hint: 'Rows, charts, images',
      mode: 'list',
      children: {
        'r': { label: 'Row Above',    action: menu('Insert', 'Rows', 'Insert 1 row above') },
        'w': { label: 'Row Below',    action: menu('Insert', 'Rows', 'Insert 1 row below') },
        'c': { label: 'Column Left',  action: menu('Insert', 'Column', 'Insert Column left') },
        's': { label: 'New Sheet',    action: menu('Insert', 'Sheet') },
        'h': { label: 'Chart',        action: menu('Insert', 'Chart') },
        'f': { label: 'Function',     action: menu('Insert', 'Function') },
        'l': { label: 'Link',         action: menu('Insert', 'Link') },
        'n': { label: 'Note',         action: menu('Insert', 'Note') },
        'o': { label: 'Comment',      action: menu('Insert', 'Comment') },
        'p': { label: 'Pivot Table',  action: menu('Insert', 'Pivot table') },
        'd': { label: 'Drawing',      action: menu('Insert', 'Drawing') },
        'i': { label: 'Image',        action: menu('Insert', 'Image') },
      }
    },

    // ═══════════════════════════════════════════════════════════
    // M — Formulas (Excel: Alt+M)
    // ═══════════════════════════════════════════════════════════
    'm': {
      label: 'Formulas', hint: 'Trace, audit, names',
      mode: 'list',
      children: {
        't': { label: 'Trace Precedents',  action: () => VexcelTraceArrows.tracePrecedents() },
        'd': { label: 'Trace Dependents',  action: () => VexcelTraceArrows.traceDependents() },
        'a': { label: 'Remove Arrows',     action: () => VexcelTraceArrows.clearArrows() },
        'h': { label: 'Show Formulas',     action: menu('View', 'Show formulas') },
        'n': { label: 'Named Ranges',      action: menu('Data', 'Named ranges') },
        'f': { label: 'Insert Function',   action: menu('Insert', 'Function') },
      }
    },

    // ═══════════════════════════════════════════════════════════
    // A — Data (Excel: Alt+A)
    // ═══════════════════════════════════════════════════════════
    'a': {
      label: 'Data', hint: 'Sort, filter, validate',
      mode: 'list',
      children: {
        's': { label: 'Sort Range',         action: menu('Data', 'Sort range') },
        'f': { label: 'Create Filter',     ...tbOrMenu(TB.filter, 'Data', 'Filter') },
        'v': { label: 'Data Validation',   action: menu('Data', 'Data validation') },
        'p': { label: 'Pivot Table',       action: menu('Insert', 'Pivot table') },
        'r': { label: 'Remove Duplicates', action: menu('Data', 'Data cleanup', 'Remove duplicates') },
        't': { label: 'Text to Columns',   action: menu('Data', 'Split text to columns') },
        'g': { label: 'Group Rows',        action: menu('Data', 'Group rows') },
        'u': { label: 'Ungroup Rows',      action: menu('Data', 'Ungroup rows') },
        'n': { label: 'Named Ranges',      action: menu('Data', 'Named ranges') },
      }
    },

    // ═══════════════════════════════════════════════════════════
    // W — View (Excel: Alt+W)
    // ═══════════════════════════════════════════════════════════
    'w': {
      label: 'View', hint: 'Freeze, zoom, display',
      mode: 'list',
      children: {
        'f': {
          label: 'Freeze', mode: 'list',
          children: {
            '1': { label: 'Freeze 1 Row',     action: menu('View', 'Freeze', '1 row') },
            '2': { label: 'Freeze 2 Rows',    action: menu('View', 'Freeze', '2 rows') },
            'c': { label: 'Freeze 1 Column',  action: menu('View', 'Freeze', '1 column') },
            'r': { label: 'No Freeze (Rows)',  action: menu('View', 'Freeze', 'No rows') },
            'l': { label: 'No Freeze (Cols)',  action: menu('View', 'Freeze', 'No columns') },
          }
        },
        'n': { label: 'Next Sheet',      action: () => VexcelNavigation.goNextSheet() },
        'v': { label: 'Prev Sheet',     action: () => VexcelNavigation.goPrevSheet() },
        'g': { label: 'Gridlines',       action: menu('View', 'Show', 'Gridlines') },
        'h': { label: 'Row/Col Headers', action: menu('View', 'Show', 'Row and column headers') },
        'z': { label: 'Zoom',            action: menu('View', 'Zoom') },
        'p': { label: 'Full Screen',     action: menu('View', 'Full screen') },
        'b': { label: 'Formula Bar',     action: menu('View', 'Show', 'Formula bar') },
      }
    },

    // ═══════════════════════════════════════════════════════════
    // F — File (Excel: Alt+F)
    // ═══════════════════════════════════════════════════════════
    'f': {
      label: 'File', hint: 'Save, print, download',
      mode: 'list',
      children: {
        's': { label: 'Save',             action: () => {
          // Google Sheets auto-saves; just trigger Cmd+S behavior
          const doc = (() => { try { return window.top.document; } catch(e) { return document; } })();
          doc.dispatchEvent(new KeyboardEvent('keydown', {
            key: 's', code: 'KeyS', keyCode: 83, metaKey: true, bubbles: true
          }));
        }},
        'p': { label: 'Print',            action: menu('File', 'Print') },
        'd': { label: 'Download',         action: menu('File', 'Download') },
        'x': { label: 'Download as Excel', action: menu('File', 'Download', 'Microsoft Excel') },
        'c': { label: 'Download as CSV',  action: menu('File', 'Download', 'Comma-separated values') },
        'f': { label: 'Download as PDF',  action: menu('File', 'Download', 'PDF document') },
        'r': { label: 'Rename',           action: menu('File', 'Rename') },
        'h': { label: 'Version History',  action: menu('File', 'Version history', 'See version history') },
        'a': { label: 'Share',            action: menu('File', 'Share', 'Share with others') },
        'e': { label: 'Email',            action: menu('File', 'Email', 'Email this file') },
      }
    },

    // ═══════════════════════════════════════════════════════════
    // E — Edit (Excel: no direct equivalent, useful addition)
    // ═══════════════════════════════════════════════════════════
    'e': {
      label: 'Edit', hint: 'Undo, copy, paste, find',
      mode: 'list',
      children: {
        'u': { label: 'Undo',             action: () => {
          const doc = (() => { try { return window.top.document; } catch(e) { return document; } })();
          doc.dispatchEvent(new KeyboardEvent('keydown', {
            key: 'z', code: 'KeyZ', keyCode: 90, metaKey: true, bubbles: true
          }));
        }},
        'r': { label: 'Redo',             action: () => {
          const doc = (() => { try { return window.top.document; } catch(e) { return document; } })();
          doc.dispatchEvent(new KeyboardEvent('keydown', {
            key: 'z', code: 'KeyZ', keyCode: 90, metaKey: true, shiftKey: true, bubbles: true
          }));
        }},
        'f': { label: 'Find',             action: menu('Edit', 'Find and replace') },
        's': { label: 'Paste Special',    action: menu('Edit', 'Paste special') },
        'v': { label: 'Paste Values',     action: menu('Edit', 'Paste special', 'Values only') },
        'a': { label: 'Select All',       action: () => {
          const doc = (() => { try { return window.top.document; } catch(e) { return document; } })();
          doc.dispatchEvent(new KeyboardEvent('keydown', {
            key: 'a', code: 'KeyA', keyCode: 65, metaKey: true, bubbles: true
          }));
        }},
      }
    },

  };

  return { getTree: () => tree };

})();
