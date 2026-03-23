// Finance-oriented formatting helpers for Google Sheets

const VexcelFinanceTools = (() => {
  const FORMAT_CYCLE = [
    { label: 'Automatic', query: 'Automatic', path: ['Format', 'Number', 'Automatic'] },
    { label: 'Number', query: 'Number', path: ['Format', 'Number', 'Number'] },
    { label: 'Percent', query: 'Percent', path: ['Format', 'Number', 'Percent'] },
    { label: 'Currency', query: 'Currency', path: ['Format', 'Number', 'Currency'] },
    { label: 'Date', query: 'Date', path: ['Format', 'Number', 'Date'] }
  ];

  const FONT_COLOR_CYCLE = [
    { label: 'Hardcode Blue', hex: '#1155cc' },
    { label: 'Formula Black', hex: '#000000' },
    { label: 'Link Green', hex: '#38761d' },
    { label: 'Alert Red', hex: '#cc0000' }
  ];

  const AUTO_COLOR_MAP = {
    hardcode: { label: 'Hardcode Blue', hex: '#1155cc' },
    formula: { label: 'Formula Black', hex: '#000000' },
    crossSheet: { label: 'Cross-Sheet Green', hex: '#38761d' }
  };

  const MAX_AUTO_COLOR_CELLS = 250;
  let formatCycleIdx = -1;
  let fontColorCycleIdx = -1;

  function topDoc() {
    try { return window.top.document; }
    catch (e) { return document; }
  }

  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  async function cycleFormat() {
    formatCycleIdx = (formatCycleIdx + 1) % FORMAT_CYCLE.length;
    const next = FORMAT_CYCLE[formatCycleIdx];
    const outcome = await VexcelMenuNavigator.runCommand(next.query, next.path, {
      prefer: 'toolFinder'
    });

    if (outcome.ok) {
      return {
        ok: true,
        strategy: outcome.strategy,
        verified: true,
        message: `Format: ${next.label}`
      };
    }

    return {
      ok: false,
      message: `Format cycle failed on ${next.label}`,
      reason: outcome.reason || 'format command did not resolve'
    };
  }

  async function cycleFontColor() {
    fontColorCycleIdx = (fontColorCycleIdx + 1) % FONT_COLOR_CYCLE.length;
    const next = FONT_COLOR_CYCLE[fontColorCycleIdx];
    const ok = await VexcelColorPicker.applyColor('Text color', next.hex);

    return ok ? {
      ok: true,
      strategy: 'colorPicker',
      verified: true,
      message: `Font Color: ${next.label}`
    } : {
      ok: false,
      message: `Font color cycle failed on ${next.label}`,
      reason: 'text color palette did not resolve'
    };
  }

  async function autoColorSelection() {
    const originalSelection = getCurrentSelectionAddress();
    if (!originalSelection) {
      return { ok: false, message: 'Auto-color failed: no active selection' };
    }

    const addresses = expandSelection(originalSelection);
    if (!addresses) {
      return { ok: false, message: 'Auto-color supports one contiguous range at a time' };
    }

    if (addresses.length > MAX_AUTO_COLOR_CELLS) {
      return {
        ok: false,
        message: `Auto-color limit: ${addresses.length} cells selected`,
        reason: `selection exceeds ${MAX_AUTO_COLOR_CELLS} cells`
      };
    }

    const counts = { hardcode: 0, formula: 0, crossSheet: 0, skipped: 0 };
    const failures = [];
    const doc = topDoc();
    let lastFormulaBarValue = readFormulaBar(doc);
    const classified = [];

    for (const addr of addresses) {
      const navigated = await navigateToAddress(addr, doc);
      if (!navigated) {
        failures.push(addr);
        continue;
      }

      const contents = await waitForFormulaBarUpdate(doc, lastFormulaBarValue);
      lastFormulaBarValue = contents;
      const bucket = classifyCell(contents);

      if (!bucket) {
        counts.skipped++;
        continue;
      }
      classified.push({ addr, bucket });
    }

    const runs = buildColorRuns(classified);
    let lastAppliedBucket = '';
    for (const run of runs) {
      const selected = await navigateToAddress(run.range, doc);
      if (!selected) {
        failures.push(run.range);
        continue;
      }

      let applied;
      if (run.bucket === lastAppliedBucket) {
        applied = VexcelColorPicker.applyLastUsedColor('Text color');
      } else {
        applied = await VexcelColorPicker.applyColor('Text color', AUTO_COLOR_MAP[run.bucket].hex);
        if (applied) lastAppliedBucket = run.bucket;
      }

      if (applied) {
        counts[run.bucket] += run.count;
      } else {
        failures.push(run.range);
      }
      await sleep(10);
    }

    if (originalSelection) {
      await navigateToAddress(originalSelection, doc);
      await sleep(12);
    }

    const colored = counts.hardcode + counts.formula + counts.crossSheet;
    if (!colored) {
      return {
        ok: false,
        message: failures.length ? 'Auto-color could not classify or color the selection' : 'Auto-color found nothing to color',
        reason: failures.length ? `failed cells: ${failures.slice(0, 5).join(', ')}` : 'selection contained only blanks'
      };
    }

    const parts = [];
    if (counts.hardcode) parts.push(`${counts.hardcode} hardcodes`);
    if (counts.formula) parts.push(`${counts.formula} formulas`);
    if (counts.crossSheet) parts.push(`${counts.crossSheet} cross-sheet`);
    if (failures.length) parts.push(`${failures.length} failed`);

    return {
      ok: true,
      strategy: 'selectionAudit',
      verified: failures.length === 0,
      message: `Auto-colored ${colored} cells: ${parts.join(', ')}`
    };
  }

  function buildColorRuns(classified) {
    const runs = [];
    let current = null;

    for (const entry of classified) {
      const parsed = parseCellAddress(entry.addr);
      if (!parsed) continue;

      if (
        current &&
        current.bucket === entry.bucket &&
        current.sheetPrefix === parsed.sheetPrefix &&
        current.row === parsed.row &&
        parsed.col === current.endCol + 1
      ) {
        current.endCol = parsed.col;
        current.count++;
        current.range = formatRange(current);
      } else {
        current = {
          bucket: entry.bucket,
          sheetPrefix: parsed.sheetPrefix,
          row: parsed.row,
          startCol: parsed.col,
          endCol: parsed.col,
          count: 1
        };
        current.range = formatRange(current);
        runs.push(current);
      }
    }

    return runs;
  }

  function classifyCell(text) {
    const value = (text || '').trim();
    if (!value) return null;
    if (!value.startsWith('=')) return 'hardcode';
    if (value.includes('!')) return 'crossSheet';
    return 'formula';
  }

  function getCurrentSelectionAddress() {
    const doc = topDoc();
    const selectors = [
      '#t-name-box input',
      '.jfk-textinput[aria-label="Name Box"]',
      'input[aria-label="Name Box"]',
      '#t-name-box'
    ];

    for (const sel of selectors) {
      const el = doc.querySelector(sel);
      if (!el) continue;
      const val = (el.value || el.textContent || '').trim();
      if (val) return val;
    }

    return '';
  }

  function expandSelection(selection) {
    const parsed = parseRange(selection);
    if (!parsed) return null;

    const addresses = [];
    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
      for (let col = parsed.startCol; col <= parsed.endCol; col++) {
        addresses.push(`${parsed.sheetPrefix}${numToCol(col)}${row}`);
      }
    }
    return addresses;
  }

  function parseRange(selection) {
    if (!selection || selection.includes(',')) return null;

    const match = selection.match(/^(?:(.*)!)?(\$?[A-Z]{1,3}\$?\d{1,7})(?::(\$?[A-Z]{1,3}\$?\d{1,7}))?$/i);
    if (!match) return null;

    const sheetPrefix = match[1] ? `${match[1]}!` : '';
    const start = parseCellRef(match[2]);
    const end = parseCellRef(match[3] || match[2]);
    if (!start || !end) return null;

    return {
      sheetPrefix,
      startCol: Math.min(start.col, end.col),
      endCol: Math.max(start.col, end.col),
      startRow: Math.min(start.row, end.row),
      endRow: Math.max(start.row, end.row)
    };
  }

  function parseCellRef(ref) {
    const clean = ref.replace(/\$/g, '').toUpperCase();
    const match = clean.match(/^([A-Z]{1,3})(\d{1,7})$/);
    if (!match) return null;

    return {
      col: colToNum(match[1]),
      row: parseInt(match[2], 10)
    };
  }

  function parseCellAddress(address) {
    const match = address.match(/^(?:(.*)!)?([A-Z]{1,3})(\d{1,7})$/i);
    if (!match) return null;
    return {
      sheetPrefix: match[1] ? `${match[1]}!` : '',
      col: colToNum(match[2].toUpperCase()),
      row: parseInt(match[3], 10)
    };
  }

  function formatRange(run) {
    const start = `${run.sheetPrefix}${numToCol(run.startCol)}${run.row}`;
    const end = `${numToCol(run.endCol)}${run.row}`;
    return run.startCol === run.endCol ? start : `${start}:${end}`;
  }

  function colToNum(col) {
    let num = 0;
    for (const ch of col) {
      num = (num * 26) + (ch.charCodeAt(0) - 64);
    }
    return num;
  }

  function numToCol(num) {
    let col = '';
    let current = num;
    while (current > 0) {
      const rem = (current - 1) % 26;
      col = String.fromCharCode(65 + rem) + col;
      current = Math.floor((current - 1) / 26);
    }
    return col;
  }

  async function navigateToAddress(address, doc) {
    const nameBox = findNameBox(doc);
    if (!nameBox) return false;

    clickElement(nameBox);
    nameBox.focus();
    await sleep(6);

    const nativeSet = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
    if (nativeSet && nativeSet.set) {
      nativeSet.set.call(nameBox, address);
    } else {
      nameBox.value = address;
    }
    nameBox.dispatchEvent(new Event('input', { bubbles: true }));
    nameBox.dispatchEvent(new Event('change', { bubbles: true }));

    const enterOpts = {
      key: 'Enter',
      code: 'Enter',
      keyCode: 13,
      which: 13,
      bubbles: true,
      cancelable: true
    };
    nameBox.dispatchEvent(new KeyboardEvent('keydown', enterOpts));
    nameBox.dispatchEvent(new KeyboardEvent('keypress', enterOpts));
    nameBox.dispatchEvent(new KeyboardEvent('keyup', enterOpts));
    return true;
  }

  async function waitForFormulaBarUpdate(doc, previousValue) {
    for (let attempt = 0; attempt < 6; attempt++) {
      const value = readFormulaBar(doc);
      if (value && value !== previousValue) return value;
      await sleep(12);
    }
    return readFormulaBar(doc);
  }

  function findNameBox(doc) {
    return doc.querySelector('#t-name-box input, .jfk-textinput[aria-label="Name Box"], input[aria-label="Name Box"]');
  }

  function readFormulaBar(doc) {
    const selectors = [
      '#t-formula-bar-input',
      '.cell-input[aria-label="Formula Bar"]',
      '.formulabar-input'
    ];

    for (const sel of selectors) {
      const el = doc.querySelector(sel);
      if (!el) continue;
      const text = (el.value || el.textContent || el.innerText || '').trim();
      if (text) return text;
    }

    const inputs = doc.querySelectorAll('.cell-input');
    for (const el of inputs) {
      const rect = el.getBoundingClientRect();
      if (rect.top < 80 && rect.width > 100 && rect.height > 0) {
        const text = (el.value || el.textContent || el.innerText || '').trim();
        if (text) return text;
      }
    }

    return '';
  }

  function clickElement(el) {
    const rect = el.getBoundingClientRect();
    const x = rect.left + rect.width / 2;
    const y = rect.top + rect.height / 2;
    const win = el.ownerDocument.defaultView || window;
    const opts = { bubbles: true, cancelable: true, view: win, clientX: x, clientY: y };
    el.dispatchEvent(new MouseEvent('mousedown', opts));
    el.dispatchEvent(new MouseEvent('mouseup', opts));
    el.dispatchEvent(new MouseEvent('click', opts));
  }

  return { cycleFormat, cycleFontColor, autoColorSelection };
})();
