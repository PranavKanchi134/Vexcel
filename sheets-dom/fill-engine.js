// Native fill helpers that bypass slow menu automation.

const VexcelFillEngine = (() => {
  const MAX_DIRECT_FILL_CELLS = 120;
  const selectionCache = new Map();
  const disabledDirections = {
    down: false,
    right: false
  };

  function topDoc() {
    try { return window.top.document; }
    catch (e) { return document; }
  }

  function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  async function fillDown() {
    return fill('down');
  }

  async function fillRight() {
    return fill('right');
  }

  async function fill(direction) {
    if (disabledDirections[direction]) {
      return fail(
        `Fill ${direction === 'down' ? 'Down' : 'Right'} direct mode unavailable`,
        `direct fill-${direction} disabled for this tab`
      );
    }

    const startedAt = now();
    const selection = getCurrentSelectionAddress();
    const parsed = parseRange(selection);
    const label = direction === 'down' ? 'Fill Down' : 'Fill Right';

    if (!selection || !parsed) {
      return fail(`${label} failed: no contiguous selection`);
    }

    const cellCount = ((parsed.endRow - parsed.startRow) + 1) * ((parsed.endCol - parsed.startCol) + 1);
    if (cellCount > MAX_DIRECT_FILL_CELLS) {
      return fail(`${label} direct mode supports up to ${MAX_DIRECT_FILL_CELLS} cells`, `selection has ${cellCount} cells`);
    }

    const doc = topDoc();
    const originalSelection = selection;
    let lastFormulaValue = readFormulaBar(doc);
    const operations = buildOperations(parsed, direction);
    if (!operations.length) {
      return fail(`${label} failed`, 'selection has no fill targets');
    }

    const sourceValues = new Map();
    for (const operation of operations) {
      if (sourceValues.has(operation.sourceAddress)) continue;

      const navigated = await navigateToAddress(operation.sourceAddress, doc);
      if (!navigated) {
        await restoreSelection(originalSelection, doc);
        return fail(`${label} failed to read the source cells`, `could not navigate to ${operation.sourceAddress}`);
      }

      const sourceValue = await waitForFormulaBarUpdate(doc, lastFormulaValue);
      lastFormulaValue = sourceValue;
      sourceValues.set(operation.sourceAddress, sourceValue);
    }

    for (let index = 0; index < operations.length; index++) {
      const operation = operations[index];
      const sourceValue = sourceValues.get(operation.sourceAddress);
      const targetValue = shiftCellContent(sourceValue, operation.rowDelta, operation.colDelta);
      const shouldVerify = index === 0;
      const wrote = await writeCell(operation.targetAddress, targetValue, doc, shouldVerify);
      if (!wrote) {
        if (shouldVerify) disabledDirections[direction] = true;
        await restoreSelection(originalSelection, doc);
        return fail(`${label} failed while writing cells`, `could not write ${operation.targetAddress}`);
      }
    }

    await restoreSelection(originalSelection, doc);
    disabledDirections[direction] = false;
    return {
      ok: true,
      strategy: direction === 'down' ? 'nativeFillDown' : 'nativeFillRight',
      verified: true,
      durationMs: Math.round(now() - startedAt),
      message: label
    };
  }

  function buildOperations(parsed, direction) {
    const ops = [];
    const rowCount = (parsed.endRow - parsed.startRow) + 1;
    const colCount = (parsed.endCol - parsed.startCol) + 1;
    const singleCell = rowCount === 1 && colCount === 1;

    if (direction === 'down') {
      if (singleCell) {
        if (parsed.startRow <= 1) return [];
        ops.push({
          sourceAddress: formatAddress(parsed.sheetPrefix, parsed.startCol, parsed.startRow - 1),
          targetAddress: formatAddress(parsed.sheetPrefix, parsed.startCol, parsed.startRow),
          rowDelta: 1,
          colDelta: 0
        });
        return ops;
      }

      if (rowCount < 2) return [];
      for (let col = parsed.startCol; col <= parsed.endCol; col++) {
        const sourceAddress = formatAddress(parsed.sheetPrefix, col, parsed.startRow);
        for (let row = parsed.startRow + 1; row <= parsed.endRow; row++) {
          ops.push({
            sourceAddress,
            targetAddress: formatAddress(parsed.sheetPrefix, col, row),
            rowDelta: row - parsed.startRow,
            colDelta: 0
          });
        }
      }
      return ops;
    }

    if (singleCell) {
      if (parsed.startCol <= 1) return [];
      ops.push({
        sourceAddress: formatAddress(parsed.sheetPrefix, parsed.startCol - 1, parsed.startRow),
        targetAddress: formatAddress(parsed.sheetPrefix, parsed.startCol, parsed.startRow),
        rowDelta: 0,
        colDelta: 1
      });
      return ops;
    }

    if (colCount < 2) return [];
    for (let row = parsed.startRow; row <= parsed.endRow; row++) {
      const sourceAddress = formatAddress(parsed.sheetPrefix, parsed.startCol, row);
      for (let col = parsed.startCol + 1; col <= parsed.endCol; col++) {
        ops.push({
          sourceAddress,
          targetAddress: formatAddress(parsed.sheetPrefix, col, row),
          rowDelta: 0,
          colDelta: col - parsed.startCol
        });
      }
    }

    return ops;
  }

  async function writeCell(address, text, doc, verify = false) {
    const navigated = await navigateToAddress(address, doc);
    if (!navigated) return false;
    await sleep(8);
    if (!VexcelCellEditor.setFormulaBarValue(text)) return false;
    await sleep(6);
    if (!VexcelCellEditor.commitFormulaBar('enter')) return false;
    await sleep(10);
    if (!verify) return true;

    const returned = await navigateToAddress(address, doc);
    if (!returned) return false;
    const confirmed = await waitForFormulaBarUpdate(doc, '');
    return normalizeCellText(confirmed) === normalizeCellText(text);
  }

  function normalizeCellText(text) {
    return `${text || ''}`.replace(/\r\n/g, '\n').trim();
  }

  function shiftCellContent(value, rowDelta, colDelta) {
    const text = `${value || ''}`;
    if (!text.startsWith('=')) return text;
    return shiftFormula(text, rowDelta, colDelta);
  }

  function shiftFormula(formula, rowDelta, colDelta) {
    const segments = formula.split(/("(?:[^"]|"")*")/g);
    return segments.map((segment, index) => (
      index % 2 === 1 ? segment : shiftFormulaSegment(segment, rowDelta, colDelta)
    )).join('');
  }

  function shiftFormulaSegment(segment, rowDelta, colDelta) {
    return segment.replace(
      /(^|[^A-Z0-9_])((?:'[^']+'|[A-Za-z0-9_]+)!)?(\$?)([A-Z]{1,3})(\$?)(\d{1,7})(?=$|[^A-Z0-9_])/gi,
      (full, prefix, sheet, absCol, colLetters, absRow, rowDigits) => {
        const shiftedCol = absCol === '$'
          ? colLetters.toUpperCase()
          : numToCol(colToNum(colLetters.toUpperCase()) + colDelta);
        const shiftedRow = absRow === '$'
          ? rowDigits
          : String(parseInt(rowDigits, 10) + rowDelta);
        return `${prefix || ''}${sheet || ''}${absCol || ''}${shiftedCol}${absRow || ''}${shiftedRow}`;
      }
    );
  }

  function formatAddress(sheetPrefix, col, row) {
    return `${sheetPrefix}${numToCol(col)}${row}`;
  }

  function getCurrentSelectionAddress() {
    const cached = selectionCache.get('current');
    if (cached && cached.selection && Date.now() - cached.at < 1500) {
      return cached.selection;
    }

    const doc = topDoc();
    const nameBox = VexcelSelectorCache.get('name-box', () => (
      doc.querySelector('#t-name-box input, .jfk-textinput[aria-label="Name Box"], input[aria-label="Name Box"]')
    ));
    const selection = nameBox ? (nameBox.value || nameBox.textContent || '').trim() : '';
    if (selection) selectionCache.set('current', { selection, at: Date.now() });
    return selection;
  }

  function parseRange(selection) {
    if (!selection || selection.includes(',')) return null;

    const match = selection.match(/^(?:(.*)!)?(\$?[A-Z]{1,3}\$?\d{1,7})(?::(\$?[A-Z]{1,3}\$?\d{1,7}))?$/i);
    if (!match) return null;

    const start = parseCellRef(match[2]);
    const end = parseCellRef(match[3] || match[2]);
    if (!start || !end) return null;

    return {
      sheetPrefix: match[1] ? `${match[1]}!` : '',
      startCol: Math.min(start.col, end.col),
      endCol: Math.max(start.col, end.col),
      startRow: Math.min(start.row, end.row),
      endRow: Math.max(start.row, end.row)
    };
  }

  function parseCellRef(ref) {
    const clean = `${ref || ''}`.replace(/\$/g, '').toUpperCase();
    const match = clean.match(/^([A-Z]{1,3})(\d{1,7})$/);
    if (!match) return null;

    return {
      col: colToNum(match[1]),
      row: parseInt(match[2], 10)
    };
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
    const nameBox = VexcelSelectorCache.get('name-box', () => (
      doc.querySelector('#t-name-box input, .jfk-textinput[aria-label="Name Box"], input[aria-label="Name Box"]')
    ));
    if (!nameBox) return false;

    clickElement(nameBox);
    nameBox.focus();
    if (nameBox.select) nameBox.select();
    await sleep(4);

    const nativeSet = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
    if (nativeSet && nativeSet.set) {
      nativeSet.set.call(nameBox, address);
    } else {
      nameBox.value = address;
    }

    nameBox.dispatchEvent(new Event('input', { bubbles: true }));
    nameBox.dispatchEvent(new Event('change', { bubbles: true }));

    const options = {
      key: 'Enter',
      code: 'Enter',
      keyCode: 13,
      which: 13,
      bubbles: true,
      cancelable: true
    };
    nameBox.dispatchEvent(new KeyboardEvent('keydown', options));
    nameBox.dispatchEvent(new KeyboardEvent('keypress', options));
    nameBox.dispatchEvent(new KeyboardEvent('keyup', options));
    selectionCache.set('current', { selection: address, at: Date.now() });
    return true;
  }

  async function restoreSelection(selection, doc) {
    if (!selection) return;
    await navigateToAddress(selection, doc);
    await sleep(12);
    selectionCache.set('current', { selection, at: Date.now() });
  }

  async function waitForFormulaBarUpdate(doc, previousValue) {
    for (let attempt = 0; attempt < 7; attempt++) {
      const nextValue = readFormulaBar(doc);
      if (attempt === 0 && nextValue && nextValue !== previousValue) return nextValue;
      if (attempt > 1 && nextValue) return nextValue;
      await sleep(attempt < 3 ? 10 : 16);
    }
    return readFormulaBar(doc);
  }

  function readFormulaBar(doc) {
    const formulaBar = VexcelSelectorCache.get('formula-bar', () => (
      doc.querySelector('#t-formula-bar-input, .cell-input[aria-label="Formula Bar"], .formulabar-input')
    ));
    if (formulaBar) {
      return (formulaBar.value || formulaBar.textContent || formulaBar.innerText || '').trim();
    }
    return '';
  }

  function clickElement(el) {
    const rect = el.getBoundingClientRect();
    const x = rect.left + rect.width / 2;
    const y = rect.top + rect.height / 2;
    const win = el.ownerDocument.defaultView || window;
    const options = { bubbles: true, cancelable: true, view: win, clientX: x, clientY: y };
    el.dispatchEvent(new MouseEvent('mousedown', options));
    el.dispatchEvent(new MouseEvent('mouseup', options));
    el.dispatchEvent(new MouseEvent('click', options));
  }

  function fail(message, reason = '') {
    return { ok: false, message, reason };
  }

  function now() {
    return (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
  }

  return { fillDown, fillRight };
})();
