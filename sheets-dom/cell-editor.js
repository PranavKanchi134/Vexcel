// Google Sheets cell editor interaction

const VexcelCellEditor = (() => {

  function topDoc() {
    try { return window.top.document; }
    catch (e) { return document; }
  }

  /**
   * Check if the cell editor is currently active (in edit mode).
   * The cell editor can be in either the current frame or the top frame.
   */
  function isEditing() {
    // Check current document first (iframe where grid lives)
    const editor = getEditorElement();
    if (editor && (document.activeElement === editor || editor.contains(document.activeElement))) {
      return true;
    }
    // Also check top frame
    const top = topDoc();
    if (top !== document) {
      const topEditor = top.querySelector('.cell-input, #waffle-cell-editor');
      if (topEditor && (top.activeElement === topEditor || topEditor.contains(top.activeElement))) {
        return true;
      }
    }
    return false;
  }

  /**
   * Get the cell editor element.
   * Searches current frame first, then top frame.
   */
  function getEditorElement() {
    const el = document.querySelector('.cell-input, #waffle-cell-editor, [role="textbox"].cell-input');
    if (el) return el;
    const top = topDoc();
    if (top !== document) {
      return top.querySelector('.cell-input, #waffle-cell-editor, [role="textbox"].cell-input');
    }
    return null;
  }

  /**
   * Get the formula bar input element (always in top frame).
   */
  function getFormulaBar() {
    const top = topDoc();
    return top.querySelector('#t-formula-bar-input, .cell-input.formula-bar-input');
  }

  function setInputValue(input, text) {
    if (!input) return false;
    const proto = input instanceof HTMLTextAreaElement
      ? HTMLTextAreaElement.prototype
      : HTMLInputElement.prototype;
    const nativeSet = Object.getOwnPropertyDescriptor(proto, 'value');
    if (nativeSet && nativeSet.set) {
      nativeSet.set.call(input, text);
    } else {
      input.value = text;
    }
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new InputEvent('input', {
      bubbles: true,
      inputType: 'insertText',
      data: text
    }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  }

  /**
   * Insert text at the current cursor position in the cell editor.
   * @param {string} text - Text to insert
   */
  function insertText(text) {
    const editor = getEditorElement();
    if (!editor) return false;

    editor.focus();

    // Try execCommand first (still works in most browsers, supports undo)
    const doc = editor.ownerDocument;
    if (doc.execCommand && doc.queryCommandSupported && doc.queryCommandSupported('insertText')) {
      doc.execCommand('insertText', false, text);
      return true;
    }

    // Fallback: InputEvent-based insertion
    const inputEvent = new InputEvent('beforeinput', {
      inputType: 'insertText',
      data: text,
      bubbles: true,
      cancelable: true,
    });
    editor.dispatchEvent(inputEvent);
    editor.dispatchEvent(new InputEvent('input', {
      inputType: 'insertText',
      data: text,
      bubbles: true,
    }));
    return true;
  }

  /**
   * Set the entire cell content (replaces current content).
   * @param {string} text - New cell content
   */
  function setCellContent(text) {
    const editor = getEditorElement();
    if (!editor) return false;

    editor.focus();
    const doc = editor.ownerDocument;

    // Select all then insert
    if (doc.execCommand && doc.queryCommandSupported && doc.queryCommandSupported('selectAll')) {
      doc.execCommand('selectAll', false, null);
      doc.execCommand('insertText', false, text);
    } else {
      // Fallback: set textContent directly
      editor.textContent = text;
      editor.dispatchEvent(new InputEvent('input', {
        inputType: 'insertText',
        data: text,
        bubbles: true,
      }));
    }
    return true;
  }

  function setFormulaBarValue(text) {
    const formulaBar = getFormulaBar();
    if (!formulaBar) return false;
    formulaBar.focus();
    return setInputValue(formulaBar, text);
  }

  function commitFormulaBar(move = 'enter') {
    const formulaBar = getFormulaBar();
    if (!formulaBar) return false;

    const key = move === 'tab' ? 'Tab' : 'Enter';
    const keyCode = move === 'tab' ? 9 : 13;
    const options = {
      key,
      code: key,
      keyCode,
      which: keyCode,
      bubbles: true,
      cancelable: true
    };

    formulaBar.dispatchEvent(new KeyboardEvent('keydown', options));
    formulaBar.dispatchEvent(new KeyboardEvent('keypress', options));
    formulaBar.dispatchEvent(new KeyboardEvent('keyup', options));
    return true;
  }

  /**
   * Focus the name box (for Go To / cell reference navigation).
   * Name box is always in the top frame.
   */
  function focusNameBox() {
    const nameBox = VexcelDom.getNameBox();
    if (nameBox) {
      nameBox.focus();
      nameBox.select();
      return true;
    }
    return false;
  }

  return {
    isEditing,
    getEditorElement,
    getFormulaBar,
    insertText,
    setCellContent,
    setFormulaBarValue,
    commitFormulaBar,
    focusNameBox
  };
})();
