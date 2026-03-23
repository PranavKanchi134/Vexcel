// Navigation shortcuts: Go To, sheet tabs, etc.

const VexcelNavigation = (() => {
  const shortcuts = {
    // Go To (Name Box focus) - Cmd+G
    'cmd+g': {
      label: 'Go To',
      action: () => {
        VexcelCellEditor.focusNameBox();
      },
      overridesChrome: true
    },

    // Go To (Name Box focus) - Ctrl+G (Excel alternative)
    'ctrl+g': {
      label: 'Go To',
      action: () => {
        VexcelCellEditor.focusNameBox();
      }
    },

    // Next sheet tab - Option+Right
    'option+right': {
      label: 'Next Sheet',
      action: () => navigateSheetTab(1)
    },

    // Previous sheet tab - Option+Left
    'option+left': {
      label: 'Previous Sheet',
      action: () => navigateSheetTab(-1)
    },

    // F5 — Go To (Excel standard)
    'f5': {
      label: 'Go To',
      action: () => VexcelCellEditor.focusNameBox()
    },

    // Ctrl+Shift+; — Insert Time (Excel alternative on Mac)
    'cmd+shift+;': {
      label: 'Insert Time',
      action: () => {
        const now = new Date();
        const hours = now.getHours();
        const minutes = String(now.getMinutes()).padStart(2, '0');
        const ampm = hours >= 12 ? 'PM' : 'AM';
        const h12 = hours % 12 || 12;
        const timeStr = `${h12}:${minutes} ${ampm}`;
        if (!VexcelCellEditor.isEditing()) {
          const editor = VexcelCellEditor.getEditorElement();
          if (editor) editor.focus();
        }
        VexcelCellEditor.insertText(timeStr);
      }
    },

    // Go to beginning of row - Home
    // Go to beginning of sheet - Cmd+Home
    // These pass through natively in Google Sheets
  };

  /**
   * Navigate to the next or previous sheet tab.
   * @param {number} direction - 1 for next, -1 for previous
   */
  function navigateSheetTab(direction) {
    const tabs = VexcelDom.getSheetTabs();
    const activeTab = VexcelDom.getActiveSheetTab();

    if (!tabs.length || !activeTab) return;

    const currentIndex = tabs.indexOf(activeTab);
    if (currentIndex === -1) return;

    const newIndex = currentIndex + direction;
    if (newIndex >= 0 && newIndex < tabs.length) {
      tabs[newIndex].click();
    }
  }

  function goNextSheet() { navigateSheetTab(1); }
  function goPrevSheet() { navigateSheetTab(-1); }

  return { shortcuts, goNextSheet, goPrevSheet };
})();
