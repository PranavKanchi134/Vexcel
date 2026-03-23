// Row and column operation shortcuts
// Note: Hide/unhide rows and columns are context-menu-only operations in Google Sheets.
// We simulate right-click on the row/column header to access these.

const VexcelRowColumn = (() => {
  const shortcuts = {
    // Insert rows/columns - Ctrl+Shift+= (Shift+= produces + on US keyboards)
    'ctrl+shift++': {
      label: 'Insert Rows/Columns',
      action: () => VexcelMenuNavigator.clickMenuPath(['Insert', 'Rows', 'Insert 1 row above'])
    },

    // Delete rows/columns - Cmd+-
    'cmd+-': {
      label: 'Delete Rows/Columns',
      action: () => VexcelMenuNavigator.clickMenuPath(['Edit', 'Delete', 'row'])
    },

    // Hide rows - Ctrl+9 (context menu operation)
    'ctrl+9': {
      label: 'Hide Rows',
      action: () => VexcelContextMenu.hideRows()
    },

    // Hide columns - Ctrl+0 (context menu operation)
    'ctrl+0': {
      label: 'Hide Columns',
      action: () => VexcelContextMenu.hideColumns()
    },

    // Unhide rows - Ctrl+Shift+9 (Shift+9 produces '(' on US keyboards)
    'ctrl+shift+(': {
      label: 'Unhide Rows',
      action: () => VexcelContextMenu.unhideRows()
    },

    // Unhide columns - Ctrl+Shift+0 (Shift+0 produces ')' on US keyboards)
    'ctrl+shift+)': {
      label: 'Unhide Columns',
      action: () => VexcelContextMenu.unhideColumns()
    },

    // Auto-fit row height - Cmd+Option+R (no Chrome conflict)
    'cmd+option+r': {
      label: 'Auto-fit Row Height',
      action: () => VexcelContextMenu.autoResizeRows()
    },

    // Auto-fit column width - Cmd+Option+C (no Chrome conflict)
    'cmd+option+c': {
      label: 'Auto-fit Column Width',
      action: () => VexcelContextMenu.autoResizeColumns()
    }
  };

  return { shortcuts };
})();
