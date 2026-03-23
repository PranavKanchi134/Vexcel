// Formula-related shortcuts

const VexcelFormula = (() => {
  const shortcuts = {
    // Toggle formula view - Ctrl+` (backtick)
    'ctrl+`': {
      label: 'Toggle Formula View',
      action: () => VexcelMenuNavigator.clickMenuPath(['View', 'Show formulas'])
    },

    // Array formula entry - Cmd+Shift+Enter
    'cmd+shift+enter': {
      label: 'Array Formula',
      action: () => {},
      passThrough: true
    },

    // Trace Precedents - Ctrl+[ (Excel: navigate to precedent cells)
    'ctrl+[': {
      label: 'Trace Precedents',
      action: () => VexcelTraceArrows.tracePrecedents()
    },

    // Trace Dependents - Ctrl+] (Excel: navigate to dependent cells)
    'ctrl+]': {
      label: 'Trace Dependents',
      action: () => VexcelTraceArrows.traceDependents()
    }
  };

  return { shortcuts };
})();
