// Selection shortcuts: autofilter toggle, etc.

const VexcelSelection = (() => {
  const shortcuts = {
    // Toggle autofilter - Cmd+Shift+L
    'cmd+shift+l': {
      label: 'Toggle Autofilter',
      action: () => VexcelMenuNavigator.clickMenuPath(['Data', 'Filter'])
    }

    // Select row (Shift+Space) - passes through natively
    // Select column (Ctrl+Space) - passes through natively
    // Extend selection (Shift+arrows) - passes through natively
    // Select to end of data (Cmd+Shift+arrows) - passes through natively
  };

  return { shortcuts };
})();
