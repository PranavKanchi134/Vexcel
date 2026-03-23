// Formatting shortcuts: number formats, clear formatting, etc.

const VexcelFormatting = (() => {
  function runFormatCommand(label, query, path) {
    return VexcelMenuNavigator.runCommand(query, path, { prefer: 'toolFinder' }).then(outcome => (
      outcome.ok
        ? { ok: true, strategy: outcome.strategy, verified: true, message: label }
        : { ok: false, message: `${label} failed`, reason: outcome.reason }
    ));
  }

  const shortcuts = {
    // Format Cells dialog - Cmd+1
    'cmd+1': {
      label: 'Format Cells',
      action: () => runFormatCommand('Format Cells', 'Custom number format', ['Format', 'Number', 'Custom number format'])
    },

    // Clear formatting - Cmd+Shift+X (note: Cmd+\ also works in GSheets natively)
    'cmd+shift+x': {
      label: 'Clear Formatting',
      action: () => runFormatCommand('Clear Formatting', 'Clear formatting', ['Format', 'Clear formatting'])
    },

    // General format - Cmd+Shift+~ (tilde)
    'cmd+shift+~': {
      label: 'General Format',
      action: () => runFormatCommand('General Format', 'Automatic', ['Format', 'Number', 'Automatic'])
    },

    // Currency format - Cmd+Shift+$
    'cmd+shift+$': {
      label: 'Currency Format',
      action: () => runFormatCommand('Currency Format', 'Currency', ['Format', 'Number', 'Currency'])
    },

    // Date format - Cmd+Shift+#
    'cmd+shift+#': {
      label: 'Date Format',
      action: () => runFormatCommand('Date Format', 'Date', ['Format', 'Number', 'Date'])
    },

    // Number format - Cmd+Shift+!
    'cmd+shift+!': {
      label: 'Number Format',
      action: () => runFormatCommand('Number Format', 'Number', ['Format', 'Number', 'Number'])
    },

    // Scientific format - Cmd+Shift+^
    'cmd+shift+^': {
      label: 'Scientific Format',
      action: () => runFormatCommand('Scientific Format', 'Scientific', ['Format', 'Number', 'Scientific'])
    },

    // Increase decimal places - Cmd+Shift+] (not standard Excel but intuitive)
    'cmd+shift+]': {
      label: 'Increase Decimals',
      action: () => VexcelToolbar.clickButton('Increase decimal places')
    },

    // Decrease decimal places - Cmd+Shift+[
    'cmd+shift+[': {
      label: 'Decrease Decimals',
      action: () => VexcelToolbar.clickButton('Decrease decimal places')
    },

    'ctrl+option+f': {
      label: 'Format Cycle',
      commandId: 'formatCycle',
      action: () => VexcelFinanceTools.cycleFormat()
    },

    'ctrl+option+c': {
      label: 'Font Color Cycle',
      commandId: 'fontColorCycle',
      action: () => VexcelFinanceTools.cycleFontColor()
    },

    'ctrl+option+a': {
      label: 'Auto-Color Selection',
      commandId: 'autoColorSelection',
      action: () => VexcelFinanceTools.autoColorSelection()
    },

    // Bold (Cmd+B), Italic (Cmd+I), Underline (Cmd+U) - pass through natively
    // Percent (Cmd+Shift+%) - pass through natively
  };

  return { shortcuts };
})();
