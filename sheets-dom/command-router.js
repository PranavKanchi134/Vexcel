// Maps high-frequency shortcuts to aggressive execution plans.

const VexcelCommandRouter = (() => {
  function resolve(commandId, options = {}) {
    const plans = {
      fillDown: {
        id: 'fillDown',
        query: 'Fill down',
        path: [],
        allowlist: ['fill down'],
        prefer: 'toolFinder'
      },
      fillRight: {
        id: 'fillRight',
        query: 'Fill right',
        path: [],
        allowlist: ['fill right'],
        prefer: 'toolFinder'
      },
      pasteValues: {
        id: 'pasteValues',
        query: 'Values only',
        path: ['Edit', 'Paste special', 'Values only'],
        allowlist: ['values only'],
        prefer: 'toolFinder'
      },
      findReplace: {
        id: 'findReplace',
        query: 'Find and replace',
        path: ['Edit', 'Find and replace'],
        allowlist: ['find and replace'],
        prefer: 'menu'
      },
      formatCycle: {
        id: 'formatCycle',
        prefer: 'toolFinder'
      },
      fontColorCycle: {
        id: 'fontColorCycle',
        prefer: 'direct'
      },
      autoColorSelection: {
        id: 'autoColorSelection',
        prefer: 'direct'
      }
    };

    const resolved = plans[commandId] || {
      id: commandId,
      query: options.query || '',
      path: options.path || [],
      allowlist: options.allowlist || [],
      prefer: options.prefer || 'toolFinder'
    };

    return {
      ...resolved,
      commandKey: `${resolved.id}::${resolved.query || ''}::${(resolved.path || []).join('>')}`
    };
  }

  return { resolve };
})();
