// Shortcut registry - central Map<combo, action> lookup

const VexcelShortcutRegistry = (() => {
  /** @type {Map<string, object>} */
  const registry = new Map();

  /**
   * Register all shortcuts from category modules.
   */
  function initialize() {
    const categories = [
      VexcelCellOperations,
      VexcelFormatting,
      VexcelNavigation,
      VexcelSelection,
      VexcelEditing,
      VexcelRowColumn,
      VexcelFormula
    ];

    for (const category of categories) {
      if (category && category.shortcuts) {
        for (const [combo, shortcut] of Object.entries(category.shortcuts)) {
          registry.set(combo, shortcut);
        }
      }
    }

    console.log(`[Vexcel] Registered ${registry.size} shortcuts`);
  }

  /**
   * Look up a shortcut by its combo string.
   * @param {string} combo - Normalized combo string, e.g. "cmd+r"
   * @returns {object|null} - The shortcut definition, or null if not found
   */
  function lookup(combo) {
    return registry.get(combo) || null;
  }

  /**
   * Check if a combo is registered.
   */
  function has(combo) {
    return registry.has(combo);
  }

  /**
   * Get all registered shortcuts (for debugging/settings UI).
   */
  function getAll() {
    return new Map(registry);
  }

  /**
   * Check if a shortcut requires opt-in.
   */
  function isOptIn(combo) {
    const shortcut = registry.get(combo);
    return shortcut ? !!shortcut.optIn : false;
  }

  return { initialize, lookup, has, getAll, isOptIn };
})();
