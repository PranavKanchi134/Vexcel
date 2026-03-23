// Shortcut action execution layer

const VexcelShortcutActions = (() => {
  function normalizeResult(result) {
    if (result && typeof result === 'object' && Object.prototype.hasOwnProperty.call(result, 'ok')) {
      return result;
    }
    if (result === false) return { ok: false };
    return { ok: true };
  }

  /**
   * Execute a shortcut action.
   * @param {object} shortcut - The shortcut definition from the registry
   * @param {KeyboardEvent} event - The original keyboard event
   * @returns {Promise<object>} - Normalized execution outcome
   */
  async function execute(shortcut, event) {
    if (!shortcut || !shortcut.action) return { ok: false };

    // If this shortcut should pass through, don't prevent default
    if (shortcut.passThrough) return { ok: false, reason: 'passThrough' };

    try {
      const result = shortcut.action(event);
      // Handle both sync and async actions
      if (result instanceof Promise) {
        return normalizeResult(await result);
      }
      return normalizeResult(result);
    } catch (err) {
      console.error(`[Vexcel] Error executing shortcut "${shortcut.label}":`, err);
      return { ok: false, reason: err.message || 'exception' };
    }
  }

  return { execute };
})();
