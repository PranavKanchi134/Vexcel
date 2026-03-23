// Key event normalization utilities

const VexcelKeyUtils = (() => {
  /**
   * Normalize a KeyboardEvent into a canonical combo string.
   * Examples: "cmd+r", "cmd+shift+$", "ctrl+;", "f2", "option+right"
   */
  function normalizeEvent(e) {
    const parts = [];

    if (e.metaKey) parts.push('cmd');
    if (e.ctrlKey) parts.push('ctrl');
    if (e.altKey) parts.push('option');
    if (e.shiftKey) parts.push('shift');

    const key = normalizeKey(e);
    if (key && !['Meta', 'Control', 'Alt', 'Shift'].includes(e.key)) {
      parts.push(key);
    }

    return parts.join('+');
  }

  /**
   * Normalize the key portion of the event.
   */
  function normalizeKey(e) {
    const key = e.key;

    // Function keys
    if (/^F\d+$/.test(key)) return key.toLowerCase();

    // Arrow keys
    const arrowMap = {
      ArrowUp: 'up',
      ArrowDown: 'down',
      ArrowLeft: 'left',
      ArrowRight: 'right'
    };
    if (arrowMap[key]) return arrowMap[key];

    // Special keys
    const specialMap = {
      Enter: 'enter',
      Escape: 'escape',
      Backspace: 'backspace',
      Delete: 'delete',
      Tab: 'tab',
      ' ': 'space',
      Home: 'home',
      End: 'end',
      PageUp: 'pageup',
      PageDown: 'pagedown'
    };
    if (specialMap[key] !== undefined) return specialMap[key];

    // On Mac, Alt/Option + letter produces special characters (e.g., Option+V = √).
    // When Alt is held, fall back to e.code to get the underlying letter key.
    if (e.altKey && e.code && e.code.startsWith('Key')) {
      return e.code.slice(3).toLowerCase(); // 'KeyV' -> 'v'
    }
    if (e.altKey && e.code && e.code.startsWith('Digit')) {
      return e.code.slice(5); // 'Digit1' -> '1'
    }

    // For shifted symbols, use the actual character produced
    // e.g., Shift+4 produces '$', Shift+` produces '~'
    if (key.length === 1) return key.toLowerCase();

    return key.toLowerCase();
  }

  /**
   * Check if the Option key was tapped (pressed and released quickly
   * without other keys being pressed).
   */
  function isOptionTap(downEvent, upEvent, maxDuration = 500) {
    if (downEvent.key !== 'Alt') return false;
    if (upEvent.key !== 'Alt') return false;
    return (upEvent.timeStamp - downEvent.timeStamp) < maxDuration;
  }

  /**
   * Check if a modifier-only event (no printable key).
   */
  function isModifierOnly(e) {
    return ['Meta', 'Control', 'Alt', 'Shift'].includes(e.key);
  }

  return { normalizeEvent, normalizeKey, isOptionTap, isModifierOnly };
})();
