# File Map

## Root

- [`manifest.json`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/manifest.json): extension definition and load order.
- [`content.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/content.js): initialization and keyboard interception entry point.
- [`service-worker.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/service-worker.js): background state and browser-action updates.

## `accelerator/`

- [`accelerator/accelerator-overlay.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/accelerator/accelerator-overlay.js): Excel-style `Option` accelerator overlay, navigation state, and rendering.

## `icons/`

- Active and inactive icon assets used by the browser action.

## `popup/`

- [`popup/popup.html`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/popup/popup.html): popup structure.
- [`popup/popup.css`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/popup/popup.css): popup styling.
- [`popup/popup.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/popup/popup.js): popup state sync and toggle handlers.

## `shortcuts/`

- [`shortcuts/shortcut-registry.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/shortcuts/shortcut-registry.js): merges category definitions into a single lookup map.
- [`shortcuts/shortcut-actions.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/shortcuts/shortcut-actions.js): executes shortcut handlers safely.

### `shortcuts/categories/`

- [`shortcuts/categories/cell-operations.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/shortcuts/categories/cell-operations.js): fill, paste, find, and a few browser-override shortcuts.
- [`shortcuts/categories/formatting.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/shortcuts/categories/formatting.js): number-format and formatting commands.
- [`shortcuts/categories/navigation.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/shortcuts/categories/navigation.js): go-to and sheet navigation.
- [`shortcuts/categories/selection.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/shortcuts/categories/selection.js): filter-related selection tooling.
- [`shortcuts/categories/editing.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/shortcuts/categories/editing.js): date and time insertion helpers.
- [`shortcuts/categories/row-column.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/shortcuts/categories/row-column.js): row or column insertion, deletion, hide, unhide, and auto-fit actions.
- [`shortcuts/categories/formula.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/shortcuts/categories/formula.js): formula-view and trace helpers.
- [`shortcuts/categories/accelerator.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/shortcuts/categories/accelerator.js): ribbon-style accelerator tree.

## `sheets-dom/`

- [`sheets-dom/menu-navigator.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/sheets-dom/menu-navigator.js): menu traversal and tool-finder fallback.
- [`sheets-dom/toolbar-actions.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/sheets-dom/toolbar-actions.js): toolbar button discovery and interaction.
- [`sheets-dom/cell-editor.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/sheets-dom/cell-editor.js): formula bar, name box, and editor helpers.
- [`sheets-dom/context-menu.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/sheets-dom/context-menu.js): right-click flows for row and column operations on canvas-based headers.
- [`sheets-dom/color-picker.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/sheets-dom/color-picker.js): color selection helpers.
- [`sheets-dom/trace-arrows.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/sheets-dom/trace-arrows.js): custom precedent and dependent visualization.

## `utils/`

- [`utils/key-utils.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/utils/key-utils.js): event normalization and modifier helpers.
- [`utils/dom-utils.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/utils/dom-utils.js): centralized selectors and interception guards.
- [`utils/messaging.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/utils/messaging.js): message helpers for background communication.
