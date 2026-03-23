# Architecture

## Runtime Model

Vexcel is a browser extension with three main runtime surfaces:

1. Background service worker
2. Popup UI
3. Content scripts injected into Google Sheets

The manifest loads a long ordered content-script chain at `document_start`, which lets Vexcel install its capture-phase keyboard listeners before Google Sheets registers many of its own handlers.

## Execution Flow

1. [`manifest.json`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/manifest.json) injects utility modules, Sheets DOM helpers, shortcut categories, the accelerator overlay, and finally [`content.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/content.js).
2. [`content.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/content.js) initializes the shortcut registry, overlay, trace-arrow UI, and global key listeners.
3. A keydown event is normalized by [`utils/key-utils.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/utils/key-utils.js), looked up in [`shortcuts/shortcut-registry.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/shortcuts/shortcut-registry.js), filtered through DOM context checks from [`utils/dom-utils.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/utils/dom-utils.js), and executed through [`shortcuts/shortcut-actions.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/shortcuts/shortcut-actions.js).
4. Execution usually delegates to a Google Sheets interaction layer in `sheets-dom/`.

## Major Components

### Background State

[`service-worker.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/service-worker.js) stores `vexcel_settings` in `chrome.storage.local`.

It is responsible for:

- Initializing defaults on install
- Toggling enabled or disabled state
- Updating optional shortcut overrides
- Broadcasting state changes to open Google Sheets tabs
- Changing the extension icon and badge text between `ON` and `OFF`

### Popup UI

[`popup/popup.html`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/popup/popup.html), [`popup/popup.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/popup/popup.js), and [`popup/popup.css`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/popup/popup.css) provide a lightweight control panel.

The popup currently supports:

- Master enable or disable toggle
- Opt-in overrides for `Cmd+N` and `Cmd+T`
- A short quick-reference list

### Content Script Pipeline

[`content.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/content.js) is the orchestrator.

Key responsibilities:

- Register all shortcut categories
- Install capture-phase `keydown` and `keyup` listeners
- Coordinate top-frame and iframe accelerator state
- Fetch settings from the service worker
- Respect live settings updates while Sheets tabs are already open
- Block browser-native actions when a mapped shortcut should win

### Cross-Frame Handling

Google Sheets often places the editing grid in an iframe while menus, toolbar controls, and overlay UI live in the top frame.

Vexcel handles that split by:

- Running content scripts in all frames
- Keeping the accelerator overlay only in the top frame
- Forwarding `Option` tap activation and accelerator keystrokes across frames via `postMessage`
- Broadcasting overlay active state back to iframes

This is one of the more important architectural choices in the package.

### Sheets Interaction Layer

The `sheets-dom/` folder acts as the adapter between shortcut intent and Google Sheets UI behavior.

- [`sheets-dom/menu-navigator.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/sheets-dom/menu-navigator.js): opens top-level menus, finds visible menu items, hovers submenus, and clicks final actions.
- [`sheets-dom/toolbar-actions.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/sheets-dom/toolbar-actions.js): finds and clicks toolbar buttons by label.
- [`sheets-dom/cell-editor.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/sheets-dom/cell-editor.js): works with the active editor and the name box.
- [`sheets-dom/context-menu.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/sheets-dom/context-menu.js): performs row or column operations that require coordinate-based right-click behavior.
- [`sheets-dom/color-picker.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/sheets-dom/color-picker.js): opens and navigates Sheets color controls.
- [`sheets-dom/trace-arrows.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/sheets-dom/trace-arrows.js): custom precedent and dependent tracing helpers.

### Shortcut System

Shortcuts are defined in category modules under `shortcuts/categories/` and merged into a central map by [`shortcuts/shortcut-registry.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/shortcuts/shortcut-registry.js).

Registered standard shortcut categories:

- Cell operations
- Formatting
- Navigation
- Selection
- Editing
- Row and column operations
- Formula tools

The accelerator feature is defined separately in [`shortcuts/categories/accelerator.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/shortcuts/categories/accelerator.js) and rendered by [`accelerator/accelerator-overlay.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/accelerator/accelerator-overlay.js).

## Default Settings

The default extension state is:

- `enabled: true`
- `cmd+n: false`
- `cmd+t: false`

Those two shortcuts are disabled by default because they override core Chrome actions.

## Notable Implementation Strategies

- Menu-first execution: many actions try the real Google Sheets menu structure first.
- Fallbacks when needed: tool finder search, toolbar clicks, synthetic keyboard events, and canvas-coordinate context-menu actions.
- Safe interception rules: shortcuts are skipped when dialogs or non-grid text inputs are active.
- Minimal messaging surface: state fetch and state update messages are centralized through [`utils/messaging.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/utils/messaging.js).
