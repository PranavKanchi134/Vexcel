# About Vexcel

Vexcel is a Manifest V3 Chrome extension that brings Excel-style keyboard shortcuts to Google Sheets on macOS.

The extension does four main things:

- Intercepts keyboard input inside Google Sheets before Sheets or Chrome handle it.
- Maps Excel-like shortcuts to Google Sheets menu, toolbar, editor, and context-menu actions.
- Provides an Excel-style accelerator overlay triggered by tapping `Option`.
- Exposes a popup for enabling or disabling shortcuts and opting into browser-conflicting overrides.

## Package Snapshot

- Extension name: `Vexcel - Excel Shortcuts for Google Sheets`
- Version: `1.0.0`
- Runtime target: Google Sheets at `https://docs.google.com/spreadsheets/*`
- Manifest: [`manifest.json`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/manifest.json)

## Start Here

- Read [`about/architecture.md`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/about/architecture.md) for the runtime flow.
- Read [`about/shortcuts.md`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/about/shortcuts.md) for the supported shortcut inventory.
- Read [`about/file-map.md`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/about/file-map.md) for a folder-by-folder guide.
- Read [`about/codex-google-sheets.md`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/about/codex-google-sheets.md) for the direct Codex-to-Google-Sheets workflow.

## Core Entry Points

- [`manifest.json`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/manifest.json): declares the extension, content scripts, popup, icons, and background service worker.
- [`content.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/content.js): keyboard interception pipeline and top-level extension initialization.
- [`service-worker.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/service-worker.js): persistent settings, state updates, and icon badge handling.
- [`popup/popup.html`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/popup/popup.html): user-facing popup UI.

## Behavior Summary

- Most shortcuts are implemented by clicking real Google Sheets menus or toolbar buttons.
- Some actions fall back to synthetic keyboard events, tool-finder searches, or canvas-style context-menu interactions when Sheets does not expose a simple DOM control.
- The accelerator overlay lives in the top frame, while the editable grid often lives in an iframe, so the package includes explicit cross-frame message passing for `Option` tap activation and follow-up key presses.

## Current Scope

- Focused on Google Sheets only.
- Focused on macOS-style Excel shortcut expectations.
- Includes standard shortcuts plus a larger accelerator tree for ribbon-style workflows.
