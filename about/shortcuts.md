# Shortcut Inventory

This file summarizes the main registered shortcut set in the package. It does not attempt to fully flatten the much larger accelerator tree, but it does describe its structure and purpose.

## Standard Registered Shortcuts

### Cell Operations

- `Cmd+D`: Fill Down
- `Cmd+R`: Fill Right
- `Cmd+Option+V`: Paste Special
- `Cmd+Shift+V`: Paste Values
- `Cmd+H`: Find and Replace
- `Cmd+N`: New Sheet, opt-in because it overrides Chrome new window
- `Cmd+T`: Alternating Colors, opt-in because it overrides Chrome new tab

### Formatting

- `Cmd+1`: Format Cells
- `Cmd+Shift+X`: Clear Formatting
- `Cmd+Shift+~`: General Format
- `Cmd+Shift+$`: Currency Format
- `Cmd+Shift+#`: Date Format
- `Cmd+Shift+!`: Number Format
- `Cmd+Shift+^`: Scientific Format
- `Cmd+Shift+]`: Increase Decimals
- `Cmd+Shift+[` : Decrease Decimals

### Navigation

- `Cmd+G`: Go To
- `Ctrl+G`: Go To
- `Option+Right`: Next Sheet
- `Option+Left`: Previous Sheet
- `F5`: Go To
- `Cmd+Shift+;`: Insert Time

### Editing

- `Ctrl+;`: Insert Date
- `Cmd+;`: Insert Time

### Row and Column Operations

- `Ctrl+Shift++`: Insert Rows or Columns
- `Cmd+-`: Delete Rows or Columns
- `Ctrl+9`: Hide Rows
- `Ctrl+0`: Hide Columns
- `Ctrl+Shift+(`: Unhide Rows
- `Ctrl+Shift+)`: Unhide Columns
- `Cmd+Option+R`: Auto-fit Row Height
- `Cmd+Option+C`: Auto-fit Column Width

### Selection

- `Cmd+Shift+L`: Toggle Autofilter

### Formula Tools

- `Ctrl+\``: Toggle Formula View
- `Cmd+Shift+Enter`: Array Formula, passed through natively
- `Ctrl+[`: Trace Precedents
- `Ctrl+]`: Trace Dependents

## Accelerator Overlay

The accelerator system is Vexcel's Excel-ribbon-inspired mode.

Activation:

- Tap `Option` by itself

Top-level branches shown by the overlay:

- `H`: Home
- `N`: Insert
- `M`: Formulas
- `A`: Data
- `W`: View

Behavior:

- Some accelerator steps render badges directly on real Google Sheets toolbar buttons.
- Other steps render a compact list panel for deeper menu-driven actions.
- When a leaf node is selected, Vexcel tries a toolbar click first when possible and falls back to a menu or tool-finder action if needed.

Primary implementation files:

- [`shortcuts/categories/accelerator.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/shortcuts/categories/accelerator.js)
- [`accelerator/accelerator-overlay.js`](/Users/pranavkanchi/Desktop/projects/coding/Vexcel/accelerator/accelerator-overlay.js)

## Shortcut Handling Rules

- Shortcuts are normalized into canonical combos such as `cmd+r` or `ctrl+;`.
- Modifier-only presses are ignored.
- Shortcuts do not fire when dialogs or non-grid text inputs are active.
- Browser-conflicting shortcuts can be blocked and replaced by Vexcel behavior.
- A few shortcuts are marked `passThrough` so Google Sheets can handle them natively.
