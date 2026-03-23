// Editing shortcuts: insert date/time, etc.

const VexcelEditing = (() => {
  const shortcuts = {
    // Insert current date - Ctrl+;
    'ctrl+;': {
      label: 'Insert Date',
      action: () => {
        const today = new Date();
        const month = String(today.getMonth() + 1).padStart(2, '0');
        const day = String(today.getDate()).padStart(2, '0');
        const year = today.getFullYear();
        const dateStr = `${month}/${day}/${year}`;
        // Enter edit mode first if not already editing
        if (!VexcelCellEditor.isEditing()) {
          const editor = VexcelCellEditor.getEditorElement();
          if (editor) editor.focus();
        }
        VexcelCellEditor.insertText(dateStr);
      }
    },

    // Insert current time - Cmd+;
    'cmd+;': {
      label: 'Insert Time',
      action: () => {
        const now = new Date();
        const hours = now.getHours();
        const minutes = String(now.getMinutes()).padStart(2, '0');
        const ampm = hours >= 12 ? 'PM' : 'AM';
        const h12 = hours % 12 || 12;
        const timeStr = `${h12}:${minutes} ${ampm}`;
        // Enter edit mode first if not already editing
        if (!VexcelCellEditor.isEditing()) {
          const editor = VexcelCellEditor.getEditorElement();
          if (editor) editor.focus();
        }
        VexcelCellEditor.insertText(timeStr);
      }
    }

    // F2 (edit cell) - passes through natively
    // Enter (confirm edit) - passes through natively
    // Escape (cancel edit) - passes through natively
    // F4 (toggle absolute reference) - passes through natively
  };

  return { shortcuts };
})();
