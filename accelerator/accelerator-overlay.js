// Vexcel Accelerator Overlay
// Excel-style ribbon KeyTips: Option tap → [H][N][M][A][W] tab bar
// → toolbar badge mode (yellow badges on actual Google Sheets toolbar buttons)
// → list panel mode (compact key-hint grid for menu-driven actions)

const VexcelAcceleratorOverlay = (() => {

  // ── State ─────────────────────────────────────────────────
  let shadowHost = null;
  let shadowRoot = null;
  let contentEl  = null;
  let active         = false;
  let currentPath    = [];  // e.g. [] | ['h'] | ['h','a']
  let optionKeyDown  = null;
  let otherKeyPressed = false;

  // ── Init ──────────────────────────────────────────────────
  function init() {
    const existing = document.getElementById('vexcel-accelerator-host');
    if (existing) existing.remove();

    shadowHost = document.createElement('div');
    shadowHost.id = 'vexcel-accelerator-host';
    shadowHost.style.cssText =
      'position:fixed;top:0;left:0;width:100%;height:100%;' +
      'z-index:2147483647;pointer-events:none;display:none;';

    const attach = () => {
      document.body.appendChild(shadowHost);
      shadowRoot = shadowHost.attachShadow({ mode: 'open' });

      const style = document.createElement('style');
      style.textContent = CSS;
      shadowRoot.appendChild(style);

      contentEl = document.createElement('div');
      contentEl.id = 'vc-root';
      shadowRoot.appendChild(contentEl);
    };

    if (document.body) attach();
    else document.addEventListener('DOMContentLoaded', attach);
  }

  // ── Key handlers ──────────────────────────────────────────
  function handleKeyDown(e) {

    // ─── When accelerator is active, ALWAYS handle keys ─────
    // Skip shouldIntercept() — the Alt keydown may have caused
    // Google Sheets to show its own menu bar, which triggers
    // isDialogOpen(). We must not let that deactivate us.
    if (active) {
      e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();

      if (e.key === 'Escape' || e.key === 'Backspace') {
        if (currentPath.length > 0) {
          // Go up one level (both Escape and Backspace)
          currentPath.pop();
          render();
        } else {
          // At top level — close entirely
          deactivate();
        }
        return true;
      }

      // Ignore bare modifier presses while active
      if (['Alt', 'Meta', 'Control', 'Shift'].includes(e.key)) return true;

      navigateKey(e.key.toLowerCase());
      return true;
    }

    // ─── Not active — track Option key tap ───────────────────
    // Do NOT gate on shouldIntercept() here — the accelerator must
    // always be activatable, even if Sheets has menus/dialogs open
    // (we dismiss those on activation anyway).

    // Intercept the Alt/Option keydown so that Google Sheets does
    // NOT activate its own menu-bar alt behaviour.
    if (e.key === 'Alt' && !e.metaKey && !e.ctrlKey && !e.shiftKey) {
      optionKeyDown = e.timeStamp;
      otherKeyPressed = false;
      e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
      return true;
    }

    // Any other key while Option is held → not a bare tap
    if (optionKeyDown) otherKeyPressed = true;

    return false;
  }

  function handleKeyUp(e) {
    if (e.key === 'Alt' && optionKeyDown && !otherKeyPressed && !active) {
      const dur = e.timeStamp - optionKeyDown;
      optionKeyDown = null;
      if (dur < 500) {
        activate();
        e.preventDefault();
        return true;
      }
    }
    if (e.key === 'Alt') optionKeyDown = null;
    return false;
  }

  // ── Activation ────────────────────────────────────────────
  function activate() {
    active = true;
    currentPath = [];
    console.log('[Vexcel] Accelerator activated');
    render();
  }

  function deactivate() {
    active = false;
    currentPath = [];
    if (shadowHost) shadowHost.style.display = 'none';
    if (contentEl)  contentEl.innerHTML = '';
  }

  // ── Navigation ────────────────────────────────────────────
  function navigateKey(key) {
    const map = getNodeAtPath(currentPath);
    const target = map ? map[key] : null;

    if (!target) {
      console.log(`[Vexcel] Accelerator: no match for '${key}' at path [${currentPath}]`);
      deactivate();
      return;
    }

    console.log(`[Vexcel] Accelerator: ${key} → ${target.label}`);

    if (target.children) {
      currentPath.push(key);
      render();
    } else if (target.toolbar) {
      // Try toolbar click first, fall back to action if button not found
      console.log(`[Vexcel] Clicking toolbar: ${target.toolbar}`);
      deactivate();
      const clicked = VexcelToolbar.clickButton(target.toolbar);
      if (!clicked && target.action) {
        console.log(`[Vexcel] Toolbar fallback to action: ${target.label}`);
        execAction(target);
      }
    } else if (target.action) {
      console.log(`[Vexcel] Executing action: ${target.label}`);
      deactivate();
      execAction(target);
    } else {
      deactivate();
    }
  }

  function execAction(target) {
    const result = target.action();
    if (result && result.then) {
      result.then(ok => {
        if (!ok) console.warn(`[Vexcel] Action failed: ${target.label}`);
      }).catch(err => {
        console.error(`[Vexcel] Action error: ${target.label}`, err);
      });
    }
  }

  // Walk the tree to the children-map at the current path
  function getNodeAtPath(path) {
    const tree = VexcelAcceleratorKeys.getTree();
    let node = tree;
    for (const key of path) {
      if (!node[key] || !node[key].children) return null;
      node = node[key].children;
    }
    return node;
  }

  // Return { node (children map), mode, breadcrumbs }
  function getContext() {
    const tree = VexcelAcceleratorKeys.getTree();
    if (currentPath.length === 0) return { node: tree, mode: 'tabs', breadcrumbs: [] };

    let node = tree;
    let mode = 'list';
    const breadcrumbs = [];

    for (const key of currentPath) {
      if (!node[key]) return null;
      breadcrumbs.push({ key, label: node[key].label });
      mode = node[key].mode || 'list';
      node = node[key].children || {};
    }
    return { node, mode, breadcrumbs };
  }

  // ── Rendering ─────────────────────────────────────────────
  function render() {
    if (!contentEl) return;
    contentEl.innerHTML = '';
    shadowHost.style.display = 'block';

    const ctx = getContext();
    if (!ctx) { deactivate(); return; }

    if (ctx.mode === 'tabs') {
      renderTabBar(ctx.node);
    } else if (ctx.mode === 'toolbar') {
      renderBreadcrumb(ctx.breadcrumbs);
      renderToolbarBadges(ctx.node);
    } else {
      renderBreadcrumb(ctx.breadcrumbs);
      renderListPanel(ctx.node);
    }
  }

  // ─ Tab bar (level 0) ──────────────────────────────────────
  function renderTabBar(tabs) {
    const bar = mkEl('div', 'vc-tabbar');

    const logo = mkEl('span', 'vc-logo'); logo.textContent = '⌥ Vexcel'; bar.appendChild(logo);
    const div  = mkEl('span', 'vc-divider'); bar.appendChild(div);

    for (const [key, tab] of Object.entries(tabs)) {
      const t = mkEl('div', 'vc-tab');
      t.appendChild(badge(key));
      const lbl = mkEl('span', 'vc-tab-label'); lbl.textContent = tab.label; t.appendChild(lbl);
      if (tab.hint) { const h = mkEl('span', 'vc-tab-hint'); h.textContent = tab.hint; t.appendChild(h); }
      t.addEventListener('click', () => navigateKey(key));
      bar.appendChild(t);
    }

    const esc = mkEl('span', 'vc-esc'); esc.textContent = 'ESC'; bar.appendChild(esc);
    contentEl.appendChild(bar);
  }

  // ─ Breadcrumb strip ───────────────────────────────────────
  function renderBreadcrumb(crumbs) {
    const bar = mkEl('div', 'vc-tabbar');

    const logo = mkEl('span', 'vc-logo'); logo.textContent = '⌥ Vexcel'; bar.appendChild(logo);
    const div  = mkEl('span', 'vc-divider'); bar.appendChild(div);

    crumbs.forEach((c, i) => {
      if (i > 0) { const sep = mkEl('span','vc-sep'); sep.textContent='›'; bar.appendChild(sep); }
      const crumb = mkEl('span', i === crumbs.length - 1 ? 'vc-crumb-active' : 'vc-crumb');
      crumb.textContent = `[${c.key.toUpperCase()}] ${c.label}`;
      bar.appendChild(crumb);
    });

    if (currentPath.length > 0) {
      const back = mkEl('span', 'vc-back');
      back.textContent = '← Back';
      back.addEventListener('click', () => { currentPath.pop(); render(); });
      bar.appendChild(back);
    }

    const esc = mkEl('span', 'vc-esc'); esc.textContent = 'ESC'; bar.appendChild(esc);
    contentEl.appendChild(bar);
  }

  // ─ Toolbar badge mode ─────────────────────────────────────
  function renderToolbarBadges(items) {
    const overflow = [];

    for (const [key, node] of Object.entries(items)) {
      if (node.toolbar) {
        const btn = findToolbarButton(node.toolbar);
        if (btn) {
          const rect = btn.getBoundingClientRect();
          if (rect.width > 0) {
            const bdg = mkEl('div', 'vc-badge');
            bdg.textContent = node.children ? key.toUpperCase() + '…' : key.toUpperCase();
            const bw = node.children ? 26 : 18;
            bdg.style.top  = (rect.bottom + 3) + 'px';
            bdg.style.left = (rect.left + rect.width / 2 - bw / 2) + 'px';
            bdg.title = node.label;
            bdg.addEventListener('click', () => navigateKey(key));
            contentEl.appendChild(bdg);
            continue;
          }
        }
      }
      overflow.push([key, node]);
    }

    if (overflow.length > 0) renderLegend(overflow);
  }

  // ─ Legend panel (overflow from toolbar badge mode) ────────
  function renderLegend(items) {
    const panel = mkEl('div', 'vc-legend');
    for (const [key, node] of items) {
      const item = mkEl('div', 'vc-legend-item');
      item.appendChild(badge(key));
      const lbl = mkEl('span', 'vc-item-label');
      lbl.textContent = node.label;
      item.appendChild(lbl);
      item.addEventListener('click', () => navigateKey(key));
      panel.appendChild(item);
    }
    contentEl.appendChild(panel);
  }

  // ─ List panel ─────────────────────────────────────────────
  function renderListPanel(items) {
    const panel = mkEl('div', 'vc-list');
    for (const [key, node] of Object.entries(items)) {
      const item = mkEl('div', 'vc-item');
      item.appendChild(badge(key));
      const lbl = mkEl('span', 'vc-item-label'); lbl.textContent = node.label; item.appendChild(lbl);
      if (node.children) { const arr = mkEl('span', 'vc-arrow'); arr.textContent = '▸'; item.appendChild(arr); }
      item.addEventListener('click', () => navigateKey(key));
      panel.appendChild(item);
    }
    contentEl.appendChild(panel);
  }

  // ── Toolbar button finder ──────────────────────────────────
  function findToolbarButton(label) {
    const candidates = document.querySelectorAll('[aria-label]');
    for (const candidate of candidates) {
      if (!inToolbar(candidate)) continue;
      if (candidate.getAttribute('aria-label') === label) return candidate;
    }
    for (const candidate of candidates) {
      if (!inToolbar(candidate)) continue;
      if (candidate.getAttribute('aria-label').includes(label)) return candidate;
    }
    return null;
  }

  function inToolbar(node) {
    const r = node.getBoundingClientRect();
    return r.top < 180 && r.height > 4 && r.height < 64 && r.width > 4;
  }

  // ── DOM helpers ───────────────────────────────────────────
  // Renamed from el() to mkEl() to avoid shadowing with for-of loops
  function mkEl(tag, cls) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    return e;
  }

  function badge(key) {
    const b = mkEl('span', 'vc-key');
    b.textContent = key.toUpperCase();
    return b;
  }

  function isActive()   { return active; }

  // ── CSS ───────────────────────────────────────────────────
  const CSS = `
    :host { all: initial; }
    #vc-root { font-family: -apple-system, 'Segoe UI', sans-serif; }

    .vc-key {
      display: inline-flex; align-items: center; justify-content: center;
      min-width: 16px; height: 16px; padding: 0 3px;
      background: #f5e642; border: 1px solid #b8a800; border-radius: 2px;
      color: #111; font-size: 10px; font-weight: 800;
      box-shadow: 0 1px 2px rgba(0,0,0,.35); flex-shrink: 0;
      letter-spacing: .3px;
    }

    .vc-tabbar {
      position: fixed; top: 0; left: 0; right: 0; height: 34px;
      display: flex; align-items: center; gap: 2px;
      background: #1c1c2e; border-bottom: 2px solid #4a8fdb;
      padding: 0 10px; z-index: 1;
      box-shadow: 0 2px 10px rgba(0,0,0,.55); pointer-events: auto;
    }
    .vc-logo { font-size: 11px; font-weight: 700; color: #4a8fdb; padding-right: 10px; flex-shrink: 0; }
    .vc-divider { width: 1px; height: 18px; background: #383850; margin: 0 4px; flex-shrink: 0; }
    .vc-tab {
      display: flex; align-items: center; gap: 5px;
      padding: 3px 9px; border-radius: 3px; cursor: pointer;
      transition: background .1s; flex-shrink: 0;
    }
    .vc-tab:hover { background: rgba(255,255,255,.1); }
    .vc-tab-label { font-size: 11px; color: #ddd; font-weight: 500; }
    .vc-tab-hint { font-size: 10px; color: #666; }
    .vc-sep   { font-size: 11px; color: #444; margin: 0 3px; }
    .vc-crumb { font-size: 11px; color: #888; }
    .vc-crumb-active { font-size: 11px; color: #fff; font-weight: 600; }
    .vc-back {
      font-size: 10px; color: #4a8fdb; cursor: pointer;
      padding: 2px 7px; border: 1px solid #4a8fdb; border-radius: 3px;
      margin-left: 6px; transition: background .1s; flex-shrink: 0;
      pointer-events: auto;
    }
    .vc-back:hover { background: rgba(74,143,219,.2); }
    .vc-esc {
      margin-left: auto; font-size: 10px; color: #555;
      padding: 2px 5px; border: 1px solid #383850; border-radius: 2px; flex-shrink: 0;
    }

    .vc-badge {
      position: fixed;
      display: inline-flex; align-items: center; justify-content: center;
      min-width: 16px; height: 15px; padding: 0 3px;
      background: #f5e642; border: 1px solid #b8a800; border-radius: 2px;
      color: #111; font-size: 10px; font-weight: 800;
      box-shadow: 0 2px 6px rgba(0,0,0,.45);
      cursor: pointer; pointer-events: auto; white-space: nowrap;
      letter-spacing: .3px; z-index: 2; transition: transform .08s;
    }
    .vc-badge:hover { transform: scale(1.15); }

    .vc-legend {
      position: fixed; bottom: 8px; left: 50%; transform: translateX(-50%);
      background: #1c1c2e; border: 1px solid #383860;
      border-radius: 6px; padding: 6px 10px;
      display: flex; flex-wrap: wrap; gap: 4px; justify-content: center;
      box-shadow: 0 4px 20px rgba(0,0,0,.6); pointer-events: auto;
      max-width: 560px;
    }

    .vc-list {
      position: fixed; top: 42px; left: 50%; transform: translateX(-50%);
      background: #1c1c2e; border: 1px solid #383860; border-radius: 7px;
      padding: 8px;
      display: grid; grid-template-columns: repeat(auto-fill, minmax(140px, 1fr));
      gap: 3px; min-width: 300px; max-width: 580px;
      box-shadow: 0 6px 24px rgba(0,0,0,.65); pointer-events: auto;
    }

    .vc-legend-item, .vc-item {
      display: flex; align-items: center; gap: 6px;
      padding: 5px 7px; border-radius: 4px; cursor: pointer;
      transition: background .1s; min-width: 0;
    }
    .vc-legend-item:hover, .vc-item:hover { background: rgba(74,143,219,.22); }
    .vc-item-label {
      font-size: 11px; color: #ccc; flex: 1; overflow: hidden;
      text-overflow: ellipsis; white-space: nowrap;
    }
    .vc-arrow { font-size: 9px; color: #555; margin-left: auto; flex-shrink: 0; }
  `;

  return { init, handleKeyDown, handleKeyUp, isActive, activate, deactivate };

})();
