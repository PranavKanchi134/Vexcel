// Local performance timing and rolling command metrics.

const VexcelPerf = (() => {
  const active = new Map();
  const stats = new Map();
  const MAX_SAMPLES = 20;
  const SLOW_MS = 450;

  function now() {
    return (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
  }

  function start(commandId, meta = {}) {
    active.set(commandId, {
      commandId,
      startAt: now(),
      marks: [],
      meta,
      strategies: []
    });
  }

  function mark(commandId, phase, extra = {}) {
    const run = active.get(commandId);
    if (!run) return;
    run.marks.push({ phase, at: now(), ...extra });
    if (extra.strategy) run.strategies.push(extra.strategy);
  }

  function finish(commandId, result = {}) {
    const run = active.get(commandId);
    if (!run) return null;
    active.delete(commandId);

    const durationMs = Math.round(now() - run.startAt);
    const record = {
      commandId,
      durationMs,
      ok: !!result.ok,
      strategy: result.strategy || run.strategies[run.strategies.length - 1] || '',
      phases: run.marks,
      slow: durationMs >= SLOW_MS
    };

    const entries = stats.get(commandId) || [];
    entries.push(record);
    while (entries.length > MAX_SAMPLES) entries.shift();
    stats.set(commandId, entries);
    publishSummary();
    return record;
  }

  function getSummary() {
    const commands = {};
    for (const [commandId, entries] of stats.entries()) {
      const total = entries.reduce((sum, entry) => sum + entry.durationMs, 0);
      const last = entries[entries.length - 1];
      commands[commandId] = {
        avgMs: Math.round(total / entries.length),
        lastMs: last.durationMs,
        lastStrategy: last.strategy || '',
        samples: entries.length,
        slowCount: entries.filter(entry => entry.slow).length
      };
    }
    return { commands, updatedAt: Date.now() };
  }

  function publishSummary() {
    try {
      chrome.runtime.sendMessage({
        type: 'PERF_EVENT',
        summary: getSummary()
      });
    } catch (err) {}
  }

  return { start, mark, finish, getSummary };
})();
