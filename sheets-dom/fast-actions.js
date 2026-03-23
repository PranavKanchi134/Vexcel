// Dedicated fast paths for latency-sensitive shortcuts.

const VexcelFastActions = (() => {
  function fillDown() {
    return VexcelFillEngine.fillDown().then(result => (
      result.ok ? result : executeCommandPlan('fillDown', 'Fill Down')
    ));
  }

  function fillRight() {
    return VexcelFillEngine.fillRight().then(result => (
      result.ok ? result : executeCommandPlan('fillRight', 'Fill Right')
    ));
  }

  function pasteValues() {
    return executeCommandPlan('pasteValues', 'Paste Values');
  }

  function findReplace() {
    return executeCommandPlan('findReplace', 'Find & Replace');
  }

  async function executeCommandPlan(commandId, label, fallback) {
    const plan = VexcelCommandRouter.resolve(commandId);
    const outcome = await VexcelMenuNavigator.runCommand(plan.query, plan.path, {
      prefer: plan.prefer,
      allowlist: plan.allowlist,
      commandKey: plan.commandKey
    });

    if (outcome.ok) {
      return {
        ok: true,
        strategy: outcome.strategy,
        verified: outcome.verified !== false,
        durationMs: outcome.durationMs,
        message: label
      };
    }

    if (fallback) return fallback();
    return {
      ok: false,
      message: `${label} failed`,
      reason: outcome.reason || 'command did not resolve'
    };
  }

  return { fillDown, fillRight, pasteValues, findReplace };
})();
