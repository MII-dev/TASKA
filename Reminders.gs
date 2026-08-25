const DIGEST_TRIGGER_HANDLER = 'sendDailyDigest';
const DIGEST_HOUR = 9;

/**
 * Whether the daily digest trigger is currently installed.
 */
function hasDigestTrigger() {
  return ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === DIGEST_TRIGGER_HANDLER;
  });
}

/**
 * Whether the daily digest is on — lets the client render the bell
 * without touching ScriptApp itself.
 */
function getReminderStatus() {
  return { enabled: hasDigestTrigger() };
}

/**
 * Installs the time-based trigger, if not already installed.
 */
function enableDailyDigest() {
  if (!hasDigestTrigger()) {
    ScriptApp.newTrigger(DIGEST_TRIGGER_HANDLER)
      .timeBased()
      .atHour(DIGEST_HOUR)
      .everyDays(1)
      .create();
  }
  return { enabled: true };
}

/**
 * Removes every trigger pointing at sendDailyDigest — defensive against
 * duplicates rather than assuming there is exactly one.
 */
function disableDailyDigest() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === DIGEST_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(t);
    }
  });
  return { enabled: false };
}

function toggleDailyDigest() {
  return hasDigestTrigger() ? disableDailyDigest() : enableDailyDigest();
}

/**
 * Fired daily by the trigger installed in enableDailyDigest(). Emails a summary
 * of overdue and due-today tasks to the trigger owner. Sends nothing when
 * there is nothing to report, so an empty inbox stays empty.
 */
function sendDailyDigest() {
  const tasks = getTasks();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const active = tasks.filter(function (t) { return t.Статус !== 'Виконано'; });

  const overdue = [];
  const dueToday = [];

  active.forEach(function (t) {
    if (!t.Дедлайн) return;
    const d = new Date(t.Дедлайн);
    if (isNaN(d)) return;
    d.setHours(0, 0, 0, 0);
    if (d < today) overdue.push(t);
    else if (d.getTime() === today.getTime()) dueToday.push(t);
  });

  const nudges = getProactiveNudges();
  const awaitingReply = getClickUpAwaitingReplyTasks();

  if (overdue.length === 0 && dueToday.length === 0 && nudges.length === 0 && awaitingReply.length === 0) return;

  const lines = [];
  const briefing = getAiProactiveBriefing();
  if (briefing) {
    lines.push(briefing);
  } else {
    lines.push('Доброго ранку! Ось стан твоїх задач на сьогодні.');
  }
  lines.push('');

  if (overdue.length > 0) {
    lines.push('ПРОСТРОЧЕНО (' + overdue.length + '):');
    overdue.forEach(function (t) {
      lines.push('- ' + t.Назва + ' (дедлайн: ' + formatDateShort(t.Дедлайн) + ')');
    });
    lines.push('');
  }

  if (dueToday.length > 0) {
    lines.push('ДЕДЛАЙН СЬОГОДНІ (' + dueToday.length + '):');
    dueToday.forEach(function (t) {
      lines.push('- ' + t.Назва);
    });
    lines.push('');
  }

  if (nudges.length > 0) {
    lines.push('🧠 ПІДКАЗКИ:');
    nudges.forEach(function (n) {
      lines.push('- ' + n.text);
    });
    lines.push('');
  }

  if (awaitingReply.length > 0) {
    lines.push('💬 ЧЕКАЮТЬ ТВОЄЇ ВІДПОВІДІ У CLICKUP (' + awaitingReply.length + '):');
    awaitingReply.forEach(function (t) {
      lines.push('- ' + t.name + ' (останнє слово за ' + t.lastCommenter + ')');
    });
    lines.push('');
  }

  let subject = '📋 TaskApp: ' + overdue.length + ' прострочено, ' + dueToday.length + ' на сьогодні';
  if (overdue.length === 0 && dueToday.length === 0) {
    const signals = nudges.length + awaitingReply.length;
    subject = '📋 TaskApp: ' + signals + ' сигнал' + (signals === 1 ? '' : signals < 5 ? 'и' : 'ів') + ' на сьогодні';
  }

  MailApp.sendEmail({
    to: Session.getActiveUser().getEmail(),
    subject: subject,
    body: lines.join('\n')
  });
}
