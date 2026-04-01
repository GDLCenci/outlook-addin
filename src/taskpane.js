/* global Office */

// State
var emailContext = {
  subject: '',
  from: '',
  fromEmail: '',
  body: '',
  conversationId: '',
  itemId: ''
};

// ── Office.js Init ──────────────────────────────────────────────
Office.onReady(function (info) {
  var log = document.getElementById('diagnosticLog');
  if (log) log.textContent = 'Office.onReady fired. Host: ' + (info.host || 'null') + ', Platform: ' + (info.platform || 'null');

  loadEmailContext();
  initTabs();
  initForm();
  initFollowup();
  runDiagnostics();
});

function loadEmailContext() {
  var item = Office.context.mailbox.item;
  if (!item) return;

  emailContext.subject = item.subject || '';
  emailContext.conversationId = item.conversationId || '';
  emailContext.itemId = item.itemId || '';

  if (item.from) {
    emailContext.from = item.from.displayName || '';
    emailContext.fromEmail = item.from.emailAddress || '';
  }

  document.getElementById('title').value = emailContext.subject;

  item.body.getAsync(Office.CoercionType.Text, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      emailContext.body = result.value.substring(0, 2000);
    }
  });
}

// ── Diagnostics: test what works ────────────────────────────────
function runDiagnostics() {
  var log = document.getElementById('diagnosticLog');
  if (!log) return;

  var results = [];

  // Test 0: mailbox and item exist?
  results.push('mailbox: ' + (Office.context.mailbox ? 'OK' : 'NULL'));
  var item = Office.context.mailbox ? Office.context.mailbox.item : null;
  results.push('item: ' + (item ? 'OK' : 'NULL'));

  if (!item) {
    results.push('STOP: no item available');
    log.textContent = results.join('\n');
    return;
  }

  // Test 1: Basic item read
  results.push('Subject: ' + (item.subject || '(empty)'));
  results.push('ItemId: ' + (item.itemId ? 'OK (' + item.itemId.substring(0, 20) + '...)' : 'FAIL'));
  results.push('ConversationId: ' + (item.conversationId ? 'OK' : 'FAIL'));
  log.textContent = results.join('\n');

  // Check what methods exist
  results.push('--- Methods available ---');
  results.push('categories obj: ' + (item.categories ? 'YES' : 'NO'));
  results.push('categories.addAsync: ' + (item.categories && item.categories.addAsync ? 'YES' : 'NO'));
  results.push('categories.getAsync: ' + (item.categories && item.categories.getAsync ? 'YES' : 'NO'));
  results.push('loadCustomPropertiesAsync: ' + (item.loadCustomPropertiesAsync ? 'YES' : 'NO'));
  results.push('getCallbackTokenAsync: ' + (Office.context.mailbox.getCallbackTokenAsync ? 'YES' : 'NO'));
  results.push('displayReplyForm: ' + (item.displayReplyForm ? 'YES' : 'NO'));
  results.push('displayReplyAllForm: ' + (item.displayReplyAllForm ? 'YES' : 'NO'));
  log.textContent = results.join('\n');

  // Test async APIs with timeout tracking
  results.push('');
  results.push('--- Async tests (wait 10s) ---');
  log.textContent = results.join('\n');

  var pending = 3;
  function checkDone() {
    pending--;
    if (pending <= 0) {
      results.push('--- All tests complete ---');
      log.textContent = results.join('\n');
    }
  }

  // Test A: categories.getAsync (read only, safer)
  if (item.categories && item.categories.getAsync) {
    try {
      item.categories.getAsync(function(r) {
        results.push('categories.getAsync: ' + r.status + (r.value ? ' (' + r.value.length + ' cats)' : ''));
        log.textContent = results.join('\n');
        checkDone();
      });
    } catch(e) {
      results.push('categories.getAsync: EXCEPTION ' + e.message);
      log.textContent = results.join('\n');
      checkDone();
    }
  } else { results.push('categories.getAsync: SKIP'); checkDone(); }

  // Test B: customProperties
  if (item.loadCustomPropertiesAsync) {
    try {
      item.loadCustomPropertiesAsync(function(r) {
        results.push('customProperties.load: ' + r.status);
        log.textContent = results.join('\n');
        checkDone();
      });
    } catch(e) {
      results.push('customProperties: EXCEPTION ' + e.message);
      log.textContent = results.join('\n');
      checkDone();
    }
  } else { results.push('customProperties: SKIP'); checkDone(); }

  // Test C: REST token
  if (Office.context.mailbox.getCallbackTokenAsync) {
    try {
      Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(r) {
        results.push('REST token: ' + r.status + (r.error ? ' - ' + r.error.message : ''));
        log.textContent = results.join('\n');
        checkDone();
      });
    } catch(e) {
      results.push('REST token: EXCEPTION ' + e.message);
      log.textContent = results.join('\n');
      checkDone();
    }
  } else { results.push('REST token: SKIP'); checkDone(); }

  // Timeout fallback
  setTimeout(function() {
    if (pending > 0) {
      results.push('TIMEOUT: ' + pending + ' async tests did not respond in 10s');
      log.textContent = results.join('\n');
    }
  }, 10000);
}

// ── Tabs ────────────────────────────────────────────────────────
function initTabs() {
  document.querySelectorAll('.tab').forEach(function (tab) {
    tab.addEventListener('click', function () {
      document.querySelectorAll('.tab').forEach(function (t) { t.classList.remove('active'); });
      document.querySelectorAll('.tab-content').forEach(function (c) { c.classList.remove('active'); });
      tab.classList.add('active');
      document.getElementById('tab-' + tab.dataset.tab).classList.add('active');
    });
  });
}

// ── Submit: try all methods ─────────────────────────────────────
function initForm() {
  document.getElementById('task-form').addEventListener('submit', function (e) {
    e.preventDefault();
    submitTask();
  });
}

function submitTask() {
  var btn = document.getElementById('submitBtn');
  var msg = document.getElementById('statusMsg');

  btn.disabled = true;
  btn.textContent = 'Invio in corso...';
  showStatus(msg, 'loading', 'Tentativo salvataggio...');

  var formData = {
    title: document.getElementById('title').value.trim(),
    area: document.getElementById('area').value,
    priority: document.getElementById('priority').value,
    dueDate: document.getElementById('dueDate').value,
    status: document.getElementById('status').value,
    assignee: document.getElementById('assignee').value,
    notes: document.getElementById('notes').value.trim()
  };

  // Build structured draft content
  var lines = ['// TASK (from Outlook Add-in)', ''];
  if (formData.title) lines.push('Titolo: ' + formData.title);
  if (formData.area) lines.push('Area: ' + formData.area);
  if (formData.priority) lines.push('Priorita: ' + formData.priority);
  if (formData.dueDate) lines.push('Scadenza: ' + formData.dueDate);
  if (formData.status) lines.push('Status: ' + formData.status);
  if (formData.assignee) lines.push('Assignee: ' + formData.assignee);
  if (formData.notes) { lines.push(''); lines.push('Note: ' + formData.notes); }

  var htmlBody = '<pre>' + lines.join('\n') + '</pre>';

  // Open reply form with pre-filled content — user saves as draft
  var item = Office.context.mailbox.item;
  item.displayReplyForm(htmlBody);

  showStatus(msg, 'success', 'Bozza aperta — salva senza inviare, poi il sync partirà (~1-2 min).');
  updateBadge('processing');

  // Trigger sync agent via poke (fire and forget, may be blocked by CORS)
  try {
    var xhr = new XMLHttpRequest();
    xhr.open('POST', 'https://claude.ai/api/v1/code/triggers/trig_016ZHRaD7zHHqwHHphSsBPGZ/poke?token=4n2Ex94JVRenNyLj660q-WkGvguMx8XlGgGkqT_EUdU');
    xhr.timeout = 10000;
    xhr.send();
  } catch(e) { /* ignore */ }

  btn.disabled = false;
  btn.textContent = 'Crea Task';
  document.getElementById('notes').value = '';
}

function tryCategories(formData, callback) {
  var item = Office.context.mailbox.item;
  if (!item.categories || !item.categories.addAsync) {
    callback(false);
    return;
  }

  var cats = [{displayName: 'Task', color: Office.MailboxEnums.CategoryColor.None}];
  if (formData.area) {
    cats.push({displayName: formData.area, color: Office.MailboxEnums.CategoryColor.None});
  }

  item.categories.addAsync(cats, function(result) {
    callback(result.status === Office.AsyncResultStatus.Succeeded);
  });
}

function tryCustomProperties(formData, callback) {
  var item = Office.context.mailbox.item;

  item.loadCustomPropertiesAsync(function(result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      callback(false);
      return;
    }
    var props = result.value;
    props.set('taskTitle', formData.title || emailContext.subject);
    props.set('taskArea', formData.area || '');
    props.set('taskPriority', formData.priority || '');
    props.set('taskDueDate', formData.dueDate || '');
    props.set('taskStatus', formData.status || 'Not started');
    props.set('taskAssignee', formData.assignee || 'Giuseppe');
    props.set('taskNotes', formData.notes || '');
    props.set('taskConversationId', emailContext.conversationId);
    props.set('taskFrom', emailContext.from);
    props.set('taskFromEmail', emailContext.fromEmail);
    props.set('taskCreatedAt', new Date().toISOString());
    props.set('taskPending', 'true');

    props.saveAsync(function(sr) {
      callback(sr.status === Office.AsyncResultStatus.Succeeded);
    });
  });
}

// ── Follow-up ───────────────────────────────────────────────────
function initFollowup() {
  document.querySelectorAll('.btn-followup').forEach(function (btn) {
    btn.addEventListener('click', function () {
      submitFollowup(parseInt(btn.dataset.days));
    });
  });
}

function submitFollowup(days) {
  var msg = document.getElementById('followupMsg');
  showStatus(msg, 'loading', 'Creazione follow-up...');

  var dueDate = new Date();
  dueDate.setDate(dueDate.getDate() + days);
  var dueDateStr = dueDate.toISOString().split('T')[0];

  var formData = {
    title: 'Follow-up: ' + emailContext.subject,
    area: '',
    priority: 'Medium',
    dueDate: dueDateStr,
    status: 'Not started',
    assignee: 'Giuseppe',
    notes: 'Follow-up su email da ' + emailContext.from
  };

  var lines = ['// TASK (from Outlook Add-in)', ''];
  lines.push('Titolo: ' + formData.title);
  lines.push('Priorita: ' + formData.priority);
  lines.push('Scadenza: ' + formData.dueDate);
  lines.push('Status: ' + formData.status);
  lines.push('Assignee: ' + formData.assignee);
  if (formData.notes) lines.push('Note: ' + formData.notes);

  var htmlBody = '<pre>' + lines.join('\n') + '</pre>';

  var item = Office.context.mailbox.item;
  item.displayReplyForm(htmlBody);

  // Trigger sync
  try {
    var xhr = new XMLHttpRequest();
    xhr.open('POST', 'https://claude.ai/api/v1/code/triggers/trig_016ZHRaD7zHHqwHHphSsBPGZ/poke?token=4n2Ex94JVRenNyLj660q-WkGvguMx8XlGgGkqT_EUdU');
    xhr.timeout = 10000;
    xhr.send();
  } catch(e) { /* ignore */ }

  showStatus(msg, 'success', 'Bozza follow-up aperta — salva senza inviare.');
}

// ── Helpers ─────────────────────────────────────────────────────
function showStatus(el, type, text) {
  el.className = 'status-msg ' + type;
  el.textContent = text;
}

function updateBadge(state) {
  var badge = document.getElementById('badge');
  if (state === 'linked') {
    badge.className = 'badge badge-linked';
    badge.textContent = 'Task collegato';
  } else if (state === 'processing') {
    badge.className = 'badge badge-processing';
    badge.textContent = 'In elaborazione...';
  } else {
    badge.className = 'badge badge-none';
    badge.textContent = 'Nessun task';
  }
}
