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

  var item = Office.context.mailbox.item;
  var results = [];

  // Test 1: Basic item read
  results.push('Subject: ' + (item.subject ? 'OK' : 'FAIL'));
  results.push('ItemId: ' + (item.itemId ? 'OK' : 'FAIL'));

  // Test 2: categories.addAsync
  if (item.categories && item.categories.addAsync) {
    results.push('categories.addAsync: EXISTS');
    item.categories.addAsync([{displayName: 'TestDiag', color: Office.MailboxEnums.CategoryColor.None}], function(r) {
      if (r.status === Office.AsyncResultStatus.Succeeded) {
        results.push('categories.addAsync: OK');
        // Clean up
        if (item.categories.removeAsync) {
          item.categories.removeAsync(['TestDiag'], function() {});
        }
      } else {
        results.push('categories.addAsync: FAIL - ' + (r.error ? r.error.message : 'unknown'));
      }
      log.textContent = results.join('\n');
    });
  } else {
    results.push('categories.addAsync: NOT AVAILABLE');
  }

  // Test 3: customProperties
  item.loadCustomPropertiesAsync(function(r) {
    if (r.status === Office.AsyncResultStatus.Succeeded) {
      var props = r.value;
      props.set('diagTest', 'hello');
      props.saveAsync(function(sr) {
        if (sr.status === Office.AsyncResultStatus.Succeeded) {
          results.push('customProperties: OK (save works)');
        } else {
          results.push('customProperties: SAVE FAIL - ' + (sr.error ? sr.error.message : 'unknown'));
        }
        log.textContent = results.join('\n');
      });
    } else {
      results.push('customProperties: LOAD FAIL - ' + (r.error ? r.error.message : 'unknown'));
      log.textContent = results.join('\n');
    }
  });

  // Test 4: getCallbackTokenAsync
  Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(r) {
    if (r.status === Office.AsyncResultStatus.Succeeded) {
      results.push('REST token: OK');
    } else {
      results.push('REST token: FAIL - ' + (r.error ? r.error.message : 'unknown'));
    }
    log.textContent = results.join('\n');
  });

  // Test 5: displayReplyFormAsync (just check if exists)
  results.push('displayReplyForm: ' + (item.displayReplyForm ? 'EXISTS' : 'NOT AVAILABLE'));
  results.push('displayReplyAllFormAsync: ' + (item.displayReplyAllFormAsync ? 'EXISTS' : 'NOT AVAILABLE'));

  log.textContent = results.join('\n');
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

  // Try methods in order of preference
  tryCategories(formData, function(catOk) {
    tryCustomProperties(formData, function(propOk) {
      var methods = [];
      if (catOk) methods.push('categoria');
      if (propOk) methods.push('proprietà');

      if (methods.length > 0) {
        showStatus(msg, 'success', 'Salvato via: ' + methods.join(' + ') + '. Sync processerà al prossimo ciclo.');
      } else {
        showStatus(msg, 'error', 'Nessun metodo ha funzionato. Verifica il pannello diagnostica.');
      }

      btn.disabled = false;
      btn.textContent = 'Crea Task';
      document.getElementById('notes').value = '';
    });
  });
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

  var done = false;
  setTimeout(function() { if (!done) { done = true; callback(false); } }, 5000);

  try {
    item.categories.addAsync(cats, function(result) {
      if (!done) {
        done = true;
        callback(result.status === Office.AsyncResultStatus.Succeeded);
      }
    });
  } catch(e) {
    if (!done) { done = true; callback(false); }
  }
}

function tryCustomProperties(formData, callback) {
  var item = Office.context.mailbox.item;
  var done = false;
  setTimeout(function() { if (!done) { done = true; callback(false); } }, 5000);

  try {
    item.loadCustomPropertiesAsync(function(result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        if (!done) { done = true; callback(false); }
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
        if (!done) {
          done = true;
          callback(sr.status === Office.AsyncResultStatus.Succeeded);
        }
      });
    });
  } catch(e) {
    if (!done) { done = true; callback(false); }
  }
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

  tryCategories(formData, function(catOk) {
    tryCustomProperties(formData, function(propOk) {
      if (catOk || propOk) {
        showStatus(msg, 'success', 'Follow-up creato per ' + dueDateStr);
      } else {
        showStatus(msg, 'error', 'Errore nel salvare il follow-up.');
      }
    });
  });
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
