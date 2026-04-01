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
var restToken = null;
var restUrl = '';
var taskQueueFolderId = null;

// ── Office.js Init ──────────────────────────────────────────────
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    loadEmailContext();
    initTabs();
    initForm();
    initFollowup();
    acquireToken();
  }
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

function acquireToken() {
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      restToken = result.value;
      restUrl = Office.context.mailbox.restUrl;
      ensureTaskQueueFolder();
    }
  });
}

// ── Task Queue Folder ───────────────────────────────────────────
function apiCall(method, path, body, callback) {
  var xhr = new XMLHttpRequest();
  xhr.open(method, restUrl + '/v2.0/me' + path);
  xhr.setRequestHeader('Authorization', 'Bearer ' + restToken);
  xhr.setRequestHeader('Content-Type', 'application/json');
  xhr.onload = function () {
    if (xhr.status >= 200 && xhr.status < 300) {
      callback(null, xhr.responseText ? JSON.parse(xhr.responseText) : null);
    } else {
      callback(new Error(xhr.status + ': ' + xhr.responseText));
    }
  };
  xhr.onerror = function () { callback(new Error('Errore di rete')); };
  xhr.send(body ? JSON.stringify(body) : null);
}

function ensureTaskQueueFolder() {
  // Check if "Task Queue" folder exists under Inbox
  apiCall('GET', '/mailfolders/inbox/childfolders?$filter=displayName eq \'Task Queue\'', null, function (err, data) {
    if (err || !data || !data.value || data.value.length === 0) {
      // Create it
      apiCall('POST', '/mailfolders/inbox/childfolders', { DisplayName: 'Task Queue' }, function (err2, folder) {
        if (!err2 && folder) {
          taskQueueFolderId = folder.Id;
        }
      });
    } else {
      taskQueueFolderId = data.value[0].Id;
    }
  });
}

// ── Categories ──────────────────────────────────────────────────
function addCategoriesToEmail(categories, callback) {
  // Get current categories first
  apiCall('GET', '/messages/' + emailContext.itemId + '?$select=categories', null, function (err, msg) {
    if (err) { callback(err); return; }

    var current = (msg && msg.Categories) || [];
    var merged = current.slice();
    categories.forEach(function (cat) {
      if (merged.indexOf(cat) === -1) merged.push(cat);
    });

    apiCall('PATCH', '/messages/' + emailContext.itemId, { Categories: merged }, function (err2) {
      callback(err2);
    });
  });
}

// ── Build Task Queue Message ────────────────────────────────────
function buildTaskData(formData) {
  var lines = ['// TASK (from Outlook Add-in)', ''];

  if (formData.title) lines.push('Titolo: ' + formData.title);
  if (formData.area) lines.push('Area: ' + formData.area);
  if (formData.priority) lines.push('Priorita: ' + formData.priority);
  if (formData.dueDate) lines.push('Scadenza: ' + formData.dueDate);
  if (formData.status) lines.push('Status: ' + formData.status);
  if (formData.assignee) lines.push('Assignee: ' + formData.assignee);

  lines.push('');
  lines.push('--- Contesto email ---');
  lines.push('Da: ' + emailContext.from + ' <' + emailContext.fromEmail + '>');
  lines.push('Oggetto: ' + emailContext.subject);
  lines.push('Conversation ID: ' + emailContext.conversationId);
  lines.push('Message ID: ' + emailContext.itemId);

  if (formData.notes) {
    lines.push('');
    lines.push('--- Note/Istruzioni ---');
    lines.push(formData.notes);
  }

  if (emailContext.body) {
    lines.push('');
    lines.push('--- Corpo email (estratto) ---');
    lines.push(emailContext.body.substring(0, 1000));
  }

  return lines.join('\n');
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

// ── Form Submit ─────────────────────────────────────────────────
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
  showStatus(msg, 'loading', 'Creazione task...');

  var formData = {
    title: document.getElementById('title').value.trim(),
    area: document.getElementById('area').value,
    priority: document.getElementById('priority').value,
    dueDate: document.getElementById('dueDate').value,
    status: document.getElementById('status').value,
    assignee: document.getElementById('assignee').value,
    notes: document.getElementById('notes').value.trim()
  };

  // Step 1: Add categories to the email
  var categories = ['Task'];
  if (formData.area) categories.push(formData.area);

  addCategoriesToEmail(categories, function (catErr) {
    if (catErr) {
      showStatus(msg, 'error', 'Errore categorie: ' + catErr.message);
      btn.disabled = false;
      btn.textContent = 'Crea Task';
      return;
    }

    // Step 2: If there are notes/overrides, create message in Task Queue
    var hasExtras = formData.notes || formData.priority || formData.dueDate ||
                    formData.assignee !== 'Giuseppe' || formData.status !== 'Not started';

    if (hasExtras && taskQueueFolderId) {
      var taskBody = buildTaskData(formData);
      var queueMsg = {
        Subject: '// TASK — ' + (formData.title || emailContext.subject),
        Body: { ContentType: 'Text', Content: taskBody },
        Categories: ['Task']
      };

      // Create in Task Queue folder
      apiCall('POST', '/mailfolders/' + taskQueueFolderId + '/messages', queueMsg, function (qErr) {
        if (qErr) {
          showStatus(msg, 'error', 'Errore Task Queue: ' + qErr.message);
        } else {
          showStatus(msg, 'success', 'Task creato (categoria + dettagli in Task Queue)');
          updateBadge('processing');
        }
        btn.disabled = false;
        btn.textContent = 'Crea Task';
      });
    } else {
      // No extras — category alone is enough
      showStatus(msg, 'success', 'Categoria "Task" aggiunta. Il sync agent creera il task.');
      updateBadge('processing');
      btn.disabled = false;
      btn.textContent = 'Crea Task';
    }

    document.getElementById('notes').value = '';
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

  // Add categories
  addCategoriesToEmail(['Task', 'Follow-up'], function (catErr) {
    if (catErr) {
      showStatus(msg, 'error', 'Errore: ' + catErr.message);
      return;
    }

    if (taskQueueFolderId) {
      var formData = {
        title: 'Follow-up: ' + emailContext.subject,
        area: '',
        priority: 'Medium',
        dueDate: dueDateStr,
        status: 'Not started',
        assignee: 'Giuseppe',
        notes: 'Follow-up automatico su email da ' + emailContext.from
      };
      var taskBody = buildTaskData(formData);
      var queueMsg = {
        Subject: '// TASK — Follow-up: ' + emailContext.subject,
        Body: { ContentType: 'Text', Content: taskBody },
        Categories: ['Task', 'Follow-up']
      };
      apiCall('POST', '/mailfolders/' + taskQueueFolderId + '/messages', queueMsg, function (qErr) {
        if (qErr) {
          showStatus(msg, 'error', 'Errore: ' + qErr.message);
        } else {
          showStatus(msg, 'success', 'Follow-up creato per ' + dueDateStr);
        }
      });
    } else {
      showStatus(msg, 'success', 'Follow-up segnato (categoria). Scadenza: ' + dueDateStr);
    }
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
