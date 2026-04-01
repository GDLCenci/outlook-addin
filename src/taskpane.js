/* global Office */

// State
var emailContext = {
  subject: '',
  from: '',
  fromEmail: '',
  body: '',
  conversationId: '',
  itemId: '',
  restItemId: ''
};
var accessToken = null;

// ── Office.js Init ──────────────────────────────────────────────
Office.onReady(function (info) {
  if (info.host === Office.HostType.Outlook) {
    loadEmailContext();
    initTabs();
    initForm();
    initFollowup();
    getToken();
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

function getToken() {
  Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      accessToken = result.value;
      // Convert item ID to REST format
      if (Office.context.mailbox.convertToRestId) {
        emailContext.restItemId = Office.context.mailbox.convertToRestId(
          emailContext.itemId, Office.MailboxEnums.RestSource.Id
        );
      } else {
        emailContext.restItemId = emailContext.itemId;
      }
    }
  });
}

// ── Graph API calls ─────────────────────────────────────────────
function graphCall(method, path, body, callback) {
  // Use restUrl from Office.js (points to correct Exchange endpoint)
  var baseUrl = Office.context.mailbox.restUrl || 'https://outlook.office.com/api';
  var url = baseUrl + '/v2.0/me' + path;

  var xhr = new XMLHttpRequest();
  xhr.open(method, url);
  xhr.setRequestHeader('Authorization', 'Bearer ' + accessToken);
  xhr.setRequestHeader('Content-Type', 'application/json');
  xhr.onload = function () {
    if (xhr.status >= 200 && xhr.status < 300) {
      callback(null, xhr.responseText ? JSON.parse(xhr.responseText) : null);
    } else {
      callback(new Error('API ' + xhr.status));
    }
  };
  xhr.onerror = function () { callback(new Error('Errore di rete')); };
  xhr.send(body ? JSON.stringify(body) : null);
}

// ── Categories on current email ─────────────────────────────────
function addCategoriesToEmail(categories, callback) {
  if (!accessToken) {
    callback(new Error('Token non disponibile. Riprova tra qualche secondo.'));
    return;
  }

  var id = emailContext.restItemId || emailContext.itemId;
  graphCall('GET', '/messages/' + id + '?$select=Categories', null, function (err, msg) {
    if (err) { callback(err); return; }

    var current = (msg && msg.Categories) || [];
    var merged = current.slice();
    categories.forEach(function (cat) {
      if (merged.indexOf(cat) === -1) merged.push(cat);
    });

    graphCall('PATCH', '/messages/' + id, { Categories: merged }, function (err2) {
      callback(err2);
    });
  });
}

// ── Task Queue Folder ───────────────────────────────────────────
var taskQueueFolderId = null;

function ensureTaskQueueFolder(callback) {
  if (taskQueueFolderId) { callback(null); return; }

  graphCall('GET', "/mailfolders/inbox/childfolders?$filter=displayName eq 'Task Queue'", null, function (err, data) {
    if (!err && data && data.value && data.value.length > 0) {
      taskQueueFolderId = data.value[0].Id;
      callback(null);
    } else {
      graphCall('POST', '/mailfolders/inbox/childfolders', { DisplayName: 'Task Queue' }, function (err2, folder) {
        if (!err2 && folder) {
          taskQueueFolderId = folder.Id;
          callback(null);
        } else {
          callback(err2 || new Error('Impossibile creare Task Queue folder'));
        }
      });
    }
  });
}

function saveToTaskQueue(formData, callback) {
  ensureTaskQueueFolder(function (err) {
    if (err) { callback(err); return; }

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

    var msg = {
      Subject: '// TASK — ' + (formData.title || emailContext.subject),
      Body: { ContentType: 'Text', Content: lines.join('\n') },
      Categories: ['Task']
    };

    graphCall('POST', '/mailfolders/' + taskQueueFolderId + '/messages', msg, callback);
  });
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

    // Step 2: If there are extras, save to Task Queue
    var hasExtras = formData.notes || formData.priority || formData.dueDate ||
                    formData.assignee !== 'Giuseppe' || formData.status !== 'Not started';

    if (hasExtras) {
      saveToTaskQueue(formData, function (qErr) {
        if (qErr) {
          showStatus(msg, 'error', 'Errore Task Queue: ' + qErr.message);
        } else {
          showStatus(msg, 'success', 'Task creato! Il sync agent lo processera al prossimo ciclo.');
          updateBadge('processing');
        }
        btn.disabled = false;
        btn.textContent = 'Crea Task';
      });
    } else {
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

  var categories = ['Task', 'Follow-up'];
  addCategoriesToEmail(categories, function (catErr) {
    if (catErr) {
      showStatus(msg, 'error', 'Errore: ' + catErr.message);
      return;
    }

    var formData = {
      title: 'Follow-up: ' + emailContext.subject,
      area: '',
      priority: 'Medium',
      dueDate: dueDateStr,
      status: 'Not started',
      assignee: 'Giuseppe',
      notes: 'Follow-up automatico su email da ' + emailContext.from
    };

    saveToTaskQueue(formData, function (qErr) {
      if (qErr) {
        showStatus(msg, 'error', 'Errore: ' + qErr.message);
      } else {
        showStatus(msg, 'success', 'Follow-up creato per ' + dueDateStr);
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
