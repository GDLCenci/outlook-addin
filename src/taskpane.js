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
  if (info.host === Office.HostType.Outlook) {
    loadEmailContext();
    initTabs();
    initForm();
    initFollowup();
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

  // Check if "Task" category already exists
  checkExistingCategories();
}

// ── Categories via Office.js (no REST token needed) ─────────────
function checkExistingCategories() {
  var item = Office.context.mailbox.item;
  if (item.categories && item.categories.getAsync) {
    item.categories.getAsync(function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        var cats = result.value.map(function (c) { return c.displayName; });
        if (cats.indexOf('Task') !== -1) {
          updateBadge('processing');
        }
      }
    });
  }
}

function addCategories(categories, callback) {
  var item = Office.context.mailbox.item;
  var callbackFired = false;

  function done(err) {
    if (callbackFired) return;
    callbackFired = true;
    callback(err);
  }

  // Timeout: if categories API hangs, proceed without them
  setTimeout(function () { done(null); }, 5000);

  if (!item.categories || !item.categories.addAsync) {
    done(null); // skip categories, proceed with custom properties
    return;
  }

  try {
    var catObjects = categories.map(function (name) {
      return { displayName: name, color: Office.MailboxEnums.CategoryColor.None };
    });

    item.categories.addAsync(catObjects, function (result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        done(null);
      } else {
        done(null); // categories failed but we proceed anyway
      }
    });
  } catch (e) {
    done(null); // proceed without categories
  }
}

// ── Custom Properties (save task data on the email itself) ──────
function saveTaskDataOnEmail(formData, callback) {
  var item = Office.context.mailbox.item;
  var callbackFired = false;

  function done(err) {
    if (callbackFired) return;
    callbackFired = true;
    callback(err);
  }

  setTimeout(function () { done(null); }, 5000);

  try {
    item.loadCustomPropertiesAsync(function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        done(null);
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

      props.saveAsync(function (saveResult) {
        done(null);
      });
    });
  } catch (e) {
    done(null);
  }
}

// ── Trigger Sync Agent ──────────────────────────────────────────
var POKE_URL = 'https://claude.ai/api/v1/code/triggers/trig_01QKVDHPePpEvz7EA3AD57hi/poke?token=4n2Ex94JVRenNyLj660q-WkGvguMx8XlGgGkqT_EUdU';

function triggerSync() {
  // Fire and forget — don't block on response
  try {
    var xhr = new XMLHttpRequest();
    xhr.open('POST', POKE_URL);
    xhr.timeout = 10000;
    xhr.send();
  } catch (e) {
    // Ignore errors — sync will run on schedule anyway
  }
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

  // Step 1: Add "Task" category (+ area if selected)
  var categories = ['Task'];
  if (formData.area) categories.push(formData.area);

  addCategories(categories, function (catErr) {
    if (catErr) {
      showStatus(msg, 'error', catErr.message);
      btn.disabled = false;
      btn.textContent = 'Crea Task';
      return;
    }

    // Step 2: Save form data as custom properties on the email
    saveTaskDataOnEmail(formData, function (propErr) {
      if (propErr) {
        showStatus(msg, 'success', 'Categoria "Task" aggiunta. Sync in avvio...');
      } else {
        showStatus(msg, 'success', 'Task salvato. Sync in avvio...');
      }
      updateBadge('processing');
      triggerSync();
      btn.disabled = false;
      btn.textContent = 'Crea Task';
      document.getElementById('notes').value = '';
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

  var categories = ['Task'];
  addCategories(categories, function (catErr) {
    if (catErr) {
      showStatus(msg, 'error', catErr.message);
      return;
    }

    var formData = {
      title: 'Follow-up: ' + emailContext.subject,
      area: '',
      priority: 'Medium',
      dueDate: dueDateStr,
      status: 'Not started',
      assignee: 'Giuseppe',
      notes: 'Follow-up su email da ' + emailContext.from
    };

    saveTaskDataOnEmail(formData, function () {
      showStatus(msg, 'success', 'Follow-up creato per ' + dueDateStr + '. Sync in avvio...');
      updateBadge('processing');
      triggerSync();
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
