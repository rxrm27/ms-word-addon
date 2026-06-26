// src/taskpane.js
'use strict';

var scanResult = null;

// ── Bootstrap ────────────────────────────────────────────────────────────────

if (typeof Office !== 'undefined') {
  Office.onReady(function (info) {
    if (info.host === Office.HostType.Word) {
      document.getElementById('btn-scan').onclick = scanDocument;
      document.getElementById('btn-mark').onclick = markConflicts;
      document.getElementById('btn-table').onclick = generateTable;
    } else {
      setStatus('This add-in requires Microsoft Word.', 'error');
    }
  });
} else {
  setStatus('Running outside Word — Office.js not available.', 'error');
}

// ── Status helper ────────────────────────────────────────────────────────────

function setStatus(msg, cls) {
  var el = document.getElementById('status');
  el.textContent = msg;
  el.className = cls || '';
}

// ── Scan ─────────────────────────────────────────────────────────────────────

function scanDocument() {
  setStatus('Scanning document...');
  Word.run(function (ctx) {
    var body = ctx.document.body;
    body.load('text');
    return ctx.sync().then(function () {
      scanResult = NumberingEngine.analyze(body.text);
      renderResults(scanResult);
      var msg = scanResult.conflicts.length === 0
        ? 'No conflicts found. ' + scanResult.referenceTable.length + ' references detected.'
        : scanResult.conflicts.length + ' conflict(s) found. ' + scanResult.referenceTable.length + ' total references.';
      setStatus(msg, scanResult.conflicts.length === 0 ? 'success' : 'error');
      document.getElementById('btn-mark').disabled = scanResult.conflicts.length === 0;
      document.getElementById('btn-table').disabled = scanResult.referenceTable.length === 0;
    });
  }).catch(function (err) {
    setStatus('Error scanning document: ' + err.message, 'error');
  });
}

// ── Render results ───────────────────────────────────────────────────────────

function renderResults(result) {
  renderConflicts(result.conflicts);
  renderDictionary(result.referenceTable);
}

function renderConflicts(conflicts) {
  var section = document.getElementById('conflicts-section');
  var list = document.getElementById('conflicts-list');
  var badge = document.getElementById('conflict-count');

  list.innerHTML = '';
  if (conflicts.length === 0) {
    section.classList.add('hidden');
    return;
  }

  section.classList.remove('hidden');
  badge.textContent = conflicts.length;

  conflicts.forEach(function (c) {
    var li = document.createElement('li');
    li.className = 'conflict-item';
    li.textContent = c.description;
    list.appendChild(li);
  });
}

function renderDictionary(table) {
  var section = document.getElementById('dictionary-section');
  var tbody = document.getElementById('dict-body');
  var badge = document.getElementById('dict-count');

  tbody.innerHTML = '';
  if (table.length === 0) {
    section.classList.add('hidden');
    return;
  }

  section.classList.remove('hidden');
  badge.textContent = table.length;

  table.forEach(function (row) {
    var tr = document.createElement('tr');
    var tdNum = document.createElement('td');
    var tdPhrase = document.createElement('td');
    tdNum.textContent = row.number;
    tdPhrase.textContent = row.phrase;
    tr.appendChild(tdNum);
    tr.appendChild(tdPhrase);
    tbody.appendChild(tr);
  });
}

// ── Mark conflicts — stub, replaced in Task 7 ────────────────────────────────
function markConflicts() { setStatus('Mark Conflicts not yet implemented.', 'error'); }

// ── Generate table — stub, replaced in Task 8 ────────────────────────────────
function generateTable() { setStatus('Generate Table not yet implemented.', 'error'); }
