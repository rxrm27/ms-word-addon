// src/taskpane.js
'use strict';

var scanResult = null;
var aliasMap = {};  // { "100": ["information management system", "system"] }

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

function refreshStatus() {
  if (!scanResult) return;
  var filtered = filterConflicts(scanResult.conflicts);
  var msg = filtered.length === 0
    ? 'No conflicts found. ' + scanResult.referenceTable.length + ' references detected.'
    : filtered.length + ' conflict(s) found. ' + scanResult.referenceTable.length + ' total references.';
  setStatus(msg, filtered.length === 0 ? 'success' : 'error');
  document.getElementById('btn-mark').disabled = filtered.length === 0;
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
      refreshStatus();
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

// Conflicts where unresolved phrase count > 1 (respects aliasMap).
function filterConflicts(conflicts) {
  return conflicts.filter(function (c) {
    if (c.type !== 'number_reuse') return true;
    var accepted = aliasMap[c.number] || [];
    var remaining = c.phrases.filter(function (p) { return accepted.indexOf(p) === -1; });
    return remaining.length > 1;
  });
}

function renderConflicts(conflicts) {
  var section = document.getElementById('conflicts-section');
  var container = document.getElementById('conflicts-list');
  var badge = document.getElementById('conflict-count');

  var visible = filterConflicts(conflicts);
  container.innerHTML = '';
  badge.textContent = visible.length;

  if (visible.length === 0) {
    section.classList.add('hidden');
    return;
  }

  section.classList.remove('hidden');

  var table = document.createElement('table');
  table.className = 'conflict-table';

  visible.forEach(function (c) {
    var tr = document.createElement('tr');
    tr.className = 'conflict-row';

    var tdLabel = document.createElement('td');
    tdLabel.className = 'conflict-label-cell';
    var tdChips = document.createElement('td');
    tdChips.className = 'conflict-chips-cell';

    if (c.type === 'number_reuse') {
      var numSpan = document.createElement('span');
      numSpan.className = 'conflict-num';
      numSpan.textContent = c.number;
      tdLabel.appendChild(numSpan);

      var accepted = aliasMap[c.number] || [];
      var displayPhrases = accepted.length > 0
        ? c.phrases.filter(function (p) { return accepted.indexOf(p) === -1; })
        : c.phrases;

      displayPhrases.forEach(function (p) {
        var chip = document.createElement('span');
        chip.className = 'phrase-chip';
        chip.textContent = p;
        tdChips.appendChild(chip);
      });
    } else {
      var phraseEm = document.createElement('em');
      phraseEm.className = 'conflict-phrase';
      phraseEm.textContent = '"' + c.phrase + '"';
      tdLabel.appendChild(phraseEm);

      c.numbers.forEach(function (n) {
        var chip = document.createElement('span');
        chip.className = 'num-chip';
        chip.textContent = n;
        tdChips.appendChild(chip);
      });
    }

    tr.appendChild(tdLabel);
    tr.appendChild(tdChips);
    table.appendChild(tr);
  });

  container.appendChild(table);
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
    var canonical = row.phrase;

    var tr = document.createElement('tr');
    tr.dataset.number = row.number;

    var tdNum = document.createElement('td');
    tdNum.textContent = row.number;

    var tdPhrase = document.createElement('td');
    tdPhrase.textContent = canonical;

    var tdAliases = document.createElement('td');
    tdAliases.className = 'alias-cell';
    tdAliases.id = 'aliases-' + row.number;
    renderAliasChips(tdAliases, row.number, canonical);

    var tdEdit = document.createElement('td');
    tdEdit.className = 'edit-cell';
    var editBtn = document.createElement('button');
    editBtn.className = 'icon-btn';
    editBtn.title = 'Add alias for ' + row.number;
    editBtn.textContent = '✏';
    editBtn.onclick = (function (num) {
      return function () { toggleAliasInput(num); };
    })(row.number);
    tdEdit.appendChild(editBtn);

    tr.appendChild(tdNum);
    tr.appendChild(tdPhrase);
    tr.appendChild(tdAliases);
    tr.appendChild(tdEdit);
    tbody.appendChild(tr);

    // Alias input row (hidden by default)
    var inputRow = document.createElement('tr');
    inputRow.id = 'alias-input-row-' + row.number;
    inputRow.className = 'alias-input-row hidden';

    var inputTd = document.createElement('td');
    inputTd.colSpan = 4;

    var input = document.createElement('input');
    input.type = 'text';
    input.className = 'alias-input';
    input.placeholder = 'e.g. the system';
    input.id = 'alias-text-' + row.number;
    input.onkeydown = (function (num, canon) {
      return function (e) { if (e.key === 'Enter') doAddAlias(num, canon); };
    })(row.number, canonical);

    var addBtn = document.createElement('button');
    addBtn.className = 'alias-add-btn';
    addBtn.textContent = 'Add';
    addBtn.onclick = (function (num, canon) {
      return function () { doAddAlias(num, canon); };
    })(row.number, canonical);

    inputTd.appendChild(input);
    inputTd.appendChild(addBtn);
    inputRow.appendChild(inputTd);
    tbody.appendChild(inputRow);
  });
}

function renderAliasChips(container, number, canonical) {
  container.innerHTML = '';
  var aliases = (aliasMap[number] || []).filter(function (a) { return a !== canonical; });
  aliases.forEach(function (alias) {
    var chip = document.createElement('span');
    chip.className = 'alias-chip';
    chip.appendChild(document.createTextNode(alias));

    var x = document.createElement('span');
    x.className = 'chip-remove';
    x.textContent = '×';
    x.title = 'Remove alias';
    x.onclick = (function (num, a, canon) {
      return function () { doRemoveAlias(num, a, canon); };
    })(number, alias, canonical);

    chip.appendChild(x);
    container.appendChild(chip);
  });
}

function toggleAliasInput(number) {
  var row = document.getElementById('alias-input-row-' + number);
  if (!row) return;
  row.classList.toggle('hidden');
  if (!row.classList.contains('hidden')) {
    var inp = document.getElementById('alias-text-' + number);
    if (inp) inp.focus();
  }
}

function doAddAlias(number, canonical) {
  var inputEl = document.getElementById('alias-text-' + number);
  if (!inputEl) return;
  var val = inputEl.value.trim().toLowerCase();
  if (!val) return;

  if (!aliasMap[number]) aliasMap[number] = [];
  if (aliasMap[number].indexOf(canonical) === -1) aliasMap[number].push(canonical);
  if (aliasMap[number].indexOf(val) === -1) aliasMap[number].push(val);

  inputEl.value = '';

  var cell = document.getElementById('aliases-' + number);
  if (cell) renderAliasChips(cell, number, canonical);

  if (scanResult) { renderConflicts(scanResult.conflicts); refreshStatus(); }
}

function doRemoveAlias(number, phrase, canonical) {
  if (!aliasMap[number]) return;
  aliasMap[number] = aliasMap[number].filter(function (a) { return a !== phrase; });
  var nonCanonical = (aliasMap[number] || []).filter(function (a) { return a !== canonical; });
  if (nonCanonical.length === 0) delete aliasMap[number];

  var cell = document.getElementById('aliases-' + number);
  if (cell) renderAliasChips(cell, number, canonical);

  if (scanResult) { renderConflicts(scanResult.conflicts); refreshStatus(); }
}

// ── Mark conflicts ────────────────────────────────────────────────────────────

function markConflicts() {
  if (!scanResult) { setStatus('Run Scan first.', 'error'); return; }

  var activeConflicts = filterConflicts(scanResult.conflicts);
  if (!activeConflicts.length) { setStatus('No active conflicts to mark.', 'error'); return; }

  setStatus('Adding comments to document...');

  var searches = [];

  activeConflicts.forEach(function (conflict) {
    if (conflict.type === 'number_reuse') {
      var accepted = aliasMap[conflict.number] || [];
      var phrases = accepted.length > 0
        ? conflict.phrases.filter(function (p) { return accepted.indexOf(p) === -1; })
        : conflict.phrases;
      phrases.forEach(function (phrase) {
        var otherPhrases = phrases.filter(function (p) { return p !== phrase; });
        searches.push({
          term: phrase + ' ' + conflict.number,
          comment: 'CONFLICT: number ' + conflict.number + ' also used for "' + otherPhrases.join('", "') + '"'
        });
      });
    } else if (conflict.type === 'phrase_reuse') {
      conflict.numbers.forEach(function (num) {
        var otherNums = conflict.numbers.filter(function (n) { return n !== num; });
        searches.push({
          term: conflict.phrase + ' ' + num,
          comment: 'CONFLICT: "' + conflict.phrase + '" also numbered as ' + otherNums.join(', ')
        });
      });
    }
  });

  var totalMarked = 0;
  var promise = Word.run(function (ctx) { return ctx.sync(); });

  searches.forEach(function (s) {
    promise = promise.then(function () {
      return Word.run(function (innerCtx) {
        var ranges = innerCtx.document.body.search(s.term, { matchCase: false });
        ranges.load('items');
        return innerCtx.sync().then(function () {
          ranges.items.forEach(function (range) {
            range.insertComment(s.comment);
            totalMarked++;
          });
          return innerCtx.sync();
        });
      });
    });
  });

  promise.then(function () {
    setStatus('Added comments to ' + totalMarked + ' location(s). Review in the Comments pane.', 'success');
  }).catch(function (err) {
    setStatus('Error marking conflicts: ' + err.message, 'error');
  });
}

// ── Generate reference table ──────────────────────────────────────────────────

function generateTable() {
  if (!scanResult) { setStatus('Run Scan first.', 'error'); return; }

  var tableData = scanResult.referenceTable;
  if (tableData.length === 0) { setStatus('No references found to tabulate.', 'error'); return; }

  setStatus('Inserting reference table...');

  Word.run(function (ctx) {
    var body = ctx.document.body;

    var heading = body.insertParagraph('Reference Numerals', Word.InsertLocation.end);
    heading.styleBuiltIn = Word.Style.heading2;

    var values = [['Reference Number', 'Component Name']];
    tableData.forEach(function (row) { values.push([row.number, row.phrase]); });

    var table = body.insertTable(values.length, 2, Word.InsertLocation.end, values);
    table.style = 'Table Grid';
    table.rows.getFirst().font.bold = true;

    return ctx.sync().then(function () {
      setStatus('Reference table inserted at end of document (' + tableData.length + ' entries).', 'success');
    });
  }).catch(function (err) {
    setStatus('Error inserting table: ' + err.message, 'error');
  });
}
