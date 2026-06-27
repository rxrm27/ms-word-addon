// src/taskpane.js
'use strict';

var scanResult        = null;
var aliasMap          = {};   // manually added by user
var autoAliasMap      = {};   // auto-computed: phrases sharing a content word with canonical
var canonicalOverride = {};   // user-edited canonical phrase per number
var dismissedConflicts = {};  // conflicts user has dismissed { "num:100": true, "ph:map": true }

// ── Bootstrap ────────────────────────────────────────────────────────────────

if (typeof Office !== 'undefined') {
  Office.onReady(function (info) {
    loadState();  // restore persisted state before wiring buttons
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

// ── Persistence (localStorage, 30-day TTL) ───────────────────────────────────

function saveState() {
  try {
    localStorage.setItem('patsync_v1', JSON.stringify({
      ts: Date.now(),
      dismissed: dismissedConflicts,
      aliasMap: aliasMap,
      canonicalOverride: canonicalOverride
    }));
  } catch (e) {}
}

function loadState() {
  try {
    var raw = localStorage.getItem('patsync_v1');
    if (!raw) return;
    var s = JSON.parse(raw);
    if (Date.now() - s.ts > 30 * 86400 * 1000) { localStorage.removeItem('patsync_v1'); return; }
    dismissedConflicts = s.dismissed      || {};
    aliasMap           = s.aliasMap       || {};
    canonicalOverride  = s.canonicalOverride || {};
  } catch (e) {}
}

// ── Status helper ────────────────────────────────────────────────────────────

function setStatus(msg, cls) {
  var el = document.getElementById('status');
  el.textContent = msg;
  el.className = cls || '';
}

function getAllConflicts() {
  if (!scanResult) return [];
  return scanResult.conflicts.concat(detectOverrideConflicts());
}

function refreshStatus() {
  if (!scanResult) return;
  var filtered = filterConflicts(getAllConflicts());
  var msg = filtered.length === 0
    ? 'No conflicts found. ' + scanResult.referenceTable.length + ' references detected.'
    : filtered.length + ' conflict(s) found. ' + scanResult.referenceTable.length + ' total references.';
  setStatus(msg, filtered.length === 0 ? 'success' : 'error');
  document.getElementById('btn-mark').disabled = filtered.length === 0;
}

// ── Auto-alias: content-word matching ────────────────────────────────────────

var CONTENT_STOP = /^(?:a|an|the|and|or|but|nor|for|so|as|to|from|of|in|on|at|by|with|not|is|are|was|were|be|been|being|also|both|either|neither|such|each|every|any|all|some|then|further|thus|hence|comprises?|includes?|has|have|had|contain|which|that|this|these|those|when|if|limited|no|its|their|our|said|may|can|via|into|onto|about|within|between|through|across|among|upon|after|before|during|whether|wherein|thereby|whereby|thereof|therein|therefor|therefore)$/i;

function stemWord(w) {
  if (w.length > 4 && w.endsWith('ies')) return w.slice(0, -3) + 'y';
  if (w.length > 4 && w.endsWith('ses')) return w.slice(0, -2);
  if (w.length > 3 && w.endsWith('s') && !w.endsWith('ss')) return w.slice(0, -1);
  return w;
}

function getContentWords(phrase) {
  return phrase.toLowerCase().split(/\s+/).filter(function (w) {
    return w.length > 1 && !CONTENT_STOP.test(w);
  }).map(stemWord);
}

function sharesContentWord(phrase, canonical) {
  var phraseWords    = getContentWords(phrase);
  var canonicalWords = getContentWords(canonical);
  if (phraseWords.length === 0) return true;
  return phraseWords.some(function (w) { return canonicalWords.indexOf(w) !== -1; });
}

function getOriginalCanonical(number) {
  if (!scanResult) return '';
  var row = scanResult.referenceTable.filter(function (r) { return r.number === number; })[0];
  return row ? row.phrase : '';
}

function getCanonical(number) {
  return canonicalOverride[number] || getOriginalCanonical(number);
}

function saveCanonical(number, newPhrase) {
  var oldCanonical = getCanonical(number);
  canonicalOverride[number] = newPhrase;
  if (aliasMap[number]) {
    var idx = aliasMap[number].indexOf(oldCanonical);
    if (idx !== -1) aliasMap[number][idx] = newPhrase;
  }
  saveState();
  if (scanResult) { buildAutoAliases(scanResult); renderResults(scanResult); refreshStatus(); }
}

function buildAutoAliases(result) {
  autoAliasMap = {};
  var n2p = result.dictionary.numberToPhrase;
  Object.keys(n2p).forEach(function (num) {
    var phrases   = n2p[num];
    if (phrases.length <= 1) return;
    var canonical = getCanonical(num);
    if (!canonical) return;
    phrases.forEach(function (p) {
      if (p === canonical) return;
      if (sharesContentWord(p, canonical)) {
        if (!autoAliasMap[num]) autoAliasMap[num] = [canonical];
        if (autoAliasMap[num].indexOf(p) === -1) autoAliasMap[num].push(p);
      }
    });
  });
}

function getAllAccepted(number) {
  var manual   = aliasMap[number]     || [];
  var auto     = autoAliasMap[number] || [];
  var combined = manual.slice();
  auto.forEach(function (p) { if (combined.indexOf(p) === -1) combined.push(p); });
  return combined;
}

// ── Override-conflict detection ───────────────────────────────────────────────

// Returns synthetic phrase_reuse conflicts when two numbers share the same canonicalOverride.
function detectOverrideConflicts() {
  var phraseToNums = {};
  Object.keys(canonicalOverride).forEach(function (num) {
    var phrase = canonicalOverride[num];
    if (!phraseToNums[phrase]) phraseToNums[phrase] = [];
    if (phraseToNums[phrase].indexOf(num) === -1) phraseToNums[phrase].push(num);
  });
  var enginePhrases = {};
  if (scanResult) {
    scanResult.conflicts.forEach(function (c) {
      if (c.type === 'phrase_reuse') enginePhrases[c.phrase] = true;
    });
  }
  var out = [];
  Object.keys(phraseToNums).forEach(function (phrase) {
    if (phraseToNums[phrase].length > 1 && !enginePhrases[phrase]) {
      out.push({ type: 'phrase_reuse', phrase: phrase, numbers: phraseToNums[phrase], synthetic: true });
    }
  });
  return out;
}

// ── Dismiss helpers ───────────────────────────────────────────────────────────

function conflictKey(c) {
  return c.type === 'number_reuse' ? 'num:' + c.number : 'ph:' + c.phrase;
}

function dismissConflict(c) {
  dismissedConflicts[conflictKey(c)] = true;
  saveState();
  renderConflicts(getAllConflicts());
  refreshStatus();
}

// ── Scan ─────────────────────────────────────────────────────────────────────

function scanDocument() {
  setStatus('Scanning document...');
  Word.run(function (ctx) {
    var body = ctx.document.body;
    body.load('text');
    return ctx.sync().then(function () {
      scanResult = NumberingEngine.analyze(body.text);
      buildAutoAliases(scanResult);
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
  renderConflicts(getAllConflicts());
  renderDictionary(result.referenceTable);
}

// Filter: dismissed conflicts hidden; number_reuse resolved by aliases; phrase_reuse by override.
function filterConflicts(conflicts) {
  return conflicts.filter(function (c) {
    if (dismissedConflicts[conflictKey(c)]) return false;
    if (c.type === 'number_reuse') {
      var accepted  = getAllAccepted(c.number);
      var remaining = c.phrases.filter(function (p) { return accepted.indexOf(p) === -1; });
      return remaining.length > 1;
    }
    if (c.type === 'phrase_reuse') {
      if (c.synthetic) return true;  // override conflicts: all numbers are active
      var active = c.numbers.filter(function (num) {
        return !canonicalOverride[num] || canonicalOverride[num] === c.phrase;
      });
      return active.length > 1;
    }
    return true;
  });
}

function renderConflicts(conflicts) {
  var section   = document.getElementById('conflicts-section');
  var container = document.getElementById('conflicts-list');
  var badge     = document.getElementById('conflict-count');

  var visible = filterConflicts(conflicts);
  container.innerHTML = '';
  badge.textContent = visible.length;

  if (visible.length === 0) { section.classList.add('hidden'); return; }
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
    var tdDismiss = document.createElement('td');
    tdDismiss.className = 'conflict-dismiss-cell';

    // Dismiss button
    var dismissBtn = document.createElement('button');
    dismissBtn.className = 'dismiss-btn';
    dismissBtn.textContent = '×';
    dismissBtn.title = 'Dismiss this conflict';
    dismissBtn.onclick = (function (conflict) {
      return function () { dismissConflict(conflict); };
    })(c);
    tdDismiss.appendChild(dismissBtn);

    if (c.type === 'number_reuse') {
      var numSpan = document.createElement('span');
      numSpan.className = 'conflict-num';
      numSpan.textContent = c.number;
      tdLabel.appendChild(numSpan);

      var accepted = getAllAccepted(c.number);
      var displayPhrases = accepted.length > 0
        ? c.phrases.filter(function (p) { return accepted.indexOf(p) === -1; })
        : c.phrases;

      displayPhrases.forEach(function (p) {
        var chip = document.createElement('span');
        chip.className = 'phrase-chip clickable-chip';
        chip.textContent = p;
        chip.title = 'Click to scroll to in document';
        chip.onclick = (function (phrase, num) {
          return function () { scrollToPhrase(phrase, num); };
        })(p, c.number);
        tdChips.appendChild(chip);
      });
    } else {
      var phraseEm = document.createElement('em');
      phraseEm.className = 'conflict-phrase';
      phraseEm.textContent = '"' + c.phrase + '"';
      tdLabel.appendChild(phraseEm);

      var displayNums = c.synthetic
        ? c.numbers
        : c.numbers.filter(function (n) {
            return !canonicalOverride[n] || canonicalOverride[n] === c.phrase;
          });

      displayNums.forEach(function (n) {
        var chip = document.createElement('span');
        chip.className = 'num-chip clickable-chip';
        chip.textContent = n;
        chip.title = 'Click to scroll to in document';
        chip.onclick = (function (phrase, num) {
          return function () { scrollToPhrase(phrase, num); };
        })(c.phrase, n);
        tdChips.appendChild(chip);
      });
    }

    tr.appendChild(tdLabel);
    tr.appendChild(tdChips);
    tr.appendChild(tdDismiss);
    table.appendChild(tr);
  });

  container.appendChild(table);
}

function renderDictionary(table) {
  var section = document.getElementById('dictionary-section');
  var tbody   = document.getElementById('dict-body');
  var badge   = document.getElementById('dict-count');

  tbody.innerHTML = '';
  if (table.length === 0) { section.classList.add('hidden'); return; }
  section.classList.remove('hidden');
  badge.textContent = table.length;

  table.forEach(function (row) {
    var canonical = getCanonical(row.number);

    var tr = document.createElement('tr');
    tr.dataset.number = row.number;

    var tdNum = document.createElement('td');
    tdNum.textContent = row.number;

    var tdPhrase = document.createElement('td');
    tdPhrase.textContent = canonical;
    tdPhrase.contentEditable = true;
    tdPhrase.className = 'component-editable';
    tdPhrase.title = 'Click to edit component name';
    tdPhrase.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); this.blur(); }
      if (e.key === 'Escape') { this.textContent = getCanonical(row.number); this.blur(); }
    });
    tdPhrase.addEventListener('blur', (function (num) {
      return function () {
        var newVal = this.textContent.trim().toLowerCase();
        if (newVal && newVal !== getCanonical(num)) saveCanonical(num, newVal);
      };
    })(row.number));

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
  if (aliasMap[number].indexOf(val)       === -1) aliasMap[number].push(val);
  inputEl.value = '';
  var cell = document.getElementById('aliases-' + number);
  if (cell) renderAliasChips(cell, number, canonical);
  saveState();
  if (scanResult) { renderConflicts(getAllConflicts()); refreshStatus(); }
}

function doRemoveAlias(number, phrase, canonical) {
  if (!aliasMap[number]) return;
  aliasMap[number] = aliasMap[number].filter(function (a) { return a !== phrase; });
  var nonCanonical = aliasMap[number].filter(function (a) { return a !== canonical; });
  if (nonCanonical.length === 0) delete aliasMap[number];
  var cell = document.getElementById('aliases-' + number);
  if (cell) renderAliasChips(cell, number, canonical);
  saveState();
  if (scanResult) { renderConflicts(getAllConflicts()); refreshStatus(); }
}

// ── Scroll to phrase in document ─────────────────────────────────────────────

function scrollToPhrase(phrase, number) {
  var term = phrase + ' ' + number;
  setStatus('Scrolling to: ' + term + '...');
  Word.run(function (ctx) {
    var ranges = ctx.document.body.search(term, { matchCase: false });
    ranges.load('items');
    return ctx.sync().then(function () {
      if (ranges.items.length > 0) {
        ranges.items[0].select();
        return ctx.sync().then(function () { setStatus('Found: ' + term, 'success'); });
      } else {
        setStatus('Not found in document: ' + term, 'error');
      }
    });
  }).catch(function (err) {
    setStatus('Error scrolling: ' + err.message, 'error');
  });
}

// ── Mark conflicts ────────────────────────────────────────────────────────────

function markConflicts() {
  if (!scanResult) { setStatus('Run Scan first.', 'error'); return; }

  var activeConflicts = filterConflicts(getAllConflicts());
  if (!activeConflicts.length) { setStatus('No active conflicts to mark.', 'error'); return; }

  setStatus('Adding comments to document...');
  var searches = [];

  activeConflicts.forEach(function (conflict) {
    if (conflict.type === 'number_reuse') {
      var accepted = getAllAccepted(conflict.number);
      var phrases  = accepted.length > 0
        ? conflict.phrases.filter(function (p) { return accepted.indexOf(p) === -1; })
        : conflict.phrases;
      phrases.forEach(function (phrase) {
        var others = phrases.filter(function (p) { return p !== phrase; });
        searches.push({
          term:    phrase + ' ' + conflict.number,
          comment: 'CONFLICT: number ' + conflict.number + ' also used for "' + others.join('", "') + '"'
        });
      });
    } else if (conflict.type === 'phrase_reuse') {
      var nums = conflict.synthetic
        ? conflict.numbers
        : conflict.numbers.filter(function (n) {
            return !canonicalOverride[n] || canonicalOverride[n] === conflict.phrase;
          });
      nums.forEach(function (num) {
        var others = nums.filter(function (n) { return n !== num; });
        searches.push({
          term:    conflict.phrase + ' ' + num,
          comment: 'CONFLICT: "' + conflict.phrase + '" also numbered as ' + others.join(', ')
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
    tableData.forEach(function (row) { values.push([row.number, getCanonical(row.number)]); });
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
