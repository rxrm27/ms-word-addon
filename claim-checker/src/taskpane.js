/* taskpane.js — Claim Checker Word add-in controller */
(function () {
  'use strict';

  var STORAGE_KEY = 'claimchecker_v1';

  var CATEGORIES = [
    { key: 'antecedentBasis', label: 'ANTECEDENT BASIS' },
    { key: 'format',          label: 'FORMAT' },
    { key: 'grammar',         label: 'GRAMMAR' },
    { key: 'scope',           label: 'SCOPE' },
    { key: 'dependency',      label: 'DEPENDENCY' },
    { key: 'elements',        label: 'ELEMENTS' },
    { key: 'specialFormats',  label: 'SPECIAL FORMATS' }
  ];

  var TYPE_LABELS = {
    method:      'Method',
    system:      'System',
    apparatus:   'Apparatus',
    crm:         'CRM',
    composition: 'Composition',
    use:         'Use',
    other:       'Other'
  };

  var _lastScanFn = null;
  var DISMISSED_KEY = 'claimchecker_dismissed_v1';
  var _dismissed = loadDismissed();

  // ── Dismissed findings persistence ────────────────────────────────────────

  function loadDismissed() {
    try {
      var raw = localStorage.getItem(DISMISSED_KEY);
      return raw ? new Set(JSON.parse(raw)) : new Set();
    } catch (e) { return new Set(); }
  }

  function saveDismissed() {
    try { localStorage.setItem(DISMISSED_KEY, JSON.stringify(Array.from(_dismissed))); } catch (e) {}
  }

  function djb2(str) {
    var h = 5381;
    for (var i = 0; i < str.length; i++) h = ((h << 5) + h) ^ str.charCodeAt(i);
    return (h >>> 0).toString(36);
  }

  function findingId(f) {
    return f.ruleCode + '_' + (f.claimNumber || 0) + '_' + djb2(f.message);
  }

  // ── State ──────────────────────────────────────────────────────────────────

  function saveState(jurisdiction) {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify({ jurisdiction: jurisdiction })); } catch (e) {}
  }

  function loadState() {
    try {
      var s = localStorage.getItem(STORAGE_KEY);
      return s ? JSON.parse(s) : {};
    } catch (e) { return {}; }
  }

  // ── DOM helpers ────────────────────────────────────────────────────────────

  function el(id) { return document.getElementById(id); }

  function setStatus(msg, isError) {
    var bar      = el('statusBar');
    var rescanEl = el('rescanBtn');
    // preserve the re-scan button node if it exists inside the bar
    bar.childNodes.forEach(function (n) { if (n !== rescanEl) bar.removeChild(n); });
    var txt = document.createTextNode(msg);
    bar.insertBefore(txt, rescanEl);
    bar.className = 'status-bar' + (isError ? ' error' : '');
    bar.classList.remove('hidden');
    if (rescanEl) {
      rescanEl.classList.toggle('hidden', !!isError);
    }
  }

  // ── Collapsible sections ───────────────────────────────────────────────────

  function setupToggles() {
    document.querySelectorAll('.cat-header').forEach(function (hdr) {
      hdr.addEventListener('click', function () {
        var section = hdr.parentElement;
        section.classList.toggle('collapsed');
      });
    });
  }

  function setBadge(catKey, findings) {
    var badge = el('badge-' + catKey);
    if (!badge) return;
    if (findings.length === 0) {
      badge.textContent = '✓';
      badge.className   = 'cat-badge badge-ok';
      return;
    }
    var hasError   = findings.some(function (f) { return f.severity === 'error'; });
    var hasWarning = findings.some(function (f) { return f.severity === 'warning'; });
    if (hasError) {
      badge.textContent = findings.length + ' error' + (findings.length > 1 ? 's' : '');
      badge.className   = 'cat-badge badge-error';
    } else if (hasWarning) {
      badge.textContent = findings.length + ' warn' + (findings.length > 1 ? 's' : '');
      badge.className   = 'cat-badge badge-warn';
    } else {
      badge.textContent = findings.length + ' info';
      badge.className   = 'cat-badge badge-info';
    }
  }

  // ── Render findings ────────────────────────────────────────────────────────

  function refreshBadge(catKey, body) {
    var rows    = Array.from(body.querySelectorAll('.finding-row:not(.dismissed)'));
    var badge   = el('badge-' + catKey);
    var section = el('cat-' + catKey);
    if (!badge) return;
    if (rows.length === 0) {
      badge.textContent = '✓';
      badge.className   = 'cat-badge badge-ok';
      if (!body.querySelector('.empty-cat')) {
        var empty = document.createElement('div');
        empty.className   = 'empty-cat';
        empty.textContent = 'No issues found';
        body.appendChild(empty);
      }
      if (section) section.classList.add('collapsed');
    } else {
      var hasError = rows.some(function (r) { return r.classList.contains('sev-error'); });
      var hasWarn  = rows.some(function (r) { return r.classList.contains('sev-warning'); });
      if (hasError) {
        badge.textContent = rows.length + ' error' + (rows.length > 1 ? 's' : '');
        badge.className   = 'cat-badge badge-error';
      } else if (hasWarn) {
        badge.textContent = rows.length + ' warn' + (rows.length > 1 ? 's' : '');
        badge.className   = 'cat-badge badge-warn';
      } else {
        badge.textContent = rows.length + ' info';
        badge.className   = 'cat-badge badge-info';
      }
    }
  }

  function renderFinding(catKey, finding) {
    var row = document.createElement('div');
    row.className = 'finding-row sev-' + finding.severity;

    if (finding.claimNumber > 0) {
      var pill = document.createElement('span');
      pill.className   = 'claim-pill sev-' + finding.severity;
      pill.textContent = 'Cl.' + finding.claimNumber;
      row.appendChild(pill);
    }

    var rule = document.createElement('span');
    rule.className   = 'rule-pill';
    rule.textContent = finding.ruleCode;
    row.appendChild(rule);

    var msg = document.createElement('span');
    msg.className   = 'finding-msg';
    msg.textContent = finding.message;
    row.appendChild(msg);

    var dismiss = document.createElement('button');
    dismiss.className = 'btn-dismiss';
    dismiss.title     = 'Dismiss';
    dismiss.textContent = '×';
    dismiss.addEventListener('click', function (e) {
      e.stopPropagation();
      _dismissed.add(findingId(finding));
      saveDismissed();
      row.classList.add('dismissed');
      var body = row.closest('.cat-body');
      if (body) refreshBadge(catKey, body);
    });
    row.appendChild(dismiss);

    return row;
  }

  function renderCategory(catKey, findings) {
    var body = el('body-' + catKey);
    body.innerHTML = '';

    var section = el('cat-' + catKey);

    // Filter out findings the user previously dismissed
    var active = findings.filter(function (f) { return !_dismissed.has(findingId(f)); });

    setBadge(catKey, active);

    if (active.length === 0) {
      var empty = document.createElement('div');
      empty.className   = 'empty-cat';
      empty.textContent = 'No issues found';
      body.appendChild(empty);
      section.classList.add('collapsed');
      return;
    }

    var sorted = active.slice().sort(function (a, b) {
      var order = { error: 0, warning: 1, info: 2 };
      var diff  = (order[a.severity] || 2) - (order[b.severity] || 2);
      if (diff !== 0) return diff;
      return (a.claimNumber || 0) - (b.claimNumber || 0);
    });

    sorted.forEach(function (f) {
      body.appendChild(renderFinding(catKey, f));
    });

    section.classList.remove('collapsed');
  }

  // ── Metrics ────────────────────────────────────────────────────────────────

  function formatClaimNums(nums) {
    if (!nums || nums.length === 0) return '';
    // Compact into ranges: [1,2,3,5,6] → "1–3, 5–6"
    var sorted = nums.slice().sort(function (a, b) { return a - b; });
    var ranges = [];
    var start = sorted[0], end = sorted[0];
    for (var i = 1; i < sorted.length; i++) {
      if (sorted[i] === end + 1) {
        end = sorted[i];
      } else {
        ranges.push(start === end ? String(start) : start + '–' + end);
        start = end = sorted[i];
      }
    }
    ranges.push(start === end ? String(start) : start + '–' + end);
    return 'Cl. ' + ranges.join(', ');
  }

  function renderMetrics(metrics) {
    el('mIndependent').textContent = metrics.independent;
    el('mDependent').textContent   = metrics.dependent;
    el('mTotal').textContent       = metrics.total;

    var indNums = el('mIndependentNums');
    var depNums = el('mDependentNums');
    if (indNums) indNums.textContent = formatClaimNums(metrics.independentNums);
    if (depNums) depNums.textContent = formatClaimNums(metrics.dependentNums);

    var chips = el('typeChips');
    chips.innerHTML = '';
    Object.keys(metrics.byType).forEach(function (t) {
      var count = metrics.byType[t];
      if (count === 0) return;
      var chip = document.createElement('span');
      chip.className   = 'type-chip ' + t;
      chip.textContent = (TYPE_LABELS[t] || t) + ' ' + count;
      chips.appendChild(chip);
    });

    el('metricsSection').classList.remove('hidden');
  }

  // ── Paragraph reconstruction ───────────────────────────────────────────────
  // Word numbered lists don't include the number in paragraph.text —
  // they're formatting metadata. We reconstruct "N. text" for level-0 list items.

  function buildTextFromParagraphs(items) {
    var listCounters = [0, 0, 0, 0];
    return items.map(function (p) {
      var li = p.listItemOrNullObject;
      if (li && !li.isNullObject) {
        var level = (typeof li.level === 'number') ? li.level : 0;
        // Reset deeper level counters when shallower level increments
        for (var k = level + 1; k < listCounters.length; k++) listCounters[k] = 0;
        listCounters[level]++;
        if (level === 0) {
          // Top-level list item → prepend claim number
          return listCounters[0] + '. ' + p.text;
        }
        // Nested items (claim body elements) — keep as indented text
        return '  ' + p.text;
      }
      return p.text;
    }).join('\n');
  }

  // ── Render results ─────────────────────────────────────────────────────────

  function renderResult(result) {
    if (result.error === 'no_claims_section') {
      setStatus(
        'Could not detect Claims section. Place your cursor on the first claim line and click "Analyze from Cursor".',
        true
      );
      el('cursorHint').classList.remove('hidden');
      return;
    }
    if (result.error === 'no_claims_found') {
      setStatus('Claims section found but no numbered claims detected (e.g. "1. An apparatus…").', true);
      return;
    }

    renderMetrics(result.metrics);

    CATEGORIES.forEach(function (cat) {
      renderCategory(cat.key, result.findings[cat.key] || []);
    });

    var total = Object.keys(result.findings).reduce(function (sum, k) {
      return sum + result.findings[k].length;
    }, 0);

    el('cursorHint').classList.add('hidden');
    setStatus('Found ' + result.metrics.total + ' claims · ' + total + ' finding' + (total !== 1 ? 's' : ''));
    el('findingsArea').classList.remove('hidden');
  }

  function resetUI(scanningLabel) {
    var btn = el('scanBtn');
    btn.disabled    = true;
    btn.textContent = scanningLabel || 'Scanning…';
    var curBtn = el('cursorBtn');
    if (curBtn) { curBtn.disabled = true; }
    el('metricsSection').classList.add('hidden');
    el('findingsArea').classList.add('hidden');
    el('statusBar').classList.add('hidden');
    el('cursorHint').classList.add('hidden');
  }

  function restoreUI() {
    var btn = el('scanBtn');
    btn.disabled    = false;
    btn.textContent = 'Scan Claims';
    var curBtn = el('cursorBtn');
    if (curBtn) { curBtn.disabled = false; }
  }

  // ── Auto scan (full document) ──────────────────────────────────────────────

  function scanClaims() {
    _lastScanFn = scanClaims;
    var jurisdiction = el('jurisdictionSelect').value;
    saveState(jurisdiction);
    resetUI('Scanning…');

    Word.run(function (context) {
      var paragraphs = context.document.body.paragraphs;
      paragraphs.load('text, listItemOrNullObject/level');

      return context.sync().then(function () {
        var text   = buildTextFromParagraphs(paragraphs.items);
        var result = ClaimEngine.analyze(text, jurisdiction);
        renderResult(result);
      });
    }).catch(function (err) {
      setStatus('Error: ' + (err.message || err), true);
    }).finally(restoreUI);
  }

  // ── Cursor scan (from insertion point to end) ──────────────────────────────

  function scanFromCursor() {
    _lastScanFn = scanFromCursor;
    var jurisdiction = el('jurisdictionSelect').value;
    saveState(jurisdiction);
    resetUI('Scanning from cursor…');

    Word.run(function (context) {
      var sel        = context.document.getSelection();
      var allParas   = context.document.body.paragraphs;
      var selParas   = sel.paragraphs;

      allParas.load('text, listItemOrNullObject/level');
      selParas.load('text');

      return context.sync().then(function () {
        // Find the paragraph index where the cursor/selection starts
        var selFirstText = selParas.items.length > 0 ? selParas.items[0].text : '';
        var startIdx     = 0;

        for (var i = 0; i < allParas.items.length; i++) {
          if (allParas.items[i].text === selFirstText) {
            startIdx = i;
            break;
          }
        }

        // Reconstruct numbered text from cursor position onward
        var fromHere  = allParas.items.slice(startIdx);
        // Reset list counters so claim numbering starts fresh from cursor
        var listCounters = [0, 0, 0, 0];
        var claimsText = fromHere.map(function (p) {
          var li = p.listItemOrNullObject;
          if (li && !li.isNullObject) {
            var level = (typeof li.level === 'number') ? li.level : 0;
            for (var k = level + 1; k < listCounters.length; k++) listCounters[k] = 0;
            listCounters[level]++;
            if (level === 0) return listCounters[0] + '. ' + p.text;
            return '  ' + p.text;
          }
          return p.text;
        }).join('\n');

        // analyzeFromText skips heading detection — treats input as already claims
        var result = ClaimEngine.analyzeFromText(claimsText, jurisdiction);
        renderResult(result);
      });
    }).catch(function (err) {
      setStatus('Error: ' + (err.message || err), true);
    }).finally(restoreUI);
  }

  // ── Init ───────────────────────────────────────────────────────────────────

  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Word) return;

    var state = loadState();
    if (state.jurisdiction) {
      var sel = el('jurisdictionSelect');
      if (sel) sel.value = state.jurisdiction;
    }

    setupToggles();
    el('scanBtn').addEventListener('click', scanClaims);

    var curBtn = el('cursorBtn');
    if (curBtn) curBtn.addEventListener('click', scanFromCursor);

    var rescanBtn = el('rescanBtn');
    if (rescanBtn) rescanBtn.addEventListener('click', function () {
      if (_lastScanFn) _lastScanFn();
    });
  });

})();
