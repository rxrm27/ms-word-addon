/* taskpane.js — Claim Checker Word add-in controller */
(function () {
  'use strict';

  var STORAGE_KEY = 'claimchecker_v1';

  var CATEGORIES = [
    { key: 'format',         label: 'FORMAT' },
    { key: 'grammar',        label: 'GRAMMAR' },
    { key: 'scope',          label: 'SCOPE' },
    { key: 'dependency',     label: 'DEPENDENCY' },
    { key: 'elements',       label: 'ELEMENTS' },
    { key: 'specialFormats', label: 'SPECIAL FORMATS' }
  ];

  var TYPE_LABELS = {
    method:      'Method',
    apparatus:   'Apparatus',
    crm:         'CRM',
    composition: 'Composition',
    use:         'Use',
    other:       'Other'
  };

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
    var bar = el('statusBar');
    bar.textContent = msg;
    bar.className   = 'status-bar' + (isError ? ' error' : '');
    bar.classList.remove('hidden');
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

  function renderFinding(finding) {
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

    return row;
  }

  function renderCategory(catKey, findings) {
    var body = el('body-' + catKey);
    body.innerHTML = '';
    setBadge(catKey, findings);

    var section = el('cat-' + catKey);

    if (findings.length === 0) {
      var empty = document.createElement('div');
      empty.className   = 'empty-cat';
      empty.textContent = 'No issues found';
      body.appendChild(empty);
      section.classList.add('collapsed');
      return;
    }

    // Sort: errors first, then warnings, then info
    var sorted = findings.slice().sort(function (a, b) {
      var order = { error: 0, warning: 1, info: 2 };
      var diff  = (order[a.severity] || 2) - (order[b.severity] || 2);
      if (diff !== 0) return diff;
      return (a.claimNumber || 0) - (b.claimNumber || 0);
    });

    sorted.forEach(function (f) {
      body.appendChild(renderFinding(f));
    });

    section.classList.remove('collapsed');
  }

  // ── Metrics ────────────────────────────────────────────────────────────────

  function renderMetrics(metrics) {
    el('mIndependent').textContent = metrics.independent;
    el('mDependent').textContent   = metrics.dependent;
    el('mTotal').textContent       = metrics.total;

    var chips = el('typeChips');
    chips.innerHTML = '';
    var types = Object.keys(metrics.byType);
    types.forEach(function (t) {
      var count = metrics.byType[t];
      if (count === 0) return;
      var chip = document.createElement('span');
      chip.className   = 'type-chip ' + t;
      chip.textContent = (TYPE_LABELS[t] || t) + ' ' + count;
      chips.appendChild(chip);
    });

    el('metricsSection').classList.remove('hidden');
  }

  // ── Main scan ──────────────────────────────────────────────────────────────

  function scanClaims() {
    var jurisdiction = el('jurisdictionSelect').value;
    saveState(jurisdiction);

    var btn = el('scanBtn');
    btn.disabled = true;
    btn.textContent = 'Scanning…';
    el('metricsSection').classList.add('hidden');
    el('findingsArea').classList.add('hidden');
    el('statusBar').classList.add('hidden');

    Word.run(function (context) {
      var body = context.document.body;
      body.load('text');
      return context.sync().then(function () {
        var text   = body.text;
        var result = ClaimEngine.analyze(text, jurisdiction);

        if (result.error === 'no_claims_section') {
          setStatus('No "Claims" section found. Add a paragraph reading exactly "Claims" before your numbered claims.', true);
          return;
        }
        if (result.error === 'no_claims_found') {
          setStatus('Found a Claims section but no numbered claims (e.g., "1. An apparatus…").', true);
          return;
        }

        renderMetrics(result.metrics);

        CATEGORIES.forEach(function (cat) {
          renderCategory(cat.key, result.findings[cat.key] || []);
        });

        var total = Object.keys(result.findings).reduce(function (sum, k) {
          return sum + result.findings[k].length;
        }, 0);

        setStatus('Found ' + result.metrics.total + ' claims · ' + total + ' finding' + (total !== 1 ? 's' : ''));
        el('findingsArea').classList.remove('hidden');
      });
    }).catch(function (err) {
      setStatus('Error: ' + (err.message || err), true);
    }).finally(function () {
      btn.disabled    = false;
      btn.textContent = 'Scan Claims';
    });
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
  });

})();
