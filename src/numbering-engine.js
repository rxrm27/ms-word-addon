// src/numbering-engine.js
// Patent reference number conflict detector — pure JS, no dependencies.
(function (global) {
  'use strict';

  // Patterns to replace with spaces BEFORE regex matching.
  // Order matters: FIG/Figure refs first, then patent nos, then dates.
  var SKIP = [
    /\bFIG\.?\s*[\d(][^\s,;.]*/gi,           // FIG. 1, FIG.2(A), FIG.17(B)
    /\bFigure\s+[\d(][^\s,;.]*/gi,           // Figure 1, Figure 10
    /\bNo\.\s+[\d][^,;\s]*/gi,               // No. 18/413,382
    /\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b/g, // dates 01/16/2024
    /\b(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}\b/gi
  ];

  // Capture: up to 6 words (letters/hyphens) then a 3-4 digit ref number (with optional sub-part).
  // Negative lookahead prevents matching numbers immediately followed by / . , or another digit.
  var REF_PAT = /\b([A-Za-z][A-Za-z\-]*(?:\s+[A-Za-z][A-Za-z\-]*){0,5})\s+(\d{3,4}(?:-\d+)?)\b(?![\/.,\d])/g;

  function cleanText(text) {
    var s = text;
    for (var i = 0; i < SKIP.length; i++) s = s.replace(SKIP[i], ' ');
    return s;
  }

  function normalizePhrase(raw) {
    return raw
      .toLowerCase()
      .replace(/^\s*(?:a|an|the)\s+/, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  // Returns array of { phrase, number, rawMatch, position }
  function parse(text) {
    var cleaned = cleanText(text);
    var entries = [];
    var m;
    REF_PAT.lastIndex = 0;
    while ((m = REF_PAT.exec(cleaned)) !== null) {
      var phrase = normalizePhrase(m[1]);
      if (!phrase) continue;
      entries.push({ phrase: phrase, number: m[2], rawMatch: m[0], position: m.index });
    }
    return entries;
  }

  // Returns { numberToPhrase: { "102": ["interface","module",...] },
  //           phraseToNumber: { "interface": ["102","202",...] } }
  function buildDictionary(entries) {
    var n2p = {}, p2n = {};
    entries.forEach(function (e) {
      if (!n2p[e.number]) n2p[e.number] = [];
      if (n2p[e.number].indexOf(e.phrase) === -1) n2p[e.number].push(e.phrase);
      if (!p2n[e.phrase]) p2n[e.phrase] = [];
      if (p2n[e.phrase].indexOf(e.number) === -1) p2n[e.phrase].push(e.number);
    });
    return { numberToPhrase: n2p, phraseToNumber: p2n };
  }

  // Returns array of conflict objects.
  function detectConflicts(dictionary) {
    var n2p = dictionary.numberToPhrase, p2n = dictionary.phraseToNumber;
    var conflicts = [], seen = {};

    Object.keys(n2p).forEach(function (num) {
      if (n2p[num].length > 1 && !seen['num:' + num]) {
        seen['num:' + num] = true;
        conflicts.push({
          type: 'number_reuse',
          number: num,
          phrases: n2p[num].slice(),
          description: 'Number ' + num + ' used for: ' + n2p[num].join(' AND ')
        });
      }
    });

    Object.keys(p2n).forEach(function (phrase) {
      if (p2n[phrase].length > 1 && !seen['ph:' + phrase]) {
        seen['ph:' + phrase] = true;
        conflicts.push({
          type: 'phrase_reuse',
          phrase: phrase,
          numbers: p2n[phrase].slice(),
          description: '"' + phrase + '" numbered as: ' + p2n[phrase].join(' AND ')
        });
      }
    });

    return conflicts;
  }

  // Returns array of { number, phrase } sorted numerically by number.
  function getReferenceDictionary(dictionary) {
    var n2p = dictionary.numberToPhrase;
    var rows = Object.keys(n2p).map(function (num) {
      return { number: num, phrase: n2p[num][0] };
    });
    rows.sort(function (a, b) {
      var an = parseInt(a.number.split('-')[0], 10);
      var bn = parseInt(b.number.split('-')[0], 10);
      return an !== bn ? an - bn : a.number.localeCompare(b.number);
    });
    return rows;
  }

  // Main entry point.
  function analyze(text) {
    var entries = parse(text);
    var dictionary = buildDictionary(entries);
    var conflicts = detectConflicts(dictionary);
    var referenceTable = getReferenceDictionary(dictionary);
    return { entries: entries, dictionary: dictionary, conflicts: conflicts, referenceTable: referenceTable };
  }

  // UMD export: works in browser (sets window.NumberingEngine) and Node.js.
  var exports = { parse: parse, buildDictionary: buildDictionary, detectConflicts: detectConflicts, getReferenceDictionary: getReferenceDictionary, analyze: analyze };
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = exports;
  } else {
    global.NumberingEngine = exports;
  }
}(typeof window !== 'undefined' ? window : global));
