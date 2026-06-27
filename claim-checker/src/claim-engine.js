/* claim-engine.js — Pure JS patent claim parser + rule checker
 * No DOM, no Office.js dependencies — fully testable in browser or Node.
 * Rules derived from WIPO Patent Drafting Manual (2nd ed. 2023), pp. 42–71.
 */
var ClaimEngine = (function () {

  // ── Constants ──────────────────────────────────────────────────────────────

  var TRANSITIONS = [
    'consisting essentially of',
    'consisting of',
    'wherein the improvement comprises',
    'characterized in that',
    'characterized by',
    'comprising',
    'including',
    'containing'
  ];

  var CATEGORY_PATTERNS = {
    crm:         /non.transitory|computer.readable\s+medium|storage\s+medium|computer\s+program\s+product|medium\s+storing/i,
    method:      /\b(method|process|procedure|technique)\b/i,
    system:      /\b(system|systems)\b/i,
    apparatus:   /\b(apparatus|device|machine|assembly|arrangement|unit|circuit|controller|module|sensor|network|platform|interface)\b/i,
    composition: /\b(composition|compound|mixture|formulation|preparation|blend|alloy|polymer)\b/i,
    use:         /\buse\s+of\b/i
  };

  // Words that may legitimately appear with "the" in claims without prior "a/an" introduction
  var KNOWN_THE_WORDS = new Set([
    'invention', 'claim', 'claims', 'art', 'prior', 'following', 'present',
    'disclosure', 'description', 'specification', 'embodiment', 'embodiments',
    'example', 'examples', 'above', 'below', 'foregoing', 'scope', 'spirit',
    'purpose', 'basis', 'case', 'manner', 'way', 'same', 'fact', 'time',
    'group', 'formula', 'range', 'form', 'ratio', 'rate', 'amount', 'order',
    // Quantifier words — "the one or more X", "the at least one X", "the plurality of X"
    'one', 'least', 'more', 'plurality', 'number', 'set',
    // Structural/list words common in CRM claims — "perform the steps of", "the following"
    'steps', 'acts', 'operations', 'instructions', 'following'
  ]);

  // Adjectives that modify claim nouns but are not the head noun
  var ORDINAL_WORDS = new Set([
    'first', 'second', 'third', 'fourth', 'fifth', 'further', 'additional',
    'another', 'other', 'said', 'predetermined', 'given', 'respective',
    'associated', 'corresponding', 'aforementioned', 'aforesaid'
  ]);

  // Words that terminate a noun phrase — used in negative lookahead in NP regexes
  var _NP_STOP = [
    // Prepositions
    'to','in','of','for','with','from','on','at','by','as','via','into','upon',
    'within','without','during','after','before','across','above','below',
    'between','among','over','under','around','per','plus','about',
    'according','toward','towards','against','despite','except','besides',
    // Conjunctions / relative
    'or','and','but','nor','that','which','when','where','while','if','unless',
    'until','since','though','although','because','whether',
    // Auxiliaries / copulae
    'is','are','was','were','has','have','had','do','does','did',
    'will','would','can','could','may','might','must','shall','should',
    'be','been','being',
    // Determiners / pronouns
    'not','no','its','their','this','these','those','such','each','any',
    'both','some','what','whose',
    // Claim verbs — 3rd-person singular (-s) AND base forms
    'comprises','comprise','includes','include','contains','contain',
    'receives','receive','provides','provide','stores','store',
    'generates','generate','performs','perform','enables','enable',
    'determines','determine','activates','activate','computes','compute',
    'calculates','calculate','transmits','transmit','processes','process',
    'communicates','communicate','outputs','output','reads','read',
    'writes','write','retrieves','retrieve','identifies','identify',
    'detects','detect','monitors','monitor','controls','control',
    'manages','manage','maintains','maintain','updates','update',
    'creates','create','displays','display','uses','use',
    'produces','produce','sends','send','operates','operate',
    'executes','execute','functions','function','runs','run',
    'applies','apply','converts','convert','measures','measure',
    'evaluates','evaluate','supplies','supply','gives','give',
    'makes','make','takes','take','sets','set','works','work',
    'analyzes','analyze','classifies','classify','initiates','initiate',
    'selects','select','assigns','assign','allocates','allocate','couples','couple',
    // Claim verbs — gerund / present-participle forms (-ing)
    'comprising','including','containing','consisting','characterized',
    'receiving','providing','storing','generating','performing','enabling',
    'determining','activating','computing','calculating','transmitting',
    'processing','communicating','outputting','reading','writing','retrieving',
    'identifying','detecting','monitoring','controlling','managing','maintaining',
    'updating','creating','displaying','using','producing','sending','operating',
    'executing','applying','converting','measuring','evaluating','supplying',
    'classifying','initiating','selecting','assigning','allocating','coupling',
    // Patent connectives
    'wherein','whereby','thereby','therefore','thus','hence','whereas','such',
    'based','using','used','having','formed','adapted','configured',
    'coupled','connected','attached','mounted','disposed','corresponding',
    'responsive','selected','designated','associated','related','derived',
    'obtained','established','defined','described','disclosed','claimed'
  ].join('|');

  // Noun phrase optional-extra: additional word that is NOT a stop word
  var _NP_EXTRA = '(?:\\s+(?!(?:' + _NP_STOP + ')\\b)[a-z][a-z\\-]+){0,2}';

  // Regex sources for antecedent-basis checks (recreated per call to reset lastIndex)
  var _INTRO_SRC    = '\\b(?:a|an)\\s+(?:(?:first|second|third|fourth|further|additional)\\s+)?(?:plurality\\s+of\\s+)?([a-z][a-z\\-]+' + _NP_EXTRA + ')';
  var _ALT_INTRO_SRC = '\\b(?:one\\s+or\\s+more|at\\s+least\\s+one)\\s+([a-z][a-z\\-]+' + _NP_EXTRA + ')';
  var _BACK_SRC     = '\\b(?:the|said)\\s+(?:(?:first|second|third|fourth|further|additional)\\s+)?([a-z][a-z\\-]+' + _NP_EXTRA + ')';

  // ── Section extraction ──────────────────────────────────────────────────────

  function extractClaimsSection(text) {
    // Heading variants (case-insensitive):
    //   "Claims" / "Claim" / "I Claim" / "We Claim" / "I/We Claim"
    //   "What is claimed is" / "What I/We claim is"
    var HEADING_RE = /(?:^|\r?\n)\s*(?:i\/we\s+claim|we\s+claim|i\s+claim|what\s+(?:is\s+claimed|(?:i|we)\s+claim(?:ed)?)\s+is|claims?)\s*[:\r\n]/i;
    var m = text.match(HEADING_RE);
    if (m) return text.substring(m.index + m[0].length);

    // Fallback: any line starting "1." or "1)" before a capital letter
    var m2 = text.match(/(?:^|\r?\n)\s*1[.)]\s+[A-Z]/);
    if (m2) return text.substring(m2.index);

    return null;
  }

  // Called by taskpane.js to skip heading detection entirely (cursor mode)
  function parseClaimsDirectly(text, jurisdiction) {
    jurisdiction = jurisdiction || 'US';
    var claims   = parseClaims(text);
    if (claims.length === 0) return { error: 'no_claims_found', claims: [], metrics: null, findings: null };
    var metrics  = computeMetrics(claims);
    var findings = applyAllRules(claims, jurisdiction);
    return { error: null, claims: claims, metrics: metrics, findings: findings };
  }

  // ── Claim parsing ───────────────────────────────────────────────────────────

  function parseClaims(text) {
    // Truncate at post-claims sections that may bleed in
    var endM = text.match(/(?:^|\r?\n)\s*(?:ABSTRACT|DRAWINGS?|BRIEF\s+DESCRIPTION|DETAILED\s+DESCRIPTION|BACKGROUND|FIELD\s+OF(?:\s+THE)?\s+INVENTION|SUMMARY\s+OF(?:\s+THE)?\s+INVENTION|DESCRIPTION\s+OF(?:\s+(?:THE\s+)?(?:PREFERRED\s+)?EMBODIMENTS?)?|Dated?\s+this\b|Date\s*:\s*\w|Signed?\s+(?:this|by)\b|Signature\s+of\b|Attorney\s+(?:for|of\s+record)\b|IN\s+WITNESS\s+WHEREOF\b)\s*(?:\r?\n|$)/im);
    if (endM) text = text.substring(0, endM.index);

    // Normalize claim-number artifacts from Word/PDF extraction:
    // "6 . Text"  → "6. Text"  (space between digit and period)
    // "1 0. Text" → "10. Text" (space inside two-digit number)
    text = text.replace(/(^|\n)(\s*)(\d+)\s+([.)]\s)/gm, function (m, nl, lead, num, sep) {
      return nl + lead + num + sep;
    });
    text = text.replace(/(^|\n)(\s*)(\d)\s(\d+)([.)]\s)/gm, function (m, nl, lead, d1, d2, sep) {
      return nl + lead + d1 + d2 + sep;
    });

    var claims = [];
    // Match numbered claims: "N." or "N)" at line start, then text until next claim or end
    var re = /(?:^|\r?\n)\s*(\d+)[.)]\s+([\s\S]+?)(?=\r?\n\s*\d+[.)]\s|\s*$)/g;
    var m;
    while ((m = re.exec(text)) !== null) {
      var num = parseInt(m[1], 10);
      // Normalize whitespace and remove bullet characters used for element lists
      var raw = m[2]
        .replace(/\r?\n\s*/g, ' ')
        .replace(/\s*[•·◦]\s*/g, ' ')
        .replace(/\s{2,}/g, ' ')
        .trim();

      var claim = {
        number:       num,
        rawText:      raw,
        preamble:     '',
        transition:   '',
        body:         '',
        isDependent:  false,
        parentClaims: [],
        isMultiDep:   false,
        claimType:    'other',
        _introduced:  new Set()   // populated by antecedent-basis check
      };

      parseParts(claim);
      detectDependency(claim);
      detectClaimType(claim);
      claims.push(claim);
    }
    return claims;
  }

  function parseParts(claim) {
    var text  = claim.rawText;
    var lower = text.toLowerCase();

    for (var i = 0; i < TRANSITIONS.length; i++) {
      var t   = TRANSITIONS[i];
      var idx = lower.indexOf(t);
      if (idx === -1) continue;

      claim.preamble   = text.substring(0, idx).replace(/,\s*$/, '').trim();
      claim.transition = t;
      var afterT       = text.substring(idx + t.length).trimStart();
      claim.body       = afterT.startsWith(':') ? afterT.substring(1).trimStart() : afterT;
      return;
    }
    // No recognized transition — store full text as preamble
    claim.preamble = text;
  }

  function detectDependency(claim) {
    var text = claim.rawText;

    // Multiple-dependency patterns
    var isMulti = /\bany\s+one\s+of\s+claims?\s+/i.test(text)
               || /\bclaims?\s+\d+(?:\s*[,/]\s*\d+)+/i.test(text)
               || /\bclaims?\s+\d+\s+(?:or|and|through|to)\s+\d+\b/i.test(text);
    if (isMulti) claim.isMultiDep = true;

    // Collect ALL referenced claim numbers — include forward refs and non-existent ones.
    // D-02 flags forward refs; D-06 flags non-existent refs. No pre-filtering here.
    var parents = [];
    var re = /\bclaims?\s+(\d+)\b/gi;
    var m;
    while ((m = re.exec(text)) !== null) {
      var n = parseInt(m[1], 10);
      if (n !== claim.number && !parents.includes(n)) parents.push(n);
    }
    // Also handle "claim N or M" / "claim N and M"
    var altM = text.match(/\bclaims?\s+(\d+)\s+(?:or|and)\s+(\d+)\b/i);
    if (altM) {
      [parseInt(altM[1], 10), parseInt(altM[2], 10)].forEach(function (n) {
        if (n !== claim.number && !parents.includes(n)) parents.push(n);
      });
    }

    if (parents.length > 0) {
      claim.isDependent  = true;
      claim.parentClaims = parents;
    }
  }

  function detectClaimType(claim) {
    var preamble = claim.preamble || claim.rawText;

    // CRM always wins (specific term)
    if (CATEGORY_PATTERNS.crm.test(preamble)) { claim.claimType = 'crm'; return; }

    // For independent claims: try to extract leading noun from "A/An [adj]* NOUN" to get
    // the most specific label (system, method, apparatus, device…) from the actual language.
    if (!claim.isDependent) {
      var m = preamble.match(/^(?:A|An)\s+(?:[a-zA-Z][a-zA-Z0-9\-]*\s+){0,5}([a-zA-Z][a-zA-Z0-9\-]+)/i);
      if (m) {
        var noun = m[1].toLowerCase();
        if (/^(method|process|procedure|technique)$/.test(noun)) { claim.claimType = 'method'; return; }
        if (/^(composition|compound|mixture|formulation|preparation|blend|alloy|polymer)$/.test(noun)) { claim.claimType = 'composition'; return; }
        if (/^(use)$/.test(noun) && /\buse\s+of\b/i.test(preamble)) { claim.claimType = 'use'; return; }
        if (/^(system|systems)$/.test(noun)) { claim.claimType = 'system'; return; }
        if (/^(apparatus|device|machine|assembly|arrangement|unit|circuit|controller|module|sensor|network|platform|interface)$/.test(noun)) { claim.claimType = 'apparatus'; return; }
      }
    }

    // Fallback keyword matching (handles dependent claims and unusual preambles)
    var keys = ['method', 'system', 'apparatus', 'composition', 'use'];
    for (var i = 0; i < keys.length; i++) {
      if (CATEGORY_PATTERNS[keys[i]].test(preamble)) {
        claim.claimType = keys[i];
        return;
      }
    }
    claim.claimType = 'other';
  }

  // ── Metrics ─────────────────────────────────────────────────────────────────

  function computeMetrics(claims) {
    var m = {
      total:           claims.length,
      independent:     0,
      dependent:       0,
      independentNums: [],
      dependentNums:   [],
      byType:          { method: 0, system: 0, apparatus: 0, crm: 0, composition: 0, use: 0, other: 0 }
    };
    claims.forEach(function (c) {
      if (c.isDependent) {
        m.dependent++;
        m.dependentNums.push(c.number);
      } else {
        m.independent++;
        m.independentNums.push(c.number);
      }
      m.byType[c.claimType] = (m.byType[c.claimType] || 0) + 1;
    });
    return m;
  }

  // ── Finding helpers ─────────────────────────────────────────────────────────

  function f(ruleCode, severity, claimNum, message) {
    return { ruleCode: ruleCode, severity: severity, claimNumber: claimNum, message: message };
  }

  // ── FORMAT rules ────────────────────────────────────────────────────────────

  function checkF01(c) {
    // Multiple sentence-ending periods in claim (not abbreviations or decimals)
    var cleaned = c.rawText
      .replace(/\.\s*$/, '')                                  // strip terminal period
      .replace(/\b(e\.g\.|i\.e\.|etc\.|et al\.|vs\.|no\.|fig\.|art\.)\s/gi, ' ')
      .replace(/\b[A-Z]\.\s/g, ' ')                          // "U.S. ", "e.g. "
      .replace(/\d+\.\d+/g, '#')                             // decimals
      .replace(/\.\)/g, ')');                                 // parenthesized refs like (2.)

    if (/\.\s+[A-Z]/.test(cleaned)) {
      return f('F-01', 'error', c.number,
        'Multiple sentences detected — a claim must be written as a single sentence');
    }
    return null;
  }

  function checkF02(c) {
    // Simple past tense auxiliaries (excludes past-participle adjectives like "attached")
    if (/\b(was|were|has been|have been|had been|would have)\b/i.test(c.rawText)) {
      return f('F-02', 'warning', c.number,
        'Past-tense verb detected ("was"/"were"/etc.) — claims should use present tense');
    }
    return null;
  }

  function checkF03(c) {
    // Dependent claims use "wherein" to add limitations — no formal transition required
    if (c.isDependent) return null;
    if (!c.transition) {
      return f('F-03', 'error', c.number,
        'No transitional phrase found (comprising / including / consisting of / etc.)');
    }
    return null;
  }

  function checkF04(c) {
    if (c.isDependent || !c.preamble) return null;
    var hasCategory = /\b(apparatus|device|system|machine|assembly|method|process|composition|compound|mixture|use|medium|product|module|unit|circuit|controller|arrangement|sensor|network|platform|interface|material|structure|element|article|formulation)\b/i
      .test(c.preamble);
    if (!hasCategory) {
      return f('F-04', 'warning', c.number,
        'Preamble may lack a clear category noun (apparatus / method / system / composition / etc.)');
    }
    return null;
  }

  function checkF05(claims) {
    // Consecutive numbering check — report first gap only
    for (var i = 0; i < claims.length; i++) {
      if (claims[i].number !== i + 1) {
        return f('F-05', 'error', claims[i].number,
          'Claim numbering gap — expected claim ' + (i + 1) + ', found claim ' + claims[i].number);
      }
    }
    return null;
  }

  function checkF06(c) {
    // Single-element independent claim body
    if (c.isDependent || !c.body) return null;
    var parts = c.body.split(/;\s*(?:and\s+)?|[•·◦]/).filter(function (p) { return p.trim().length > 5; });
    if (parts.length < 2) {
      return f('F-06', 'warning', c.number,
        'Independent claim has only one element — consider whether scope is defensibly narrow');
    }
    return null;
  }

  function checkF07(c) {
    // Body elements appear disconnected (no relational language)
    if (c.isDependent || !c.body || c.body.length < 30) return null;
    var hasRelational = /\b(configured\s+to|coupled|connected|attached|mounted|disposed|operably|operatively|in\s+communication|electrically|thermally|in\s+fluid|wherein|such\s+that|so\s+as\s+to|adapted\s+to|operable\s+to|stored\s+in|received\s+from|transmitted\s+to|linked\s+to|integrated\s+with)\b/i
      .test(c.body);
    if (!hasRelational) {
      return f('F-07', 'warning', c.number,
        'Claim elements may not be connected — elements must interrelate, not merely be listed');
    }
    return null;
  }

  function checkF08(claims) {
    if (claims.length > 0 && claims[0].isDependent) {
      return f('F-08', 'error', 1,
        'Claim 1 must be independent — it currently references another claim');
    }
    return null;
  }

  // ── GRAMMAR rules ───────────────────────────────────────────────────────────

  function singularize(w) {
    if (!w) return w;
    if (w.endsWith('ies') && w.length > 4) return w.slice(0, -3) + 'y';
    if (w.endsWith('ses') || w.endsWith('xes') || w.endsWith('zes')) return w.slice(0, -2);
    if (w.endsWith('s') && !w.endsWith('ss') && w.length > 3) return w.slice(0, -1);
    return w;
  }

  var _STOP_WORDS_SET = new Set(_NP_STOP.split('|'));

  function getHeadNoun(phrase) {
    var words = phrase.toLowerCase().trim().split(/\s+/);
    for (var i = words.length - 1; i >= 0; i--) {
      var w = words[i];
      if (!ORDINAL_WORDS.has(w) && !_STOP_WORDS_SET.has(w) && w.length > 1) return w;
    }
    return words[words.length - 1];
  }

  function checkAntecedentBasis(c, claims) {
    var findings = [];
    var text     = c.rawText;

    // Start with introduced set — inherit union of all parents (handles single-dep and multi-dep)
    var introduced = new Set();
    if (c.isDependent && c.parentClaims.length > 0) {
      c.parentClaims.forEach(function (pNum) {
        for (var pi = 0; pi < claims.length; pi++) {
          if (claims[pi].number === pNum && claims[pi]._introduced) {
            claims[pi]._introduced.forEach(function (n) { introduced.add(n); });
            break;
          }
        }
      });
    }

    // Collect introductions — rebuild regexes each call to reset lastIndex
    var introRe    = new RegExp(_INTRO_SRC,     'gi');
    var altIntroRe = new RegExp(_ALT_INTRO_SRC, 'gi');
    // Gerund direct-object pattern: "obtaining spatial data", "receiving a request" etc.
    // Captures bare nouns (mass nouns) used as direct objects of method/CRM step verbs.
    var gerundRe = /\b(?:obtaining|generating|receiving|storing|computing|creating|collecting|determining|extracting|identifying|detecting|processing|transmitting|calculating|fetching|reading|writing|selecting|retrieving|encoding|decoding|converting|analyzing|evaluating|monitoring|measuring|comparing|allocating|scheduling|authenticating|validating|compiling|executing|initiating|performing|applying)\s+([a-z][a-z\-]+(?:\s+(?!of\b|from\b|in\b|to\b|with\b|by\b|for\b|at\b|and\b|or\b|via\b|into\b|on\b)[a-z][a-z\-]+){0,2})/gi;
    var m;

    function addPhrase(raw) {
      var phrase = raw.toLowerCase().trim();
      introduced.add(phrase);
      introduced.add(getHeadNoun(phrase));
      introduced.add(singularize(getHeadNoun(phrase)));
    }

    // "a/an X" introductions
    while ((m = introRe.exec(text)) !== null)   addPhrase(m[1]);
    // "one or more X" / "at least one X" introductions
    while ((m = altIntroRe.exec(text)) !== null) addPhrase(m[1]);
    // Gerund bare-noun introductions (mass nouns in method/CRM steps)
    while ((m = gerundRe.exec(text)) !== null)  addPhrase(m[1]);

    c._introduced = new Set(introduced);

    function isIntroduced(phrase) {
      var lp   = phrase.toLowerCase().trim();
      var head = getHeadNoun(lp);
      var sing = singularize(head);
      if (KNOWN_THE_WORDS.has(head) || KNOWN_THE_WORDS.has(lp)) return true;
      if (introduced.has(lp) || introduced.has(head) || introduced.has(sing)) return true;
      for (var n of introduced) {
        var nHead = getHeadNoun(n);
        if (nHead === head || nHead === sing) return true;
      }
      return false;
    }

    // Check back-references: "the X" / "said X"
    var flagged = new Set();
    var backRe  = new RegExp(_BACK_SRC, 'gi');
    while ((m = backRe.exec(text)) !== null) {
      var noun = m[1].toLowerCase().trim();
      if (flagged.has(noun)) continue;
      if (isIntroduced(noun)) continue;
      flagged.add(noun);
      findings.push(f('G-01', 'error', c.number,
        '"the ' + noun + '" used without prior introduction with "a" or "an"'));
    }

    return findings;
  }

  function checkG03(c) {
    if (!c.transition) return null;
    var idx = c.rawText.toLowerCase().indexOf(c.transition);
    if (idx === -1) return null;
    var after = c.rawText.substring(idx + c.transition.length).trimStart();
    if (!after.startsWith(':')) {
      return f('G-03', 'warning', c.number,
        '"' + c.transition + '" should be followed by a colon ":"');
    }
    return null;
  }

  function checkG04(c) {
    if (!c.transition) return null;
    var idx = c.rawText.toLowerCase().indexOf(c.transition);
    if (idx === -1) return null;
    var before = c.rawText.substring(0, idx).trimEnd();
    if (!before.endsWith(',')) {
      return f('G-04', 'warning', c.number,
        'Preamble should end with a comma before "' + c.transition + '"');
    }
    return null;
  }

  function checkG05(c) {
    if (!c.rawText.trim().endsWith('.')) {
      return f('G-05', 'warning', c.number, 'Claim does not end with a period');
    }
    return null;
  }

  function checkG06(c) {
    // Same noun introduced more than once with "a" without first/second disambiguation
    // Uses head noun as key so "a conversation" + "a conversation type" don't collide
    var text   = c.rawText.toLowerCase();
    var counts = {};
    var re     = new RegExp(_INTRO_SRC, 'gi');
    var m;
    while ((m = re.exec(text)) !== null) {
      var noun    = getHeadNoun(m[1]);  // head noun as dedup key
      var context = text.substring(Math.max(0, m.index - 25), m.index);
      if (/\b(first|second|third|fourth|another|further)\b/.test(context)) continue;
      counts[noun] = (counts[noun] || 0) + 1;
    }
    var findings = [];
    Object.keys(counts).forEach(function (noun) {
      if (counts[noun] >= 2) {
        findings.push(f('G-06', 'error', c.number,
          '"a ' + noun + '" introduced more than once — use "first ' + noun +
          '" / "second ' + noun + '" or "another ' + noun + '" to distinguish them'));
      }
    });
    return findings;
  }

  function checkG07(c) {
    if (!c.body) return [];
    // "Wherein" clauses should not introduce new structural elements
    var findings = [];
    var whereins = c.body.match(/\bwherein\b[^;.]*/gi) || [];
    whereins.forEach(function (wc) {
      if (/\b(a|an)\s+[a-z][a-z\-]+/.test(wc.toLowerCase())) {
        findings.push(f('G-07', 'warning', c.number,
          '"wherein" clause introduces a new element — "wherein" should describe properties of elements already recited'));
      }
    });
    return findings;
  }

  // ── SCOPE rules ─────────────────────────────────────────────────────────────

  function checkS01(c) {
    var text = c.rawText;
    // Skip if "or" only appears in dependency references
    if (/\bclaim[s]?\s+\d+\s+or\s+\d+/i.test(text) && !/ or /i.test(text.replace(/\bclaim[s]?\s+[\d\s,\-–orand]+/gi, ''))) {
      return null;
    }
    // Look for bare "or" connecting noun-like alternatives (not conjunctions with prepositions)
    var bareOr = text.replace(/\bclaim[s]?\s+[\d\s,\-–orand]+/gi, '').match(/\b([a-z]{2,})\s+or\s+([a-z]{2,})\b/gi);
    if (!bareOr) return null;
    var stopWords = new Set(['that', 'which', 'this', 'where', 'when', 'if', 'as', 'its', 'any', 'each', 'other', 'not', 'and', 'but', 'nor']);
    var hasBare = bareOr.some(function (match) {
      var parts = match.toLowerCase().split(/\s+or\s+/);
      return parts.every(function (p) { return !stopWords.has(p.trim()); });
    });
    if (hasBare) {
      return f('S-01', 'warning', c.number,
        'Bare "or" connecting alternatives may be unclear — consider Markush format ("selected from the group consisting of A, B and C") or "and/or"');
    }
    return null;
  }

  function checkS02(c) {
    if (c.transition === 'consisting of') {
      return f('S-02', 'info', c.number,
        '"Consisting of" creates a closed claim — competitors may avoid infringement by adding one additional element');
    }
    return null;
  }

  function checkS03(c) {
    if (c.transition !== 'consisting of') return null;
    // Check if percentages are present and attempt to sum them
    var percents = [];
    var re = /(\d+(?:\.\d+)?)\s*%/g;
    var m;
    while ((m = re.exec(c.rawText)) !== null) percents.push(parseFloat(m[1]));
    if (percents.length >= 2) {
      var total = percents.reduce(function (a, b) { return a + b; }, 0);
      if (Math.abs(total - 100) > 2) {
        return f('S-03', 'error', c.number,
          'Closed "consisting of" claim with percentages — values sum to ' + total.toFixed(1) +
          '% (should sum to 100%)');
      }
    }
    return null;
  }

  function checkS04(c) {
    if (c.isDependent || !c.body) return null;
    var words = c.body.split(/\s+/).filter(function (w) { return w.length > 2; });
    if (words.length < 4) {
      return f('S-04', 'warning', c.number,
        'Claim body is very short — may be impossibly broad or missing key elements');
    }
    return null;
  }

  function checkS05(c) {
    // Method claims should have gerund-form steps
    if (c.claimType !== 'method' || c.isDependent || !c.body) return null;
    if (!/\b[a-z]{3,}ing\b/.test(c.body)) {
      return f('S-05', 'warning', c.number,
        'Method claim body has no gerund-form steps (-ing verbs) — use "receiving," "determining," "transmitting," etc.');
    }
    return null;
  }

  // ── DEPENDENCY rules ─────────────────────────────────────────────────────────

  function checkD01(c) {
    // Claim text looks dependent but no prior claim number was parsed
    var looksDependent = /^the\s+[a-z]+\s+(of|according\s+to|recited\s+in|as\s+claimed\s+in|as\s+defined\s+in)\b/i
      .test(c.rawText);
    if (looksDependent && !c.isDependent) {
      return f('D-01', 'error', c.number,
        'Claim looks dependent but no prior claim number reference was found');
    }
    return null;
  }

  function checkD02(claims) {
    var findings = [];
    claims.forEach(function (c) {
      if (!c.isDependent) return;
      c.parentClaims.forEach(function (p) {
        if (p >= c.number) {
          findings.push(f('D-02', 'error', c.number,
            'Forward reference — claim ' + c.number + ' references claim ' + p + ' which appears later in the set'));
        }
      });
    });
    return findings;
  }

  function checkD03(c) {
    if (!c.isDependent) return null;
    // After the dependency reference phrase, is there any new limitation?
    var text  = c.rawText;
    // Strip the dependency reference portion
    var stripped = text.replace(/\bthe\s+[a-z\s]+(?:of|according\s+to|recited\s+in|as\s+claimed\s+in|as\s+defined\s+in)\s+(?:any\s+one\s+of\s+)?claims?\s+[\d\s,\-–orand]+[,;]?\s*/i, '').trim();
    if (stripped.replace(/[.,;]/g, '').trim().length < 8) {
      return f('D-03', 'error', c.number,
        'Dependent claim adds no new limitation — every dependent claim must add at least one restriction');
    }
    return null;
  }

  function checkD04(c) {
    // Multiple-dep should use "any one of claims" format
    if (!c.isMultiDep) return null;
    if (!/\bany\s+one\s+of\s+claims?\b/i.test(c.rawText) &&
        !/\bclaims?\s+\d+(?:\s*[,/]\s*\d+)+/i.test(c.rawText)) {
      return f('D-04', 'warning', c.number,
        'Multiple dependency should use standard format: "any one of claims N–M" or "any one of claims N, M and P"');
    }
    return null;
  }

  function checkD05(c, claims, jurisdiction) {
    if (jurisdiction !== 'US' || !c.isMultiDep) return null;
    var claimMap = {};
    claims.forEach(function (x) { claimMap[x.number] = x; });
    var multiDepParent = c.parentClaims.find(function (p) {
      return claimMap[p] && claimMap[p].isMultiDep;
    });
    if (multiDepParent) {
      return f('D-05', 'error', c.number,
        'Multiple-multiple dependency not allowed in US practice — claim ' + c.number +
        ' is a multiple-dep that references claim ' + multiDepParent + ' (also a multiple-dep)');
    }
    return null;
  }

  function checkD06(claims) {
    var findings = [];
    var claimNums = new Set(claims.map(function (c) { return c.number; }));
    claims.forEach(function (c) {
      if (!c.isDependent) return;
      c.parentClaims.forEach(function (p) {
        if (!claimNums.has(p)) {
          findings.push(f('D-06', 'error', c.number,
            'References claim ' + p + ' which does not exist in the claim set'));
        }
      });
    });
    return findings;
  }

  function checkD07(claims, jurisdiction) {
    var findings = [];
    var total       = claims.length;
    var multiDepCnt = claims.filter(function (c) { return c.isMultiDep; }).length;
    if (jurisdiction === 'US' && total > 20) {
      findings.push(f('D-07', 'info', 0,
        total + ' claims — USPTO charges excess fees for claims beyond 20 independent + total thresholds'));
    } else if (jurisdiction === 'EPO' && total > 10) {
      findings.push(f('D-07', 'info', 0,
        total + ' claims — EPO charges excess claim fees for claims 11 and above'));
    }
    if (jurisdiction === 'US' && multiDepCnt > 0) {
      findings.push(f('D-07', 'info', 0,
        multiDepCnt + ' multiple dependent claim(s) — USPTO charges an additional fee per multiple-dep claim'));
    }
    return findings;
  }

  // ── ELEMENTS rules ───────────────────────────────────────────────────────────

  function checkE01(c) {
    if (c.claimType !== 'method' || c.isDependent || !c.body) return null;
    // Only flag infinitive form at the START of a body segment (after ';' or line break).
    // "to determine" mid-sentence is a purpose clause, not a method step.
    // Only flag when a body segment STARTS with "to + verb" — purpose clauses like
    // "...to determine X" mid-sentence are valid and should not be flagged.
    var INFINITIVE_AT_START = /^to\s+(receive|send|transmit|process|calculate|determine|generate|store|retrieve|compare|detect|identify|select|output|input|perform|execute|analyze|compute|measure|apply|convert|display|update|configure|allocate|schedule|authenticate|validate|encode|decode)\b/i;
    var segments = c.body.split(/[;\n]/);
    var badSeg = segments.some(function (seg) {
      return INFINITIVE_AT_START.test(seg.trim().replace(/^[\s\-–•]+/, ''));
    });
    if (badSeg) {
      return f('E-01', 'warning', c.number,
        'Method step uses infinitive form ("to receive") — prefer gerund form ("receiving")');
    }
    return null;
  }

  function checkE02(c) {
    if (/\bmeans\s+for\b/i.test(c.rawText)) {
      return f('E-02', 'info', c.number,
        '"Means for" language detected — ensure corresponding structure is described in the specification (§112(f)/Rule 19)');
    }
    return null;
  }

  function checkE03(c) {
    if (c.isDependent || !c.body) return null;
    var elements = c.body.split(/;\s*/).filter(function (e) { return e.trim().length > 5; });
    if (elements.length < 2) return null;
    var allMeansFor = elements.every(function (el) {
      return /^\s*(means\s+for|step\s+for)\b/i.test(el.trim());
    });
    if (allMeansFor) {
      return f('E-03', 'error', c.number,
        'All elements defined as "means for" — risk of rejection as a mere desideratum (desired result without structure)');
    }
    return null;
  }

  function checkE04(c) {
    // Numeric ranges should include explicit endpoints
    var openRanges = c.rawText.match(/\bbetween\s+\d[\d.]*\s+and\s+\d[\d.]*\b/gi);
    if (openRanges) {
      return f('E-04', 'warning', c.number,
        'Numeric range uses "between X and Y" — clarify whether endpoints are inclusive; prefer "from X to Y"');
    }
    return null;
  }

  function checkE05(c, jurisdiction) {
    if (jurisdiction !== 'US') return null;
    if (/\(\d{1,4}\)/.test(c.rawText)) {
      return f('E-05', 'warning', c.number,
        'Reference numerals in parentheses (e.g., "(100)") are generally avoided in US claims — risk of limiting scope');
    }
    return null;
  }

  // ── SPECIAL FORMAT rules ─────────────────────────────────────────────────────

  function checkSF01(c) {
    var twoPart = ['characterized in that', 'characterized by', 'wherein the improvement comprises'];
    if (twoPart.indexOf(c.transition) !== -1) {
      return f('SF-01', 'info', c.number,
        'Two-part (Jepson/improvement) claim — the preamble constitutes an admission that those features are known prior art');
    }
    return null;
  }

  function checkSF02(c) {
    if (/selected\s+from\s+the\s+group\s+comprising/i.test(c.rawText)) {
      return f('SF-02', 'error', c.number,
        'Markush format error — use "selected from the group consisting of" (not "comprising"); "comprising" creates an open group');
    }
    return null;
  }

  function checkSF03(c) {
    if (!/selected\s+from\s+the\s+group\s+consisting\s+of/i.test(c.rawText)) return null;
    if (!/combinations?\s+thereof/i.test(c.rawText)) {
      return f('SF-03', 'info', c.number,
        'Markush group without "and combinations thereof" — claim covers only one member at a time; add "and combinations thereof" if multiple simultaneous selections are intended');
    }
    return null;
  }

  function checkSF04(c) {
    if (/\bmeans\s+for\b/i.test(c.rawText)) {
      return f('SF-04', 'info', c.number,
        'Means-plus-function language — claim scope is limited to structures disclosed in the specification and their equivalents');
    }
    return null;
  }

  function checkSF05(c, jurisdiction) {
    // EPO prefers two-part format for improvement inventions
    if (jurisdiction !== 'EPO') return null;
    if (c.isDependent || !c.preamble) return null;
    // If claim describes an improvement over a known base, EPO expects two-part format
    // We flag if the word "improved" or "enhanced" appears in preamble (suggests improvement)
    if (/\b(improved|enhanced|novel|new)\b/i.test(c.preamble)) {
      return f('SF-05', 'info', c.number,
        'EPO prefers two-part format for improvement claims — consider "characterized in that" to separate known features from novel ones');
    }
    return null;
  }

  // ── Main runner ──────────────────────────────────────────────────────────────

  function applyAllRules(claims, jurisdiction) {
    var out = {
      format:          [],
      antecedentBasis: [],
      grammar:         [],
      scope:           [],
      dependency:      [],
      elements:        [],
      specialFormats:  []
    };

    // Set-level format checks
    var f05 = checkF05(claims);
    if (f05) out.format.push(f05);
    var f08 = checkF08(claims);
    if (f08) out.format.push(f08);

    // Set-level dependency checks
    out.dependency = out.dependency.concat(checkD02(claims));
    out.dependency = out.dependency.concat(checkD06(claims));
    out.dependency = out.dependency.concat(checkD07(claims, jurisdiction));

    // Per-claim checks (must run in claim order for antecedent basis inheritance)
    claims.forEach(function (c) {
      var fi;

      // FORMAT
      fi = checkF01(c); if (fi) out.format.push(fi);
      fi = checkF02(c); if (fi) out.format.push(fi);
      fi = checkF03(c); if (fi) out.format.push(fi);
      fi = checkF04(c); if (fi) out.format.push(fi);
      fi = checkF06(c); if (fi) out.format.push(fi);
      fi = checkF07(c); if (fi) out.format.push(fi);

      // ANTECEDENT BASIS — must run before other grammar checks (populates c._introduced)
      out.antecedentBasis = out.antecedentBasis.concat(checkAntecedentBasis(c, claims));
      fi = checkG03(c); if (fi) out.grammar.push(fi);
      fi = checkG04(c); if (fi) out.grammar.push(fi);
      fi = checkG05(c); if (fi) out.grammar.push(fi);
      out.grammar = out.grammar.concat(checkG06(c));
      out.grammar = out.grammar.concat(checkG07(c));

      // SCOPE
      fi = checkS01(c); if (fi) out.scope.push(fi);
      fi = checkS02(c); if (fi) out.scope.push(fi);
      fi = checkS03(c); if (fi) out.scope.push(fi);
      fi = checkS04(c); if (fi) out.scope.push(fi);
      fi = checkS05(c); if (fi) out.scope.push(fi);

      // DEPENDENCY (per-claim)
      fi = checkD01(c); if (fi) out.dependency.push(fi);
      fi = checkD03(c); if (fi) out.dependency.push(fi);
      fi = checkD04(c); if (fi) out.dependency.push(fi);
      fi = checkD05(c, claims, jurisdiction); if (fi) out.dependency.push(fi);

      // ELEMENTS
      fi = checkE01(c); if (fi) out.elements.push(fi);
      fi = checkE02(c); if (fi) out.elements.push(fi);
      fi = checkE03(c); if (fi) out.elements.push(fi);
      fi = checkE04(c); if (fi) out.elements.push(fi);
      fi = checkE05(c, jurisdiction); if (fi) out.elements.push(fi);

      // SPECIAL FORMATS
      fi = checkSF01(c); if (fi) out.specialFormats.push(fi);
      fi = checkSF02(c); if (fi) out.specialFormats.push(fi);
      fi = checkSF03(c); if (fi) out.specialFormats.push(fi);
      fi = checkSF04(c); if (fi) out.specialFormats.push(fi);
      fi = checkSF05(c, jurisdiction); if (fi) out.specialFormats.push(fi);
    });

    // Strip any nulls that slipped in via .concat() on functions that returned null
    Object.keys(out).forEach(function (k) { out[k] = out[k].filter(Boolean); });

    return out;
  }

  // ── Public API ───────────────────────────────────────────────────────────────

  return {
    analyze: function (text, jurisdiction) {
      jurisdiction = jurisdiction || 'US';

      var claimsText = extractClaimsSection(text);
      if (!claimsText) {
        return { error: 'no_claims_section', claims: [], metrics: null, findings: null };
      }

      var claims = parseClaims(claimsText);
      if (claims.length === 0) {
        return { error: 'no_claims_found', claims: [], metrics: null, findings: null };
      }

      var metrics  = computeMetrics(claims);
      var findings = applyAllRules(claims, jurisdiction);

      return { error: null, claims: claims, metrics: metrics, findings: findings };
    },

    // Called from cursor mode — skips heading detection, treats text as already the claims section
    analyzeFromText: parseClaimsDirectly,

    // Exposed for unit tests
    _extractClaimsSection: extractClaimsSection,
    _parseClaims:          parseClaims,
    _computeMetrics:       computeMetrics
  };

})();
