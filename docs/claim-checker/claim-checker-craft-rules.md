# Claim Checker — Craft Rules
## Rules for automated claim validation in Word add-in
## Derived from: WIPO Patent Drafting Manual (2nd ed., 2023), pp. 42–71

---

## RULE CATEGORIES

1. **FORMAT** — Structural/syntactic rules about claim shape
2. **GRAMMAR** — Antecedent basis, punctuation, connective words
3. **SCOPE** — Breadth, transitional phrases, open/closed claims
4. **DEPENDENCY** — Claim set structure and relationships
5. **ELEMENTS** — Types and connections of claim elements
6. **SPECIAL FORMATS** — Two-part, Markush, M+F, multiple dependency

---

## CATEGORY 1: FORMAT RULES

### F-01 Single Sentence Rule
**Rule:** Each claim must be written as a single sentence.
**Check:** Detect multiple period-terminated sentences within a single numbered claim.
**Severity:** Error

### F-02 Present Tense Rule
**Rule:** Claims must be written in present tense.
**Check:** Flag past tense verbs (e.g., "was," "were," "provided," "attached [past]") in claim body.
**Severity:** Warning
**Exception:** "wherein" clauses describing relationships may use past participles as adjectives ("attached to," "connected to") — these are acceptable.

### F-03 Three-Part Structure Rule
**Rule:** Every claim must have: Preamble → Transitional Phrase → Body.
**Check:** Detect claims lacking a recognized transitional phrase (comprising/including/consisting of/characterized by/wherein the improvement comprises).
**Severity:** Error

### F-04 Preamble Must Identify Category
**Rule:** Preamble must identify claim category (apparatus, method, system, composition, compound, device, process, use).
**Check:** Flag if preamble does not contain one of these (or recognized synonym).
**Severity:** Warning

### F-05 Numbering Rule
**Rule:** Claims must be consecutively numbered starting at 1.
**Check:** Detect gaps in claim numbering or non-sequential numbering.
**Severity:** Error

### F-06 Claim Starts on New Line
**Rule:** Each claim begins on a new paragraph/line.
**Check:** Structural check on paragraph breaks between numbered claims.
**Severity:** Warning

### F-07 Body Must Have Connected Elements
**Rule:** Claim body cannot be a mere unconnected list; elements must be related to each other.
**Check:** Flag claims where all elements lack relational terms (configured to, attached to, connected to, coupled to, electrically connected to, in communication with, disposed within, etc.).
**Severity:** Warning

### F-08 Multiple Elements Required
**Rule:** Claims generally require at least two elements (single-element claims are presumptively too broad).
**Check:** Flag independent claims with only one body element.
**Severity:** Warning

---

## CATEGORY 2: GRAMMAR RULES

### G-01 Antecedent Basis — First Introduction
**Rule:** Every element must be introduced with "a" or "an" on its first occurrence in a claim.
**Check:** Detect elements preceded by "the" or "said" without prior introduction with "a/an" in same claim.
**Severity:** Error (§112(b) / Art. 84 EPC defect)

### G-02 Antecedent Basis — Back-Reference
**Rule:** After first introduction with "a/an", subsequent references to same element must use "the" or "said".
**Check:** Detect elements re-introduced with "a/an" after already being referenced with "the/said" in same claim.
**Severity:** Warning

### G-03 Antecedent Basis — Cross-Claim Independence
**Rule:** Antecedent basis must be re-established in each claim independently. Claim N cannot rely on antecedent from Claim N-1 for a NEW element.
**Check:** Each numbered claim is parsed independently for antecedent basis.
**Severity:** Error

### G-04 Punctuation — Preamble Separator
**Rule:** Preamble must be separated from transitional phrase by a comma.
**Check:** Flag if no comma found between preamble text and "comprising" / "consisting" / etc.
**Severity:** Warning

### G-05 Punctuation — Transitional Phrase Separator
**Rule:** Transitional phrase must be followed by a colon before body elements.
**Check:** Flag if "comprising" / "consisting of" / etc. not followed by colon.
**Severity:** Warning

### G-06 Punctuation — Body Element Separators
**Rule:** Body elements separated by semicolons; last element ends with period.
**Check:** Flag if elements separated by commas only (ambiguous parsing) or if claim does not end with period.
**Severity:** Warning

### G-07 "Said" vs "The"
**Rule:** "Said" and "the" are interchangeable for back-reference; "said" is archaic but not incorrect.
**Check:** Flag excessive "said" usage as style warning (not error) — prefer "the" for plain language.
**Severity:** Info

### G-08 Wherein Clause Logic
**Rule:** "Wherein" clause must describe result/property flowing from previously recited structure — not introduce new structural elements not already in the claim.
**Check:** Flag "wherein" clauses that introduce new nouns not previously established in the claim.
**Severity:** Warning

### G-09 First/Second Disambiguation
**Rule:** When a claim recites two instances of the same element, they must be distinguished as "first [element]" and "second [element]" (or "another [element]").
**Check:** Detect two uses of "a [same noun]" in one claim without first/second or another disambiguation.
**Severity:** Error

---

## CATEGORY 3: SCOPE / TRANSITIONAL PHRASE RULES

### S-01 Recognize Open-Ended Transitions
**Rule:** "Comprising," "including," "containing," "characterized by" = open-ended (allows additional elements).
**Check:** Flag no action; detect for S-02 / S-03 comparisons.
**Informational tag:** OPEN

### S-02 Recognize Closed Transitions
**Rule:** "Consisting of" = closed (NO additional elements allowed beyond those listed).
**Check:** Flag closed claims with WARNING that competitors can avoid infringement by adding one element.
**Severity:** Warning (strategic note, not defect)

### S-03 Closed Claim Chemical Percentages
**Rule:** In "consisting of" chemical composition claims, all component percentages must sum to 100%.
**Check:** If claim uses "consisting of" AND contains percentage values, verify they sum to 100 (± rounding tolerance).
**Severity:** Error if sum ≠ 100

### S-04 "Or" Ambiguity
**Rule:** Use of plain "or" in claims is generally considered unclear in patent language.
**Check:** Flag bare "or" connecting technical alternatives (e.g., "metal or plastic") without Markush format.
**Severity:** Warning
**Exception:** "or" in "any one of Claims X or Y" dependency references is standard and acceptable.

### S-05 Open Phrase in Preamble vs. Body Consistency
**Rule:** If body uses "consisting of" (closed), ensure preamble is not so broad as to be contradicted by closed body.
**Check:** Flag conflict between open preamble breadth claim and closed body — raise strategic warning.
**Severity:** Info

---

## CATEGORY 4: DEPENDENCY RULES

### D-01 Dependency Reference Format
**Rule:** Dependent claim must explicitly reference at least one prior claim number.
**Valid patterns:**
- "The [X] according to claim N, wherein..."
- "The [X] of claim N, further comprising..."
- "The [X] recited in claim N, wherein..."
- "The method of claim N, wherein the step of..."
**Check:** Flag dependent-looking claims (not Claim 1 type) that lack "claim N" reference.
**Severity:** Error

### D-02 Dependency Hierarchy — No Forward References
**Rule:** A dependent claim can only refer to a PRIOR (lower-numbered) claim.
**Check:** Flag any claim N that references claim M where M > N.
**Severity:** Error

### D-03 Dependent Claim Adds New Limitation
**Rule:** Every dependent claim must add at least one limitation not present in the referenced claim.
**Check:** Flag dependent claims that appear identical to parent claim body (no additional text beyond reference phrase).
**Severity:** Error

### D-04 Multiple Dependent Claim Format
**Rule:** Multiple dependent claims must use "any one of Claims X–Y" or "any one of Claims X, Y and Z" format.
**Check:** Validate format of multiple dependency references.
**Severity:** Warning if format is non-standard.

### D-05 Multiple-Multiple Dependency Warning (US)
**Rule (US only):** A multiple dependent claim cannot itself serve as the base for another multiple dependent claim.
**Check (US mode):** Flag Claim N that is a multiple dependent claim AND is itself referenced by another multiple dependent claim M.
**Severity:** Error (US mode) / Info (EPO/PCT mode)

### D-06 Independent Claim Starts Each Set
**Rule:** A claim set must begin with an independent claim (Claim 1 must be independent).
**Check:** Flag if Claim 1 contains a dependency reference.
**Severity:** Error

### D-07 Dependency Chain Validity
**Rule:** Every claim in a dependency chain must be traceable back to an independent claim.
**Check:** Detect circular dependencies or broken chains.
**Severity:** Error

### D-08 Claim Count Cost Warning
**Rule (informational):** Many patent offices charge extra fees for claims beyond 20 (US) or 10 (EPO). Multiple dependent claims may incur additional fees (US).
**Check:** Count total claims and multiple dependent claims; warn at threshold.
**Severity:** Info

---

## CATEGORY 5: ELEMENT TYPE RULES

### E-01 Structural Elements Must Be Named
**Rule:** Structural elements should be defined by what they are (noun phrase).
**Check:** Flag structural claims where elements are defined only by result (no structural noun).
**Severity:** Warning

### E-02 Functional Elements Must Have Spec Support Reminder
**Rule:** Functional elements ("a heater," "configured to," "means for") must have corresponding structural description in specification.
**Check:** Log functional language for manual spec-support verification; flag as Info.
**Severity:** Info (requires manual verification)

### E-03 Activity Elements — Gerund Form
**Rule:** Process/method claim steps must use gerund form (-ing verbs): "receiving," "determining," "transmitting."
**Check:** Flag method claim steps using infinitive ("to receive," "to determine") or imperative form.
**Severity:** Warning

### E-04 Parametric Elements — Range Completeness
**Rule:** Numeric ranges must include endpoints and ideally an example midpoint.
**Check:** Flag ranges stated as "between X and Y" or "X to Y" without explicit endpoint inclusion language ("from X to Y, inclusive" or "X–Y").
**Severity:** Warning

### E-05 Intentional Elements ("for") — Scope Caution
**Rule:** "For [purpose]" language in preamble may be limiting; in body as intentional element, may narrow scope.
**Check:** Flag "for [purpose]" language in preamble; flag in body with scope note.
**Severity:** Info

---

## CATEGORY 6: SPECIAL FORMAT RULES

### SF-01 Two-Part Claim Detection
**Rule:** Two-part / improvement claim uses "characterized by," "characterized in that," or "wherein the improvement comprises" as transitional phrase.
**Check:** Detect these phrases; apply two-part claim validation (admission check).
**Note:** Presence of two-part format = explicit admission that pre-characterizing portion is prior art.
**Severity:** Info (flag for attorney review)

### SF-02 Two-Part Claim Inapplicability Warning
**Rule:** Two-part format is inappropriate for:
(a) New chemical compounds/groups
(b) Combination of equal-status known integers where inventive step lies solely in combination
(c) Complex systems with changes in several interconnected parts
**Check:** If two-part format detected, flag above scenarios for manual review.
**Severity:** Warning

### SF-03 Markush Claim Format
**Rule:** Markush groups must follow format: "selected from the group consisting of X, Y and Z"
**Check:** Flag "selected from the group comprising" (should be "consisting of" for proper Markush).
**Severity:** Error

### SF-04 Markush — Single Selection Default
**Rule:** Without "combinations thereof," Markush group = one member selected at a time.
**Check:** If multiple selections may be intended, flag if "combinations thereof" is absent.
**Severity:** Warning

### SF-05 Markush — Novelty Risk
**Rule:** If any single alternative in Markush group is in prior art, entire Markush claim lacks novelty.
**Check:** Flag this as strategic risk note whenever Markush claim detected.
**Severity:** Info

### SF-06 Means-Plus-Function Detection
**Rule:** M+F claims use "means for [function]" or "step for [function]" format.
**Check:** Detect "means for" / "step for" patterns.
**Severity:** Info (flag for spec-support verification)

### SF-07 Means-Plus-Function — Spec Support Warning
**Rule:** M+F claims require clearly defined corresponding structure in specification.
**Check:** Flag M+F claims with reminder that spec must identify the structure performing each recited function.
**Severity:** Warning

### SF-08 All-Functional Claim Warning
**Rule:** If ALL elements in a claim are defined functionally (no structural nouns), the claim risks rejection as "mere desideratum."
**Check:** Detect claims where every element uses only functional language with no structural anchor.
**Severity:** Error

### SF-09 Reference Numeral — US Mode Warning
**Rule (US mode):** Reference numerals in parentheses after claim elements risk limiting claim scope.
**Check:** Flag "(123)" or similar numeral patterns in claims when US jurisdiction selected.
**Severity:** Warning (US mode) / Info (EPO/PCT mode)

---

## JURISDICTION MODES

The Claim Checker should operate in one of these modes (user-selectable):

| Mode | Key Differences |
|---|---|
| **US (USPTO)** | No reference numerals; no multiple-multiple dependencies; extra fees for >20 claims and multiple deps; Jepson claim optional |
| **EPO** | Two-part claim preferred; multiple-multiple deps allowed; reference numerals encouraged; Art. 84 clarity enforced strictly |
| **PCT** | Reference numerals in parentheses preferred; multiple deps allowed; broadest drafting for later national phase |
| **Generic** | Apply all rules; flag jurisdiction-specific differences as Info |

---

## RULE SEVERITY LEVELS

| Level | Meaning | UI Treatment |
|---|---|---|
| **Error** | Clear violation — claim will likely be rejected or is legally defective | Red highlight |
| **Warning** | Strategic/style issue — may cause problems in prosecution or litigation | Orange highlight |
| **Info** | Informational / requires manual attorney review | Blue highlight |

---

## CLAIM SET HEALTH METRICS

The checker should also report:

1. **Total claim count** — with cost threshold warnings
2. **Independent claim count** — flag if > 3 (high cost, unity risk)
3. **Dependent claim count** — flag if proportion of fallback claims is low
4. **Multiple dependent claim count** — count and flag US mode cost
5. **Claim type distribution** — apparatus vs. method vs. CRM vs. composition
6. **Max dependency depth** — warn if chain > 5 (claim may be overly narrow)
7. **Missing claim types** — e.g., method claimed but no apparatus/CRM claim (missed coverage)

---

## CROSS-REFERENCE TO WIPO MANUAL

| Rule Code | WIPO Section | Page |
|---|---|---|
| F-01 to F-08 | Module IV §2 | pp. 42–46 |
| G-01 to G-09 | Module IV §2.2 | pp. 46–49 |
| S-01 to S-05 | Module IV §2.1 | pp. 44–46 |
| D-01 to D-08 | Module IV §3.1–3.3 | pp. 55–64 |
| E-01 to E-05 | Module IV §2, §2.5 | pp. 42–54 |
| SF-01 to SF-09 | Module IV §2.3–2.5 | pp. 49–54 |
