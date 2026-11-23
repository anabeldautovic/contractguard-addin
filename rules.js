// ContractGuard review logic: read clauses, compare with the template, check the rulebook, decide.
// Fixed rules only, no AI model. Draft rulebook, NOT approved by OMV Legal.
//   template.js = the contract (what OMV and the supplier agree on)
//   rules.js    = the rulebook (how OMV Legal reacts when a supplier changes it)
(function (root) {
  "use strict";

  const RULEBOOK_VERSION = "v0.2-draft (not approved by OMV Legal)";

  // ---------------- the template ----------------
  // A baseline is what the supplier's version is compared with: an approved template, or an uploaded earlier round.
  const makeBase = (clauses) => clauses.map((c) => ({ num: c.num, head: c.head, paras: c.paras, body: c.paras.join(" ").replace(/\s+/g, " ").trim() }));
  const TEMPLATES = root.OMV_TEMPLATES.map((t) => ({ id: t.id, name: t.name, short: t.short, source: t.source, base: makeBase(t.clauses) }));
  const TEMPLATE_NAME = TEMPLATES[0].name;
  const OMV_STANDARD = TEMPLATES[0].base; // the current template; the rulebook is written for it

  // ---------------- the rulebook ----------------
  // Each rule covers one clause (found by words in its heading) and checks one point.
  // "scope" says which sentences the rule is about: if anything outside that changed too,
  // ContractGuard says so and never accepts on its own.
  // A rule answers: ok (accept), counter (reject: put OMV wording back), route / esc (escalate),
  // or same (its point is unchanged).
  const yrs = (n) => (n === 1 ? "1 year" : Number.isInteger(n) ? `${n} years` : `${Math.round(n * 12)} months`);
  const SRC = (n) => `OMV GPC Jan 2026, clause ${n}`;

  // Read a period however it's written: "sixty (60) days", "30 calendar days", "one month", "4 weeks", "twelve months".
  const WORDNUM = { one: 1, two: 2, three: 3, four: 4, five: 5, six: 6, seven: 7, eight: 8, nine: 9, ten: 10, eleven: 11, twelve: 12,
    fourteen: 14, fifteen: 15, twenty: 20, thirty: 30, "forty-five": 45, "forty five": 45, sixty: 60, ninety: 90, "one hundred twenty": 120 };
  const NUM = `(?:(\\d+)|\\b(${Object.keys(WORDNUM).sort((a, b) => b.length - a.length).join("|")})\\b)(?:\\s*\\((\\d+)\\))?`;
  const UNIT = `(?:calendar\\s+|working\\s+|business\\s+)?(days?|weeks?|months?|years?)`;
  // Finds "<number> <unit> <after>" and returns the length in `want` units ("days" or "years").
  function period(t, after, want) {
    const m = t.match(new RegExp(`${NUM}\\s*${UNIT}\\s+${after}`, "i"));
    if (!m) return null;
    const n = +(m[3] || m[1] || WORDNUM[m[2].toLowerCase()]), u = m[4].toLowerCase();
    if (want === "years") return n / (u.startsWith("year") ? 1 : u.startsWith("month") ? 12 : u.startsWith("week") ? 52 : 365);
    return n * (u.startsWith("week") ? 7 : u.startsWith("month") ? 30 : u.startsWith("year") ? 365 : 1);
  }

  // Two kinds of rules:
  //  - HARD CHECKS (kind "hard"): coded, for the few critical points with numbers or exact facts.
  //  - PLAYBOOK entries (kind "playbook"): one plain-language sentence the lawyers write and edit.
  //    Any change to that topic gets the playbook's suggestion; the sentence is shown as the reason.
  // "trigger" finds the topic by meaning in OTHER clauses too (e.g. a liability cap hidden in the
  // insurance clause). In this prototype that uses key phrases; at OMV a local model would do it.
  const CAP = /(shall not exceed|limited to|\bcap(ped)?\b|maximum (aggregate |total )?liability|in no event|not be liable|exclud\w*[^.]*(liabilit|damage)|indirect|consequential)/i;
  const playbook = (o) => ({ ...o, kind: "playbook", scope: /./, test: () => ({ a: o.suggest, to: o.to, why: `Playbook: ${o.text}` }) });

  const RULES = [
    { id: "R-PAY-01", kind: "hard", keys: ["payment"], scope: /\b(days?|weeks?|months?)\b/i, standard: "OMV pays 60 days after receiving a proper invoice.", source: SRC("6.4"),
      test(t) {
        const d = period(t, "(?:from|after)\\s+(?:the\\s+)?receipt", "days");
        if (d === null) return { a: "counter", why: "Couldn't find the 60-day payment period any more. Put OMV's wording back, or read the clause." };
        if (d < 60) return { a: "route", to: "Finance / Procurement", why: `Payment period is ${d} days instead of 60. Paying ${60 - d} days earlier costs OMV working capital, so this is a commercial decision rather than a legal one.` };
        if (d > 60) return { a: "ok", why: `Payment period is ${d} days, longer than the standard 60. This favours OMV.` };
        return { a: "same", why: "The payment period is still 60 days." };
      } },
    { id: "R-WAR-01", kind: "hard", keys: ["warranty"], scope: /warranty period|\byears?\b|\bmonths?\b/i, standard: "Warranty of 2 years for movables, 3 years for immovables.", source: SRC("7.2"),
      test(t) {
        const mov = period(t, "for\\s+movables", "years"), imm = period(t, "for\\s+immovables", "years");
        if (mov === null || imm === null) return { a: "counter", why: "The warranty periods are gone or reworded. Put OMV's periods back (2 years movables, 3 years immovables)." };
        if (mov < 2 || imm < 3) return { a: "counter", why: `The supplier shortens the warranty to ${yrs(mov)} for movables and ${yrs(imm)} for immovables. OMV's standard is 2 and 3 years. Reject to put OMV's periods back.` };
        if (mov > 2 || imm > 3) return { a: "ok", why: "The warranty is longer than OMV's standard. This favours OMV." };
        return { a: "same", why: "Warranty periods are unchanged." };
      } },
    playbook({ id: "P-INS-01", keys: ["insurance"], topic: "insurance", standard: "Liability insurance proportionate to price and risk, with proof before work starts.", source: SRC("11"), suggest: "counter",
      text: "The supplier keeps liability insurance that fits the price and the risk, and shows proof before starting. If the supplier changes this: Reject." }),
    { id: "R-LIA-01", kind: "hard", keys: ["liability"], topic: "liability", scope: /./, standard: "The Contractor is liable for all damage; no cap and no exclusions.", source: SRC("8"),
      trigger: /liabilit\w*[^.]*(shall not exceed|limited to|\bcap(ped)?\b|maximum|in no event|exclud)|not be liable|(shall not exceed|in no event)[^.]*liabilit/i, crossA: "esc", crossTo: "Senior lawyer", crossWhy: "OMV’s standard has no cap or exclusion of liability, wherever it is written.",
      test(t, omv) {
        if (CAP.test(t) && !CAP.test(omv)) return { a: "esc", to: "Senior lawyer", why: "The supplier caps or excludes its liability. OMV's standard has no cap." };
        if (!/liable for any damages/i.test(t)) return { a: "esc", to: "Senior lawyer", why: "The basic liability sentence was changed." };
        return { a: "same", why: "Liability is still unlimited." };
      } },
    playbook({ id: "P-CONF-01", keys: ["confidential"], topic: "confidentiality", standard: "Confidential with no end date; disclosing the contract needs OMV's consent; survives termination.", source: SRC("14"), suggest: "counter",
      text: "Everything OMV shares stays confidential with no end date, and the contract may only be mentioned with OMV’s consent. If the supplier changes this: Reject; escalate if they insist.",
      trigger: /confidential[^.]*\b(\d+\)?\s*(years?|months?)|until|expire)/i, crossA: "counter", crossWhy: "Confidentiality has no end date in OMV’s standard." }),
    playbook({ id: "P-COMP-01", keys: ["code of conduct", "sanctions"], topic: "Code of Conduct and sanctions", standard: "Code of Conduct and sanctions duties stay as they are.", source: SRC("16-17"), suggest: "esc", to: "Compliance",
      text: "OMV’s Code of Conduct and sanctions duties are not negotiated by the reviewing lawyer. Any change: Escalate to Compliance." }),
    { id: "R-LAW-01", kind: "hard", keys: ["applicable law"], scope: /./, standard: "Austrian law, UN sales convention (CISG) excluded.", source: SRC("19"),
      test(t) {
        if (/law of austria|austrian law/i.test(t) && /(convention on contracts for the inter-?national sales? of goods|cisg)[^.]*excluded/i.test(t)) return { a: "same", why: "Still Austrian law, CISG excluded." };
        return { a: "esc", to: "Senior lawyer", why: "The supplier changed the governing law." };
      } },
    { id: "R-JUR-01", kind: "hard", keys: ["jurisdiction"], scope: /./, standard: "Exclusive jurisdiction of the court for Vienna, Inner City.", source: SRC("20"),
      test(t) {
        if (/vienna/i.test(t) && /exclusive jurisdiction/i.test(t)) return { a: "same", why: "Still the Vienna court, exclusively." };
        return { a: "esc", to: "Senior lawyer", why: "The supplier changed the court for disputes." };
      } },
    playbook({ id: "P-AMD-01", keys: ["amendment"], topic: "amendments", standard: "Amendments only in signed writing (DocuSign counts). Plain e-mail does not count.", source: SRC("22"), suggest: "counter",
      text: "Changes to the contract only count when signed in writing (DocuSign counts); plain e-mail does not. If the supplier changes this: Reject.",
      trigger: /e-?mail[^.]*(satisf|counts?|suffic|valid|binding)/i, crossA: "counter", crossWhy: "Only signed writing counts as an amendment in OMV’s standard." }),
  ];

  // Plain-language summary of each rule, shown in the panel's Rulebook tab.
  // Playbook entries ARE their plain-language sentence; hard checks list what they check.
  const OUTCOMES = {
    "R-PAY-01": ["Fewer than 60 days → Escalate (a commercial question, not a legal one)", "More than 60 days → Accept (favours OMV)", "60-day period gone → Reject"],
    "R-WAR-01": ["Shorter than 2 years (movables) or 3 years (immovables) → Reject", "Longer → Accept (favours OMV)", "Periods gone or reworded → Reject"],
    "R-LIA-01": ["Any cap or exclusion of liability → Escalate, also when it sits in another clause", "Basic liability sentence changed → Escalate"],
    "R-LAW-01": ["Not Austrian law, or the UN sales convention no longer excluded → Escalate"],
    "R-JUR-01": ["Not the Vienna court, or no longer exclusive → Escalate"],
  };
  RULES.forEach((r) => { if (r.kind === "playbook") OUTCOMES[r.id] = [r.text]; });
  // Quick picks when a lawyer escalates (the lawyer chooses; ContractGuard doesn't suggest who).
  // Placeholder addresses: OMV's real ones would go here.
  const CONTACTS = { "Finance / Procurement": "finance-review@omv.example", "Senior lawyer": "senior-legal@omv.example", "Compliance": "compliance-review@omv.example" };
  const RULEBOOK_OWNER = { name: "Legal Ops (rulebook owner)", email: "legal-ops@omv.example" };

  // What ContractGuard suggests. Same words as the buttons in the panel.
  // "review" = no approved rule covers the change, so ContractGuard makes no suggestion.
  const LABEL = { escalate: "Escalate", reject: "Reject", review: "You decide", accept: "Accept" };
  const FROM_RULE = { ok: "accept", counter: "reject", route: "escalate", esc: "escalate" };
  const findRule = (head) => { const h = head.toLowerCase(); return RULES.find((r) => r.keys.some((k) => h.includes(k))); };

  // ---------------- 1. read ----------------
  // paras: [{text, isHeading}] in document order. Returns clauses with paragraph indexes and texts.
  const HEAD_RE = /^\s*(\d+)\.?\s+(.{1,70}?)\s*$/;
  function splitClauses(paras) {
    const out = []; let cur = null, auto = 0;
    paras.forEach((p, i) => {
      const text = (p.text || "").trim();
      const m = text.match(HEAD_RE);
      const looksLikeHeading = m && !/[.;:]$/.test(text);
      const styledHeading = p.isHeading && !p.isTitle && text && text.length < 80;
      if (looksLikeHeading || styledHeading) {
        auto++;
        cur = { num: m ? m[1] : String(auto), head: m ? m[2] : text, headIdx: i, bodyIdx: [], paras: [], body: "" };
        out.push(cur);
      } else if (cur && text) {
        cur.bodyIdx.push(i);
        cur.paras.push(text);
        cur.body += (cur.body ? " " : "") + text;
      }
    });
    out.forEach((c) => (c.body = c.body.replace(/\s+/g, " ").trim()));
    return out;
  }
  const norm = (s) => s.toLowerCase().replace(/[^a-z0-9%]+/g, " ").trim();
  function sim(a, b) { const A = new Set(norm(a).split(" ")), B = new Set(norm(b).split(" ")); let i = 0; A.forEach((w) => B.has(w) && i++); return i / Math.max(1, A.size + B.size - i); }

  // ---------------- 2. compare ----------------
  function matchClauses(base, cp) {
    const used = new Set(), pairs = [];
    for (const b of base) {
      let best = null, score = 0;
      cp.forEach((c, j) => { if (used.has(j)) return; const s = norm(c.head) === norm(b.head) ? 1 : sim(b.head, c.head) * 0.8; if (s > score) { score = s; best = j; } });
      if (best !== null && score >= 0.4) { used.add(best); pairs.push({ b, c: cp[best] }); } else pairs.push({ b, c: null });
    }
    cp.forEach((c, j) => { if (!used.has(j)) pairs.push({ b: null, c }); });
    return pairs;
  }
  // Difference between two lists (of words or paragraphs). Returns [{t: "=", "-" or "+", w: item}].
  function diffWords(a, b) { return diffSeq(a.split(/\s+/).filter(Boolean), b.split(/\s+/).filter(Boolean)); }
  function diffSeq(A, B) {
    const n = A.length, m = B.length;
    const L = Array.from({ length: n + 1 }, () => new Int32Array(m + 1));
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i][j] = A[i] === B[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
    const ops = []; let i = 0, j = 0;
    while (i < n && j < m) {
      if (A[i] === B[j]) { ops.push({ t: "=", w: A[i] }); i++; j++; }
      else if (L[i + 1][j] >= L[i][j + 1]) ops.push({ t: "-", w: A[i++] });
      else ops.push({ t: "+", w: B[j++] });
    }
    while (i < n) ops.push({ t: "-", w: A[i++] });
    while (j < m) ops.push({ t: "+", w: B[j++] });
    return ops;
  }
  // The same diff as phrases: {removed:[...], added:[...]}. "added" is used to highlight words in Word.
  function wordChanges(a, b) {
    const removed = [], added = []; let r = [], ad = [];
    const flush = () => { if (r.length) removed.push(r.join(" ")); if (ad.length) added.push(ad.join(" ")); r = []; ad = []; };
    for (const o of diffWords(a, b)) { if (o.t === "=") flush(); else (o.t === "-" ? r : ad).push(o.w); }
    flush();
    return { removed, added };
  }
  // Only the sentences that differ, so a comment shows full sentences instead of word fragments.
  const sentences = (t) => t.split(/(?<=[.!?])\s+(?=[A-Z0-9"“(])/).map((x) => x.trim()).filter(Boolean);
  // a and b are lists of paragraphs; each paragraph is split into sentences on its own.
  const allSentences = (paras) => paras.flatMap(sentences);
  function sentenceDiff(a, b) {
    const A = allSentences(a), B = allSentences(b), sa = new Set(A), sb = new Set(B);
    return { omv: A.filter((x) => !sb.has(x)), supplier: B.filter((x) => !sa.has(x)) };
  }
  function changedSentences(a, b) {
    const d = sentenceDiff(a, b);
    return { omv: d.omv.join(" "), supplier: d.supplier.join(" ") };
  }

  // A changed definition quietly changes every clause that uses the defined term.
  // Find which terms changed (e.g. "Force Majeure") and where OMV's text uses them.
  function changedTerms(paras, other) {
    const keep = new Set(other), terms = new Set();
    let cur = [];
    for (const p of paras) {
      if (/^\d+\.\d+\s+["“]/.test(p)) cur = [...p.matchAll(/["“]([^"”]+)["”]/g)].map((m) => m[1]);
      if (!keep.has(p)) cur.forEach((t) => terms.add(t));
    }
    return terms;
  }
  function definitionNote(b, c, base) {
    const terms = new Set([...changedTerms(c.paras, b.paras), ...changedTerms(b.paras, c.paras)]);
    const notes = [];
    for (const t of terms) {
      const used = base.filter((x) => x !== b && x.body.includes(t)).map((x) => x.num);
      if (used.length) notes.push(`“${t}” is used in clause${used.length > 1 ? "s" : ""} ${used.join(", ")}`);
    }
    return notes.length ? ` Careful: a changed definition also changes the clauses that use it. ${notes.join("; ")}.` : "";
  }

  // Is this document based on OMV's template at all? Share of OMV clauses found, and how similar they are.
  function templateMatch(base, cp) {
    const matched = matchClauses(base, cp).filter((p) => p.b && p.c);
    const avg = matched.length ? matched.reduce((s, p) => s + sim(p.b.body, p.c.body), 0) / matched.length : 0;
    const share = matched.length / base.length;
    return { share, avg, ok: share >= 0.6 && avg >= 0.6 };
  }
  // Which template is this supplier version based on? Scores every template, best first.
  // Ties go to the newer (first) template.
  function detectTemplate(cp) {
    return TEMPLATES.map((t, k) => ({ t, k, m: templateMatch(t.base, cp) }))
      .sort((x, y) => (y.m.share * y.m.avg) - (x.m.share * x.m.avg) || x.k - y.k);
  }

  // Does the supplier's new text touch a rule's topic, even outside that rule's own clause?
  // E.g. a liability cap written into the insurance clause. The stronger suggestion wins.
  // Prototype: key phrases ("trigger"). At OMV: a local model reads the text and picks the rule.
  const RANK = { accept: 0, review: 1, reject: 2, escalate: 3 };
  const EXCEPTION = /\b(unless|except|save for|provided that|subject to|other than|außer|ausgenommen|sofern|soweit|es sei denn|vorbehaltlich)\b/i;
  function matchByMeaning(f) {
    const added = f.b ? sentenceDiff(f.b.paras, f.c.paras).supplier : allSentences(f.c.paras);
    for (const r of RULES) {
      if (!r.trigger || r === f.rule || !added.some((s) => r.trigger.test(s))) continue;
      const a = FROM_RULE[r.crossA];
      f.matched = r;
      const note = `It also touches ${r.topic} (rule ${r.id}): ${r.crossWhy}`;
      if (RANK[a] > RANK[f.a]) { f.a = a; f.to = r.crossTo || f.to; f.why = `${note} ${f.why}`; }
      else f.why += ` ${note}`;
    }
  }

  // ---------------- 3-5. check + decide ----------------
  function review(base, cp) {
    const findings = [], same = [];
    const pairs = matchClauses(base, cp);
    for (const { b, c } of pairs) {
      if (b && c && b.body === c.body) { same.push({ b, c }); continue; }
      const head = (c || b).head, rule = findRule(head);
      const f = { b, c, head, rule, to: null, cosmetic: false };
      if (!c) { f.a = "review"; f.why = "The supplier deleted this whole clause. No rule covers a deletion, so you decide whether OMV can live without it."; }
      else if (!b) { f.a = "review"; f.why = "The supplier added a clause that is not in OMV's standard. No rule covers it, so you decide."; }
      else if (norm(b.body) === norm(c.body)) { f.a = "accept"; f.cosmetic = true; f.why = "Only punctuation, capitals or spacing changed. No change in meaning, so the suggestion is Accept. Please confirm."; }
      else if (!rule) { f.a = "review"; f.why = "This clause changed, but no approved rule covers it. ContractGuard doesn't guess, so you decide."; }
      else {
        const r = rule.test(c.body, b.body);
        // Did anything change outside what this rule looks at?
        const d = sentenceDiff(b.paras, c.paras);
        const extra = d.omv.concat(d.supplier).some((x) => !rule.scope.test(x));
        f.to = r.to || null;
        if (r.a === "same") {
          f.a = "review";
          f.why = `Rule ${rule.id} checks one point and it is unchanged (${r.why.replace(/\.$/, "")}). But other wording in this clause changed, and no rule covers that, so you decide.`;
        } else {
          f.a = r.a === "esc" && !f.to ? "review" : FROM_RULE[r.a];
          f.why = r.why;
          if (extra) {
            if (f.a === "accept") f.a = "review";
            f.why += " Other wording in this clause changed too, which no rule covers. Check “Show the change”.";
          }
        }
      }
      if (b && c) f.changes = wordChanges(b.body, c.body);
      if (b && c && /definition/i.test(head)) f.why += definitionNote(b, c, base);
      if (c && !f.cosmetic) matchByMeaning(f);
      // Never suggest accepting new text that contains an exception ("unless", "except", "außer"...).
      if (f.a === "accept" && c && !f.cosmetic) {
        const added = b ? sentenceDiff(b.paras, c.paras).supplier : allSentences(c.paras);
        const ex = added.join(" ").match(EXCEPTION);
        if (ex) { f.a = "review"; f.why += ` But the new text contains an exception (“${ex[0]}”), so it isn’t suggested for acceptance. You decide.`; }
      }
      findings.push(f);
    }
    // Clause order, as in the document. A deleted clause sits where it was in OMV's standard.
    const pos = (f) => {
      if (f.c) return cp.indexOf(f.c);
      const k = base.indexOf(f.b);
      for (const s of base.slice(k + 1)) { const p = pairs.find((x) => x.b === s && x.c); if (p) return cp.indexOf(p.c) - 0.5; }
      return cp.length;
    };
    findings.sort((x, y) => pos(x) - pos(y));
    return { findings, same, total: base.length };
  }

  const TAG = "ContractGuard ·";
  const isOurs = (text) => /^(ContractGuard ·|\[ContractGuard\])/.test(text || "");
  const suggestion = (f) => LABEL[f.a];
  // A short quote around the changed words, with a few words of context: "…the Agreement sixty (60) days from receipt…".
  function snippet(a, b) {
    const ops = diffWords(a, b), C = 5;
    let first = -1, last = -1;
    ops.forEach((o, k) => { if (o.t !== "=") { if (first < 0) first = k; last = k; } });
    if (first < 0) return { omv: a, sup: b };
    const from = Math.max(0, first - C), to = Math.min(ops.length, last + C + 1);
    const side = (t) => (from > 0 ? "…" : "") + ops.slice(from, to).filter((o) => o.t === "=" || o.t === t).map((o) => o.w).join(" ") + (to < ops.length ? "…" : "");
    return { omv: side("-"), sup: side("+") };
  }
  const cut = (t, n = 320) => (t.length > n ? t.slice(0, n).replace(/\s+\S*$/, "") + "…" : t);
  const ICON = { accept: "✔", reject: "✖", escalate: "⚑", review: "?" };

  // The comment text. Word comments from add-ins can't be reliably formatted,
  // so the layout uses symbols, short quotes and blank lines instead of bold or italic.
  function commentText(f) {
    const q = (t) => `“${cut(t)}”`;
    const head = f.cosmetic ? "✔ Suggest: Accept (wording only)" : f.a === "review" ? "? No rule: you decide" : `${ICON[f.a]} Suggest: ${suggestion(f)}`;
    const L = [`${TAG} ${head}`, "", f.why];
    if (!f.cosmetic) {
      let omv, sup;
      if (f.b && f.c) {
        const d = sentenceDiff(f.b.paras, f.c.paras);
        if (d.omv.length && d.omv.length === d.supplier.length) {
          const s = d.omv.map((x, k) => snippet(x, d.supplier[k]));
          omv = s.map((x) => q(x.omv)).join(" / "); sup = s.map((x) => q(x.sup)).join(" / ");
        } else {
          omv = d.omv.length ? d.omv.map(q).join(" ") : "(nothing: the supplier added this)";
          sup = d.supplier.length ? d.supplier.map(q).join(" ") : "(deleted by the supplier)";
        }
      } else if (!f.c) { omv = q(f.b.body); sup = "(whole clause deleted)"; }
      else { omv = "(no such clause)"; sup = q(f.c.body); }
      L.push("", `${f.label || "OMV"} ▸ ${omv}`, `Supplier ▸ ${sup}`);
    }
    if (f.b && f.c && f.b.num !== f.c.num) L.push("", `Note: OMV clause ${f.b.num} is clause ${f.c.num} here.`);
    const kind = (r) => (r.kind === "hard" ? "hard check" : "playbook");
    if (f.rule && !f.cosmetic && f.a !== "review") L.push("", `Rule ${f.rule.id} (${kind(f.rule)}) · ${f.rule.source}`);
    if (f.matched) L.push(`Matched to rule ${f.matched.id}.`);
    return L.join("\n");
  }

  root.ContractGuard = { templateMatch, detectTemplate, TEMPLATES, makeBase, RULEBOOK_VERSION, TEMPLATE_NAME, OMV_STANDARD, diffSeq, RULES, OUTCOMES, CONTACTS, RULEBOOK_OWNER, findRule, sentenceDiff, LABEL, TAG, isOurs, suggestion, splitClauses, review, commentText, diffWords };
})(typeof window !== "undefined" ? window : globalThis);
