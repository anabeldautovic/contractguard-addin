// ContractGuard review logic: read clauses, compare with the template, check the rulebook, decide.
// Fixed rules only, no AI model. Draft rulebook, NOT approved by OMV Legal.
//   template.js = the contract (what OMV and the supplier agree on)
//   rules.js    = the rulebook (how OMV Legal reacts when a supplier changes it)
(function (root) {
  "use strict";

  const RULEBOOK_VERSION = "v0.2-draft (not approved by OMV Legal)";

  // ---------------- the template ----------------
  const T = root.OMV_TEMPLATE;
  const TEMPLATE_NAME = T.name;
  const OMV_STANDARD = T.clauses.map((c) => ({ num: c.num, head: c.head, paras: c.paras, body: c.paras.join(" ").replace(/\s+/g, " ").trim() }));

  // ---------------- the rulebook ----------------
  // Each rule covers one clause (found by words in its heading) and checks one point.
  // "scope" says which sentences the rule is about: if anything outside that changed too,
  // ContractGuard says so and never accepts on its own.
  // A rule answers: ok (accept), counter (reject: put OMV wording back), route / esc (escalate),
  // or same (its point is unchanged).
  const num = (t, re) => { const m = t.match(re); return m ? +m[1] : null; };
  const yrs = (n) => (n === 1 ? "1 year" : Number.isInteger(n) ? `${n} years` : `${Math.round(n * 12)} months`);
  const SRC = (n) => `OMV GPC Jan 2026, clause ${n}`;
  const RULES = [
    { id: "R-PAY-01", keys: ["payment"], scope: /\bdays?\b/i, standard: "OMV pays 60 days after receiving a proper invoice.", source: SRC("6.4"),
      test(t) {
        const d = num(t, /(\d+)\)?\s*(?:calendar\s+)?days\s+(?:from|after)\s+(?:the\s+)?receipt/i);
        if (d === null) return { a: "counter", why: "Couldn't find the 60-day payment period any more. Put OMV's wording back, or read the clause." };
        if (d < 60) return { a: "route", to: "Finance / Procurement", why: `Payment period is ${d} days instead of 60. Paying ${60 - d} days earlier costs OMV working capital, so this is a money decision for Finance or Procurement, not a legal one.` };
        if (d > 60) return { a: "ok", why: `Payment period is ${d} days, longer than the standard 60. This favours OMV.` };
        return { a: "same", why: "The payment period is still 60 days." };
      } },
    { id: "R-WAR-01", keys: ["warranty"], scope: /warranty period|\byears?\b|\bmonths?\b/i, standard: "Warranty of 2 years for movables, 3 years for immovables.", source: SRC("7.2"),
      test(t) {
        const toY = (m) => (m ? (/month/i.test(m[2]) ? +m[1] / 12 : +m[1]) : null);
        const mov = toY(t.match(/(\d+)\)?\s*(years?|months?)\s+for movables/i)), imm = toY(t.match(/(\d+)\)?\s*(years?|months?)\s+for immovables/i));
        if (mov === null || imm === null) return { a: "counter", why: "The warranty periods are gone or reworded. Put OMV's periods back (2 years movables, 3 years immovables)." };
        if (mov < 2 || imm < 3) return { a: "counter", why: `The supplier shortens the warranty to ${yrs(mov)} for movables and ${yrs(imm)} for immovables. OMV's standard is 2 and 3 years. Reject to put OMV's periods back.` };
        if (mov > 2 || imm > 3) return { a: "ok", why: "The warranty is longer than OMV's standard. This favours OMV." };
        return { a: "same", why: "Warranty periods are unchanged." };
      } },
    { id: "R-INS-01", keys: ["insurance"], scope: /./, standard: "Liability insurance proportionate to price and risk, with proof before work starts.", source: SRC("11"),
      test(t) {
        if (/proportionate/i.test(t) && /proofs? of (such )?insurance/i.test(t)) return { a: "same", why: "Still requires proportionate insurance with proof." };
        return { a: "counter", why: "The insurance duty was weakened. Reject to put OMV's wording back." };
      } },
    { id: "R-LIA-01", keys: ["liability"], scope: /./, standard: "The Contractor is liable for all damage; no cap and no exclusions.", source: SRC("8"),
      test(t, omv) {
        const CAP = /(shall not exceed|limited to|\bcap(ped)?\b|maximum (aggregate |total )?liability|in no event|not be liable|exclud\w*[^.]*(liabilit|damage)|indirect|consequential)/i;
        if (CAP.test(t) && !CAP.test(omv)) return { a: "esc", to: "Senior lawyer", why: "The supplier caps or excludes its liability. OMV's standard has no cap, so a senior lawyer decides." };
        if (!/liable for any damages/i.test(t)) return { a: "esc", to: "Senior lawyer", why: "The basic liability sentence was changed. A senior lawyer decides." };
        return { a: "same", why: "Liability is still unlimited." };
      } },
    { id: "R-CONF-01", keys: ["confidential"], scope: /./, standard: "Confidential with no end date; disclosing the contract needs OMV's consent; survives termination.", source: SRC("14"),
      test(t, omv) {
        const P = /(\d+)\)?\s*(years?|months?)/i, m = t.match(P);
        if (m && !P.test(omv)) return { a: "counter", why: `The supplier wants confidentiality to end after ${m[1]} ${m[2]}. OMV's standard has no end date. Reject to put OMV's wording back; escalate if they insist.` };
        if (!/shall survive/i.test(t)) return { a: "counter", why: "Confidentiality would no longer survive the end of the contract. Reject to put OMV's wording back." };
        if (!/existence, content or progress/i.test(t)) return { a: "counter", why: "The supplier could now talk about the contract without OMV's consent. Reject to put OMV's wording back." };
        return { a: "same", why: "Confidentiality still has no end date." };
      } },
    { id: "R-COMP-01", keys: ["code of conduct", "sanctions"], scope: /./, standard: "Code of Conduct and sanctions duties: every change goes to Compliance.", source: SRC("16-17"),
      test() { return { a: "esc", to: "Compliance", why: "The supplier changed a Code of Conduct or sanctions duty. Compliance must look at every change here." }; } },
    { id: "R-LAW-01", keys: ["applicable law"], scope: /./, standard: "Austrian law, UN sales convention (CISG) excluded.", source: SRC("19"),
      test(t) {
        if (/law of austria|austrian law/i.test(t) && /(convention on contracts for the inter-?national sales? of goods|cisg)[^.]*excluded/i.test(t)) return { a: "same", why: "Still Austrian law, CISG excluded." };
        return { a: "esc", to: "Senior lawyer", why: "The supplier changed the governing law. This needs a senior lawyer." };
      } },
    { id: "R-JUR-01", keys: ["jurisdiction"], scope: /./, standard: "Exclusive jurisdiction of the court for Vienna, Inner City.", source: SRC("20"),
      test(t) {
        if (/vienna/i.test(t) && /exclusive jurisdiction/i.test(t)) return { a: "same", why: "Still the Vienna court, exclusively." };
        return { a: "esc", to: "Senior lawyer", why: "The supplier changed the court for disputes. This needs a senior lawyer." };
      } },
    { id: "R-AMD-01", keys: ["amendment"], scope: /./, standard: "Amendments only in signed writing (DocuSign counts). Plain e-mail does not count.", source: SRC("22"),
      test(t) {
        const emailOk = /e-?mail/i.test(t) && !/e-?mail[^.]*(shall not|does not|do not|not satisfy|insufficient)/i.test(t);
        if (emailOk) return { a: "counter", why: "The supplier wants plain e-mail to count as a written amendment. Under OMV's standard only signed writing counts. Reject to put OMV's wording back." };
        if (!/signed/i.test(t)) return { a: "counter", why: "The signature requirement was removed. Reject to put OMV's wording back." };
        return { a: "same", why: "Signed writing is still required." };
      } },
  ];

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
      else if (norm(b.body) === norm(c.body)) { f.a = "accept"; f.cosmetic = true; f.why = "Only punctuation, capitals or spacing changed. No change in meaning, so ContractGuard accepted it."; }
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
    const head = f.cosmetic ? "✔ Accepted (wording only)" : f.a === "review" ? "? No rule: you decide" : `${ICON[f.a]} Suggest: ${suggestion(f)}`;
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
      L.push("", `OMV ▸ ${omv}`, `Supplier ▸ ${sup}`);
    }
    if (f.b && f.c && f.b.num !== f.c.num) L.push("", `Note: OMV clause ${f.b.num} is clause ${f.c.num} here.`);
    if (f.rule && !f.cosmetic && f.a !== "review") L.push("", `Rule ${f.rule.id} · ${f.rule.source}`);
    return L.join("\n");
  }

  root.ContractGuard = { RULEBOOK_VERSION, TEMPLATE_NAME, OMV_STANDARD, diffSeq, RULES, LABEL, TAG, isOurs, suggestion, splitClauses, review, commentText, diffWords };
})(typeof window !== "undefined" ? window : globalThis);
