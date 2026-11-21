// ContractGuard review logic: read clauses, compare, check rules, decide.
// Fixed rules only, no AI model. Draft rulebook, NOT approved by OMV Legal.
(function (root) {
  "use strict";

  const RULEBOOK_VERSION = "v0.1-draft (not approved by OMV Legal)";

  // Practice text: simplified paraphrase of the public OMV General Purchase Conditions
  // (Austrian law, Nov 2025). Not the real OMV template.
  const OMV_STANDARD = [
    ["1", "Scope", "These General Purchase Conditions apply to all purchases of goods and services by OMV from the Contractor."],
    ["2", "Delivery and Acceptance", "The Contractor shall deliver on the agreed date. Risk passes to OMV only on acceptance at the agreed place of delivery."],
    ["3", "Prices and Payment", "Prices are fixed. OMV shall pay within 60 days after receipt of a proper invoice. Payment is made in OMV's regular payment run."],
    ["4", "Liability", "The Contractor is liable for all damage under the statutory provisions. The Contractor's liability is not limited or excluded."],
    ["5", "Insurance", "The Contractor shall maintain insurance cover that is proportionate to the risks of the contract and provide proof on request."],
    ["6", "Subcontracting", "The Contractor may use subcontractors only with OMV's prior written approval. For each breach the Contractor shall pay a contractual penalty of 1% of the contract price, up to a maximum of 5%."],
    ["7", "Confidentiality", "The Contractor shall keep all information received from OMV confidential without time limit. The existence of this contract is also confidential."],
    ["8", "Compliance", "The Contractor shall comply with the OMV Code of Conduct and all anti-corruption and sanctions laws, and shall keep related records for 5 years."],
    ["9", "Amendments", "Amendments must be in writing and signed by both parties. Signature via DocuSign counts as written form. Email does not satisfy written form."],
    ["10", "Governing Law and Jurisdiction", "Austrian law applies, excluding the UN Convention on Contracts for the International Sale of Goods (CISG). The courts of Vienna, Inner City, have exclusive jurisdiction."],
  ].map(([num, head, body]) => ({ num, head, body }));

  // ---------------- rulebook ----------------
  const RULES = [
    { id: "R-PAY-01", keys: ["payment", "prices"], standard: "OMV pays 60 days after a proper invoice, in the automated payment run.",
      source: "GPC Nov 2025, payment terms; Annual Report 2025",
      test(t) {
        const m = t.match(/(\d+)\s*(?:calendar\s+)?days/i);
        if (!m) return { a: "esc", why: "Couldn't find a payment period in days. A lawyer should read this clause." };
        const d = +m[1];
        if (d < 60) return { a: "route", to: "Finance / Procurement", why: `Payment period is ${d} days instead of 60. Paying ${60 - d} days earlier costs OMV working capital, so this is a money decision for Finance or Procurement, not a legal one.` };
        if (d > 60) return { a: "ok", why: `Payment period is ${d} days, longer than the standard 60. This favours OMV.` };
        return { a: "ok", why: "Still 60 days." };
      } },
    { id: "R-LIA-01", keys: ["liability"], standard: "The Contractor's liability is not limited or excluded.", source: "GPC Nov 2025, liability",
      test(t) {
        const rest = t.replace(/not limited or excluded/i, "");
        const cap = /(shall not exceed|limited to|\bcap(ped)?\b|maximum (total )?liability|in no event|exclud\w* .*liabilit|not be liable)/i.test(rest);
        if (cap || !/not limited or excluded/i.test(t)) return { a: "esc", to: "Senior lawyer", why: "The supplier limits or excludes its liability. OMV's standard has no cap, so a senior lawyer decides." };
        return { a: "ok", why: "Liability is still unlimited." };
      } },
    { id: "R-INS-01", keys: ["insurance"], standard: "Insurance proportionate to the contract's risks, proof on request.", source: "GPC Nov 2025, insurance",
      test(t) {
        if (/proportionate/i.test(t) && /proof/i.test(t)) return { a: "ok", why: "Still requires proportionate insurance with proof on request." };
        return { a: "esc", why: "The insurance duty was weakened or reworded. A lawyer should check it." };
      } },
    { id: "R-SUB-01", keys: ["subcontract"], standard: "Subcontractors only with prior written approval; penalty 1% per breach, max 5%.", source: "GPC Nov 2025, subcontracting",
      test(t) {
        if (/prior written approval/i.test(t) && /\b1\s?%/.test(t) && /\b5\s?%/.test(t)) return { a: "ok", why: "Approval duty and penalty are unchanged." };
        return { a: "counter", why: "The approval duty or the penalty was changed. Propose OMV's standard wording back." };
      } },
    { id: "R-CONF-01", keys: ["confidential"], standard: "Confidential without time limit; the contract's existence is also confidential.", source: "GPC Nov 2025, confidentiality",
      test(t) {
        const m = t.match(/(\d+)\s*(years?|months?)/i);
        if (m) return { a: "counter", why: `The supplier wants confidentiality to end after ${m[1]} ${m[2]}. OMV's standard has no end date. Propose 'without time limit'; escalate if they insist.` };
        if (/\buntil\b/i.test(t) || !/without time limit/i.test(t)) return { a: "counter", why: "The 'without time limit' wording is gone. Propose it back." };
        if (!/existence/i.test(t)) return { a: "counter", why: "The sentence making the contract's existence confidential was removed." };
        return { a: "ok", why: "Still confidential without time limit." };
      } },
    { id: "R-COMP-01", keys: ["compliance"], standard: "Code of Conduct, anti-corruption, sanctions; records kept 5 years.", source: "GPC Nov 2025, Code of Conduct / sanctions",
      test(t) {
        const y = t.match(/(\d+)\s*years/i);
        if (/code of conduct/i.test(t) && /corruption/i.test(t) && /sanction/i.test(t) && y && +y[1] >= 5) return { a: "ok", why: "All compliance duties still in place." };
        return { a: "esc", to: "Compliance", why: "A compliance duty was removed or shortened. Compliance must look at this." };
      } },
    { id: "R-AMD-01", keys: ["amendment"], standard: "Amendments only in signed writing (DocuSign counts). Email does not count.", source: "GPC Nov 2025, amendments",
      test(t) {
        const emailOk = /e-?mail/i.test(t) && !/e-?mail does not|not .*e-?mail|e-?mail .*(not|insufficient)/i.test(t);
        if (emailOk) return { a: "counter", why: "The supplier wants email to count as a written amendment. Under OMV's standard only signed writing counts. Propose the standard wording back." };
        if (!/signed/i.test(t)) return { a: "counter", why: "The signature requirement was removed." };
        return { a: "ok", why: "Signed writing is still required." };
      } },
    { id: "R-LAW-01", keys: ["governing law", "jurisdiction"], standard: "Austrian law, CISG excluded, courts of Vienna Inner City.", source: "GPC Nov 2025, governing law",
      test(t) {
        if (/austrian law/i.test(t) && /cisg/i.test(t) && /vienna/i.test(t)) return { a: "ok", why: "Still Austrian law, CISG excluded, Vienna courts." };
        return { a: "esc", to: "Senior lawyer", why: "Governing law or court was changed. This needs a senior lawyer." };
      } },
  ];
  const LABEL = { ok: "Accept", route: "Route", counter: "Counter-propose", esc: "Escalate" };
  const ORDER = { esc: 0, counter: 1, route: 2, ok: 3 };
  const findRule = (head) => { const h = head.toLowerCase(); return RULES.find((r) => r.keys.some((k) => h.includes(k))); };

  // ---------------- 1. read ----------------
  // paras: [{text, isHeading}] in document order. Returns clauses with paragraph indexes.
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
        cur = { num: m ? m[1] : String(auto), head: m ? m[2] : text, headIdx: i, bodyIdx: [], body: "" };
        out.push(cur);
      } else if (cur && text) {
        cur.bodyIdx.push(i);
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
  // Word-level diff: returns {removed:[...], added:[...]} as phrases.
  function wordChanges(a, b) {
    const A = a.split(/\s+/).filter(Boolean), B = b.split(/\s+/).filter(Boolean), n = A.length, m = B.length;
    const L = Array.from({ length: n + 1 }, () => new Int32Array(m + 1));
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i][j] = A[i] === B[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
    const removed = [], added = []; let i = 0, j = 0, r = [], ad = [];
    const flush = () => { if (r.length) removed.push(r.join(" ")); if (ad.length) added.push(ad.join(" ")); r = []; ad = []; };
    while (i < n && j < m) {
      if (A[i] === B[j]) { flush(); i++; j++; }
      else if (L[i + 1][j] >= L[i][j + 1]) r.push(A[i++]);
      else ad.push(B[j++]);
    }
    while (i < n) r.push(A[i++]); while (j < m) ad.push(B[j++]); flush();
    return { removed, added };
  }

  // ---------------- 3-5. check + decide ----------------
  function review(base, cp) {
    const findings = [], same = [];
    for (const { b, c } of matchClauses(base, cp)) {
      if (b && c && b.body === c.body) { same.push({ b, c }); continue; }
      const head = (c || b).head, rule = findRule(head);
      const f = { b, c, head, rule, to: null, cosmetic: false };
      if (!c) { f.a = "esc"; f.why = "The supplier deleted this whole clause. A lawyer must decide whether OMV can live without it."; }
      else if (!b) { f.a = "esc"; f.why = "The supplier added a clause that is not in OMV's standard. No rule covers it, so a lawyer reviews it."; }
      else if (norm(b.body) === norm(c.body)) { f.a = "ok"; f.cosmetic = true; f.why = "Only punctuation, capitals or spacing changed. No change in meaning."; }
      else if (!rule) { f.a = "esc"; f.why = "This clause changed, but the rulebook has no rule for it yet. ContractGuard doesn't guess, so a lawyer reviews it."; }
      else { const r = rule.test(c.body); f.a = r.a; f.to = r.to || null; f.why = r.why; }
      if (b && c) f.changes = wordChanges(b.body, c.body);
      findings.push(f);
    }
    findings.sort((x, y) => ORDER[x.a] - ORDER[y.a]);
    return { findings, same, total: base.length };
  }

  const TAG = "[ContractGuard]";
  function commentText(f) {
    const label = f.cosmetic ? "Wording only" : LABEL[f.a] + (f.to ? ` to ${f.to}` : "");
    const q = (x) => `“${x}”`;
    const lines = [`${TAG} ${label.toUpperCase()}`];
    if (!f.c) lines.push(`Clause ${f.b.num} (${f.head}) is missing from the supplier's version.`);
    if (f.changes && !f.cosmetic) {
      if (f.changes.removed.length) lines.push("Removed: " + f.changes.removed.map(q).join(" / "));
      if (f.changes.added.length) lines.push("Added: " + f.changes.added.map(q).join(" / "));
    }
    if (f.b && f.c && f.b.num !== f.c.num) lines.push(`Renumbered: OMV clause ${f.b.num} is clause ${f.c.num} here.`);
    lines.push("Why: " + f.why);
    if (f.rule && !f.cosmetic) lines.push(`Rule ${f.rule.id}. OMV standard: ${f.rule.standard} Source: ${f.rule.source}.`);
    lines.push("Decide with the buttons in the ContractGuard panel.");
    return lines.join("\n");
  }

  root.ContractGuard = { RULEBOOK_VERSION, OMV_STANDARD, RULES, LABEL, TAG, splitClauses, review, commentText };
})(typeof window !== "undefined" ? window : globalThis);
