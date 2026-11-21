/* ContractGuard Word add-in: the side panel. Uses Office.js to read the open document and add comments. */
(function () {
  "use strict";
  const CG = window.ContractGuard;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;");
  // Colour for each kind of suggestion (CSS classes).
  const PILL = { escalate: "esc", reject: "counter", review: "route", accept: "ok" };
  const norm = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9%]+/g, " ").trim();

  // The lawyer's three choices (to edit, the lawyer just types in the document).
  const ACTIONS = [
    { k: "accept", label: "Accept" },
    { k: "reject", label: "Reject" },
    { k: "escalate", label: "Escalate" },
  ];
  const ACT_LABEL = { accept: "Accept", reject: "Reject", escalate: "Escalate" };

  let review = null;     // last review result
  let decisions = [];    // decisions[i] = {k, text, replyId, auto, note} for finding i, or null
  let pending = null;    // {i, k}: lawyer chose something other than the suggestion; waiting for "why"
  let proposing = null;  // index of the finding whose "propose a rule" form is open
  let log = [];          // decision log, saved inside the document
  let tab = "changes";
  let currentRule = null; // rule id highlighted in the Rulebook tab
  let lastCp = null;     // clauses from the last read, to know where the cursor is
  let busy = false;

  function status(msg, isErr) { const s = $("status"); s.textContent = msg; s.className = isErr ? "err" : ""; }

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Word) { status("Open this in Word.", true); return; }
    if (!Office.context.requirements.isSetSupported("WordApi", "1.4")) {
      status("This version of Word is too old for comments from add-ins. Try Word in the browser (office.com).", true); return;
    }
    $("run").disabled = false; $("clear").disabled = false;
    $("run").onclick = onRun;
    $("clear").onclick = onClear;
    document.querySelectorAll(".tab").forEach((b) => (b.onclick = () => showTab(b.dataset.tab)));
    loadLog();
    renderRules();
    renderLog();
    status("Open a supplier's contract, then press Run review.");
    // Follow the cursor: when it moves into a clause, show that change and its rule.
    try { Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelection); } catch (e) { console.warn(e); }
    Word.run(async (ctx) => { await readClauses(ctx); }).catch(() => {});
  });

  async function guarded(fn) {
    if (busy) return;
    busy = true;
    const btns = () => document.querySelectorAll("#run, #clear, #result button");
    btns().forEach((b) => (b.disabled = true));
    try { await fn(); }
    catch (e) { console.error(e); status("Something went wrong: " + (e.message || e), true); }
    finally {
      busy = false;
      btns().forEach((b) => (b.disabled = false));
    }
  }

  // ---------------- tabs ----------------
  function showTab(t) {
    tab = t;
    document.querySelectorAll(".tab").forEach((b) => b.classList.toggle("on", b.dataset.tab === t));
    $("result").hidden = t !== "changes"; $("rules").hidden = t !== "rules"; $("log").hidden = t !== "log";
    if (t === "rules") scrollToRule();
  }
  function openRule(id) { currentRule = id; renderRules(); showTab("rules"); }
  function scrollToRule() {
    const el = currentRule && document.querySelector(`.rule[data-id="${currentRule}"]`);
    if (el) el.scrollIntoView({ block: "nearest", behavior: "smooth" });
  }

  // ---------------- the decision log (saved in the document) ----------------
  const settings = () => Office.context.document.settings;
  function loadLog() { try { log = settings().get("contractguard.log") || []; } catch (e) { log = []; } }
  function saveLog() { try { settings().set("contractguard.log", log); settings().saveAsync(() => {}); } catch (e) { console.warn(e); } }
  function addLog(f, decision, note) {
    log.push({ t: new Date().toISOString(), clause: `${(f.c || f.b).num}. ${f.head}`, suggested: f.a === "review" ? "(no rule)" : CG.LABEL[f.a], decision, note: note || "" });
    saveLog(); renderLog();
  }
  // Events that aren't about one change (review run, comments removed).
  function addEvent(what, note) {
    log.push({ t: new Date().toISOString(), clause: "(whole document)", suggested: "", decision: what, note: note || "" });
    saveLog(); renderLog();
  }
  function renderLog() {
    if (!log.length) { $("log").innerHTML = `<p class="muted">No decisions yet. Every Accept, Reject, Escalate, Go back and rule proposal is recorded here and saved inside this document.</p>`; return; }
    const time = (t) => { const d = new Date(t); return d.toLocaleDateString() + " " + d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" }); };
    let h = `<p class="muted">Saved inside this document. Your name is on each reply in the Word comments.</p><table class="logt"><tr><th>When</th><th>Clause</th><th>Suggested</th><th>Decision</th></tr>`;
    [...log].reverse().forEach((e) => {
      const over = e.suggested !== "(no rule)" && ACT_LABEL[e.decision.toLowerCase()] && e.decision !== e.suggested;
      h += `<tr><td>${esc(time(e.t))}</td><td>${esc(e.clause)}</td><td>${esc(e.suggested)}</td><td>${over ? "<b>" : ""}${esc(e.decision)}${over ? "</b>" : ""}${e.note ? `<div class="note">${esc(e.note)}</div>` : ""}</td></tr>`;
    });
    const cell = (x) => `"${String(x).replace(/"/g, '""')}"`;
    const csv = [["Time", "Clause", "Suggested", "Decision", "Note"], ...log.map((e) => [e.t, e.clause, e.suggested, e.decision, e.note])].map((r) => r.map(cell).join(",")).join("\r\n");
    h += `</table><a class="mini" download="contractguard-log.csv" href="data:text/csv;charset=utf-8,${encodeURIComponent(csv)}">Download log (CSV)</a>`;
    $("log").innerHTML = h;
  }

  // ---------------- the rulebook tab ----------------
  function renderRules() {
    let h = `<div class="rbhead"><b>Rulebook ${esc(CG.RULEBOOK_VERSION)}</b><br>Owner: ${esc(CG.RULEBOOK_OWNER.name)}. Lawyers can't change rules from Word: they propose a rule, Legal approves it, and only then ContractGuard uses it.</div>`;
    for (const r of CG.RULES) {
      const used = review ? review.findings.map((f, i) => ({ f, i })).filter((x) => x.f.rule === r) : [];
      h += `<div class="rule${r.id === currentRule ? " cur" : ""}" data-id="${r.id}"><div class="rid">${esc(r.id)} <span>· ${esc(r.source.replace("OMV GPC Jan 2026, ", ""))}</span></div>`;
      h += `<div class="rstd"><b>OMV position:</b> ${esc(r.standard)}</div><ul>` + CG.OUTCOMES[r.id].map((o) => `<li>${esc(o)}</li>`).join("") + `</ul>`;
      if (used.length) h += `<div class="used">In this contract: ` + used.map((x) => `<a href="#" class="tochange" data-i="${x.i}">${esc((x.f.c || x.f.b).num)}. ${esc(x.f.head)} → ${esc(CG.suggestion(x.f))}</a>`).join(", ") + `</div>`;
      h += `</div>`;
    }
    h += `<div class="rule"><div class="rid">Every other change</div><ul><li>No rule covers it → You decide. ContractGuard never guesses.</li><li>A rule's point is unchanged but other wording changed → You decide.</li><li>Only punctuation, capitals or spacing changed → Accepted automatically.</li></ul></div>`;
    const body = `Proposal for a new ContractGuard rule\n\nClause:\nWhen the supplier changes:\nContractGuard should suggest (Accept / Reject / Escalate to whom):\nBecause:\n\nProposed by:`;
    h += `<a class="mini" href="${mailto(CG.RULEBOOK_OWNER.email, "ContractGuard: new rule proposal", body)}" target="_blank" rel="noopener">✎ Propose a new rule</a>`;
    $("rules").innerHTML = h;
    $("rules").querySelectorAll(".tochange").forEach((a) => (a.onclick = (e) => { e.preventDefault(); showTab("changes"); focusCard(+a.dataset.i, true); }));
  }

  // ---------------- e-mails ----------------
  const mailto = (to, subject, body) => `mailto:${to}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  function docName() {
    try { const u = Office.context.document.url || ""; return decodeURIComponent(u.split(/[\\/]/).pop().split("?")[0]) || "the supplier's contract"; }
    catch (e) { return "the supplier's contract"; }
  }
  function escalationMail(f, note, to) {
    const who = to || "";
    const subject = `ContractGuard escalation: clause ${(f.c || f.b).num} ${f.head} (${docName()})`;
    const body = `Hello,\n\nplease decide on this change in ${docName()}:\n\n` + CG.commentText(f).split("\n").slice(2).join("\n") +
      (note ? `\n\nNote from the reviewing lawyer: ${note}` : "") + `\n\nThe Word comment on this clause stays open until it's decided.\n`;
    return { who, href: mailto(CG.CONTACTS[who] || "", subject, body) };
  }
  function proposalMail(f, d, when, should, because) {
    const subject = `ContractGuard: rule proposal for clause ${(f.c || f.b).num} ${f.head}`;
    const body = `Proposal for the ContractGuard rulebook\n\nClause: ${(f.c || f.b).num}. ${f.head}\nWhen the supplier: ${when}\nContractGuard should suggest: ${should}\nBecause: ${because}\n\n` +
      `Current rule: ${f.rule ? f.rule.id : "none (ContractGuard said: you decide)"}\nContractGuard suggested: ${f.a === "review" ? "nothing" : CG.LABEL[f.a]}. The lawyer decided: ${ACT_LABEL[d.k]}.\nContract: ${docName()}\n`;
    return mailto(CG.RULEBOOK_OWNER.email, subject, body);
  }

  // ---------------- reading the document ----------------
  async function deleteOurComments(ctx) {
    const cs = ctx.document.body.getComments();
    cs.load("items/content");
    await ctx.sync();
    let n = 0;
    cs.items.forEach((c) => { if (CG.isOurs(c.content)) { c.delete(); n++; } });
    await ctx.sync();
    return n;
  }

  // ---------------- ask before anything that throws away work ----------------
  function ask(msg, yesLabel, onYes) {
    const box = $("confirm");
    box.innerHTML = `<div class="ft">${msg}</div><div class="fbtns"><button class="act danger yes">${yesLabel}</button><button class="undo no">Cancel</button></div>`;
    box.hidden = false;
    box.querySelector(".yes").onclick = () => { box.hidden = true; onYes(); };
    box.querySelector(".no").onclick = () => { box.hidden = true; };
    box.scrollIntoView({ block: "nearest", behavior: "smooth" });
    box.querySelector(".no").focus();
  }
  function onRun() {
    const made = decisions.filter((d) => d && !d.auto).length;
    if (review && made) ask(`Run the review again? This replaces ContractGuard's comments, including your ${made} decision repl${made === 1 ? "y" : "ies"}. Your decisions stay in the Log.`, "Run again", () => guarded(runReview));
    else guarded(runReview);
  }
  function onClear() {
    guarded(async () => {
      let n = 0;
      await Word.run(async (ctx) => {
        const cs = ctx.document.body.getComments(); cs.load("items/content"); await ctx.sync();
        n = cs.items.filter((c) => CG.isOurs(c.content)).length;
      });
      if (!n) { status("No ContractGuard comments in this document."); return; }
      ask(`Remove all ${n} ContractGuard comment${n === 1 ? "" : "s"} from this document? Your decision replies in them are removed too. The Log stays.`, "Remove", () => guarded(clearComments));
    });
  }

  async function clearComments() {
    await Word.run(async (ctx) => {
      const n = await deleteOurComments(ctx);
      if (n) addEvent("Comments removed", `${n} ContractGuard comment${n === 1 ? "" : "s"}`);
      status(n ? `Removed ${n} ContractGuard comment${n > 1 ? "s" : ""}. Yellow highlights stay; use Undo or remove them by hand.` : "No ContractGuard comments found.");
      $("result").innerHTML = `<p class="muted">Press Run review to see the supplier's changes here.</p>`;
      review = null; decisions = []; pending = null; proposing = null;
      renderRules();
    });
  }

  // Read the document fresh and split it into clauses (paragraph numbers change after edits).
  async function readClauses(ctx) {
    const paras = ctx.document.body.paragraphs;
    paras.load("items/text,items/styleBuiltIn");
    await ctx.sync();
    const list = paras.items.map((p) => ({
      text: p.text,
      isHeading: /^Heading\d/.test(p.styleBuiltIn || ""),
      isTitle: p.styleBuiltIn === "Title",
    }));
    const cp = CG.splitClauses(list);
    lastCp = cp;
    return { paras, cp };
  }
  const locate = (cp, ref) => cp.find((c) => norm(c.head) === norm(ref.head)) || cp.find((c) => c.num === ref.num) || null;
  function clauseRange(paras, c) {
    let r = paras.items[c.headIdx].getRange("Whole");
    if (c.bodyIdx.length) r = r.expandTo(paras.items[c.bodyIdx[c.bodyIdx.length - 1]].getRange("Whole"));
    return r;
  }

  // Paragraphs of the supplier's clause that differ from OMV's text.
  function changedParaIdx(c, f) {
    const omv = new Set(f.b ? f.b.paras : []);
    return c.bodyIdx.filter((_, k) => !omv.has(c.paras[k]));
  }

  // Where a comment goes: only the changed text, not the whole clause.
  //  - the supplier's added words (from the first to the last one found), or
  //  - the changed paragraph (when words were only removed), or
  //  - the paragraph just before a deleted paragraph, or
  //  - for a deleted clause: the next clause heading; for a new clause: the whole new clause.
  function anchorRange(paras, cp, f, c, res) {
    if (!f.b && c) return clauseRange(paras, c);
    if (!c) {
      const next = CG.OMV_STANDARD.slice(CG.OMV_STANDARD.indexOf(f.b) + 1).map((x) => locate(cp, x)).find(Boolean);
      return (next ? paras.items[next.headIdx] : paras.items[0]).getRange("Whole");
    }
    const hits = res.filter((r) => r.items.length).map((r) => r.items[0]);
    if (hits.length) return hits.length > 1 ? hits[0].expandTo(hits[hits.length - 1]) : hits[0];
    const idx = changedParaIdx(c, f);
    if (idx.length) return paras.items[idx[0]].getRange("Whole");
    let before = c.headIdx, j = 0;
    for (const o of CG.diffSeq(f.b.paras, c.paras)) { if (o.t === "-") break; before = c.bodyIdx[j++]; }
    return paras.items[before].getRange("Whole");
  }

  // Search for the words the supplier added (only in changed paragraphs), to highlight them and anchor the comment.
  function findAdded(paras, c, f) {
    const out = [];
    if (!f.changes || f.cosmetic) return out;
    const changed = changedParaIdx(c, f);
    for (const phrase of f.changes.added) {
      const p = phrase.replace(/[.,;:]+$/, "");
      if (p.length < 2 || p.length > 250) continue;
      for (const idx of changed) {
        const res = paras.items[idx].search(p, { matchCase: true });
        res.load("items");
        out.push(res);
      }
    }
    return out;
  }

  // ---------------- run the review ----------------
  async function runReview() {
    status("Reading the document…");
    await Word.run(async (ctx) => {
      await deleteOurComments(ctx);
      const { paras, cp } = await readClauses(ctx);
      if (cp.length < 2) {
        status("I couldn't find numbered clauses (like \"3. Prices and Payment\") in this document, so I can't review it. Please review it by hand.", true);
        return;
      }

      status("Comparing with the OMV standard…");
      const R = CG.review(CG.OMV_STANDARD, cp);

      // First find the supplier's added words, then put each comment on just the changed text.
      const found = R.findings.map((f) => (f.c ? findAdded(paras, f.c, f) : []));
      await ctx.sync();
      const comments = R.findings.map((f, i) => {
        const cm = anchorRange(paras, cp, f, f.c, found[i]).insertComment(CG.commentText(f));
        if (f.cosmetic) cm.resolved = true; // wording-only changes are accepted automatically
        cm.load("id");
        return cm;
      });
      found.flat().forEach((res) => { if (res.items.length) res.items[0].font.highlightColor = "Yellow"; });
      await ctx.sync();
      R.findings.forEach((f, i) => (f.commentId = comments[i].id));

      review = R; pending = null; proposing = null;
      decisions = R.findings.map((f) => f.cosmetic ? { k: "accept", auto: true, text: "Accepted automatically." } : null);
      addEvent("Review run", `${R.findings.length} change${R.findings.length === 1 ? "" : "s"}, ${R.same.length} of ${R.total} clauses identical`);
      R.findings.forEach((f) => { if (f.cosmetic) addLog(f, "Accept", "Automatic: wording only"); });
      showResult(); renderRules(); showTab("changes");
      const open = decisions.filter((d) => !d).length;
      status(R.findings.length ? `Done. ${open} change${open === 1 ? "" : "s"} for you to decide.` : "Done. No changes from the OMV standard.");
    });
  }

  // ---------------- the three decisions ----------------
  async function getComment(ctx, f) {
    const cs = ctx.document.body.getComments();
    cs.load("items/id");
    await ctx.sync();
    return cs.items.find((c) => c.id === f.commentId) || null;
  }

  // A lawyer clicks a button. Escalate asks "to whom?"; choosing something other than the suggestion asks "why?".
  function choose(i, k) {
    const f = review.findings[i];
    const rec = f.a === "review" ? null : f.a;
    if (k === "escalate" || (rec && k !== rec)) {
      pending = { i, k }; showResult();
      const t = document.querySelector(".why input, .why textarea"); if (t) t.focus();
      return;
    }
    guarded(() => decide(i, k, "", ""));
  }

  async function decide(i, k, note, to) {
    const f = review.findings[i];
    const rec = f.a === "review" ? null : f.a;
    const over = rec && k !== rec ? ` (suggested: ${CG.LABEL[rec]})` : "";
    const why = note ? ` Reason: ${note}` : "";
    await Word.run(async (ctx) => {
      const cm = await getComment(ctx, f);
      const { paras, cp } = await readClauses(ctx);
      const c = f.c ? locate(cp, f.c) : null;
      let text, reply = null;

      if (k === "accept") {
        text = f.c ? "Accepted. Supplier's wording stays." : "Accepted. The clause stays deleted.";
        if (cm) { reply = cm.reply(`✔ ${text}${over}${why}`); cm.resolved = true; }
        if (c) clauseRange(paras, c).font.highlightColor = null;
      }

      else if (k === "reject") {
        // Put OMV's wording back as tracked changes, then set Track Changes back to how the lawyer had it.
        ctx.document.load("changeTrackingMode");
        await ctx.sync();
        const before = ctx.document.changeTrackingMode;
        ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        if (cm) { reply = cm.reply(`✖ Rejected. OMV standard wording put back (tracked change).${over}${why}`); cm.resolved = true; }
        if (f.b && c) {
          // Changed clause: put back only the paragraphs that differ from OMV's text.
          rewriteBody(paras, c, f.b.paras);
        } else if (!f.b && c) {
          // Clause the supplier added: take it out.
          clauseRange(paras, c).delete();
        } else if (f.b && !c) {
          // Clause the supplier deleted: put it back before the next OMV clause that is still there.
          const pos = CG.OMV_STANDARD.indexOf(f.b);
          const next = CG.OMV_STANDARD.slice(pos + 1).map((s) => locate(cp, s)).find(Boolean);
          const put = (t) => next ? paras.items[next.headIdx].insertParagraph(t, "Before") : ctx.document.body.insertParagraph(t, "End");
          put(`${f.b.num} ${f.b.head}`).styleBuiltIn = Word.Style.heading1;
          f.b.paras.forEach((t) => { const p = put(t); p.styleBuiltIn = Word.Style.normal; if (isSubhead(t)) p.font.bold = true; });
        }
        ctx.document.changeTrackingMode = before;
        text = (f.b ? "Rejected. OMV standard wording put back." : "Rejected. Supplier's new clause removed.") + " Shown as a tracked change.";
      }

      else if (k === "escalate") {
        const toWhom = to ? ` to ${to}` : "";
        if (cm) reply = cm.reply(`⚑ Escalated${toWhom}.${over}${why}`);
        text = `Escalated${toWhom}. The comment stays open.`;
      }

      if (reply) reply.load("id");
      await ctx.sync();
      decisions[i] = { k, text, note, to, replyId: reply ? reply.id : null };
    });
    pending = null;
    addLog(f, ACT_LABEL[k], [to ? `To: ${to}` : "", note].filter(Boolean).join(" · "));
    showResult(); renderRules();
    const n = decisions.filter(Boolean).length;
    status(n === decisions.length ? "All changes decided." : `${n} of ${decisions.length} decided.`);
  }

  // ---------------- change your mind ----------------
  async function undo(i) {
    const f = review.findings[i], d = decisions[i];
    let note = "Decision undone. Choose again.";
    await Word.run(async (ctx) => {
      const cm = await getComment(ctx, f);
      if (cm) {
        cm.replies.load("items/id");
        await ctx.sync();
        const r = cm.replies.items.find((x) => x.id === d.replyId);
        if (r) r.delete();
        cm.resolved = false;
      }
      await ctx.sync();

      if (d.k === "accept" && f.c) {
        // Put the yellow highlight back on the supplier's added words.
        const { paras, cp } = await readClauses(ctx);
        const c = locate(cp, f.c);
        if (c) {
          const res = findAdded(paras, c, f);
          await ctx.sync();
          res.forEach((x) => { if (x.items.length) x.items[0].font.highlightColor = "Yellow"; });
          await ctx.sync();
        }
      }

      if (d.k === "reject") note = await undoReject(ctx, f);
    });
    decisions[i] = null;
    if (proposing === i) proposing = null;
    addLog(f, "Go back", `Undid ${ACT_LABEL[d.k]}`);
    showResult(); renderRules();
    status(note);
  }

  // Make a clause's body read `want`, changing only the paragraphs that differ (others stay untouched).
  const isSubhead = (t) => /^\d+(\.\d+)+ /.test(t) && !/[.;:]$/.test(t);
  function rewriteBody(paras, c, want) {
    const ops = CG.diffSeq(want, c.paras); // "-" = wanted but missing, "+" = there but not wanted
    let j = 0, last = paras.items[c.headIdx], rem = [], add = [];
    const flush = () => {
      const n = Math.min(rem.length, add.length);
      for (let x = 0; x < n; x++) { add[x].insertText(rem[x], "Replace"); last = add[x]; }
      add.slice(n).forEach((p) => p.delete());
      for (const t of rem.slice(n)) {
        last = last.insertParagraph(t, "After");
        last.styleBuiltIn = Word.Style.normal;
        if (isSubhead(t)) last.font.bold = true;
      }
      rem = []; add = [];
    };
    for (const o of ops) {
      if (o.t === "=") { flush(); last = paras.items[c.bodyIdx[j++]]; }
      else if (o.t === "-") rem.push(o.w);
      else add.push(paras.items[c.bodyIdx[j++]]);
    }
    flush();
  }

  // Take back a Reject: get the supplier's wording back and remove the red tracked-change marks.
  // Tries three ways, from gentlest to strongest, and says which one worked.
  async function undoReject(ctx, f) {
    const can16 = Office.context.requirements.isSetSupported("WordApi", "1.6");
    const ours = new Set([...(f.b ? f.b.paras : []), ...(f.c ? f.c.paras : []), f.b ? `${f.b.num} ${f.b.head}` : "", f.c ? `${f.c.num} ${f.c.head}` : ""].map(norm).filter(Boolean));
    const countIn = async (c) => {
      const { paras } = await readClauses(ctx);
      const tc = clauseRange(paras, c).getTrackedChanges(); tc.load("items"); await ctx.sync();
      return tc;
    };
    let { cp } = await readClauses(ctx);
    let c = locate(cp, f.c || f.b);

    // 1. Undo the tracked changes inside this clause.
    if (can16 && c) {
      try {
        const tc = await countIn(c);
        if (tc.items.length) {
          const n = tc.items.length;
          tc.rejectAll(); await ctx.sync();
          ({ cp } = await readClauses(ctx)); c = locate(cp, f.c || f.b);
          if (!c || !(await countIn(c)).items.length) return `Decision undone. The supplier's wording is back (${n} tracked change${n === 1 ? "" : "s"} undone).`;
        }
      } catch (e) { console.warn("Way 1 failed", e); }

      // 2. Look through all tracked changes in the document for ours (same text as this clause).
      try {
        const all = ctx.document.body.getTrackedChanges(); all.load("items/text"); await ctx.sync();
        const mine = all.items.filter((t) => [...ours].some((o) => o.includes(norm(t.text)) && norm(t.text).length > 2));
        if (mine.length) { mine.forEach((t) => t.reject()); await ctx.sync(); return `Decision undone. The supplier's wording is back (${mine.length} tracked change${mine.length === 1 ? "" : "s"} undone).`; }
      } catch (e) { console.warn("Way 2 failed", e); }
    }

    // 3. Last resort: Track Changes off, write the supplier's original text back into the clause.
    try {
      ctx.document.load("changeTrackingMode"); await ctx.sync();
      const before = ctx.document.changeTrackingMode;
      ctx.document.changeTrackingMode = Word.ChangeTrackingMode.off;
      const { paras, cp: cp2 } = await readClauses(ctx);
      const cc = locate(cp2, f.c || f.b);
      if (!cc) { ctx.document.changeTrackingMode = before; await ctx.sync(); return "Decision undone in the panel, but I couldn't find the clause. Please use Review → Reject on the red text."; }
      if (!f.c) clauseRange(paras, cc).delete(); // we had put back a deleted clause: remove it again
      else rewriteBody(paras, cc, f.c.paras);     // change back only the paragraphs that differ
      ctx.document.changeTrackingMode = before;
      await ctx.sync();
      return "Decision undone. I wrote the supplier's wording back directly (Track Changes was paused for this).";
    } catch (e) {
      console.warn("Way 3 failed", e);
      return "Decision undone in the panel, but Word didn't let me remove the red text. Please click it and use Review → Reject. (" + (e.message || e) + ")";
    }
  }

  // ---------------- the changes tab ----------------
  // Word-by-word difference: deleted OMV words struck through, supplier's words underlined.
  function diffHtml(f) {
    if (!f.b) return `<ins>${esc(f.c.body)}</ins>`;
    if (!f.c) return `<del>${esc(f.b.body)}</del>`;
    // Long unchanged stretches are shortened to "…" so the change stands out.
    const ops = CG.diffWords(f.b.body, f.c.body), KEEP = 12, out = [];
    let run = [];
    const flushRun = (atStart, atEnd) => {
      if (run.length > 2 * KEEP) out.push(...(atStart ? [] : run.slice(0, KEEP)), "<span class=\"gap\">…</span>", ...(atEnd ? [] : run.slice(-KEEP)));
      else out.push(...run);
      run = [];
    };
    ops.forEach((o) => {
      if (o.t === "=") { run.push(esc(o.w)); return; }
      flushRun(out.length === 0, false);
      out.push(o.t === "-" ? `<del>${esc(o.w)}</del>` : `<ins>${esc(o.w)}</ins>`);
    });
    flushRun(out.length === 0, true);
    return out.join(" ").replace(/<\/del> <del>/g, " ").replace(/<\/ins> <ins>/g, " ");
  }

  // The short "when the supplier…" text for a rule proposal.
  function changeSummary(f) {
    if (!f.b) return `adds a new clause "${f.head}"`;
    if (!f.c) return `deletes clause ${f.b.num} (${f.head})`;
    const d = CG.sentenceDiff(f.b.paras, f.c.paras);
    if (!d.omv.length) return `adds "${d.supplier.join(" ")}"`;
    if (!d.supplier.length) return `deletes "${d.omv.join(" ")}"`;
    return `changes "${d.omv.join(" ")}" to "${d.supplier.join(" ")}"`;
  }

  function showResult() {
    const R = review;
    if (!R) return;
    const n = R.findings.length;
    const nd = decisions.filter(Boolean).length;
    const cnt = { escalate: 0, reject: 0, review: 0, accept: 0 };
    R.findings.forEach((f) => cnt[f.a]++);
    const moved = R.same.filter((s) => s.b.num !== s.c.num).map((s) => `${s.b.head} moved from ${s.b.num} to ${s.c.num}`);

    let h = `<div class="summary"><b>${n ? `${n} change${n > 1 ? "s" : ""} from the OMV standard` : "No changes from the OMV standard"}</b>`;
    if (n) {
      h += `<div class="pills">`;
      [["escalate", "to escalate"], ["reject", "to reject"], ["review", "you decide"], ["accept", "to accept"]].forEach(([k, t]) => { if (cnt[k]) h += `<span class="pill ${PILL[k]}">${cnt[k]} ${t}</span>`; });
      h += `</div>`;
    }
    h += `<div class="proof">✓ ${R.same.length} of ${R.total} OMV clauses are word-for-word identical.${moved.length ? " Same wording, new number: " + esc(moved.join("; ")) + "." : ""}</div>`;
    if (n) h += `<div class="progress">${nd} of ${n} decided</div><div class="bar"><i style="width:${Math.round((100 * nd) / n)}%"></i></div>`;
    h += `</div>`;

    R.findings.forEach((f, i) => {
      const d = decisions[i];
      const rec = f.a === "review" ? null : f.a;
      const tag = f.cosmetic ? "Wording only" : f.a === "review" ? "No rule: you decide" : "Suggest: " + CG.suggestion(f);
      h += `<div class="item l-${PILL[f.a]}${d ? " decided" : ""}" data-i="${i}" title="Click to show in the document">`;
      h += `<div class="top"><div class="h">${esc((f.c || f.b).num)}. ${esc(f.head)}</div><span class="tag ${PILL[f.a]}">${esc(tag)}</span></div>`;
      h += `<div class="w">${esc(f.why)}</div>`;
      h += `<div class="links">`;
      if (!f.cosmetic) h += `<details><summary>Show the change</summary><div class="diff">${diffHtml(f)}</div><div class="legend"><del>OMV wording removed</del> <ins>supplier's wording</ins></div></details>`;
      if (f.rule) h += `<a href="#" class="rulelink" data-id="${f.rule.id}">Rule ${esc(f.rule.id)} ▾</a>`;
      h += `</div>`;
      if (f.rule) {
        const r = f.rule;
        h += `<div class="rulepop" hidden><button class="x" title="Close">×</button><div class="rid">${esc(r.id)} <span>· ${esc(r.source.replace("OMV GPC Jan 2026, ", ""))}</span></div>`;
        h += `<div class="rstd"><b>OMV position:</b> ${esc(r.standard)}</div><ul>` + CG.OUTCOMES[r.id].map((o) => `<li>${esc(o)}</li>`).join("") + `</ul>`;
        h += `<a href="#" class="tobook" data-id="${r.id}">Open in Rulebook ›</a></div>`;
      }

      if (d) {
        h += `<div class="done d-${d.k}">${esc(d.text)}${d.note ? `<div class="note">Reason: ${esc(d.note)}</div>` : ""}</div>`;
        if (d.k === "escalate") {
          const m = escalationMail(f, d.note, d.to);
          h += `<a class="mini" href="${m.href}" target="_blank" rel="noopener">✉ Email draft${m.who ? " to " + esc(m.who) : ""}</a>`;
        }
        h += `<div class="after"><button class="undo" data-i="${i}">↩ Go back</button>`;
        if (!d.auto && (f.a === "review" || (rec && d.k !== rec))) h += `<button class="undo propose" data-i="${i}">✎ Propose a rule</button>`;
        h += `</div>`;
        if (proposing === i) {
          const should = ACT_LABEL[d.k] + (d.k === "escalate" ? " to " : "");
          h += `<div class="form" data-i="${i}"><div class="ft">Propose a rule to ${esc(CG.RULEBOOK_OWNER.name)}. It only goes live after Legal approves it.</div>
            <label>When the supplier…<textarea class="p-when" rows="3">${esc(changeSummary(f))}</textarea></label>
            <label>ContractGuard should suggest<input class="p-should" value="${esc(should)}"></label>
            <label>Because<textarea class="p-because" rows="2" placeholder="Why this is OMV's position">${esc(d.note || "")}</textarea></label>
            <div class="fbtns"><a class="mini send" href="#" target="_blank" rel="noopener">Send to Legal Ops</a><button class="undo cancel">Cancel</button></div></div>`;
        }
      } else if (pending && pending.i === i) {
        const isEsc = pending.k === "escalate", isOver = rec && pending.k !== rec;
        h += `<div class="form why">`;
        if (isOver) h += `<div class="ft">You're choosing <b>${ACT_LABEL[pending.k]}</b> instead of the suggested <b>${CG.LABEL[rec]}</b>.</div>`;
        if (isEsc) {
          h += `<label>Escalate to whom?<input class="p-to" placeholder="Name or team"></label>`;
          h += `<div class="chips">` + Object.keys(CG.CONTACTS).map((n) => `<button class="chip" type="button">${esc(n)}</button>`).join("") + `</div>`;
        }
        h += `<label>${isOver ? "Why? (optional: goes into the comment and the log)" : "Note (optional)"}<textarea rows="2" placeholder="${isOver ? "e.g. long-term supplier, agreed in the last negotiation" : "e.g. please decide by Friday"}"></textarea></label>`;
        h += `<div class="fbtns"><button class="act rec confirm">${ACT_LABEL[pending.k]}</button><button class="undo cancel">Cancel</button></div></div>`;
      } else {
        h += `<div class="acts">` + ACTIONS.map((a) => `<button class="act${a.k === rec ? " rec" : ""}" data-i="${i}" data-k="${a.k}"${busy ? " disabled" : ""}>${a.label}</button>`).join("") + `</div>`;
      }
      h += `</div>`;
    });
    $("result").innerHTML = h;
    wireResult();
  }

  function wireResult() {
    const R = review, root = $("result");
    // Click anywhere on a change (except its buttons and forms) to jump to it in the document.
    root.querySelectorAll(".item").forEach((el) => {
      el.onclick = (e) => {
        if (e.target.closest("button, summary, a, textarea, input, label, .form, .rulepop")) return;
        const i = +el.dataset.i;
        focusCard(i, false);
        guarded(() => jumpTo(R.findings[i]));
      };
    });
    // Rule link: show the rule in a small pop-up box inside the card; click again or × to close.
    root.querySelectorAll(".rulelink").forEach((a) => (a.onclick = (e) => {
      e.preventDefault();
      const pop = a.closest(".item").querySelector(".rulepop"), open = pop.hidden;
      pop.hidden = !open; a.textContent = a.textContent.replace(open ? "▾" : "▴", open ? "▴" : "▾");
    }));
    root.querySelectorAll(".rulepop .x").forEach((b) => (b.onclick = () => {
      const item = b.closest(".item"), a = item.querySelector(".rulelink");
      item.querySelector(".rulepop").hidden = true; a.textContent = a.textContent.replace("▴", "▾");
    }));
    root.querySelectorAll(".tobook").forEach((a) => (a.onclick = (e) => { e.preventDefault(); openRule(a.dataset.id); }));
    root.querySelectorAll(".act[data-k]").forEach((el) => (el.onclick = () => choose(+el.dataset.i, el.dataset.k)));
    root.querySelectorAll(".undo[data-i]:not(.propose)").forEach((el) => (el.onclick = () => guarded(() => undo(+el.dataset.i))));
    root.querySelectorAll(".propose").forEach((el) => (el.onclick = () => { proposing = +el.dataset.i; showResult(); }));
    const why = root.querySelector(".why");
    if (why) {
      const toBox = why.querySelector(".p-to");
      why.querySelectorAll(".chip").forEach((c) => (c.onclick = () => { toBox.value = c.textContent; why.querySelectorAll(".chip").forEach((x) => x.classList.toggle("on", x === c)); }));
      why.querySelector(".confirm").onclick = () => {
        const p = pending, note = why.querySelector("textarea").value.trim(), to = toBox ? toBox.value.trim() : "";
        guarded(() => decide(p.i, p.k, note, to));
      };
      why.querySelector(".cancel").onclick = () => { pending = null; showResult(); };
    }
    const form = root.querySelector(".form[data-i]");
    if (form) {
      const i = +form.dataset.i, f = R.findings[i], send = form.querySelector(".send");
      const update = () => { send.href = proposalMail(f, decisions[i], form.querySelector(".p-when").value, form.querySelector(".p-should").value, form.querySelector(".p-because").value); };
      form.querySelectorAll("textarea, input").forEach((x) => (x.oninput = update));
      update();
      send.onclick = () => { addLog(f, "Rule proposed", form.querySelector(".p-should").value + " · " + form.querySelector(".p-because").value); proposing = null; setTimeout(showResult, 0); status("Rule proposal opened in your e-mail. Send it to Legal Ops."); };
      form.querySelector(".cancel").onclick = () => { proposing = null; showResult(); };
    }
  }

  // Highlight a change card (and its rule), optionally scrolling the panel to it.
  function focusCard(i, scroll) {
    document.querySelectorAll("#result .item.active").forEach((x) => x.classList.remove("active"));
    const el = document.querySelector(`#result .item[data-i="${i}"]`);
    if (el) { el.classList.add("active"); if (scroll) el.scrollIntoView({ block: "nearest", behavior: "smooth" }); }
    const f = review && review.findings[i];
    if (f && f.rule && f.rule.id !== currentRule) { currentRule = f.rule.id; renderRules(); }
  }

  // Scroll the document to a change and select the changed words (or the changed paragraph).
  async function jumpTo(f) {
    await Word.run(async (ctx) => {
      const { paras, cp } = await readClauses(ctx);
      const c = f.c ? locate(cp, f.c) : null;
      const res = c ? findAdded(paras, c, f) : [];
      await ctx.sync();
      anchorRange(paras, cp, f, c, res).select();
      await ctx.sync();
    });
  }

  // ---------------- follow the cursor in the document ----------------
  let selTimer = null;
  function onSelection() {
    clearTimeout(selTimer);
    selTimer = setTimeout(() => {
      if (busy || !lastCp) return;
      Word.run(async (ctx) => {
        const p = ctx.document.getSelection().paragraphs.getFirst();
        p.load("text");
        await ctx.sync();
        const t = (p.text || "").trim();
        if (!t) return;
        const c = lastCp.find((x) => `${x.num} ${x.head}` === t || `${x.num}. ${x.head}` === t || x.head === t) || lastCp.find((x) => x.paras.includes(t));
        if (!c) return;
        const rule = CG.findRule(c.head);
        if (review) {
          const i = review.findings.findIndex((f) => f.c && norm(f.c.head) === norm(c.head));
          if (i >= 0 && tab === "changes") focusCard(i, true);
        }
        if (rule && rule.id !== currentRule) { currentRule = rule.id; renderRules(); if (tab === "rules") scrollToRule(); }
      }).catch(() => {});
    }, 250);
  }
})();
