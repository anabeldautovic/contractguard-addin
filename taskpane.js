/* ContractGuard Word add-in: the side panel. Uses Office.js to read the open document and add comments. */
(function () {
  "use strict";
  const CG = window.ContractGuard;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
  // Colour for each kind of suggestion (CSS classes).
  const PILL = { escalate: "esc", reject: "counter", review: "route", accept: "ok" };
  const norm = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9%]+/g, " ").trim();

  // The lawyer's three choices (to edit, the lawyer just types in the document).
  const ACTIONS = [
    { k: "accept", label: "Accept" },
    { k: "reject", label: "Reject" },
    { k: "escalate", label: "Escalate" },
  ];

  let review = null;     // last review result
  let decisions = [];    // decisions[i] = {k, text, replyId, auto} for finding i, or null
  let busy = false;

  function status(msg, isErr) { const s = $("status"); s.textContent = msg; s.className = isErr ? "err" : ""; }

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Word) { status("Open this in Word.", true); return; }
    if (!Office.context.requirements.isSetSupported("WordApi", "1.4")) {
      status("This version of Word is too old for comments from add-ins. Try Word in the browser (office.com).", true); return;
    }
    $("run").disabled = false; $("clear").disabled = false;
    $("run").onclick = () => guarded(runReview);
    $("clear").onclick = () => guarded(clearComments);
    status("Open a supplier's contract, then press Run review.");
  });

  async function guarded(fn) {
    if (busy) return;
    busy = true;
    $("run").disabled = $("clear").disabled = true;
    document.querySelectorAll(".act,.undo").forEach((b) => (b.disabled = true));
    try { await fn(); }
    catch (e) { console.error(e); status("Something went wrong: " + (e.message || e), true); }
    finally {
      busy = false;
      $("run").disabled = $("clear").disabled = false;
      document.querySelectorAll(".act,.undo").forEach((b) => (b.disabled = false));
    }
  }

  async function deleteOurComments(ctx) {
    const cs = ctx.document.body.getComments();
    cs.load("items/content");
    await ctx.sync();
    let n = 0;
    cs.items.forEach((c) => { if (CG.isOurs(c.content)) { c.delete(); n++; } });
    await ctx.sync();
    return n;
  }

  async function clearComments() {
    await Word.run(async (ctx) => {
      const n = await deleteOurComments(ctx);
      status(n ? `Removed ${n} ContractGuard comment${n > 1 ? "s" : ""}. Yellow highlights stay; use Undo or remove them by hand.` : "No ContractGuard comments found.");
      $("result").innerHTML = "";
      review = null; decisions = [];
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
    return { paras, cp: CG.splitClauses(list) };
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
      await ctx.sync();

      review = R;
      decisions = R.findings.map((f) => f.cosmetic ? { k: "accept", auto: true, text: "Accepted automatically." } : null);
      showResult();
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

  async function decide(i, k) {
    const f = review.findings[i];
    await Word.run(async (ctx) => {
      const cm = await getComment(ctx, f);
      const { paras, cp } = await readClauses(ctx);
      const c = f.c ? locate(cp, f.c) : null;
      let text, reply = null;

      if (k === "accept") {
        text = f.c ? "Accepted. Supplier's wording stays." : "Accepted. The clause stays deleted.";
        if (cm) { reply = cm.reply("✔ " + text); cm.resolved = true; }
        if (c) clauseRange(paras, c).font.highlightColor = null;
      }

      else if (k === "reject") {
        ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        if (cm) { reply = cm.reply("✖ Rejected. OMV standard wording put back (tracked change)."); cm.resolved = true; }
        const isSubhead = (t) => /^\d+(\.\d+)+ /.test(t) && !/[.;:]$/.test(t);
        if (f.b && c) {
          // Changed clause: go paragraph by paragraph and put back only what differs from OMV's text.
          const ops = CG.diffSeq(f.b.paras, c.paras);
          let j = 0, last = paras.items[c.headIdx], rem = [], add = [];
          const flush = () => {
            const k = Math.min(rem.length, add.length);
            for (let x = 0; x < k; x++) { add[x].insertText(rem[x], "Replace"); add[x].getRange("Whole").font.highlightColor = null; last = add[x]; }
            add.slice(k).forEach((p) => p.delete());
            for (const t of rem.slice(k)) {
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
        text = f.b ? "Rejected. OMV standard wording put back." : "Rejected. Supplier's new clause removed.";
        text += " Shown as a tracked change.";
      }

      else if (k === "escalate") {
        if (cm) reply = cm.reply("⚑ Escalated.");
        text = "Escalated. The comment stays open.";
      }

      if (reply) reply.load("id");
      await ctx.sync();
      decisions[i] = { k, text, replyId: reply ? reply.id : null };
    });
    showResult();
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

      if (d.k === "accept") {
        // Put the yellow highlight back on the supplier's added words.
        const { paras, cp } = await readClauses(ctx);
        const c = f.c ? locate(cp, f.c) : null;
        if (c) {
          const res = findAdded(paras, c, f);
          await ctx.sync();
          res.forEach((x) => { if (x.items.length) x.items[0].font.highlightColor = "Yellow"; });
          await ctx.sync();
        }
      }

      if (d.k === "reject") {
        // Undo our tracked change so the supplier's wording comes back.
        if (Office.context.requirements.isSetSupported("WordApi", "1.6")) {
          const { paras, cp } = await readClauses(ctx);
          const c = locate(cp, f.c || f.b);
          if (c) {
            clauseRange(paras, c).getTrackedChanges().rejectAll();
            await ctx.sync();
          } else {
            note = "Decision undone in the panel. Couldn't find the clause, so please undo the tracked change yourself: Review → Reject.";
          }
        } else {
          note = "Decision undone in the panel. This Word can't undo tracked changes for me, so please do it yourself: click the clause, then Review → Reject.";
        }
      }
    });
    decisions[i] = null;
    showResult();
    status(note);
  }

  // ---------------- the panel ----------------
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

  function showResult() {
    const R = review;
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
      h += `<div class="top" data-i="${i}" title="Show in document"><div class="h">${esc((f.c || f.b).num)}. ${esc(f.head)}</div><span class="tag ${PILL[f.a]}">${esc(tag)}</span></div>`;
      h += `<div class="w">${esc(f.why)}</div>`;
      if (!f.cosmetic) h += `<details><summary>Show the change</summary><div class="diff">${diffHtml(f)}</div><div class="legend"><del>OMV wording removed</del> <ins>supplier's wording</ins></div></details>`;
      if (d) {
        h += `<div class="done d-${d.k}">${esc(d.text)}</div><button class="undo" data-i="${i}">↩ Go back</button>`;
      } else {
        h += `<div class="acts">` + ACTIONS.map((a) => `<button class="act${a.k === rec ? " rec" : ""}" data-i="${i}" data-k="${a.k}"${busy ? " disabled" : ""}>${a.label}</button>`).join("") + `</div>`;
      }
      h += `</div>`;
    });
    $("result").innerHTML = h;
    // Click anywhere on a change (except its buttons) to jump to it in the document.
    $("result").querySelectorAll(".item").forEach((el) => {
      el.onclick = (e) => {
        if (e.target.closest("button, summary, a")) return;
        $("result").querySelectorAll(".item.active").forEach((x) => x.classList.remove("active"));
        el.classList.add("active");
        guarded(() => jumpTo(R.findings[+el.dataset.i]));
      };
    });
    $("result").querySelectorAll(".undo").forEach((el) => {
      el.onclick = () => guarded(() => undo(+el.dataset.i));
    });
    $("result").querySelectorAll(".act").forEach((el) => {
      el.onclick = () => guarded(() => decide(+el.dataset.i, el.dataset.k));
    });
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

})();
