/* ContractGuard Word add-in: the side panel. Uses Office.js to read the open document and add comments. */
(function () {
  "use strict";
  const CG = window.ContractGuard;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
  const PILL = { esc: "esc", counter: "counter", route: "route", ok: "ok" };
  const norm = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9%]+/g, " ").trim();

  // The lawyer's three choices (to edit, the lawyer just types in the document). "rec" = which button ContractGuard suggests for each kind of finding.
  const ACTIONS = [
    { k: "accept", label: "Accept" },
    { k: "reject", label: "Reject" },
    { k: "escalate", label: "Escalate" },
  ];
  const REC = { ok: "accept", counter: "reject", route: "escalate", esc: "escalate" };

  let review = null;     // last review result
  let decisions = [];    // decisions[i] = {k, text} for finding i, or null
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
    document.querySelectorAll(".act").forEach((b) => (b.disabled = true));
    try { await fn(); }
    catch (e) { console.error(e); status("Something went wrong: " + (e.message || e), true); }
    finally {
      busy = false;
      $("run").disabled = $("clear").disabled = false;
      document.querySelectorAll(".act").forEach((b) => (b.disabled = false));
    }
  }

  async function deleteOurComments(ctx) {
    const cs = ctx.document.body.getComments();
    cs.load("items/content");
    await ctx.sync();
    let n = 0;
    cs.items.forEach((c) => { if ((c.content || "").startsWith(CG.TAG)) { c.delete(); n++; } });
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

      // Write comments and queue highlight searches
      const searches = [], comments = [];
      for (const f of R.findings) {
        const range = f.c ? clauseRange(paras, f.c) : paras.items[0].getRange("Whole");
        if (f.c && f.changes && !f.cosmetic) {
          for (const phrase of f.changes.added) {
            const p = phrase.replace(/[.,;:]+$/, "");
            if (!p || p.length > 250) continue;
            for (const idx of f.c.bodyIdx) {
              const res = paras.items[idx].search(p, { matchCase: true });
              res.load("items");
              searches.push(res);
            }
          }
        }
        const cm = range.insertComment(CG.commentText(f));
        cm.load("id");
        comments.push(cm);
      }
      await ctx.sync();
      R.findings.forEach((f, i) => (f.commentId = comments[i].id));
      searches.forEach((res) => { if (res.items.length) res.items[0].font.highlightColor = "Yellow"; });
      await ctx.sync();

      review = R;
      decisions = R.findings.map(() => null);
      showResult();
      status("Done. Comments added. Decide each change below.");
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
      let text;

      if (k === "accept") {
        text = f.c ? "Accepted. Supplier's wording stays." : "Accepted. The clause stays deleted.";
        if (cm) { cm.reply("✔ " + text); cm.resolved = true; }
        if (c) clauseRange(paras, c).font.highlightColor = null;
      }

      else if (k === "reject") {
        ctx.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        if (cm) { cm.reply("✖ Rejected. OMV standard wording put back (tracked change)."); cm.resolved = true; }
        if (f.b && c) {
          // Changed clause: replace the supplier's text with OMV's.
          if (c.bodyIdx.length) {
            const first = paras.items[c.bodyIdx[0]];
            first.insertText(f.b.body, "Replace");
            first.getRange("Whole").font.highlightColor = null;
            c.bodyIdx.slice(1).forEach((idx) => paras.items[idx].delete());
          } else {
            paras.items[c.headIdx].insertParagraph(f.b.body, "After").styleBuiltIn = Word.Style.normal;
          }
        } else if (!f.b && c) {
          // Clause the supplier added: take it out.
          clauseRange(paras, c).delete();
        } else if (f.b && !c) {
          // Clause the supplier deleted: put it back before the next OMV clause that is still there.
          const pos = CG.OMV_STANDARD.findIndex((s) => s.num === f.b.num);
          let next = null;
          for (const s of CG.OMV_STANDARD.slice(pos + 1)) { next = locate(cp, s); if (next) break; }
          if (next) {
            const before = paras.items[next.headIdx];
            before.insertParagraph(`${f.b.num}. ${f.b.head}`, "Before");
            before.insertParagraph(f.b.body, "Before").styleBuiltIn = Word.Style.normal;
          } else {
            ctx.document.body.insertParagraph(`${f.b.num}. ${f.b.head}`, "End");
            ctx.document.body.insertParagraph(f.b.body, "End").styleBuiltIn = Word.Style.normal;
          }
        }
        text = f.b ? "Rejected. OMV standard wording put back." : "Rejected. Supplier's new clause removed.";
        text += " Shown as a tracked change.";
      }

      else if (k === "escalate") {
        const who = f.to || "Senior lawyer";
        if (cm) cm.reply(`⚑ Escalated to ${who}.`);
        text = `Escalated to ${who}. The comment stays open.`;
      }

      await ctx.sync();
      decisions[i] = { k, text };
    });
    showResult();
    const n = decisions.filter(Boolean).length;
    status(n === decisions.length ? "All changes decided." : `${n} of ${decisions.length} decided.`);
  }

  // ---------------- the panel ----------------
  function showResult() {
    const R = review;
    const cnt = { esc: 0, counter: 0, route: 0, ok: 0 };
    R.findings.forEach((f) => cnt[f.a]++);
    const n = R.findings.length;
    const nd = decisions.filter(Boolean).length;
    const moved = R.same.filter((s) => s.b.num !== s.c.num).map((s) => `${s.b.head} moved from ${s.b.num} to ${s.c.num}`);
    let h = `<div class="summary"><b>${n ? `${n} change${n > 1 ? "s" : ""}, ${cnt.esc} escalation${cnt.esc === 1 ? "" : "s"}` : "No changes found"}</b><div class="pills">`;
    ["esc", "counter", "route", "ok"].forEach((k) => { if (cnt[k]) h += `<span class="pill ${PILL[k]}">${cnt[k]} ${CG.LABEL[k].toLowerCase()}</span>`; });
    h += `<span class="pill same">${R.same.length} identical</span></div>`;
    h += `<div>Proof nothing else changed: ${R.same.length} of ${R.total} OMV clauses are word-for-word identical.${moved.length ? " Same wording, new number: " + esc(moved.join("; ")) + "." : ""}</div>`;
    if (n) h += `<div class="progress">${nd} of ${n} decided</div><div class="bar"><i style="width:${Math.round((100 * nd) / n)}%"></i></div>`;
    h += `</div>`;

    R.findings.forEach((f, i) => {
      const label = f.cosmetic ? "Wording only" : CG.LABEL[f.a] + (f.to ? ` to ${f.to}` : "");
      const d = decisions[i];
      h += `<div class="item l-${f.a}${d ? " decided" : ""}"><div class="top" data-i="${i}"><div class="h">${esc((f.c || f.b).num)}. ${esc(f.head)} · ${esc(label)}</div><div class="w">${esc(f.why)}</div></div>`;
      if (d) {
        h += `<div class="done d-${d.k}">${esc(d.text)}</div>`;
      } else {
        const rec = REC[f.a];
        h += `<div class="acts">` + ACTIONS.map((a) => `<button class="act${a.k === rec ? " rec" : ""}" data-i="${i}" data-k="${a.k}"${busy ? " disabled" : ""}>${a.label}</button>`).join("") + `</div>`;
        h += `<div class="hint">Suggested: ${ACTIONS.find((a) => a.k === rec).label}</div>`;
      }
      h += `</div>`;
    });
    $("result").innerHTML = h;
    $("result").querySelectorAll(".top").forEach((el) => {
      el.onclick = () => { const f = R.findings[+el.dataset.i]; if (f.c) guarded(() => jumpTo(f.c)); };
    });
    $("result").querySelectorAll(".act").forEach((el) => {
      el.onclick = () => guarded(() => decide(+el.dataset.i, el.dataset.k));
    });
  }

  async function jumpTo(ref) {
    await Word.run(async (ctx) => {
      const { paras, cp } = await readClauses(ctx);
      const c = locate(cp, ref);
      if (c) paras.items[c.headIdx].select();
      await ctx.sync();
    });
  }
})();
