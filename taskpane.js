/* ContractGuard Word add-in: the side panel. Uses Office.js to read the open document and add comments. */
(function () {
  "use strict";
  const CG = window.ContractGuard;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
  const PILL = { esc: "esc", counter: "counter", route: "route", ok: "ok" };
  let lastClauses = [];

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
    $("run").disabled = $("clear").disabled = true;
    try { await fn(); }
    catch (e) { console.error(e); status("Something went wrong: " + (e.message || e), true); }
    finally { $("run").disabled = $("clear").disabled = false; }
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
    });
  }

  async function runReview() {
    status("Reading the document…");
    await Word.run(async (ctx) => {
      await deleteOurComments(ctx);
      const paras = ctx.document.body.paragraphs;
      paras.load("items/text,items/styleBuiltIn");
      await ctx.sync();

      // 1. Read
      const list = paras.items.map((p) => ({
        text: p.text,
        isHeading: /^Heading\d/.test(p.styleBuiltIn || ""),
        isTitle: p.styleBuiltIn === "Title",
      }));
      const cp = CG.splitClauses(list);
      lastClauses = cp;
      if (cp.length < 2) {
        status("I couldn't find numbered clauses (like \"3. Prices and Payment\") in this document, so I can't review it. Please review it by hand.", true);
        return;
      }

      // 2-5. Compare, check, decide
      status("Comparing with the OMV standard…");
      const R = CG.review(CG.OMV_STANDARD, cp);

      // Write comments and queue highlight searches
      const searches = [];
      for (const f of R.findings) {
        let range;
        if (!f.c) {
          range = paras.items[0].getRange("Whole");
        } else {
          range = paras.items[f.c.headIdx].getRange("Whole");
          if (f.c.bodyIdx.length) range = range.expandTo(paras.items[f.c.bodyIdx[f.c.bodyIdx.length - 1]].getRange("Whole"));
          if (f.changes && !f.cosmetic) {
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
        }
        range.insertComment(CG.commentText(f));
      }
      await ctx.sync();
      searches.forEach((res) => { if (res.items.length) res.items[0].font.highlightColor = "Yellow"; });
      await ctx.sync();

      showResult(R);
      status("Done. Comments added to the document.");
    });
  }

  function showResult(R) {
    const cnt = { esc: 0, counter: 0, route: 0, ok: 0 };
    R.findings.forEach((f) => cnt[f.a]++);
    const n = R.findings.length;
    const moved = R.same.filter((s) => s.b.num !== s.c.num).map((s) => `${s.b.head} moved from ${s.b.num} to ${s.c.num}`);
    let h = `<div class="summary"><b>${n ? `${n} change${n > 1 ? "s" : ""}, ${cnt.esc} escalation${cnt.esc === 1 ? "" : "s"}` : "No changes found"}</b><div class="pills">`;
    ["esc", "counter", "route", "ok"].forEach((k) => { if (cnt[k]) h += `<span class="pill ${PILL[k]}">${cnt[k]} ${CG.LABEL[k].toLowerCase()}</span>`; });
    h += `<span class="pill same">${R.same.length} identical</span></div>`;
    h += `<div>Proof nothing else changed: ${R.same.length} of ${R.total} OMV clauses are word-for-word identical.${moved.length ? " Same wording, new number: " + esc(moved.join("; ")) + "." : ""}</div></div>`;
    R.findings.forEach((f, i) => {
      const label = f.cosmetic ? "Wording only" : CG.LABEL[f.a] + (f.to ? ` to ${f.to}` : "");
      h += `<button class="item l-${f.a}" data-i="${i}"><div class="h">${esc((f.c || f.b).num)}. ${esc(f.head)} · ${esc(label)}</div><div class="w">${esc(f.why)}</div></button>`;
    });
    $("result").innerHTML = h;
    $("result").querySelectorAll(".item").forEach((el) => {
      el.onclick = () => { const f = R.findings[+el.dataset.i]; if (f.c) guarded(() => jumpTo(f.c.headIdx)); };
    });
  }

  async function jumpTo(idx) {
    await Word.run(async (ctx) => {
      const paras = ctx.document.body.paragraphs; paras.load("items"); await ctx.sync();
      if (paras.items[idx]) paras.items[idx].select();
      await ctx.sync();
    });
  }
})();
