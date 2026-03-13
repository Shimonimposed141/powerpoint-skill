const pptxgen = require("pptxgenjs");
const fs = require("fs");
let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Bünz, Chen, DeStefano";
pres.title = "Linear-Time Permutation Check";

const T = {
  bg: { title: "1E2761", content: "F8F9FA", section: "1E2761" },
  tx: { pri: "1A1A2E", sec: "4A4A5A", mut: "8C8C8C", wh: "FFFFFF" },
  ac: { pos: "0173B2", neg: "DE8F05", emp: "029E73" },
  chart: ["0173B2", "DE8F05", "029E73", "CC78BC"],
  cardFill: "0173B215", divider: "D0D0D0"
};
const F = {
  cover:    { face: "Georgia", size: 48 },
  section:  { face: "Georgia", size: 36 },
  title:    { face: "Georgia", size: 22 },
  subtitle: { face: "Calibri", size: 18 },
  body:     { face: "Calibri", size: 15 },
  small:    { face: "Calibri", size: 12 },
  caption:  { face: "Calibri", size: 11 },
  stat:     { face: "Georgia", size: 60 },
  cardTitle:{ face: "Georgia", size: 18 },
  tblHead:  { face: "Calibri", size: 13 },
  tblCell:  { face: "Calibri", size: 12 },
};
const SW = 10, SH = 5.625, M = 0.5;
const CW = SW - 2 * M;
const CY = 0.9;
const CH = SH - CY - M;

// Slide Masters
pres.defineSlideMaster({ title: "TITLE_SLIDE", background: { color: T.bg.title },
  objects: [{ rect: { x: 0, y: 4.8, w: SW, h: 0.825, fill: { color: "000000", transparency: 30 } } }]
});
pres.defineSlideMaster({ title: "SECTION_SLIDE", background: { color: T.bg.section } });
pres.defineSlideMaster({ title: "CONTENT_SLIDE", background: { color: T.bg.content },
  objects: [{ rect: { x: 0, y: 0, w: SW, h: 0.7, fill: { color: T.bg.title } } }]
});
pres.defineSlideMaster({ title: "THANK_YOU", background: { color: T.bg.title } });

// Manifests
const manifest = JSON.parse(fs.readFileSync("formulas/manifest.json", "utf8"));
const dManifest = (() => {
  try { return JSON.parse(fs.readFileSync("diagrams/manifest.json", "utf8")); }
  catch { return {}; }
})();

// Shadow factory
const mkSh = () => ({ type: "outer", color: "000000", blur: 6, offset: 2, angle: 135, opacity: 0.15 });

// OMML placeholder
function addMathText(slide, id, x, y, w, h, opts = {}) {
  slide.addText(`{{MATH:${id}}}`, {
    x, y, w, h,
    fontFace: opts.fontFace || F.body.face,
    fontSize: opts.fontSize || F.body.size,
    color: opts.color || T.tx.pri,
    align: opts.align || "center",
    valign: opts.valign || "middle",
    margin: 0
  });
}

// Formula helpers
function addFormula(slide, id, y, targetH = 0.4) {
  const f = manifest[id];
  if (!f || f.error) return;
  if (f.render === "omml") { addMathText(slide, id, M, y, CW, targetH); return; }
  const dpi = f.dpi || 600;
  const scale = targetH / (f.height_px / dpi);
  const w = (f.width_px / dpi) * scale;
  slide.addImage({ path: f.file, x: (SW - w) / 2, y, w, h: targetH });
}
function addFormulaAt(slide, id, x, y, maxW, targetH) {
  const f = manifest[id];
  if (!f || f.error) return;
  if (f.render === "omml") { addMathText(slide, id, x, y, maxW, targetH, { align: "left" }); return; }
  const dpi = f.dpi || 600;
  let scale = targetH / (f.height_px / dpi);
  let w = (f.width_px / dpi) * scale, h = targetH;
  if (w > maxW) { const s = maxW / w; w = maxW; h *= s; }
  slide.addImage({ path: f.file, x, y, w, h });
}

// Diagram helpers
function addDiagram(slide, id, y, targetH, maxW = CW) {
  const d = dManifest[id];
  if (!d || d.error) return;
  let w, h = targetH;
  if (d.format === "svg") {
    const b64 = Buffer.from(fs.readFileSync(d.file, "utf8")).toString("base64");
    const scale = targetH / (d.height_pt / 72);
    w = (d.width_pt / 72) * scale;
    if (w > maxW) { const s = maxW / w; w = maxW; h *= s; }
    slide.addImage({ data: "image/svg+xml;base64," + b64, x: (SW - w) / 2, y, w, h });
  }
}
function addDiagramAt(slide, id, x, y, maxW, targetH) {
  const d = dManifest[id];
  if (!d || d.error) return;
  let w, h = targetH;
  if (d.format === "svg") {
    const b64 = Buffer.from(fs.readFileSync(d.file, "utf8")).toString("base64");
    const scale = targetH / (d.height_pt / 72);
    w = (d.width_pt / 72) * scale;
    if (w > maxW) { const s = maxW / w; w = maxW; h *= s; }
    slide.addImage({ data: "image/svg+xml;base64," + b64, x, y, w, h });
  }
}

// Layout helpers
function sTitle(sl, txt) {
  sl.addText(txt, { x: M, y: 0.08, w: CW, h: 0.55,
    fontFace: F.title.face, fontSize: F.title.size, bold: true,
    color: T.tx.wh, margin: [0, 6, 0, 0], shrinkText: true });
}
function sectionSlide(title) {
  const sl = pres.addSlide({ masterName: "SECTION_SLIDE" });
  sl.addText(title, { x: M, y: SH / 2 - 0.5, w: CW, h: 1,
    fontFace: F.section.face, fontSize: F.section.size, bold: true,
    color: T.tx.wh, align: "center", valign: "middle" });
  sl.addShape(pres.shapes.RECTANGLE, { x: SW / 2 - 1.5, y: SH / 2 + 0.4,
    w: 3, h: 0.04, fill: { color: T.ac.pos } });
  return sl;
}
function addBullets(sl, items, x, y, w, h, opts = {}) {
  const arr = items.map((t, i) => ({
    text: t, options: {
      bullet: true, breakLine: i < items.length - 1,
      fontSize: opts.fontSize || F.body.size,
      fontFace: opts.fontFace || F.body.face,
      color: opts.color || T.tx.pri,
    }
  }));
  sl.addText(arr, { x, y, w, h, valign: "top", margin: [4, 6, 4, 4], shrinkText: true });
}
function addCard(sl, x, y, w, h, color, title, body) {
  sl.addShape(pres.shapes.RECTANGLE, { x, y, w, h,
    fill: { color: "FFFFFF" }, shadow: mkSh() });
  sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h,
    fill: { color } });
  // px: 0.19" clear of accent bar. pw: 0.15" from right edge.
  const px = x + 0.25, pw = w - 0.4;
  sl.addText(title, { x: px, y: y + 0.1, w: pw, h: 0.4,
    fontFace: F.cardTitle.face, fontSize: F.cardTitle.size, bold: true,
    color: T.tx.pri, valign: "top", margin: [0, 0, 0, 0], shrinkText: true });
  sl.addText(body, { x: px, y: y + 0.48, w: pw, h: h - 0.62,
    fontFace: F.body.face, fontSize: F.body.size,
    color: T.tx.sec, valign: "top", margin: [2, 0, 8, 0], shrinkText: true });
}
function addCardFormula(sl, x, y, w, h, color, title, formulaId, fH) {
  sl.addShape(pres.shapes.RECTANGLE, { x, y, w, h,
    fill: { color: "FFFFFF" }, shadow: mkSh() });
  sl.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.06, h,
    fill: { color } });
  const px = x + 0.25, pw = w - 0.4;
  sl.addText(title, { x: px, y: y + 0.1, w: pw, h: 0.4,
    fontFace: F.cardTitle.face, fontSize: F.cardTitle.size, bold: true,
    color: T.tx.pri, valign: "top", margin: [0, 0, 0, 0], shrinkText: true });
  addFormulaAt(sl, formulaId, px, y + 0.55, pw, fH || (h - 0.7));
}
function addTable(sl, headers, rows, x, y, w, colW) {
  const hRow = headers.map(h => ({ text: h, options: {
    fill: { color: T.bg.title }, color: T.tx.wh,
    bold: true, fontFace: F.tblHead.face, fontSize: F.tblHead.size,
    align: "center", valign: "middle"
  }}));
  const dRows = rows.map((row, i) => row.map(cell => ({
    text: String(cell), options: {
      fill: { color: i % 2 === 0 ? "FFFFFF" : "F0F2F5" },
      fontFace: F.tblCell.face, fontSize: F.tblCell.size,
      color: T.tx.pri, valign: "middle"
    }
  })));
  sl.addTable([hRow, ...dRows], { x, y, w, colW, border: { pt: 0.5, color: "E0E0E0" } });
}
function addNumCard(sl, num, title, body, x, y, w, h, color) {
  addCard(sl, x, y, w, h, color, title, body);
  sl.addShape(pres.shapes.OVAL, { x: x + w - 0.4, y: y - 0.15,
    w: 0.35, h: 0.35, fill: { color } });
  sl.addText(String(num), { x: x + w - 0.4, y: y - 0.15, w: 0.35, h: 0.35,
    fontFace: F.title.face, fontSize: 14, bold: true,
    color: "FFFFFF", align: "center", valign: "middle", margin: 0 });
}
function cols2() {
  const gap = 0.3, colW = (CW - gap) / 2;
  return { left: { x: M, w: colW }, right: { x: M + colW + gap, w: colW }, colW, gap };
}

// ============================================================
// SLIDE 1: Title
// ============================================================
{
  const sl = pres.addSlide({ masterName: "TITLE_SLIDE" });
  sl.addText("Linear-Time\nPermutation Check", { x: M, y: 1.2, w: CW, h: 1.4,
    fontFace: F.cover.face, fontSize: 44, bold: true,
    color: T.tx.wh, margin: 0 });
  sl.addText("Improved Prover, Verifier & Soundness\nfor SNARK Permutation Arguments", {
    x: M, y: 2.6, w: CW, h: 0.8,
    fontFace: F.subtitle.face, fontSize: F.subtitle.size,
    color: T.tx.wh, margin: 0 });
  sl.addText("Benedikt B\u00FCnz, Jessica Chen, Zachary DeStefano  \u2014  New York University", {
    x: M, y: 4.85, w: CW, h: 0.5,
    fontFace: F.caption.face, fontSize: F.body.size,
    color: T.tx.wh, align: "left", margin: 0 });
}

// ============================================================
// SLIDE 2: Why Permutation Checks Matter (two-column)
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "Why Permutation Checks Matter");
  const c = cols2();

  sl.addText("Wiring check is fundamental to every SNARK:", {
    x: M, y: CY, w: CW, h: 0.35,
    fontFace: F.body.face, fontSize: F.body.size, bold: true, color: T.tx.pri });

  addFormula(sl, "f19-perm-relation", CY + 0.4, 0.35);

  addCard(sl, c.left.x, CY + 0.9, c.left.w, 1.6, T.ac.pos,
    "Gate Check",
    "Verify each gate computes correctly. Efficient via single sumcheck invocation \u2014 O(n) prover, O(log n) verifier.");

  addCard(sl, c.right.x, CY + 0.9, c.right.w, 1.6, T.ac.neg,
    "Wiring Check (Bottleneck)",
    "Prove input/output connections are consistent. Current methods commit n extra elements or use O(log\u00B2 n)-round GKR.");

  addBullets(sl, [
    "HyperPlonk, Spartan, STARK: all rely on permutation checks",
    "Memory checking in zkVMs (read/write consistency)",
    "Lookup arguments for range proofs and non-arithmetic ops",
  ], M, CY + 2.7, CW, 1.5);
}

// ============================================================
// SLIDE 3: Current Approaches & Limitations (two-column)
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "Current Approaches & Their Limitations");
  const c = cols2();

  addCard(sl, c.left.x, CY, c.left.w, 2.6, T.ac.neg,
    "Product Check (Plonk-style)",
    "Commit to n extra field elements of full width.\nSoundness error: n/|F| (linear in input size).\nFor n=2\u00B3\u00B2, |F|=2\u00B9\u00B2\u2078: security only 2\u207B\u2079\u2076.\nRecent Fiat-Shamir attacks raise further concerns.");

  addCard(sl, c.right.x, CY, c.right.w, 2.6, T.ac.pos,
    "GKR-based (Spartan-style)",
    "No extra commitments \u2014 O(n) field ops.\nBut: O(log\u00B2 n) rounds, proof size, verifier time.\nSame poor soundness: n/|F|.\nHigher prover latency, implementation complexity.");

  addCard(sl, M, CY + 2.8, CW, 1.3, T.ac.emp,
    "The Gap",
    "Neither achieves the ideal: linear prover, log verifier, polylog soundness, no extra commitments, any PCS. Our BiPerm and MulPerm close this gap.");
}

// ============================================================
// SECTION: Key Idea
// ============================================================
sectionSlide("Key Idea: Sumcheck Formulation");

// ============================================================
// SLIDE 4: Core Insight — Sumcheck Formulation
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "Reducing Permutation to Sumcheck");

  sl.addText("Rewrite the permutation claim as a sum over an indicator function:", {
    x: M, y: CY, w: CW, h: 0.35,
    fontFace: F.body.face, fontSize: F.body.size, color: T.tx.pri });

  addFormula(sl, "f01-sumcheck-core", CY + 0.45, 0.4);

  sl.addText("where the indicator function is:", {
    x: M, y: CY + 0.95, w: CW, h: 0.3,
    fontFace: F.body.face, fontSize: F.body.size, color: T.tx.sec });

  addFormula(sl, "f02-indicator", CY + 1.3, 0.55);

  addCard(sl, M, CY + 2.05, CW, 2.0, T.ac.pos,
    "Key Question: How to Arithmetize the Indicator?",
    "The efficiency of the entire protocol hinges on the degree of the arithmetization.\nHyperPlonk: degree \u03BC \u2192 O(n\u00B7\u03BC\u00B2) prover time.\nOur insight: decompose the indicator into fewer, larger parts \u2192 lower degree \u2192 faster prover.");
}

// ============================================================
// SLIDE 5: Degree Tradeoff Spectrum
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "Degree\u2013Complexity Tradeoff");

  sl.addText("Splitting the indicator into fewer parts reduces prover cost:", {
    x: M, y: CY, w: CW, h: 0.35,
    fontFace: F.body.face, fontSize: F.body.size, color: T.tx.pri });

  addDiagram(sl, "d01-degree-tradeoff", CY + 0.4, 1.0);

  addFormula(sl, "f21-degree-tradeoff", CY + 1.55, 0.45);

  const c = cols2();
  addCard(sl, c.left.x, CY + 2.15, c.left.w, 2.0, T.ac.pos,
    "BiPerm (degree 2)",
    "Split indicator into 2 halves.\nO(n) field ops, O(\u03BC/|F|) soundness.\nRequires sparse PCS (Hyrax, Dory, KZH).\nSingle degree-3 sumcheck, no grand product.");

  addCard(sl, c.right.x, CY + 2.15, c.right.w, 2.0, T.ac.emp,
    "MulPerm (degree \u2113)",
    "Split into \u2113 = \u221A\u03BC parts.\nn\u00B7\u00D5(\u221Alog n) field ops, any PCS.\nPolylog soundness: O(\u03BC\u00B3\u02C0\u00B2/|F|).\nTwo sumchecks, compatible with all PCS.");
}

// ============================================================
// SECTION: BiPerm
// ============================================================
sectionSlide("BiPerm \u2014 Linear-Time Protocol");

// ============================================================
// SLIDE 6: BiPerm — Split Indicator
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "BiPerm: Splitting the Indicator");

  sl.addText("Factor the indicator into left and right halves of the output bits:", {
    x: M, y: CY, w: CW, h: 0.35,
    fontFace: F.body.face, fontSize: F.body.size, color: T.tx.pri });

  addFormula(sl, "f05-biperm-split", CY + 0.45, 0.35);

  sl.addText("This yields a degree-3 sumcheck:", {
    x: M, y: CY + 0.9, w: CW, h: 0.3,
    fontFace: F.body.face, fontSize: F.body.size, color: T.tx.sec });

  addFormula(sl, "f06-biperm-sumcheck", CY + 1.25, 0.4);

  const c = cols2();
  addCard(sl, c.left.x, CY + 1.85, c.left.w, 2.3, T.ac.pos,
    "Advantages",
    "Single degree-3 sumcheck, no grand product.\nn-sparse polynomials of size n\u00B9\u02D9\u2075.\nO(n) prover, O(\u03BC) verifier.\nO(\u03BC/|F|) soundness, O(log n) proof.");

  addCard(sl, c.right.x, CY + 1.85, c.right.w, 2.3, T.ac.neg,
    "Limitation",
    "MLEs have n\u221An entries (n-sparse).\nNeeds PCS with sparsity-aware commit/open.\n\u2713 Hyrax, Dory, KZH.\n\u2717 KZG, FRI, Ligero.\nMotivates MulPerm for universal PCS.");
}

// ============================================================
// SLIDE 7: BiPerm Properties (stat-callout + table)
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "BiPerm: Properties & PCS Compatibility");

  addFormula(sl, "f07-biperm-props", CY + 0.1, 0.35);

  addTable(sl,
    ["PCS", "Preprocess", "Prove", "Open", "Compatible?"],
    [
      ["Hyrax/KZH", "O(n)", "O(n)", "O(\u221An)", "\u2713 Optimal"],
      ["Dory",       "O(n)", "O(n)", "O(\u221An)", "\u2713 Optimal"],
      ["KZG",        "O(n\u00B9\u02D9\u2075)", "O(n\u00B9\u02D9\u2075)", "O(n)", "\u2717 Superlinear"],
      ["FRI/WHIR",   "O(n\u00B9\u02D9\u2075)", "O(n\u00B9\u02D9\u2075)", "O(n)", "\u2717 Superlinear"],
    ],
    M, CY + 0.6, CW, [1.5, 1.8, 1.8, 1.8, 2.1]);

  addCard(sl, M, CY + 2.5, CW, 1.65, T.ac.emp,
    "Practical Takeaway",
    "Ideal with sparse-friendly PCS. Strictly O(n) prover beyond witness commitment.\nPreprocessed index: O(n\u00B9\u02D9\u2075) size, O(n) weight.\nShout (ST25) achieves same structure for lookups.");
}

// ============================================================
// SECTION: MulPerm
// ============================================================
sectionSlide("MulPerm \u2014 Universal Protocol");

// ============================================================
// SLIDE 8: MulPerm Decomposition
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "MulPerm: Multi-Part Decomposition");

  sl.addText("Generalize BiPerm: decompose indicator into \u2113 parts, each mapping \u03BC/\u2113 bits:", {
    x: M, y: CY, w: CW, h: 0.35,
    fontFace: F.body.face, fontSize: F.body.size, color: T.tx.pri });

  addFormula(sl, "f08-mulperm-decompose", CY + 0.45, 0.4);

  sl.addText("Each part is a product of eq polynomials:", {
    x: M, y: CY + 0.95, w: CW, h: 0.3,
    fontFace: F.body.face, fontSize: F.body.size, color: T.tx.sec });

  addFormula(sl, "f09-partial-prod", CY + 1.35, 0.4);

  addCard(sl, M, CY + 1.95, CW, 2.1, T.ac.pos,
    "Two-Sumcheck Architecture",
    "1st sumcheck: verify the degree-(\u2113+1) relation over the boolean hypercube.\n2nd sumcheck: verify the partial product evaluations using preprocessed eq polynomials.\nAll polynomials are multilinear and of size O(n) \u2014 compatible with any PCS.");
}

// ============================================================
// SLIDE 9: First Sumcheck
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "MulPerm: First Sumcheck");

  sl.addText("Prove the main permutation relation via a degree-(\u2113+1) sumcheck:", {
    x: M, y: CY, w: CW, h: 0.35,
    fontFace: F.body.face, fontSize: F.body.size, color: T.tx.pri });

  addFormula(sl, "f10-first-sumcheck", CY + 0.45, 0.45);

  const c = cols2();
  addCard(sl, c.left.x, CY + 1.1, c.left.w, 3.0, T.ac.pos,
    "How It Works",
    "Evaluate f and partial products p over {0,1}^\u03BC.\nRound polynomial: degree \u2113+1.\nAfter \u03BC rounds: reduces to point evals at random \u03B1.");

  addCard(sl, c.right.x, CY + 1.1, c.right.w, 3.0, T.ac.neg,
    "Prover Cost Analysis",
    "Na\u00EFve: O(n\u00B7\u2113\u00B7\u03BC) total.\nBut partial products map {0,1}\u2192{0,1}: only 4 possible functions.\nEnables bucketing for the second sumcheck.");
}

// ============================================================
// SLIDE 10: Second Sumcheck & Bucketing
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "Second Sumcheck & Bucketing Trick");

  addFormula(sl, "f11-second-sumcheck", CY + 0.05, 0.4);

  const c = cols2();
  addCard(sl, c.left.x, CY + 0.55, c.left.w, 2.2, T.ac.pos,
    "Bucketing Algorithm",
    "Round polynomial takes finitely many distinct forms.\nGroup evaluations into buckets by signature.\nEarly rounds: few buckets \u2192 fast.\nLate rounds: switch to direct evaluation.");

  addCard(sl, c.right.x, CY + 0.55, c.right.w, 2.2, T.ac.neg,
    "Bounded Function Space",
    "eq maps {0,1}\u2192{0,1}: only 4^(\u03BC/\u2113) possible round polynomial forms.\nFor \u2113>2: sublinear bucket count.\nEarly: O(n) via bucket aggregation.\nLate: direct O(n/2^k) evaluation.");

  addCard(sl, M, CY + 2.95, CW, 1.15, T.ac.emp,
    "Optimal Switch Point",
    "Switch at round k' = log \u221A(log n). Balances bucket count growth (doubly exponential) against direct cost reduction. Total: n\u00B7\u00D5(\u221Alog n) field operations.");
}

// ============================================================
// SLIDE 11: Total Cost & Parameter Setting (stat-callout)
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "MulPerm: Putting It Together");

  addFormula(sl, "f12-total-cost", CY + 0.05, 0.45);

  const c = cols2();
  // Left: stat callout
  sl.addText("n\u00B7\u00D5(\u221Alog n)", { x: c.left.x, y: CY + 0.6, w: c.left.w, h: 0.8,
    fontFace: F.stat.face, fontSize: 44, bold: true,
    color: T.ac.pos, align: "center", valign: "middle", margin: 0 });
  sl.addText("field operations (any PCS)", { x: c.left.x, y: CY + 1.35, w: c.left.w, h: 0.35,
    fontFace: F.body.face, fontSize: F.body.size,
    color: T.tx.sec, align: "center", margin: 0 });
  addFormulaAt(sl, "f13-set-ell", c.left.x, CY + 1.8, c.left.w, 0.35);

  // Right: comparison card
  addCard(sl, c.right.x, CY + 0.6, c.right.w, 1.6, T.ac.emp,
    "MulPerm vs. Prior Work",
    "vs. HyperPlonk sumcheck: prover speedup.\nvs. GKR: log\u00B2 n \u2192 2 log n rounds.\nvs. ProdCheck: n/|F| \u2192 polylog(n)/|F|.\nNo extra commitments beyond witness.");

  // Bottom: soundness
  addFormula(sl, "f14-soundness", CY + 2.45, 0.35);
  addFormula(sl, "f20-soundness-improve", CY + 2.9, 0.35);
  sl.addText("Soundness: polylogarithmic, not linear. For n=2\u00B3\u00B2: 2\u207B\u2079\u2076 \u2192 2\u207B\u00B9\u00B2\u2070", {
    x: M, y: CY + 3.35, w: CW, h: 0.3,
    fontFace: F.body.face, fontSize: F.body.size, bold: true, color: T.ac.pos, align: "center" });
}

// ============================================================
// SECTION: Applications & Extensions
// ============================================================
sectionSlide("Applications & Extensions");

// ============================================================
// SLIDE 12: Comparison Table
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "Comparison of Permutation Check PIOPs");

  addTable(sl,
    ["PIOP", "|Comm.|", "Rounds", "P field ops", "V time", "Soundness"],
    [
      ["ProdCheck",  "n+log n", "log n",     "n",               "log n",     "n/|F|"],
      ["GKR-based",  "log\u00B2 n", "log\u00B2 n",   "n",               "log\u00B2 n",   "n/|F|"],
      ["HyperPlonk", "log\u00B2 n", "2 log n",   "n\u00B7\u00D5(log n)",     "log n",     "log\u00B2 n/|F|"],
      ["BiPerm",     "log n",   "log n",     "n",               "log n",     "log n/|F|"],
      ["MulPerm",    "log\u00B9\u02D9\u2075 n","2 log n", "n\u00B7\u00D5(\u221Alog n)", "2 log n",   "log\u00B9\u02D9\u2075 n/|F|"],
    ],
    M, CY + 0.1, CW, [1.5, 1.3, 1.2, 1.8, 1.2, 2.0]);

  addCard(sl, M, CY + 2.3, CW, 1.8, T.ac.pos,
    "Key Advantages",
    "BiPerm: linear prover + log soundness (sparse PCS).\nMulPerm: near-linear prover, any PCS, polylog soundness.\nBoth eliminate n extra commitments.");
}

// ============================================================
// SLIDE 13: Applications — HyperPlonk & Spartan
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "Applications: HyperPlonk & Spartan");
  const c = cols2();

  addCard(sl, c.left.x, CY, c.left.w, 2.7, T.ac.pos,
    "Improving HyperPlonk",
    "Replace perm check with BiPerm/MulPerm.\nSingle witness commitment, no extra oracles.\nBatch gate+perm sumchecks.\nCommitment: 2nF \u2192 |w|.");

  addCard(sl, c.right.x, CY, c.right.w, 2.7, T.ac.neg,
    "Improving Spartan (SPARK)",
    "SPARK: row/col/val polynomials for sparse matrices.\nReplace GKR memory check with our lookup.\nImproved soundness, verifier time, proof size.\nNo super-constant round GKR needed.");

  sl.addText("SPARK sparse polynomial encoding:", {
    x: M, y: CY + 2.8, w: CW, h: 0.3,
    fontFace: F.cardTitle.face, fontSize: F.cardTitle.size, bold: true, color: T.tx.pri, margin: [0, 0, 0, 4], shrinkText: true });

  addFormula(sl, "f23-sparse-poly", CY + 3.15, 0.45);

  sl.addText("Structure matches our lookup argument \u2014 row, col are maps between boolean hypercubes", {
    x: M, y: CY + 3.65, w: CW, h: 0.25,
    fontFace: F.caption.face, fontSize: F.caption.size, color: T.tx.mut, shrinkText: true });
}

// ============================================================
// SLIDE 14: Lookup Extension
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "Generalization to Lookup Arguments");

  sl.addText("Replace injective permutation with arbitrary map to a table:", {
    x: M, y: CY, w: CW, h: 0.35,
    fontFace: F.body.face, fontSize: F.body.size, color: T.tx.pri });

  addFormula(sl, "f17-lookup-def", CY + 0.45, 0.4);

  const c = cols2();
  addCard(sl, c.left.x, CY + 1.0, c.left.w, 1.5, T.ac.pos,
    "MulLookup",
    "First lookup with polylog soundness.\nProver: n\u00B7\u00D5(\u221Alog T) for T<n.\nOnly commits the map, no extra data.");

  addCard(sl, c.right.x, CY + 1.0, c.right.w, 1.5, T.ac.neg,
    "Prover-Provided Maps",
    "Prove map is well-formed.\nBinary check: output bits \u2208 {0,1}.\nPermutation: commit \u03C4, verify \u03C3(\u03C4(x))=x.");

  addCard(sl, M, CY + 2.7, CW, 1.4, T.ac.emp,
    "Prover Cost & Table Restrictions",
    "T<n: n\u00B7\u00D5(\u221Alog T) ops. n<T<n^{log n}: O(n(log T\u2212log n)).\nTable poly evaluated once at random point.\nFirst lookup argument with polylog soundness error.");
}

// ============================================================
// SLIDE 15: Key Takeaways (3 numbered cards)
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "Key Takeaways");

  const gap = 0.25, cardW = (CW - 2 * gap) / 3, cardH = 2.6;
  const cardY = CY + 0.15;
  const colors = [T.ac.pos, T.ac.emp, T.ac.neg];
  const items = [
    { t: "BiPerm", b: "Strictly linear prover time.\nSingle degree-3 sumcheck.\nlog n verifier, log n/|F| soundness.\nRequires sparse-friendly PCS." },
    { t: "MulPerm", b: "n\u00B7\u00D5(\u221Alog n) prover, any PCS.\nTwo sumcheck protocols.\nPolylog soundness: \u03BC\u00B3\u02C0\u00B2/|F|.\nNo extra commitments." },
    { t: "Lookups & Apps", b: "First polylog-sound lookup.\nImproves HyperPlonk, Spartan.\nR1CS-style GKR for layered circuits.\nDrop-in replacement." },
  ];
  items.forEach((item, i) => {
    addNumCard(sl, i + 1, item.t, item.b, M + i * (cardW + gap), cardY, cardW, cardH, colors[i]);
  });

  sl.addText("Single witness commitment \u2022 No GKR needed \u2022 Polylog soundness \u2022 Drop-in replacement", {
    x: M, y: CY + 3.0, w: CW, h: 0.5,
    fontFace: F.body.face, fontSize: F.body.size, bold: true,
    color: T.ac.pos, align: "center", valign: "middle" });
}

// ============================================================
// SLIDE 16: References
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "References");
  const refs = [
    "B\u00FCnz, Chen, DeStefano \u2014 Linear-Time Permutation Check. 2025.",
    "[CBBZ23] Chen, Bowe, Bünz, Zhao \u2014 HyperPlonk. EUROCRYPT 2023.",
    "[Set20] Setty \u2014 Spartan: Efficient and General-Purpose zkSNARKs. CRYPTO 2020.",
    "[STW24] Setty, Thaler, Wahby \u2014 Lasso: Lookup Arguments. EUROCRYPT 2024.",
    "[ST25] Setty, Thaler \u2014 Shout: Efficient Range-Check Protocol. ePrint 2024.",
    "[GKR08] Goldwasser, Kalai, Rothblum \u2014 Delegating Computation. STOC 2008.",
  ];
  const refH = CH / refs.length;
  refs.forEach((r, i) => {
    sl.addText(r, { x: M, y: CY + i * refH, w: CW, h: refH,
      fontFace: F.body.face, fontSize: F.body.size, color: T.tx.pri,
      valign: "middle", margin: [0, 0, 0, 8] });
  });
}

// ============================================================
// SLIDE 17: Thank You
// ============================================================
{
  const sl = pres.addSlide({ masterName: "THANK_YOU" });
  sl.addText("Thank You", { x: M, y: SH / 2 - 0.8, w: CW, h: 1,
    fontFace: F.cover.face, fontSize: F.cover.size, bold: true,
    color: T.tx.wh, align: "center", valign: "middle" });
  sl.addText("Questions?", { x: M, y: SH / 2 + 0.3, w: CW, h: 0.5,
    fontFace: F.subtitle.face, fontSize: F.subtitle.size,
    color: T.tx.wh, align: "center" });
}

// ============================================================
// APPENDIX
// ============================================================
sectionSlide("Appendix / Backup Slides");

// ============================================================
// BACKUP 1: Prover-Provided Permutation
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "Backup: Prover-Provided Permutation");
  const c = cols2();

  addCard(sl, c.left.x, CY, c.left.w, 1.8, T.ac.pos,
    "Binary Map Check",
    "Prove each output bit of \u03C3 is in {0,1}.\nUse sumcheck to verify the constraint:");

  addFormulaAt(sl, "f16-binmap", c.left.x + 0.15, CY + 1.95, c.left.w - 0.3, 0.35);

  addCard(sl, c.right.x, CY, c.right.w, 1.8, T.ac.neg,
    "Inverse Trick",
    "To prove \u03C3 is a permutation (bijective):\nCommit to inverse \u03C4.\nRun permutation check: \u03C3(\u03C4(x)) = x.");

  addCard(sl, M, CY + 2.5, CW, 1.65, T.ac.emp,
    "Batching Trick",
    "Combine f=g and permutation check into a single sumcheck via random linear combination R:");

  addFormula(sl, "f15-batch-trick", CY + 3.6, 0.3);
}

// ============================================================
// BACKUP 2: Soundness & eq Polynomial
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "Backup: Soundness & Formal Statement");

  sl.addText("The eq polynomial (building block of all indicator functions):", {
    x: M, y: CY, w: CW, h: 0.35,
    fontFace: F.body.face, fontSize: F.body.size, color: T.tx.pri });

  addFormula(sl, "f04-eq-poly", CY + 0.45, 0.45);

  sl.addText("HyperPlonk arithmetization (degree-\u03BC indicator):", {
    x: M, y: CY + 1.05, w: CW, h: 0.3,
    fontFace: F.body.face, fontSize: F.body.size, color: T.tx.sec });

  addFormula(sl, "f03-degree-mu", CY + 1.45, 0.45);

  sl.addText("Formal claim (Lemma 4):", {
    x: M, y: CY + 2.05, w: CW, h: 0.3,
    fontFace: F.body.face, fontSize: F.body.size, bold: true, color: T.tx.pri });

  addFormula(sl, "f22-lemma4-formal", CY + 2.45, 0.45);

  addCard(sl, M, CY + 3.05, CW, 1.1, T.ac.pos,
    "Soundness Bound",
    "BiPerm: O(\u03BC/|F|). MulPerm: O(\u03BC\u00B3\u02C0\u00B2/|F|).\nBoth polylog vs linear n/|F|. n=2\u00B3\u00B2: 2\u207B\u2079\u2076 \u2192 2\u207B\u00B9\u00B2\u2070.");
}

// ============================================================
// BACKUP 3: Degree Tradeoff Detail
// ============================================================
{
  const sl = pres.addSlide({ masterName: "CONTENT_SLIDE" });
  sTitle(sl, "Backup: Full Degree\u2013Complexity Landscape");

  addFormula(sl, "f21-degree-tradeoff", CY + 0.1, 0.5);

  addTable(sl,
    ["Protocol", "Degree", "Index Size", "P field ops", "Soundness"],
    [
      ["Na\u00EFve (full MLE)", "1", "n\u00B2", "n", "log n/|F|"],
      ["BiPerm", "2", "n\u00B9\u02D9\u2075 (sparse)", "n", "\u03BC/|F|"],
      ["MulPerm (\u2113=\u221A\u03BC)", "\u221A\u03BC", "n (dense)", "n\u00B7\u00D5(\u221A\u03BC)", "\u03BC\u00B3\u02C0\u00B2/|F|"],
      ["HyperPlonk", "\u03BC", "n (dense)", "n\u00B7\u03BC\u00B2", "\u03BC\u00B2/|F|"],
    ],
    M, CY + 0.75, CW, [2.0, 1.2, 2.0, 2.0, 1.8]);

  addCard(sl, M, CY + 2.55, CW, 1.6, T.ac.emp,
    "Design Space Navigation",
    "Left\u2192right: degree up, index shrinks, prover cost grows.\nBiPerm & MulPerm: optimal tradeoff zone.\nFirst to break linear soundness barrier at near-linear prover cost.");
}

// ============================================================
pres.writeFile({ fileName: "linear_perm_slides.pptx" });
