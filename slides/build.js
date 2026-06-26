const pptxgen = require("pptxgenjs");

// ---- Palette (mirrors the Joe Doe project's own style.css) ----
const NAVY   = "2C3E50";  // --primary-color
const SLATE  = "34495E";  // --secondary-color
const CORAL  = "E74C3C";  // --accent-color
const INK    = "243340";  // dark code panel bg
const PAPER  = "FFFFFF";
const MIST   = "F4F6F8";  // light section tint
const BODY   = "44525E";
const MUTE   = "8A98A6";
const LINE   = "DDE3E8";

// code token colors (on dark panel)
const C_SEL  = "F39C12";  // selectors  (amber)
const C_PROP = "7FC7E8";  // properties (sky)
const C_VAL  = "97D085";  // values     (green)
const C_COM  = "7F8C9A";  // comments   (gray)
const C_PUN  = "ECF0F1";  // punctuation/plain (near-white)

const HEAD = "Cambria";
const SANS = "Calibri";
const MONO = "Courier New";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";       // 13.33 x 7.5
pres.author = "Joe Doe CSS Course";
pres.title  = "Styling the Web with CSS";
const W = 13.33, H = 7.5;

// ---------- helpers ----------
function shadow() {
  return { type: "outer", color: "000000", blur: 7, offset: 3, angle: 90, opacity: 0.18 };
}

// turn a CSS source line into colored rich-text runs (leading spaces preserved)
function cssRuns(line) {
  const indent = line.match(/^\s*/)[0];
  const t = line.trim();
  const lead = indent.length ? [{ text: indent.replace(/ /g, " ") }] : [];
  let runs;
  if (t === "") return [{ text: " " }];
  if (t.startsWith("/*") || t.startsWith("*") || t.startsWith("@import")) {
    runs = [{ text: t, options: { color: C_COM, italic: true } }];
  } else if (t.startsWith("@media")) {
    runs = [{ text: t, options: { color: C_SEL, bold: true } }];
  } else if (t === "}" || t === "{") {
    runs = [{ text: t, options: { color: C_PUN } }];
  } else if (t.endsWith("{")) {
    runs = [{ text: t.replace(/\s*\{$/, ""), options: { color: C_SEL, bold: true } },
            { text: " {", options: { color: C_PUN } }];
  } else if (/:/.test(t) && t.endsWith(";")) {
    const i = t.indexOf(":");
    runs = [{ text: t.slice(0, i), options: { color: C_PROP } },
            { text: ": ", options: { color: C_PUN } },
            { text: t.slice(i + 1).replace(/;$/, "").trim(), options: { color: C_VAL } },
            { text: ";", options: { color: C_PUN } }];
  } else {
    runs = [{ text: t, options: { color: C_PUN } }];
  }
  return [...lead, ...runs];
}

// dark code panel with simple CSS syntax coloring
function codePanel(slide, x, y, w, h, lines, fs = 13) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x, y, w, h, rectRadius: 0.08, fill: { color: INK }, line: { type: "none" },
    shadow: shadow(),
  });
  // little window dots
  ["E74C3C", "F1C40F", "2ECC71"].forEach((c, i) =>
    slide.addShape(pres.shapes.OVAL, { x: x + 0.22 + i * 0.22, y: y + 0.18, w: 0.11, h: 0.11, fill: { color: c }, line: { type: "none" } }));
  const runs = [];
  lines.forEach((ln, idx) => {
    const r = cssRuns(ln);
    r[r.length - 1].options = { ...(r[r.length - 1].options || {}), breakLine: true };
    if (idx === lines.length - 1) delete r[r.length - 1].options.breakLine;
    runs.push(...r);
  });
  slide.addText(runs, {
    x: x + 0.25, y: y + 0.45, w: w - 0.5, h: h - 0.6,
    fontFace: MONO, fontSize: fs, color: C_PUN, align: "left", valign: "top",
    lineSpacingMultiple: 1.12, margin: 0,
  });
}

// repeating motif: coral number chip + title
function header(slide, n, title, sub, dark) {
  slide.addShape(pres.shapes.OVAL, { x: 0.6, y: 0.5, w: 0.62, h: 0.62, fill: { color: CORAL }, line: { type: "none" }, shadow: shadow() });
  slide.addText(String(n).padStart(2, "0"), { x: 0.6, y: 0.5, w: 0.62, h: 0.62, align: "center", valign: "middle", fontFace: HEAD, fontSize: 18, bold: true, color: "FFFFFF", margin: 0 });
  slide.addText(title, { x: 1.4, y: 0.45, w: 11.3, h: 0.55, fontFace: HEAD, fontSize: 30, bold: true, color: dark ? "FFFFFF" : NAVY, margin: 0, valign: "middle" });
  if (sub) slide.addText(sub, { x: 1.42, y: 1.02, w: 11.3, h: 0.4, fontFace: SANS, fontSize: 14, italic: true, color: dark ? "CAD6DF" : MUTE, margin: 0 });
}

function bullets(slide, x, y, w, items, fs = 15) {
  slide.addText(items.map((it, i) => {
    const o = { bullet: { code: "2022", indent: 14 }, color: BODY, breakLine: true, paraSpaceAfter: 8 };
    if (i === items.length - 1) o.breakLine = false;
    if (Array.isArray(it)) return { text: it[0], options: { ...o, color: it[1] || BODY, bold: it[2] || false } };
    return { text: it, options: o };
  }), { x, y, w, h: 4.5, fontFace: SANS, fontSize: fs, align: "left", valign: "top", lineSpacingMultiple: 1.05, margin: 0 });
}

function card(slide, x, y, w, h, fill) {
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill || MIST }, line: { color: LINE, width: 1 }, shadow: shadow() });
}

function footer(slide, idx) {
  slide.addText("Joe Doe Portfolio  ·  Learning CSS", { x: 0.6, y: H - 0.42, w: 7, h: 0.3, fontFace: SANS, fontSize: 9, color: MUTE, margin: 0 });
  slide.addText(String(idx), { x: W - 1.1, y: H - 0.42, w: 0.5, h: 0.3, fontFace: SANS, fontSize: 9, color: MUTE, align: "right", margin: 0 });
}

let N = 0;
function content(dark) {
  const s = pres.addSlide();
  s.background = { color: dark ? NAVY : PAPER };
  N++;
  return s;
}

// ============================================================
// 1 — TITLE
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: 10.4, y: -1.6, w: 4.6, h: 4.6, fill: { color: SLATE }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: 11.6, y: 4.7, w: 3.2, h: 3.2, fill: { color: CORAL }, line: { type: "none" } });
  s.addText("PROJECT-BASED LEARNING", { x: 0.9, y: 1.5, w: 9, h: 0.4, fontFace: SANS, fontSize: 15, bold: true, color: CORAL, charSpacing: 4, margin: 0 });
  s.addText("Styling the Web with CSS", { x: 0.9, y: 2.0, w: 10.5, h: 1.3, fontFace: HEAD, fontSize: 50, bold: true, color: "FFFFFF", margin: 0 });
  s.addText("Turning the Joe Doe portfolio from plain HTML into a designed website — one CSS concept at a time.",
    { x: 0.92, y: 3.45, w: 9.2, h: 1.0, fontFace: SANS, fontSize: 19, color: "CAD6DF", margin: 0, lineSpacingMultiple: 1.1 });
  // mini code chip
  codePanel(s, 0.9, 4.85, 5.7, 1.55, ["nav a {", "  color: var(--primary-color);", "  text-decoration: none;", "}"], 13);
  s.addText("A 4-page CV site · HTML structure ➜ CSS presentation", { x: 7.0, y: 6.05, w: 5.6, h: 0.4, fontFace: SANS, fontSize: 13, italic: true, color: "9FB0BD", align: "right", margin: 0 });
  s.addNotes("Welcome. Over this session we style a real project — Joe Doe's CV site. Version 1 was plain HTML; today we add the CSS layer. Every concept maps to a line you'll actually write in style.css.");
}

// ============================================================
// 2 — THE PROJECT (before / after)
// ============================================================
{
  const s = content(false);
  header(s, 1, "Meet the project: the Joe Doe portfolio", "Same HTML, two very different experiences — that difference is CSS.");
  // before card
  card(s, 0.6, 1.75, 5.85, 4.9, MIST);
  s.addText("BEFORE  ·  HTML only", { x: 0.9, y: 1.95, w: 5.2, h: 0.4, fontFace: SANS, fontSize: 14, bold: true, color: MUTE, charSpacing: 2, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.9, y: 2.5, w: 5.25, h: 3.9, fill: { color: "FFFFFF" }, line: { color: LINE, width: 1 } });
  s.addText([
    { text: "Joe Doe\n", options: { fontSize: 22, bold: true, color: "000000", breakLine: true } },
    { text: "Front-End Developer\n\n", options: { fontSize: 12, color: "000000", breakLine: true } },
    { text: "Home  About  Experience  Contact\n\n", options: { fontSize: 12, color: "1A0DAB", underline: true, breakLine: true } },
    { text: "Welcome\n", options: { fontSize: 16, bold: true, color: "000000", breakLine: true } },
    { text: "I'm building my CV in HTML.", options: { fontSize: 12, color: "000000" } },
  ], { x: 1.05, y: 2.65, w: 4.95, h: 3.6, fontFace: "Times New Roman", align: "left", valign: "top", margin: 4 });

  // after card
  card(s, 6.9, 1.75, 5.85, 4.9, MIST);
  s.addText("AFTER  ·  HTML + CSS", { x: 7.2, y: 1.95, w: 5.2, h: 0.4, fontFace: SANS, fontSize: 14, bold: true, color: CORAL, charSpacing: 2, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 7.2, y: 2.5, w: 5.25, h: 3.9, fill: { color: PAPER }, line: { type: "none" } });
  // styled header band
  s.addShape(pres.shapes.OVAL, { x: 7.55, y: 2.75, w: 0.85, h: 0.85, fill: { color: SLATE }, line: { type: "none" } });
  s.addText("JD", { x: 7.55, y: 2.75, w: 0.85, h: 0.85, align: "center", valign: "middle", fontFace: HEAD, fontSize: 18, bold: true, color: "FFFFFF", margin: 0 });
  s.addText([{ text: "Joe Doe\n", options: { fontFace: HEAD, fontSize: 22, bold: true, color: NAVY, breakLine: true } },
             { text: "Front-End Developer", options: { fontFace: SANS, fontSize: 12, color: MUTE } }],
    { x: 8.55, y: 2.78, w: 3.7, h: 0.85, valign: "middle", margin: 0 });
  ["Home", "About", "Experience", "Contact"].forEach((t, i) =>
    s.addText(t, { x: 7.55 + i * 1.18, y: 3.78, w: 1.1, h: 0.3, fontFace: SANS, fontSize: 11, bold: i === 0, color: i === 0 ? CORAL : NAVY, margin: 0 }));
  // card
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.55, y: 4.25, w: 4.55, h: 1.95, rectRadius: 0.06, fill: { color: "F9F9F9" }, line: { type: "none" }, shadow: shadow() });
  s.addText([{ text: "Welcome\n", options: { fontFace: HEAD, fontSize: 15, bold: true, color: NAVY, breakLine: true } },
             { text: "I'm building my CV in HTML.", options: { fontFace: SANS, fontSize: 12, color: BODY } }],
    { x: 7.8, y: 4.45, w: 4.1, h: 1.5, valign: "top", margin: 0 });
  footer(s, N);
  s.addNotes("Drive the point home: the markup is identical. CSS adds color, spacing, a card with rounded corners and a shadow, a typeface, and layout. Presentation is a separate layer — that's the whole idea.");
}

// ============================================================
// 3 — WHAT IS CSS / HOW IT CONNECTS
// ============================================================
{
  const s = content(false);
  header(s, 2, "What is CSS, and how does it reach the page?", "Three languages, three jobs — kept in separate files.");
  // three role cards
  const roles = [
    ["HTML", "Structure", "What the content IS:\nheadings, nav, sections, the form.", NAVY],
    ["CSS", "Presentation", "How it LOOKS:\ncolor, spacing, type, layout.", CORAL],
    ["Browser", "Renders", "Combines the two and\npaints pixels on screen.", SLATE],
  ];
  roles.forEach((r, i) => {
    const x = 0.6 + i * 4.25;
    card(s, x, 1.85, 3.95, 2.35, MIST);
    s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: 2.1, w: 0.55, h: 0.55, fill: { color: r[3] }, line: { type: "none" } });
    s.addText(r[0][0], { x: x + 0.3, y: 2.1, w: 0.55, h: 0.55, align: "center", valign: "middle", fontFace: HEAD, bold: true, fontSize: 18, color: "FFFFFF", margin: 0 });
    s.addText([{ text: r[0] + "  ", options: { bold: true, color: NAVY, fontSize: 17 } }, { text: "· " + r[1], options: { color: MUTE, fontSize: 13, italic: true } }],
      { x: x + 1.0, y: 2.12, w: 2.8, h: 0.55, valign: "middle", fontFace: SANS, margin: 0 });
    s.addText(r[2], { x: x + 0.32, y: 2.85, w: 3.35, h: 1.2, fontFace: SANS, fontSize: 13, color: BODY, valign: "top", margin: 0, lineSpacingMultiple: 1.05 });
  });
  // the link line
  s.addText("CSS is attached to a page with one line in the HTML <head>:", { x: 0.6, y: 4.55, w: 12, h: 0.4, fontFace: SANS, fontSize: 15, color: BODY, margin: 0 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.0, w: 12.1, h: 1.0, rectRadius: 0.08, fill: { color: INK }, line: { type: "none" }, shadow: shadow() });
  s.addText([
    { text: "<link ", options: { color: C_PROP } },
    { text: "rel", options: { color: C_SEL } }, { text: "=", options: { color: C_PUN } }, { text: '"stylesheet" ', options: { color: C_VAL } },
    { text: "href", options: { color: C_SEL } }, { text: "=", options: { color: C_PUN } }, { text: '"css/style.css"', options: { color: C_VAL } },
    { text: ">", options: { color: C_PROP } },
  ], { x: 0.9, y: 5.0, w: 11.5, h: 1.0, fontFace: MONO, fontSize: 17, valign: "middle", margin: 0 });
  footer(s, N);
  s.addNotes("Separation of concerns. The browser fetches style.css because of this <link> in the head. Every Joe Doe page includes the same line, so one stylesheet themes the whole site. Point out: change the file once, every page updates.");
}

// ============================================================
// 4 — ANATOMY OF A RULE
// ============================================================
{
  const s = content(false);
  header(s, 3, "Anatomy of a CSS rule", "Every rule answers two questions: WHO do I style, and HOW?");
  codePanel(s, 0.6, 1.9, 6.4, 2.5, [
    "nav a {",
    "  color: var(--primary-color);",
    "  text-decoration: none;",
    "  margin-right: 1rem;",
    "}",
  ], 16);
  // labels pointing at parts
  const lab = [
    ["SELECTOR", "Who to style — here, links inside the nav.", CORAL, 1.95],
    ["PROPERTY", "Which aspect to change (color, margin…).", C_PROP === "7FC7E8" ? SLATE : SLATE, 3.0],
    ["VALUE", "The setting you give that property.", "2E8B57", 4.05],
  ];
  card(s, 7.3, 1.9, 5.45, 4.0, MIST);
  s.addText("A declaration = property : value ;", { x: 7.55, y: 2.1, w: 5, h: 0.4, fontFace: SANS, fontSize: 14, bold: true, color: NAVY, margin: 0 });
  s.addText([
    { text: "Selector", options: { bold: true, color: CORAL, breakLine: true, paraSpaceAfter: 3 } },
    { text: "The pattern that picks elements. nav a means “every <a> inside <nav>.”\n\n", options: { color: BODY, fontSize: 13, breakLine: true } },
    { text: "Declaration block { … }", options: { bold: true, color: SLATE, breakLine: true, paraSpaceAfter: 3 } },
    { text: "The curly braces hold one or more declarations.\n\n", options: { color: BODY, fontSize: 13, breakLine: true } },
    { text: "Declaration", options: { bold: true, color: "2E8B57", breakLine: true, paraSpaceAfter: 3 } },
    { text: "property: value;  — always ends with a semicolon.", options: { color: BODY, fontSize: 13 } },
  ], { x: 7.55, y: 2.55, w: 5.0, h: 3.2, fontFace: SANS, fontSize: 14, valign: "top", margin: 0, lineSpacingMultiple: 1.04 });
  s.addText("Read it aloud: “Style every link in the nav — make it navy, remove the underline, add right margin.”",
    { x: 0.6, y: 4.7, w: 6.5, h: 1.1, fontFace: SANS, fontSize: 14, italic: true, color: SLATE, valign: "top", margin: 0, lineSpacingMultiple: 1.1 });
  footer(s, N);
  s.addNotes("Slow down here — this is the grammar of all CSS. Selector, then a block of property:value declarations. Stress the semicolon and the colon. Everything else in the course is variations on this shape.");
}

// ============================================================
// 5 — SELECTORS
// ============================================================
{
  const s = content(false);
  header(s, 4, "Selectors: choosing what to style", "The Joe Doe stylesheet uses four kinds you'll meet constantly.");
  const rows = [
    ["Type / element", "body, h2, section", "Matches every element of that tag.", NAVY],
    ["Descendant", "nav a", "An <a> that sits inside a <nav>.", SLATE],
    ["Pseudo-class", "nav a:hover", "A state — here, while the mouse is over it.", CORAL],
    ["Class (you'll add)", ".card", "Any element with class=\"card\". Reusable.", "2E8B57"],
  ];
  rows.forEach((r, i) => {
    const y = 1.85 + i * 1.18;
    card(s, 0.6, y, 7.0, 1.02, i % 2 ? MIST : PAPER);
    s.addShape(pres.shapes.OVAL, { x: 0.85, y: y + 0.27, w: 0.48, h: 0.48, fill: { color: r[3] }, line: { type: "none" } });
    s.addText(r[0], { x: 1.55, y: y + 0.12, w: 5.9, h: 0.4, fontFace: SANS, fontSize: 15, bold: true, color: NAVY, margin: 0 });
    s.addText(r[2], { x: 1.55, y: y + 0.52, w: 5.9, h: 0.4, fontFace: SANS, fontSize: 12, color: BODY, margin: 0 });
    s.addText(r[1], { x: 1.55, y: y + 0.12, w: 5.9, h: 0.4, fontFace: MONO, fontSize: 13, color: CORAL, align: "right", margin: 0 });
  });
  codePanel(s, 7.9, 1.85, 4.85, 4.0, [
    "/* element */",
    "body { line-height: 1.6; }",
    "",
    "/* descendant */",
    "nav a { color: var(",
    "  --primary-color); }",
    "",
    "/* pseudo-class */",
    "nav a:hover {",
    "  color: var(--accent-color);",
    "}",
  ], 13);
  footer(s, N);
  s.addNotes("Walk down the table, matching each to the live code panel. The :hover rule is a great live demo — hover a nav link in the browser and watch it turn coral. Mention classes are what they'll add when components repeat.");
}

// ============================================================
// 6 — CASCADE & SPECIFICITY
// ============================================================
{
  const s = content(false);
  header(s, 5, "The cascade: who wins when rules collide?", "“Cascading” is in the name for a reason. Three tie-breakers, in order.");
  const steps = [
    ["1", "Importance & origin", "Author styles beat browser defaults. !important overrides — use sparingly."],
    ["2", "Specificity", "More specific selectors win. id > class > element. nav a:hover beats a."],
    ["3", "Source order", "All else equal, the LAST rule written wins. Order matters."],
  ];
  steps.forEach((st, i) => {
    const x = 0.6 + i * 4.25;
    card(s, x, 1.95, 3.95, 2.2, MIST);
    s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: 2.2, w: 0.6, h: 0.6, fill: { color: CORAL }, line: { type: "none" } });
    s.addText(st[0], { x: x + 0.3, y: 2.2, w: 0.6, h: 0.6, align: "center", valign: "middle", fontFace: HEAD, bold: true, fontSize: 22, color: "FFFFFF", margin: 0 });
    s.addText(st[1], { x: x + 1.05, y: 2.25, w: 2.8, h: 0.55, fontFace: SANS, fontSize: 15, bold: true, color: NAVY, valign: "middle", margin: 0 });
    s.addText(st[2], { x: x + 0.32, y: 2.95, w: 3.4, h: 1.1, fontFace: SANS, fontSize: 13, color: BODY, valign: "top", margin: 0, lineSpacingMultiple: 1.05 });
  });
  card(s, 0.6, 4.5, 12.15, 1.85, "FFF5F3");
  s.addText("Example from our site", { x: 0.85, y: 4.65, w: 5, h: 0.35, fontFace: SANS, fontSize: 13, bold: true, color: CORAL, margin: 0 });
  s.addText([
    { text: "nav a", options: { fontFace: MONO, color: NAVY } },
    { text: " sets links navy. ", options: { color: BODY } },
    { text: "nav a:hover", options: { fontFace: MONO, color: NAVY } },
    { text: " is more specific, so on hover the coral wins — but only while hovering. No conflict, just the cascade doing its job.", options: { color: BODY } },
  ], { x: 0.85, y: 5.05, w: 11.6, h: 1.1, fontFace: SANS, fontSize: 15, valign: "top", margin: 0, lineSpacingMultiple: 1.1 });
  footer(s, N);
  s.addNotes("The cascade is the concept students fight most. Keep it concrete: same property, two rules, the more specific or later one wins. Discourage !important early — it papers over specificity problems.");
}

// ============================================================
// 7 — COLORS & CUSTOM PROPERTIES
// ============================================================
{
  const s = content(false);
  header(s, 6, "Colors & custom properties (variables)", "Define the brand once in :root, reuse it everywhere with var().");
  codePanel(s, 0.6, 1.9, 6.5, 3.5, [
    ":root {",
    "  --primary-color: #2c3e50;",
    "  --secondary-color: #34495e;",
    "  --accent-color: #e74c3c;",
    "  --text-color: #333333;",
    "}",
    "",
    "nav a { color: var(--primary-color); }",
    "nav a:hover { color: var(--accent-color); }",
  ], 14);
  // swatches
  const sw = [["--primary-color", "2C3E50"], ["--secondary-color", "34495E"], ["--accent-color", "E74C3C"], ["--card-background", "F9F9F9"], ["--text-color", "333333"]];
  s.addText("Joe Doe's palette", { x: 7.4, y: 1.95, w: 5, h: 0.4, fontFace: SANS, fontSize: 16, bold: true, color: NAVY, margin: 0 });
  sw.forEach((c, i) => {
    const y = 2.5 + i * 0.62;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.4, y, w: 0.7, h: 0.5, rectRadius: 0.05, fill: { color: c[1] }, line: { color: LINE, width: 1 } });
    s.addText(c[0], { x: 8.25, y, w: 3.3, h: 0.5, fontFace: MONO, fontSize: 13, color: NAVY, valign: "middle", margin: 0 });
    s.addText("#" + c[1].toLowerCase(), { x: 11.4, y, w: 1.3, h: 0.5, fontFace: MONO, fontSize: 12, color: MUTE, valign: "middle", align: "right", margin: 0 });
  });
  card(s, 0.6, 5.6, 12.15, 1.05, "EAF6EF");
  s.addText([
    { text: "Why bother?  ", options: { bold: true, color: "2E8B57" } },
    { text: "Rebrand the whole 4-page site by editing ", options: { color: BODY } },
    { text: "one", options: { bold: true, color: NAVY } },
    { text: " line. Change --accent-color and every link-hover, highlight and button follows.", options: { color: BODY } },
  ], { x: 0.85, y: 5.72, w: 11.6, h: 0.8, fontFace: SANS, fontSize: 15, valign: "middle", margin: 0, lineSpacingMultiple: 1.05 });
  footer(s, N);
  s.addNotes("Custom properties are the DRY principle for design. Demo live: change --accent-color to a different hex and watch every hover state update at once. Note var() with the -- prefix and that :root is the whole document.");
}

// ============================================================
// 8 — THE BOX MODEL
// ============================================================
{
  const s = content(false);
  header(s, 7, "The box model: every element is a box", "Content, then padding, then border, then margin — from the inside out.");
  // nested diagram
  const bx = 0.7, by = 1.95;
  s.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: 6.0, h: 4.3, fill: { color: "FBE3DF" }, line: { color: CORAL, width: 1, dashType: "dash" } });
  s.addText("margin", { x: bx + 0.12, y: by + 0.06, w: 2, h: 0.3, fontFace: SANS, fontSize: 12, bold: true, color: CORAL, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: bx + 0.75, y: by + 0.5, w: 4.5, h: 3.3, fill: { color: "FCEFD6" }, line: { color: "E0A93C", width: 1 } });
  s.addText("border", { x: bx + 0.87, y: by + 0.56, w: 2, h: 0.3, fontFace: SANS, fontSize: 12, bold: true, color: "B5781C", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: bx + 1.45, y: by + 1.0, w: 3.1, h: 2.3, fill: { color: "E5F0E9" }, line: { color: "2E8B57", width: 1 } });
  s.addText("padding", { x: bx + 1.55, y: by + 1.06, w: 2, h: 0.3, fontFace: SANS, fontSize: 12, bold: true, color: "2E8B57", margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: bx + 2.15, y: by + 1.55, w: 1.7, h: 1.2, fill: { color: NAVY }, line: { type: "none" } });
  s.addText("content", { x: bx + 2.15, y: by + 1.55, w: 1.7, h: 1.2, align: "center", valign: "middle", fontFace: SANS, fontSize: 12, bold: true, color: "FFFFFF", margin: 0 });
  // explanation + code
  bullets(s, 7.2, 2.0, 5.6, [
    ["Content — your text or image.", BODY],
    ["Padding — space inside, between content and border.", BODY],
    ["Border — the edge line around the padding.", BODY],
    ["Margin — space pushing other boxes away.", BODY],
  ], 15);
  codePanel(s, 7.2, 4.25, 5.6, 2.05, [
    "* {",
    "  margin: 0;",
    "  padding: 0;",
    "  box-sizing: border-box;",
    "}",
  ], 14);
  s.addText("box-sizing: border-box makes width include padding + border — far easier to reason about.",
    { x: 0.7, y: 6.35, w: 6.0, h: 0.7, fontFace: SANS, fontSize: 12.5, italic: true, color: SLATE, valign: "top", margin: 0, lineSpacingMultiple: 1.05 });
  footer(s, N);
  s.addNotes("The box model trips up everyone. Use the nested diagram: content in the middle, padding hugs it, border wraps that, margin pushes neighbours away. The universal reset zeroes defaults and sets border-box so sizing is predictable.");
}

// ============================================================
// 9 — TYPOGRAPHY
// ============================================================
{
  const s = content(false);
  header(s, 8, "Typography: setting the voice of the site", "Import a webfont, then apply it once on the body — it inherits everywhere.");
  codePanel(s, 0.6, 1.95, 6.7, 3.0, [
    "@import url('...Inter...');",
    "",
    ":root {",
    "  --font-family: 'Inter', sans-serif;",
    "}",
    "",
    "body {",
    "  font-family: var(--font-family);",
    "  line-height: 1.6;",
    "}",
  ], 13.5);
  card(s, 7.6, 1.95, 5.15, 3.0, MIST);
  s.addText("Three ideas at work", { x: 7.85, y: 2.1, w: 4.6, h: 0.4, fontFace: SANS, fontSize: 15, bold: true, color: NAVY, margin: 0 });
  bullets(s, 7.85, 2.6, 4.7, [
    ["@import pulls the Inter webfont from Google Fonts.", BODY],
    ["font-family lists fallbacks — sans-serif if Inter fails.", BODY],
    ["line-height: 1.6 gives readable spacing between lines.", BODY],
    ["Inheritance: set on body, every child uses it.", BODY],
  ], 13.5);
  // type specimen
  card(s, 0.6, 5.15, 12.15, 1.6, PAPER);
  s.addText([
    { text: "Joe Doe  ", options: { fontFace: SANS, fontSize: 30, bold: true, color: NAVY } },
    { text: "Front-End Developer", options: { fontFace: SANS, fontSize: 16, color: MUTE } },
  ], { x: 0.85, y: 5.3, w: 11.6, h: 0.6, valign: "middle", margin: 0 });
  s.addText("I build fast, accessible websites with clean markup — a calm, readable type scale ties the pages together.",
    { x: 0.85, y: 5.95, w: 11.6, h: 0.7, fontFace: SANS, fontSize: 15, color: BODY, valign: "top", margin: 0, lineSpacingMultiple: 1.1 });
  footer(s, N);
  s.addNotes("Typography is 90% of how a site feels. One import, one font-family on body, and inheritance carries it through. line-height is the easiest readability win. Mention the comma-separated fallback stack.");
}

// ============================================================
// 10 — BUILDING THE CARD COMPONENT
// ============================================================
{
  const s = content(false);
  header(s, 9, "Composing a component: the section card", "Four properties combine into the reusable card you see on every page.");
  codePanel(s, 0.6, 1.95, 6.4, 3.7, [
    "section {",
    "  background: var(--card-background);",
    "  padding: 1rem;",
    "  border-radius: 8px;",
    "  box-shadow: 0 2px 4px",
    "              rgba(0,0,0,0.1);",
    "  margin-bottom: 1rem;",
    "}",
  ], 15);
  // live preview of the card
  s.addText("Renders as", { x: 7.4, y: 1.95, w: 5, h: 0.35, fontFace: SANS, fontSize: 13, bold: true, color: MUTE, charSpacing: 2, margin: 0 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.4, y: 2.4, w: 5.35, h: 1.7, rectRadius: 0.09, fill: { color: "F9F9F9" }, line: { type: "none" }, shadow: shadow() });
  s.addText([{ text: "About Me\n", options: { fontFace: HEAD, fontSize: 18, bold: true, color: NAVY, breakLine: true } },
             { text: "I build fast, accessible websites with clean markup.", options: { fontFace: SANS, fontSize: 13, color: BODY } }],
    { x: 7.7, y: 2.65, w: 4.8, h: 1.2, valign: "top", margin: 0, lineSpacingMultiple: 1.05 });
  // mapping list
  const map = [
    ["background", "fills the card so it lifts off the page"],
    ["padding", "breathing room inside the edge"],
    ["border-radius", "softens the corners (8px)"],
    ["box-shadow", "subtle depth — x, y, blur, color"],
  ];
  map.forEach((m, i) => {
    const y = 4.35 + i * 0.6;
    s.addText(m[0], { x: 7.4, y, w: 2.0, h: 0.5, fontFace: MONO, fontSize: 13, color: CORAL, valign: "middle", margin: 0 });
    s.addText(m[1], { x: 9.45, y, w: 3.3, h: 0.5, fontFace: SANS, fontSize: 12.5, color: BODY, valign: "middle", margin: 0 });
  });
  footer(s, N);
  s.addNotes("This is the payoff slide: individual properties students just learned combine into a real component. Because the selector is the section element, every section across all four pages becomes a card for free.");
}

// ============================================================
// 11 — LAYOUT WITH GRID
// ============================================================
{
  const s = content(false);
  header(s, 10, "Layout: arranging boxes with CSS Grid", "Turn <main> into a two-column grid with three lines of CSS.");
  codePanel(s, 0.6, 1.95, 6.3, 2.5, [
    "main {",
    "  display: grid;",
    "  grid-template-columns: repeat(2, 1fr);",
    "  gap: 1rem;",
    "}",
  ], 14);
  // grid diagram
  s.addText("1fr        +        1fr", { x: 7.2, y: 1.95, w: 5.55, h: 0.35, fontFace: MONO, fontSize: 13, color: MUTE, align: "center", margin: 0 });
  const gx = 7.2, gy = 2.4, gw = 2.6, gh = 1.3, gap = 0.35;
  const cells = ["About", "Skills", "Education", "Experience"];
  cells.forEach((c, i) => {
    const cx = gx + (i % 2) * (gw + gap);
    const cy = gy + Math.floor(i / 2) * (gh + gap);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: cy, w: gw, h: gh, rectRadius: 0.06, fill: { color: i % 2 ? MIST : "EAF1F6" }, line: { color: SLATE, width: 1 } });
    s.addText(c, { x: cx, y: cy, w: gw, h: gh, align: "center", valign: "middle", fontFace: SANS, fontSize: 14, bold: true, color: NAVY, margin: 0 });
  });
  card(s, 0.6, 4.75, 12.15, 1.9, "FFF5F3");
  bullets(s, 0.85, 4.9, 11.9, [
    ["display: grid turns <main> into a grid container; its sections become grid items.", BODY],
    ["repeat(2, 1fr) = two equal columns; 1fr means “one fraction of the free space.”", BODY],
    ["gap: 1rem spaces items evenly — no margins to fiddle with.", BODY],
  ], 14);
  footer(s, N);
  s.addNotes("Grid is the modern way to do page layout. fr units are the key idea — fractions of available space, so columns stay equal as the window resizes. gap replaces the old margin juggling. Flexbox is the 1-D cousin; mention briefly.");
}

// ============================================================
// 12 — RESPONSIVE / MEDIA QUERIES
// ============================================================
{
  const s = content(false);
  header(s, 11, "Responsive design: mobile-first media queries", "Style for small screens first, then enhance when there's room.");
  // phone frame
  s.addText("Phone  ·  one column", { x: 0.6, y: 1.95, w: 3, h: 0.35, fontFace: SANS, fontSize: 13, bold: true, color: MUTE, margin: 0 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 2.4, w: 1.9, h: 3.9, rectRadius: 0.12, fill: { color: PAPER }, line: { color: NAVY, width: 2 } });
  ["EAF1F6", "MIST", "EAF1F6"].forEach((_, i) =>
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.85, y: 2.7 + i * 1.1, w: 1.4, h: 0.9, rectRadius: 0.05, fill: { color: i % 2 ? MIST : "EAF1F6" }, line: { color: LINE, width: 1 } }));
  // desktop frame
  s.addText("Desktop  ·  two columns (≥ 640px)", { x: 3.05, y: 1.95, w: 4.5, h: 0.35, fontFace: SANS, fontSize: 13, bold: true, color: CORAL, margin: 0 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 3.05, y: 2.4, w: 4.45, h: 3.9, rectRadius: 0.08, fill: { color: PAPER }, line: { color: NAVY, width: 2 } });
  for (let i = 0; i < 4; i++) {
    const cx = 3.3 + (i % 2) * 2.05, cy = 2.7 + Math.floor(i / 2) * 1.6;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: cy, w: 1.85, h: 1.35, rectRadius: 0.05, fill: { color: i % 2 ? MIST : "EAF1F6" }, line: { color: LINE, width: 1 } });
  }
  codePanel(s, 7.85, 2.3, 4.9, 2.7, [
    "/* base: mobile, 1 col */",
    "main { }",
    "",
    "/* wider screens only */",
    "@media (min-width: 640px) {",
    "  main {",
    "    display: grid;",
    "    grid-template-columns:",
    "      repeat(2, 1fr);",
    "  }",
    "}",
  ], 12.5);
  card(s, 7.85, 5.2, 4.9, 1.45, MIST);
  s.addText([
    { text: "Mobile-first  ", options: { bold: true, color: CORAL } },
    { text: "means the simple layout is the default; the ", options: { color: BODY } },
    { text: "@media", options: { fontFace: MONO, color: NAVY } },
    { text: " block adds the grid only when the screen is at least 640px wide.", options: { color: BODY } },
  ], { x: 8.1, y: 5.35, w: 4.45, h: 1.15, fontFace: SANS, fontSize: 13, valign: "middle", margin: 0, lineSpacingMultiple: 1.08 });
  footer(s, N);
  s.addNotes("Resize the browser live across 640px and watch the layout snap from one column to two. Mobile-first = base styles target phones, media queries layer on enhancements. This is exactly the @media block at the bottom of style.css.");
}

// ============================================================
// 13 — YOUR TURN (PBL exercises)
// ============================================================
{
  const s = content(false);
  header(s, 12, "Your turn: extend the Joe Doe site", "Project-based practice — each task uses a concept from today.");
  const tasks = [
    ["Theme switch", "Add a --background-color dark value and a .dark class. Concept: variables + cascade.", CORAL],
    ["Style the form", "Give the contact inputs padding, a border, and radius; color the Send button. Concept: box model.", SLATE],
    ["Hover the cards", "Add section:hover that lifts the card with a stronger box-shadow. Concept: pseudo-classes.", "2E8B57"],
    ["3-column grid", "At ≥ 1024px, show three columns. Concept: media queries + grid.", NAVY],
  ];
  tasks.forEach((t, i) => {
    const x = 0.6 + (i % 2) * 6.2;
    const y = 1.95 + Math.floor(i / 2) * 2.25;
    card(s, x, y, 5.9, 2.0, MIST);
    s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.3, w: 0.55, h: 0.55, fill: { color: t[2] }, line: { type: "none" } });
    s.addText(String(i + 1), { x: x + 0.3, y: y + 0.3, w: 0.55, h: 0.55, align: "center", valign: "middle", fontFace: HEAD, bold: true, fontSize: 18, color: "FFFFFF", margin: 0 });
    s.addText(t[0], { x: x + 1.05, y: y + 0.32, w: 4.6, h: 0.5, fontFace: SANS, fontSize: 17, bold: true, color: NAVY, valign: "middle", margin: 0 });
    s.addText(t[1], { x: x + 0.32, y: y + 1.0, w: 5.3, h: 0.9, fontFace: SANS, fontSize: 13.5, color: BODY, valign: "top", margin: 0, lineSpacingMultiple: 1.06 });
  });
  footer(s, N);
  s.addNotes("Hand these out as lab work. They escalate in difficulty and each maps back to a slide: variables, box model, pseudo-classes, responsive grid. Encourage editing style.css live and refreshing the browser.");
}

// ============================================================
// 14 — RECAP / CLOSING
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: NAVY };
  s.addShape(pres.shapes.OVAL, { x: -1.4, y: 5.0, w: 4.2, h: 4.2, fill: { color: SLATE }, line: { type: "none" } });
  s.addShape(pres.shapes.OVAL, { x: 11.8, y: -1.4, w: 3.4, h: 3.4, fill: { color: CORAL }, line: { type: "none" } });
  s.addText("RECAP", { x: 0.9, y: 0.7, w: 5, h: 0.4, fontFace: SANS, fontSize: 15, bold: true, color: CORAL, charSpacing: 4, margin: 0 });
  s.addText("From markup to a designed site", { x: 0.9, y: 1.15, w: 11, h: 0.9, fontFace: HEAD, fontSize: 36, bold: true, color: "FFFFFF", margin: 0 });
  const recap = [
    "CSS is a separate presentation layer, linked with one line.",
    "Every rule = selector + property: value declarations.",
    "Selectors & the cascade decide who gets styled, and who wins.",
    "Variables, the box model & typography are your core toolkit.",
    "Grid + mobile-first media queries handle layout, any screen.",
  ];
  s.addText(recap.map((t, i) => ({ text: t, options: { bullet: { code: "2022", indent: 18 }, color: "DDE6EC", breakLine: i !== recap.length - 1, paraSpaceAfter: 12, fontSize: 18 } })),
    { x: 0.95, y: 2.3, w: 8.7, h: 3.2, fontFace: SANS, valign: "top", margin: 0, lineSpacingMultiple: 1.05 });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.95, y: 5.95, w: 8.5, h: 0.95, rectRadius: 0.1, fill: { color: SLATE }, line: { type: "none" } });
  s.addText([
    { text: "Next:  ", options: { bold: true, color: CORAL } },
    { text: "open css/style.css and try the four challenges — break it, fix it, make it yours.", options: { color: "FFFFFF" } },
  ], { x: 1.2, y: 5.95, w: 8.0, h: 0.95, fontFace: SANS, fontSize: 15, valign: "middle", margin: 0 });
  s.addNotes("Recap the arc: structure vs presentation, the rule grammar, selectors and cascade, the core toolkit, and layout. Send them to style.css to do the exercises. The best way to learn CSS is to change a value and refresh.");
}

pres.writeFile({ fileName: "JoeDoe-CSS.pptx" }).then(f => console.log("Wrote", f));
