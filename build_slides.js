// Build the AI Adoption Research presentation slides.
// Theme: "The Digital Divide" — navy + coral accent for the surprise finding.
const pptxgen = require("pptxgenjs");
const path = require("path");

const FIG = (name) => path.join(__dirname, "figures", name);

// -------------------- Palette --------------------
const NAVY      = "1E2761"; // primary
const ICE       = "CADCFC"; // secondary tint
const CORAL     = "F96167"; // accent (surprise findings + key stats)
const WHITE     = "FFFFFF";
const CHARCOAL  = "2D3142"; // body text
const MUTED     = "7B8794"; // captions

const FONT_HDR  = "Calibri"; // headers
const FONT_BODY = "Calibri"; // body
const FONT_NUM  = "Arial Black"; // big stats

// -------------------- Setup --------------------
const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3 x 7.5 inches
pres.author = "Mohamed Abdel-Hamid";
pres.title = "AI Adoption — Amplifying the Digital Divide";
pres.subject = "Data Preparation and Analysis, Spring 2026";

const W = 13.3, H = 7.5;

// -------------------- Helpers --------------------
function addSidebar(slide, num, label) {
  // Left navy sidebar with slide number + label
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.6, h: H, fill: { color: NAVY }, line: { color: NAVY, width: 0 },
  });
  slide.addText(String(num).padStart(2, "0"), {
    x: 0, y: 0.3, w: 0.6, h: 0.5,
    fontSize: 22, bold: true, fontFace: FONT_HDR, color: WHITE,
    align: "center", margin: 0,
  });
  slide.addShape(pres.shapes.LINE, {
    x: 0.15, y: 0.85, w: 0.3, h: 0,
    line: { color: CORAL, width: 1.5 },
  });
  // Vertical label down the sidebar
  slide.addText(label.toUpperCase(), {
    x: -0.7, y: H/2 - 0.5, w: 2, h: 0.4,
    fontSize: 10, bold: true, fontFace: FONT_HDR, color: ICE,
    rotate: 270, align: "center", charSpacing: 4, margin: 0,
  });
}

function addFooter(slide) {
  slide.addText([
    { text: "Mohamed Abdel-Hamid  ·  ", options: { color: MUTED } },
    { text: "AI Adoption Research", options: { color: NAVY, bold: true } },
    { text: "  ·  May 2026", options: { color: MUTED } },
  ], {
    x: 0.6, y: H - 0.4, w: W - 1.2, h: 0.3,
    fontSize: 9, fontFace: FONT_BODY, align: "right", margin: 0,
  });
}

function pageFrame(slide, num, label) {
  slide.background = { color: WHITE };
  addSidebar(slide, num, label);
  addFooter(slide);
}

function slideTitle(slide, title, subtitle) {
  slide.addText(title, {
    x: 1.0, y: 0.4, w: W - 1.5, h: 0.7,
    fontSize: 32, bold: true, fontFace: FONT_HDR, color: NAVY,
    align: "left", margin: 0,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 1.0, y: 1.1, w: W - 1.5, h: 0.4,
      fontSize: 14, italic: true, fontFace: FONT_BODY, color: MUTED,
      align: "left", margin: 0,
    });
  }
}

// =========================================================
// SLIDE 1 — Title
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: NAVY };

  // Decorative right block
  s.addShape(pres.shapes.RECTANGLE, {
    x: W - 4.2, y: 0, w: 4.2, h: H,
    fill: { color: "152044" }, line: { color: "152044", width: 0 },
  });
  // Coral accent stripe (thin, decorative)
  s.addShape(pres.shapes.RECTANGLE, {
    x: W - 0.25, y: 0, w: 0.25, h: H,
    fill: { color: CORAL }, line: { color: CORAL, width: 0 },
  });

  // Big number — reduced font and wider box so it doesn't wrap
  s.addText("0.72", {
    x: W - 4.0, y: 1.7, w: 3.6, h: 1.7,
    fontSize: 84, bold: true, fontFace: FONT_NUM, color: WHITE,
    align: "left", margin: 0,
  });
  s.addText([
    { text: "the country-level correlation\n", options: { color: ICE } },
    { text: "between AI adoption and GDP per capita", options: { color: ICE } },
  ], {
    x: W - 4.0, y: 3.5, w: 3.6, h: 1.0,
    fontSize: 13, italic: true, fontFace: FONT_BODY,
    align: "left", margin: 0,
  });

  s.addText("AI Adoption", {
    x: 0.7, y: 1.8, w: 8.5, h: 0.9,
    fontSize: 54, bold: true, fontFace: FONT_HDR, color: WHITE,
    align: "left", margin: 0,
  });
  s.addText("Amplifying the Digital Divide", {
    x: 0.7, y: 2.7, w: 8.5, h: 0.9,
    fontSize: 40, fontFace: FONT_HDR, color: CORAL,
    align: "left", margin: 0,
  });

  s.addShape(pres.shapes.LINE, {
    x: 0.7, y: 3.85, w: 1.2, h: 0,
    line: { color: CORAL, width: 3 },
  });

  s.addText("A cross-country and within-country exploration of generative AI use,\ncompiling 11 primary data sources across 38 countries.", {
    x: 0.7, y: 4.1, w: 9, h: 1.2,
    fontSize: 18, italic: true, fontFace: FONT_BODY, color: ICE,
    align: "left", margin: 0,
  });

  s.addText([
    { text: "Mohamed Abdel-Hamid", options: { bold: true, color: WHITE } },
    { text: "  ·  Data Preparation and Analysis  ·  Spring 2026  ·  University of St. Thomas", options: { color: ICE } },
  ], {
    x: 0.7, y: H - 1.0, w: 9, h: 0.4,
    fontSize: 13, fontFace: FONT_BODY, align: "left", margin: 0,
  });
}

// =========================================================
// SLIDE 2 — The Question
// =========================================================
{
  const s = pres.addSlide();
  pageFrame(s, 2, "Question");
  slideTitle(s, "The question", "Inside a country, and across countries");

  // Big quote
  s.addText("“Is generative AI adoption following — and amplifying — the existing digital divide?”", {
    x: 1.0, y: 1.9, w: W - 1.8, h: 1.3,
    fontSize: 28, italic: true, bold: true, fontFace: FONT_HDR, color: NAVY,
    align: "left", margin: 0,
  });

  // Two pillars
  const cardY = 3.7, cardW = 5.5, cardH = 2.5, gap = 0.4;
  const cardX1 = 1.0;
  const cardX2 = cardX1 + cardW + gap;

  // Card 1
  s.addShape(pres.shapes.RECTANGLE, {
    x: cardX1, y: cardY, w: cardW, h: cardH,
    fill: { color: ICE }, line: { color: ICE, width: 0 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: cardX1, y: cardY, w: 0.12, h: cardH,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 },
  });
  s.addText("WITHIN COUNTRIES", {
    x: cardX1 + 0.35, y: cardY + 0.25, w: cardW - 0.5, h: 0.3,
    fontSize: 11, bold: true, fontFace: FONT_HDR, color: NAVY, charSpacing: 3, margin: 0,
  });
  s.addText("Does AI use vary by\nage, income, and education?", {
    x: cardX1 + 0.35, y: cardY + 0.6, w: cardW - 0.5, h: 1.0,
    fontSize: 22, bold: true, fontFace: FONT_HDR, color: CHARCOAL, margin: 0,
  });
  s.addText("Tested with US (Pew, Brookings) and EU (Eurostat) data on adoption by demographic group.", {
    x: cardX1 + 0.35, y: cardY + 1.65, w: cardW - 0.5, h: 0.7,
    fontSize: 12, fontFace: FONT_BODY, color: CHARCOAL, italic: true, margin: 0,
  });

  // Card 2
  s.addShape(pres.shapes.RECTANGLE, {
    x: cardX2, y: cardY, w: cardW, h: cardH,
    fill: { color: ICE }, line: { color: ICE, width: 0 },
  });
  s.addShape(pres.shapes.RECTANGLE, {
    x: cardX2, y: cardY, w: 0.12, h: cardH,
    fill: { color: CORAL }, line: { color: CORAL, width: 0 },
  });
  s.addText("ACROSS COUNTRIES", {
    x: cardX2 + 0.35, y: cardY + 0.25, w: cardW - 0.5, h: 0.3,
    fontSize: 11, bold: true, fontFace: FONT_HDR, color: CORAL, charSpacing: 3, margin: 0,
  });
  s.addText("Does AI use track GDP\nand internet penetration?", {
    x: cardX2 + 0.35, y: cardY + 0.6, w: cardW - 0.5, h: 1.0,
    fontSize: 22, bold: true, fontFace: FONT_HDR, color: CHARCOAL, margin: 0,
  });
  s.addText("Tested with Microsoft AI Diffusion (34 countries) and World Bank/IMF macro indicators.", {
    x: cardX2 + 0.35, y: cardY + 1.65, w: cardW - 0.5, h: 0.7,
    fontSize: 12, fontFace: FONT_BODY, color: CHARCOAL, italic: true, margin: 0,
  });
}

// =========================================================
// SLIDE 3 — Hypothesis
// =========================================================
{
  const s = pres.addSlide();
  pageFrame(s, 3, "Hypothesis");
  slideTitle(s, "Hypothesis", "What I expected to see in the data");

  // Main hypothesis box
  s.addShape(pres.shapes.RECTANGLE, {
    x: 1.0, y: 2.0, w: W - 1.8, h: 1.6,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 },
  });
  s.addText([
    { text: "Generative AI adoption follows existing digital-divide patterns.\n", options: { bold: true, color: WHITE, fontSize: 22 } },
    { text: "AI is amplifying — not closing — the digital divide.", options: { italic: true, color: CORAL, fontSize: 18 } },
  ], {
    x: 1.4, y: 2.2, w: W - 2.2, h: 1.2,
    fontFace: FONT_HDR, align: "left", margin: 0, valign: "middle",
  });

  // Two predictions
  const py = 4.1;
  s.addText("Specifically:", {
    x: 1.0, y: py, w: 5, h: 0.4,
    fontSize: 14, italic: true, fontFace: FONT_BODY, color: MUTED, margin: 0,
  });

  const pred = [
    { num: "1.", text: "Within countries:", body: "Younger, higher-income, more-educated individuals adopt at substantially higher rates." },
    { num: "2.", text: "Across countries:", body: "Adoption correlates strongly with GDP per capita and internet penetration." },
  ];
  pred.forEach((p, i) => {
    const py2 = py + 0.5 + i * 1.1;
    s.addText(p.num, {
      x: 1.0, y: py2, w: 0.5, h: 0.6, fontSize: 28, bold: true,
      fontFace: FONT_NUM, color: CORAL, margin: 0,
    });
    s.addText([
      { text: p.text + " ", options: { bold: true, color: NAVY } },
      { text: p.body, options: { color: CHARCOAL } },
    ], {
      x: 1.6, y: py2 + 0.05, w: W - 2.5, h: 0.9,
      fontSize: 16, fontFace: FONT_BODY, margin: 0,
    });
  });
}

// =========================================================
// SLIDE 4 — Data Sources
// =========================================================
{
  const s = pres.addSlide();
  pageFrame(s, 4, "Data");
  slideTitle(s, "Eleven primary data sources", "Compiled into a 10-sheet master workbook (38 countries)");

  // Three columns of sources
  const colY = 1.9, colW = 3.85, colH = 4.7, gap = 0.25;
  const cols = [
    { title: "WITHIN-COUNTRY", color: NAVY, sources: [
      ["Pew Research", "ChatGPT use by US adults (n=5,123, 2025); by age + education"],
      ["Pew Research", "Teens AI use by household income (Dec 2025)"],
      ["Pew Research", "21% of US workers use AI on the job (Oct 2025)"],
      ["Brookings / RPS", "US generative AI use by demographics (2024)"],
      ["Eurostat", "EU GenAI use by country and age group (2025)"],
    ]},
    { title: "CROSS-COUNTRY", color: CORAL, sources: [
      ["Microsoft", "AI Diffusion Report 2025 H2 — % of working-age pop using GenAI (top 30)"],
      ["Stanford HAI", "AI Index 2025 — country-level optimism"],
      ["Ipsos", "AI Monitor 2024 — 32-country attitudes"],
      ["Anthropic", "Economic Index Sept 2025 — per-capita Claude usage"],
      ["Visual Capitalist", "GPO-AI 2024 — daily/weekly use, 21 countries"],
    ]},
    { title: "CONTROL VARIABLES", color: NAVY, sources: [
      ["IMF / World Bank", "GDP per capita (PPP), 2024"],
      ["ITU / World Bank", "Internet users (% of population), 2024"],
      ["World Bank", "Tertiary education completion rate"],
      ["", ""],
      ["", ""],
    ]},
  ];

  cols.forEach((col, ci) => {
    const x = 1.0 + ci * (colW + gap);
    s.addShape(pres.shapes.RECTANGLE, {
      x, y: colY, w: colW, h: 0.45,
      fill: { color: col.color }, line: { color: col.color, width: 0 },
    });
    s.addText(col.title, {
      x, y: colY, w: colW, h: 0.45,
      fontSize: 12, bold: true, fontFace: FONT_HDR, color: WHITE,
      align: "center", valign: "middle", charSpacing: 3, margin: 0,
    });
    col.sources.forEach((src, si) => {
      const sy = colY + 0.55 + si * 0.83;
      if (!src[0]) return;
      s.addText(src[0], {
        x: x + 0.1, y: sy, w: colW - 0.2, h: 0.3,
        fontSize: 11, bold: true, fontFace: FONT_HDR, color: NAVY, margin: 0,
      });
      s.addText(src[1], {
        x: x + 0.1, y: sy + 0.28, w: colW - 0.2, h: 0.5,
        fontSize: 9.5, fontFace: FONT_BODY, color: CHARCOAL, margin: 0,
      });
    });
  });
}

// =========================================================
// SLIDE 5 — Methods
// =========================================================
{
  const s = pres.addSlide();
  pageFrame(s, 5, "Methods");
  slideTitle(s, "How I tested it", "Compile, compare, correlate — then check the numbers two different ways");

  const stepY = 2.0, stepW = 5.5, stepH = 1.3, stepGap = 0.3;
  const steps = [
    { num: "1", title: "Compile", body: "11 sources → 10-sheet workbook with merged Master_Country sheet (38 countries)" },
    { num: "2", title: "Compare", body: "US/EU adoption rates by age, income, education vs. national average" },
    { num: "3", title: "Correlate", body: "Pearson r between country-level adoption and GDP, internet, education" },
    { num: "4", title: "Verify", body: "Cross-check Python correlations against Excel CORREL() formulas in workbook" },
  ];

  steps.forEach((step, i) => {
    const row = Math.floor(i / 2), col = i % 2;
    const x = 1.0 + col * (stepW + stepGap);
    const y = stepY + row * (stepH + stepGap);

    s.addShape(pres.shapes.RECTANGLE, {
      x, y, w: stepW, h: stepH,
      fill: { color: WHITE }, line: { color: ICE, width: 1.5 },
    });
    // Number circle
    s.addShape(pres.shapes.OVAL, {
      x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7,
      fill: { color: NAVY }, line: { color: NAVY, width: 0 },
    });
    s.addText(step.num, {
      x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7,
      fontSize: 24, bold: true, fontFace: FONT_NUM, color: WHITE,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText(step.title, {
      x: x + 1.15, y: y + 0.25, w: stepW - 1.3, h: 0.4,
      fontSize: 18, bold: true, fontFace: FONT_HDR, color: NAVY, margin: 0,
    });
    s.addText(step.body, {
      x: x + 1.15, y: y + 0.7, w: stepW - 1.3, h: 0.5,
      fontSize: 12, fontFace: FONT_BODY, color: CHARCOAL, margin: 0,
    });
  });

  // Note on source comparability
  s.addText([
    { text: "Sources use different adoption definitions (have-ever-used vs. used-in-last-3-months vs. daily). Cross-source comparisons are indicative; the Microsoft and Eurostat measures are most directly comparable.", options: { color: MUTED, italic: true } },
  ], {
    x: 1.0, y: H - 1.4, w: W - 1.8, h: 0.7,
    fontSize: 11, fontFace: FONT_BODY, margin: 0,
  });
}

// =========================================================
// SLIDE 6 — Result A: Within-country divide
// =========================================================
{
  const s = pres.addSlide();
  pageFrame(s, 6, "Result A");
  slideTitle(s, "Result A — The within-country divide", "Younger and more-educated adults use ChatGPT far more than everyone else");

  // Big stat callouts on the left — wider boxes, reduced font, "x" instead of multiply sign so it doesn't wrap
  const statsX = 0.9, statY = 1.9;

  // Stat 1: 5.8x
  s.addText("5.8x", {
    x: statsX, y: statY, w: 4.6, h: 1.1,
    fontSize: 64, bold: true, fontFace: FONT_NUM, color: NAVY,
    align: "left", margin: 0,
  });
  s.addText([
    { text: "more likely · ", options: { bold: true, color: NAVY } },
    { text: "US adults under 30 use ChatGPT than 65+\n", options: { color: CHARCOAL } },
    { text: "(58% vs. 10% — Pew, 2025)", options: { color: MUTED, italic: true, fontSize: 11 } },
  ], {
    x: statsX, y: statY + 1.05, w: 4.6, h: 1.0,
    fontSize: 13, fontFace: FONT_BODY, margin: 0,
  });

  // Stat 2: 2.9x
  s.addText("2.9x", {
    x: statsX, y: statY + 2.3, w: 4.6, h: 1.1,
    fontSize: 64, bold: true, fontFace: FONT_NUM, color: CORAL,
    align: "left", margin: 0,
  });
  s.addText([
    { text: "more likely · ", options: { bold: true, color: CORAL } },
    { text: "Postgrads use ChatGPT than HS-or-less\n", options: { color: CHARCOAL } },
    { text: "(52% vs. 18% — Pew, 2025)", options: { color: MUTED, italic: true, fontSize: 11 } },
  ], {
    x: statsX, y: statY + 3.35, w: 4.6, h: 1.0,
    fontSize: 13, fontFace: FONT_BODY, margin: 0,
  });

  // Chart on right
  s.addImage({
    path: FIG("chart1_us_adoption_by_age.png"),
    x: 5.8, y: 1.7, w: 6.7, h: 3.2,
  });
  s.addText("Figure: ChatGPT use by US adults, by age. Source: Pew Research Center, June 2025 (n=5,123).", {
    x: 5.8, y: 4.95, w: 6.7, h: 0.3,
    fontSize: 10, italic: true, fontFace: FONT_BODY, color: MUTED, align: "center", margin: 0,
  });

  // Conclusion
  s.addText("Same pattern shows up in Europe — Eurostat reports 64% of 16-24 year-olds use generative AI vs. 7% of 65-74 year-olds.", {
    x: 0.9, y: H - 1.3, w: W - 1.5, h: 0.6,
    fontSize: 14, bold: true, fontFace: FONT_HDR, color: NAVY, margin: 0,
  });
}

// =========================================================
// SLIDE 7 — Result B: Across-country divide
// =========================================================
{
  const s = pres.addSlide();
  pageFrame(s, 7, "Result B");
  slideTitle(s, "Result B — The across-country divide", "Richer countries use AI more — and the link is strong, not subtle");

  // Scatter chart
  s.addImage({
    path: FIG("chart4_adoption_vs_gdp.png"),
    x: 0.9, y: 1.7, w: 8.3, h: 4.5,
  });

  // Stats panel on right
  const px = 9.5, py = 1.9, pw = 3.2;
  s.addShape(pres.shapes.RECTANGLE, {
    x: px, y: py, w: pw, h: 4.5,
    fill: { color: NAVY }, line: { color: NAVY, width: 0 },
  });
  s.addText("KEY CORRELATIONS", {
    x: px + 0.2, y: py + 0.25, w: pw - 0.4, h: 0.35,
    fontSize: 11, bold: true, fontFace: FONT_HDR, color: ICE, charSpacing: 3, margin: 0,
  });

  const items = [
    { label: "Adoption × GDP/cap", r: "+0.72" },
    { label: "Adoption × Internet", r: "+0.61" },
    { label: "Adoption × Tertiary ed.", r: "+0.51" },
  ];
  items.forEach((it, i) => {
    const y = py + 0.85 + i * 1.15;
    s.addText(it.r, {
      x: px + 0.2, y, w: pw - 0.4, h: 0.55,
      fontSize: 32, bold: true, fontFace: FONT_NUM, color: CORAL,
      align: "left", margin: 0,
    });
    s.addText(it.label, {
      x: px + 0.2, y: y + 0.55, w: pw - 0.4, h: 0.3,
      fontSize: 11, fontFace: FONT_BODY, color: ICE, margin: 0,
    });
  });

  s.addText("n = 34 countries", {
    x: px + 0.2, y: py + 4.0, w: pw - 0.4, h: 0.3,
    fontSize: 10, italic: true, fontFace: FONT_BODY, color: ICE, align: "right", margin: 0,
  });

  // Conclusion
  s.addText("Wealth, internet access, and education all predict how much a country uses AI. Of the three, wealth is the strongest predictor.", {
    x: 0.9, y: H - 1.0, w: W - 1.5, h: 0.5,
    fontSize: 14, bold: true, fontFace: FONT_HDR, color: NAVY, margin: 0,
  });
}

// =========================================================
// SLIDE 8 — The Twist: Optimism Paradox
// =========================================================
{
  const s = pres.addSlide();
  pageFrame(s, 8, "Surprise");

  // Use a coral accent strip for this slide to signal "different"
  s.addShape(pres.shapes.RECTANGLE, {
    x: 0.6, y: 0, w: 0.08, h: H,
    fill: { color: CORAL }, line: { color: CORAL, width: 0 },
  });

  // No "THE TWIST" tag — sidebar already says "Surprise"
  slideTitle(s, "Skeptics adopt more, enthusiasts adopt less", "Where AI use is highest, public opinion of AI is actually lowest");

  // Big negative correlation — wider box & smaller font so it doesn't wrap
  s.addText("−0.67", {
    x: 0.9, y: 1.9, w: 5.0, h: 1.4,
    fontSize: 76, bold: true, fontFace: FONT_NUM, color: CORAL,
    align: "left", margin: 0,
  });
  s.addText([
    { text: "AI adoption × AI optimism\n", options: { bold: true, color: NAVY } },
    { text: "across 16 countries with both measures", options: { color: MUTED, italic: true } },
  ], {
    x: 0.9, y: 3.2, w: 4.5, h: 0.7,
    fontSize: 13, fontFace: FONT_BODY, margin: 0,
  });

  // Examples
  const exX = 0.9, exY = 4.05;
  s.addText("Most optimistic, lower adoption:", {
    x: exX, y: exY, w: 4.5, h: 0.3,
    fontSize: 11, bold: true, fontFace: FONT_HDR, color: NAVY, margin: 0,
  });
  s.addText("China 83%  ·  Indonesia 80%  ·  Thailand 77%", {
    x: exX, y: exY + 0.32, w: 4.5, h: 0.3,
    fontSize: 12, fontFace: FONT_BODY, color: CHARCOAL, margin: 0,
  });
  s.addText("Most skeptical, higher adoption:", {
    x: exX, y: exY + 0.85, w: 4.5, h: 0.3,
    fontSize: 11, bold: true, fontFace: FONT_HDR, color: CORAL, margin: 0,
  });
  s.addText("Netherlands 36%  ·  US 39%  ·  France 42%", {
    x: exX, y: exY + 1.17, w: 4.5, h: 0.3,
    fontSize: 12, fontFace: FONT_BODY, color: CHARCOAL, margin: 0,
  });

  // Chart on right
  s.addImage({
    path: FIG("chart6_optimism_vs_adoption.png"),
    x: 5.8, y: 1.7, w: 6.7, h: 4.2,
  });
  s.addText("Figure: Country-level AI adoption (Microsoft) vs. AI optimism (Stanford/Ipsos), n=16.", {
    x: 5.8, y: 5.95, w: 6.7, h: 0.3,
    fontSize: 10, italic: true, fontFace: FONT_BODY, color: MUTED, align: "center", margin: 0,
  });

  // Interpretation
  s.addText("Liking AI and using AI turn out to be different things. One reading: people who actually use the tools see the failures up close.", {
    x: 0.9, y: H - 1.0, w: W - 1.5, h: 0.5,
    fontSize: 14, bold: true, fontFace: FONT_HDR, color: NAVY, margin: 0,
  });
}

// =========================================================
// SLIDE 9 — Uncertainty & Limitations
// =========================================================
{
  const s = pres.addSlide();
  pageFrame(s, 9, "Limits");
  slideTitle(s, "Uncertainty & scope of inference", "Where these findings should and should not be applied");

  const lims = [
    { num: "01", title: "Source comparability", body: "Different surveys define \"adoption\" differently — ever-used vs. last-3-months vs. daily. Direct comparisons are indicative." },
    { num: "02", title: "Online sampling bias", body: "Country-level surveys are mostly online — likely overstating adoption in lower-internet countries." },
    { num: "03", title: "Correlational, not causal", body: "Cross-sectional Pearson r only. Reverse causation (AI → productivity → GDP) cannot be ruled out." },
    { num: "04", title: "Small n for the twist", body: "Optimism×adoption analysis rests on only 16 countries. Larger sampling needed to firmly establish the paradox." },
    { num: "05", title: "Coverage skews OECD", body: "34-country adoption sample under-represents Sub-Saharan Africa and parts of Central Asia." },
  ];

  const startY = 1.85, rowH = 0.78, leftX = 1.0;
  lims.forEach((lim, i) => {
    const y = startY + i * rowH;
    s.addText(lim.num, {
      x: leftX, y, w: 0.7, h: 0.6, fontSize: 22, bold: true,
      fontFace: FONT_NUM, color: ICE, margin: 0,
    });
    s.addText([
      { text: lim.title + " · ", options: { bold: true, color: NAVY } },
      { text: lim.body, options: { color: CHARCOAL } },
    ], {
      x: leftX + 0.8, y: y + 0.1, w: W - leftX - 1.5, h: 0.65,
      fontSize: 13, fontFace: FONT_BODY, margin: 0,
    });
  });

  // Scope of inference
  s.addShape(pres.shapes.RECTANGLE, {
    x: 1.0, y: H - 1.5, w: W - 1.8, h: 0.85,
    fill: { color: ICE }, line: { color: ICE, width: 0 },
  });
  s.addText([
    { text: "Scope of inference: ", options: { bold: true, color: NAVY } },
    { text: "Findings generalize to OECD and middle-income economies. Not generalizable to low-income contexts — those need separate data.", options: { color: CHARCOAL } },
  ], {
    x: 1.2, y: H - 1.45, w: W - 2.1, h: 0.75,
    fontSize: 12, fontFace: FONT_BODY, valign: "middle", margin: 0,
  });
}

// =========================================================
// SLIDE 10 — Refined hypothesis / Conclusion
// =========================================================
{
  const s = pres.addSlide();
  s.background = { color: NAVY };

  // Coral accent strip on the right
  s.addShape(pres.shapes.RECTANGLE, {
    x: W - 0.35, y: 0, w: 0.35, h: H,
    fill: { color: CORAL }, line: { color: CORAL, width: 0 },
  });

  s.addText("WRAPPING UP", {
    x: 0.7, y: 0.55, w: 5, h: 0.4,
    fontSize: 12, bold: true, fontFace: FONT_HDR, color: ICE, charSpacing: 4, margin: 0,
  });
  s.addText("What I learned", {
    x: 0.7, y: 0.95, w: W - 1.5, h: 0.7,
    fontSize: 38, bold: true, fontFace: FONT_HDR, color: WHITE, margin: 0,
  });
  s.addShape(pres.shapes.LINE, {
    x: 0.7, y: 1.75, w: 1.0, h: 0,
    line: { color: CORAL, width: 3 },
  });

  // Two refined statements
  s.addText("Original (supported)", {
    x: 0.7, y: 2.1, w: 5, h: 0.4,
    fontSize: 13, bold: true, fontFace: FONT_HDR, color: CORAL, charSpacing: 2, margin: 0,
  });
  s.addText("Generative AI adoption follows existing digital-divide patterns. Within countries, age and education gradients are steep. Across countries, adoption tracks GDP (r=+0.72) and internet penetration (r=+0.61). AI is amplifying the digital divide.", {
    x: 0.7, y: 2.5, w: W - 1.5, h: 1.3,
    fontSize: 16, fontFace: FONT_BODY, color: WHITE, margin: 0,
  });

  s.addText("Added refinement (new)", {
    x: 0.7, y: 4.1, w: 5, h: 0.4,
    fontSize: 13, bold: true, fontFace: FONT_HDR, color: CORAL, charSpacing: 2, margin: 0,
  });
  s.addText("Public sentiment about AI is decoupled from — and inversely related to — actual usage (r=−0.67). Adoption and optimism are distinct outcomes with potentially different drivers. Future work should treat them separately.", {
    x: 0.7, y: 4.5, w: W - 1.5, h: 1.3,
    fontSize: 16, fontFace: FONT_BODY, color: WHITE, margin: 0,
  });

  // Footer with repo
  s.addShape(pres.shapes.LINE, {
    x: 0.7, y: H - 1.05, w: W - 1.5, h: 0,
    line: { color: ICE, width: 0.5 },
  });
  s.addText([
    { text: "Repository  ", options: { color: ICE, bold: true } },
    { text: "github.com/moh-a-abde/ai-adoption-research", options: { color: WHITE } },
  ], {
    x: 0.7, y: H - 0.85, w: W - 1.5, h: 0.4,
    fontSize: 13, fontFace: FONT_BODY, align: "left", margin: 0,
  });
  s.addText("Mohamed Abdel-Hamid  ·  May 2026", {
    x: 0.7, y: H - 0.85, w: W - 1.5, h: 0.4,
    fontSize: 13, fontFace: FONT_BODY, color: ICE, align: "right", margin: 0,
  });
}

// -------------------- Write --------------------
pres.writeFile({ fileName: "/sessions/wizardly-lucid-davinci/mnt/Project/slides/AI_Adoption_Presentation.pptx" })
  .then(f => console.log("Wrote", f));
