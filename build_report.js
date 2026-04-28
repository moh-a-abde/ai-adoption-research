// Build the AI Adoption research report Word doc.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
  AlignmentType, LevelFormat, ExternalHyperlink, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageBreak, PageOrientation,
} = require('docx');

const ROOT = __dirname;
const FIGURES_DIR = path.join(ROOT, 'figures');
const REPORTS_DIR = path.join(ROOT, 'reports');
const OUT = path.join(REPORTS_DIR, 'AI_Adoption_Research_Report.docx');

fs.mkdirSync(REPORTS_DIR, { recursive: true });

// ---------- Style helpers ----------
const ARIAL = "Arial";
const COL_NAVY = "1F4E78";
const COL_BLUE = "2E75B6";
const COL_LIGHT_BLUE = "D5E8F0";
const COL_RED = "C00000";
const COL_TEXT = "333333";

const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const cellBorders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function P(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after ?? 120, before: opts.before ?? 0 },
    alignment: opts.align ?? AlignmentType.LEFT,
    children: [new TextRun({
      text,
      bold: opts.bold,
      italics: opts.italic,
      size: opts.size ?? 22,  // 11pt
      color: opts.color ?? COL_TEXT,
      font: ARIAL,
    })],
  });
}

function Runs(parts, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after ?? 120 },
    alignment: opts.align ?? AlignmentType.LEFT,
    children: parts.map(p => {
      if (typeof p === 'string') {
        return new TextRun({ text: p, size: 22, color: COL_TEXT, font: ARIAL });
      }
      return new TextRun({
        text: p.text, bold: p.bold, italics: p.italic,
        size: p.size ?? 22, color: p.color ?? COL_TEXT, font: ARIAL,
      });
    }),
  });
}

function H1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180 },
    children: [new TextRun({ text, bold: true, size: 32, color: COL_NAVY, font: ARIAL })],
  });
}

function H2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 26, color: COL_NAVY, font: ARIAL })],
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, size: 22, color: COL_TEXT, font: ARIAL })],
  });
}

function bulletRich(parts) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: parts.map(p => {
      if (typeof p === 'string') return new TextRun({ text: p, size: 22, color: COL_TEXT, font: ARIAL });
      return new TextRun({ text: p.text, bold: p.bold, italics: p.italic, size: 22, color: p.color ?? COL_TEXT, font: ARIAL });
    }),
  });
}

function image(path, w, h, alt) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 100 },
    children: [new ImageRun({
      type: 'png',
      data: fs.readFileSync(path),
      transformation: { width: w, height: h },
      altText: { title: alt, description: alt, name: alt },
    })],
  });
}

function caption(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 200 },
    children: [new TextRun({ text, italics: true, size: 18, color: "595959", font: ARIAL })],
  });
}

// ---------- Tables ----------
function tableCell(text, opts = {}) {
  return new TableCell({
    borders: cellBorders,
    width: { size: opts.width, type: WidthType.DXA },
    margins: cellMargins,
    shading: opts.shade ? { fill: opts.shade, type: ShadingType.CLEAR } : undefined,
    children: [new Paragraph({
      alignment: opts.align ?? AlignmentType.LEFT,
      children: [new TextRun({
        text, bold: opts.bold, size: 20, color: opts.color ?? COL_TEXT, font: ARIAL,
      })],
    })],
  });
}

function corrTable() {
  const widths = [4500, 1800, 3060]; // sums to 9360
  const headerShade = COL_NAVY;
  const altShade = "F2F2F2";

  const rows = [
    ["Variable Pair", "Pearson r", "Interpretation", true, null],
    ["AI Adoption × GDP per capita", "+0.717", "Strong positive — wealthier countries adopt much more.", false, null],
    ["AI Adoption × Internet Penetration", "+0.608", "Strong positive — adoption rides existing digital infrastructure.", false, altShade],
    ["AI Adoption × Tertiary Education", "+0.505", "Moderate positive — more-educated populations adopt more.", false, null],
    ["AI Adoption × AI Optimism", "−0.666", "Strong NEGATIVE — countries that USE AI more are LESS optimistic about it.", false, altShade],
    ["AI Optimism × GDP per capita", "−0.717", "Strong negative — richer countries are more skeptical of AI.", false, null],
  ];

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map(([a, b, c, isHeader, shade]) => new TableRow({
      tableHeader: isHeader,
      children: [
        tableCell(a, { width: widths[0], bold: isHeader, color: isHeader ? "FFFFFF" : COL_TEXT, shade: isHeader ? headerShade : shade }),
        tableCell(b, { width: widths[1], bold: true, color: isHeader ? "FFFFFF" : (a.includes('AI Optimism') || a.includes('Optimism')) ? COL_RED : COL_NAVY, shade: isHeader ? headerShade : shade, align: AlignmentType.CENTER }),
        tableCell(c, { width: widths[2], color: isHeader ? "FFFFFF" : COL_TEXT, shade: isHeader ? headerShade : shade }),
      ],
    })),
  });
}

function sourceTable() {
  const widths = [3000, 6360];
  const headerShade = COL_NAVY;
  const altShade = "F2F2F2";

  const rows = [
    ["Source", "Citation / URL", true],
    ["Pew Research Center (June 2025)", "\"34% of U.S. adults have used ChatGPT, about double the share in 2023\" — pewresearch.org/short-reads/2025/06/25/", false],
    ["Pew Research Center (Dec 2025)", "\"Teens, Social Media and AI Chatbots 2025\" — pewresearch.org/internet/2025/12/09/", true],
    ["Pew Research Center (Oct 2025)", "\"21% of U.S. workers now use AI in their job\" — pewresearch.org/short-reads/2025/10/06/", false],
    ["Brookings (2025)", "\"How are Americans using AI? Evidence from a nationwide survey\" — brookings.edu/articles/", true],
    ["Stanford HAI (2025)", "2025 AI Index Report, Public Opinion chapter — hai.stanford.edu/ai-index/2025-ai-index-report/public-opinion", false],
    ["Ipsos (2024)", "Ipsos AI Monitor 2024, 32-country survey — ipsos.com/en-us/ipsos-ai-monitor-2024", true],
    ["Microsoft (Jan 2026)", "AI Diffusion Report 2025 H2 — microsoft.com/en-us/research/ (PDF)", false],
    ["Eurostat (Dec 2025)", "\"32.7% of EU people used generative AI tools in 2025\" — ec.europa.eu/eurostat/", true],
    ["Anthropic (Sept 2025)", "Economic Index Report: Uneven geographic and enterprise AI adoption — anthropic.com/research/anthropic-economic-index-september-2025-report", false],
    ["Visual Capitalist / GPO-AI 2024", "\"How Often People Use ChatGPT Across 21 Countries\" — visualcapitalist.com/how-often-people-use-chatgpt-across-21-countries/", true],
    ["IMF / World Bank (2024)", "World Economic Outlook 2024 (GDP PPP); ITU/World Bank Internet Users — data.worldbank.org", false],
  ];

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map(([a, b, isHeader]) => new TableRow({
      tableHeader: isHeader,
      children: [
        tableCell(a, { width: widths[0], bold: isHeader, color: isHeader ? "FFFFFF" : COL_TEXT, shade: isHeader ? headerShade : (rows.indexOf([a,b,isHeader]) % 2 ? altShade : null) }),
        tableCell(b, { width: widths[1], bold: false, color: isHeader ? "FFFFFF" : COL_TEXT, shade: isHeader ? headerShade : null }),
      ],
    })),
  });
}

// ---------- Top countries summary table ----------
function topCountriesTable() {
  const widths = [800, 3000, 1900, 1830, 1830]; // 9360
  const headerShade = COL_NAVY;
  const rows = [
    ["#", "Country", "Adoption %", "GDP/cap PPP", "Internet %", true],
    [1, "United Arab Emirates", "64.0%", "$96,850", "100%", false],
    [2, "Singapore", "60.9%", "$132,570", "96%", false],
    [3, "Norway", "46.4%", "$92,650", "99%", false],
    [4, "Ireland", "44.6%", "$115,300", "96%", false],
    [5, "France", "44.0%", "$60,340", "86%", false],
    [6, "Spain", "41.8%", "$53,350", "95%", false],
    [7, "New Zealand", "40.5%", "$53,800", "96%", false],
    [8, "Netherlands", "38.9%", "$77,460", "99%", false],
    [9, "United Kingdom", "38.9%", "$60,620", "95%", false],
    [10, "Qatar", "38.3%", "$113,700", "100%", false],
  ];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: widths,
    rows: rows.map((r) => {
      const isHeader = r[5];
      return new TableRow({
        tableHeader: isHeader,
        children: [0,1,2,3,4].map(i => tableCell(String(r[i]), {
          width: widths[i],
          bold: isHeader,
          color: isHeader ? "FFFFFF" : COL_TEXT,
          shade: isHeader ? headerShade : null,
          align: i === 1 ? AlignmentType.LEFT : AlignmentType.CENTER,
        })),
      });
    }),
  });
}

// =============================
// Build the document
// =============================
const doc = new Document({
  creator: "Mohamed Abdel-Hamid",
  title: "AI Adoption Research — Initial Findings",
  styles: {
    default: { document: { run: { font: ARIAL, size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: ARIAL, color: COL_NAVY },
        paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: ARIAL, color: COL_NAVY },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
    },
    children: [
      // ===== Title block =====
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 100 },
        children: [new TextRun({ text: "AI Adoption Across Demographics and Countries", bold: true, size: 44, color: COL_NAVY, font: ARIAL })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
        children: [new TextRun({ text: "Initial Research & Data Exploration", italics: true, size: 26, color: COL_BLUE, font: ARIAL })],
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 360 },
        children: [new TextRun({ text: "Author: Mohamed Abdel-Hamid  |  Date: April 2026  |  Topic Check-in", size: 20, color: "595959", font: ARIAL })],
      }),

      // ===== 1. Topic & Hypothesis =====
      H1("1. Topic and Hypothesis"),
      P("This research examines generative AI adoption along two dimensions: how usage varies across demographic groups (age, education, income) within the United States, and how it varies across countries — testing whether AI adoption is closing or widening pre-existing technology gaps."),
      H2("Hypothesis"),
      Runs([
        { text: "Generative AI adoption follows existing digital-divide patterns. ", bold: true },
        "Specifically: (1) within countries, younger, higher-income, and more-educated individuals adopt at substantially higher rates than older, lower-income, less-educated individuals; and (2) across countries, adoption correlates strongly with GDP per capita and internet penetration. AI is amplifying — not closing — the digital divide.",
      ]),

      // ===== 2. Data Sources =====
      H1("2. Data Sources"),
      P("Data is compiled from eleven primary sources covering both demographic and country-level dimensions. The accompanying spreadsheet (AI_Adoption_Research_Data.xlsx) contains the raw numbers organized in nine sheets, plus a master country-level sheet for correlation analysis."),
      sourceTable(),
      P(" "),
      Runs([
        { text: "Caveat: ", bold: true, italic: true },
        { text: "different sources define \"adoption\" differently (have-ever-used vs. used-in-last-3-months vs. daily use). Cross-source comparisons should be read as indicative rather than definitive. The Microsoft AI Diffusion measure and the Eurostat measure are the most directly comparable across countries.", italic: true, color: "595959" },
      ]),

      // ===== 3. US Demographic Findings =====
      new Paragraph({ children: [new PageBreak()] }),
      H1("3. Findings — US Demographic Dimension"),

      H2("3.1 Age"),
      P("Pew Research Center (June 2025) finds that 34% of US adults have ever used ChatGPT — roughly double the share in 2023. The age gradient is steep: adults under 30 are 5.8x more likely to use ChatGPT than those 65 and over."),
      image(path.join(FIGURES_DIR, 'chart1_us_adoption_by_age.png'), 540, 305, "ChatGPT adoption by age group, US adults"),
      caption("Figure 1. ChatGPT adoption by age group, US adults. Source: Pew Research Center, June 2025 (n=5,123)."),

      H2("3.2 Education"),
      P("The education gradient is equally pronounced. Adults with a postgraduate degree are 2.9x more likely to have used ChatGPT than those with a high school education or less. Even within the bachelor's-or-higher tier, ChatGPT use is the modal experience (around half), while it remains a minority experience among less-educated adults."),
      image(path.join(FIGURES_DIR, 'chart2_us_adoption_by_education.png'), 540, 305, "ChatGPT adoption by education, US adults"),
      caption("Figure 2. ChatGPT adoption by education, US adults. Source: Pew Research Center, June 2025."),

      H2("3.3 Income (US Teens)"),
      P("Pew's December 2025 teen survey provides the cleanest income breakdown available: 66% of teens in households earning $75,000+ have used ChatGPT, vs. 56% in households earning under $30,000 — a 10-percentage-point gap among the same age cohort. The Brookings/Real-Time Population Survey similarly finds AI adoption \"led by younger, higher-income adults with college educations, with older, rural and lower-income adults lagging behind.\""),

      H2("3.4 EU Age Pattern (Cross-Reference)"),
      P("The Eurostat 2025 survey of EU residents aged 16-74 confirms the same age pattern at scale. Across the EU, 32.7% of adults used a generative AI tool in 2025, but the spread by age is enormous: 63.8% of 16-24-year-olds vs. just 7% of 65-74-year-olds — a 9.1x ratio."),
      image(path.join(FIGURES_DIR, 'chart7_eu_adoption_by_age.png'), 540, 305, "EU GenAI use by age group"),
      caption("Figure 3. EU generative AI use by age group, 2025. Source: Eurostat, December 2025. Middle brackets estimated."),

      // ===== 4. Country-Level Findings =====
      new Paragraph({ children: [new PageBreak()] }),
      H1("4. Findings — Country Dimension"),

      H2("4.1 Adoption Rankings"),
      P("Microsoft's AI Diffusion Report (Jan 2026) provides the most consistent cross-country measure: percentage of working-age population using generative AI tools. The global average in H2 2025 was 16.3%, but the spread is dramatic — from 64% in the UAE to under 2% in many low-income countries. The top 10 are dominated by wealthy, small, and digitally-mature economies:"),
      topCountriesTable(),
      P(" "),
      P("The full top-30 ranking is shown below; data for an additional 4 large economies (Japan, Brazil, China, India) is included as well for context."),
      image(path.join(FIGURES_DIR, 'chart3_country_adoption.png'), 540, 660, "Country AI adoption rankings"),
      caption("Figure 4. Generative AI adoption by country (% of working-age population). Source: Microsoft AI Diffusion Report 2025 H2 (January 2026)."),

      H2("4.2 Adoption vs. GDP per capita"),
      P("The country-level correlation is striking. Across 34 countries with both adoption and GDP data, the Pearson correlation between AI adoption and GDP per capita (PPP) is r = +0.717 — a strong positive relationship. Outliers exist (the US under-performs its income; Spain and France over-perform), but the linear fit is unmistakable."),
      image(path.join(FIGURES_DIR, 'chart4_adoption_vs_gdp.png'), 540, 360, "AI adoption vs GDP per capita scatter"),
      caption("Figure 5. AI adoption vs. GDP per capita (PPP), 34 countries. Pearson r = +0.717."),

      H2("4.3 Adoption vs. Internet Penetration"),
      P("AI adoption also correlates strongly with internet penetration (r = +0.608) — confirming that AI is layering on top of, and benefiting from, existing digital infrastructure. Countries with poor internet access are not catching up via AI; they're being left further behind."),
      image(path.join(FIGURES_DIR, 'chart5_adoption_vs_internet.png'), 540, 360, "AI adoption vs internet penetration scatter"),
      caption("Figure 6. AI adoption vs. internet penetration, 34 countries. Pearson r = +0.608."),

      // ===== 5. The counter-intuitive finding =====
      new Paragraph({ children: [new PageBreak()] }),
      H1("5. Counter-Intuitive Finding: Optimism Inversely Tracks Adoption"),
      Runs([
        "The most surprising result of this initial exploration is the ",
        { text: "inverse", bold: true, italic: true },
        " relationship between AI optimism (Stanford/Ipsos) and AI adoption (Microsoft). Across the 16 countries with both measures, the correlation is ",
        { text: "r = −0.666", bold: true, color: COL_RED },
        ".",
      ]),
      P("Countries that say AI is more beneficial than harmful — China (83%), Indonesia (80%), Thailand (77%), India (62%) — are not the countries that actually use it the most. Conversely, the most adoption-heavy populations (Netherlands, France, US, UK) are among the most skeptical."),
      image(path.join(FIGURES_DIR, 'chart6_optimism_vs_adoption.png'), 540, 360, "AI optimism vs adoption scatter"),
      caption("Figure 7. AI optimism vs. AI adoption, 16 countries with both data points. Pearson r = −0.666."),
      P("This is the kind of finding that should reshape the hypothesis. Two non-mutually-exclusive interpretations:"),
      bulletRich([
        { text: "Familiarity breeds skepticism. ", bold: true },
        "Heavy users have direct experience with AI's failures, hallucinations, and limitations; non-users hear mostly the hype.",
      ]),
      bulletRich([
        { text: "Optimism reflects aspiration, not adoption. ", bold: true },
        "Lower-income countries with rapidly-growing digital sectors may see AI as a leapfrog opportunity, even if their current adoption rates are low.",
      ]),
      P("Either way, this means \"do people like AI?\" and \"do people use AI?\" are measuring fundamentally different things at the country level — a useful refinement for any future research design."),

      // ===== 6. Correlation table =====
      H1("6. Correlation Summary"),
      P("All correlations below were computed from the merged country-level dataset (38 countries with at least one country-level metric, 34 with adoption data). Full data and Excel formulas are in the Correlations sheet of the accompanying workbook."),
      corrTable(),

      // ===== 7. Hypothesis Refined =====
      H1("7. Hypothesis Refined Based on These Findings"),
      Runs([
        { text: "Original hypothesis (still supported): ", bold: true },
        "Generative AI adoption follows existing digital-divide patterns. Within the US (and the EU), younger, more-educated, higher-income individuals adopt at much higher rates. Across countries, adoption correlates strongly with GDP per capita (r = +0.72) and internet penetration (r = +0.61).",
      ]),
      Runs([
        { text: "Added refinement: ", bold: true },
        "Public sentiment about AI is ",
        { text: "decoupled from", italic: true },
        " — and in fact inversely related to — actual usage. Wealthy, high-adoption countries are the most skeptical of AI; lower-income, lower-adoption countries are the most optimistic. Future analysis should treat \"adoption\" and \"optimism\" as distinct outcomes with potentially different drivers.",
      ]),

      // ===== 8. Limitations =====
      H1("8. Limitations and Next Steps"),
      bullet("Source comparability: Different surveys define adoption differently. The Microsoft Diffusion measure and Eurostat measure are the closest match (both are population-share, last-3-months); the Ipsos and Pew measures are have-ever-used, which inflates numbers."),
      bullet("US income data: Pew's adult-income breakdown for 2025 is not publicly tabulated in headline reports; income evidence is strongest for teens. Adult income evidence comes via Brookings (qualitative) rather than precise tabulations."),
      bullet("Selection bias: Most country-level surveys are online-only, biasing samples toward already-connected populations and likely overstating adoption in lower-internet countries."),
      bullet("Causality: All findings are correlational. The hypothesis posits that GDP and internet enable adoption, but reverse causation (early AI adoption boosting productivity, hence GDP) is plausible at the margin and cannot be ruled out from cross-section data."),
      bullet("Next analytical steps: (a) regression analysis controlling GDP, internet, and education simultaneously; (b) within-country time-series to test whether the digital divide is widening or narrowing; (c) deeper US income breakdowns from the Census HTOPS public-use file; (d) sectoral analysis from Anthropic's Economic Index to test whether high-skill occupations dominate adoption."),

      // ===== 9. Files =====
      H1("9. Accompanying Files"),
      bulletRich([{ text: "AI_Adoption_Research_Data.xlsx", bold: true }, " — full data compilation across 9 sheets, including the merged Master_Country sheet and live-formula Correlations sheet."]),
      bulletRich([{ text: "chart1–chart7 PNGs", bold: true }, " — high-resolution versions of every figure in this report."]),
      bulletRich([{ text: "build_spreadsheet.py / explore.py", bold: true }, " — reproducible Python scripts that generated the data and analysis."]),
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(OUT, buf);
  console.log("Wrote", OUT);
});
