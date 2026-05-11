# AI Adoption Research

> Is generative AI adoption following — and amplifying — the existing digital divide?

This project compiles eleven primary data sources into a single workbook to test whether generative AI adoption is reinforcing the existing digital divide, both within the United States (across age, income, and education) and across countries (against GDP, internet penetration, and tertiary education).

**Repository:** https://github.com/moh-a-abde/ai-adoption-research
**Course:** Data Preparation and Analysis (Spring 2026), University of St. Thomas
**Author:** Mo Abde

---

## Question and Hypothesis

**Question.** Is generative AI adoption uneven across demographic groups within countries, and uneven across countries — and does the unevenness track existing digital-divide indicators?

**Hypothesis.** Generative AI adoption follows existing digital-divide patterns:
1. **Within countries**, younger, higher-income, more-educated individuals adopt at substantially higher rates.
2. **Across countries**, adoption correlates strongly with GDP per capita and internet penetration.

AI is **amplifying — not closing** — the digital divide.

---

## Main Findings

| Finding | Evidence |
|---|---|
| **Strong wealth–adoption link across countries** | AI adoption × GDP per capita: **Pearson r = +0.72** (n=34 countries) |
| **Adoption rides existing internet infrastructure** | AI adoption × internet penetration: **r = +0.61** (n=34) |
| **Education matters, but less than wealth** | AI adoption × tertiary education: r = +0.51 (n=34) |
| **Steep age gradient inside countries** | US adults under 30 are **5.8x** more likely to use ChatGPT than those 65+ (58% vs. 10%, Pew 2025); EU 16-24 use AI at 64% vs. 7% for 65-74 (Eurostat 2025) |
| **Steep education gradient** | Postgrads are **2.9x** more likely to have used ChatGPT than HS-or-less adults (52% vs. 18%, Pew 2025) |
| **Surprise: optimism is *inversely* related to adoption** | AI adoption × AI optimism: **r = −0.67** (n=16). Countries that use AI most (Netherlands, France, US) are most skeptical; countries most enthusiastic (China 83%, Indonesia 80%) use it less. |

**Bottom line.** Hypothesis supported. AI adoption is amplifying the digital divide. A bonus finding emerged: public sentiment about AI and actual usage of AI are not the same thing — they correlate negatively at the country level.

See `notebooks/01_exploration.ipynb` for the full analysis and `reports/AI_Adoption_Research_Report.docx` for the narrative write-up.

---

## Repository Structure

```
.
├── data/
│   └── AI_Adoption_Research_Data.xlsx   # 10-sheet master workbook (all source data + correlations)
├── notebooks/
│   └── 01_exploration.ipynb             # Main analysis notebook (start here)
├── figures/                             # 7 chart PNGs used in the report
├── reports/
│   ├── AI_Adoption_Research_Report.docx # Narrative research report
│   └── AI_Adoption_Research_Report.pdf
├── slides/
│   └── AI_Adoption_Presentation.pptx    # Class presentation
├── build_spreadsheet.py                 # Rebuilds the master workbook from compiled source values
├── explore.py                           # Computes correlations + regenerates all 7 PNG figures
├── build_report.js                      # Regenerates the Word report from figures
├── requirements.txt                     # Python dependencies
├── package.json                         # Node dependencies (for build_report.js)
└── README.md
```

---

## Data Sources

All eleven sources are catalogued on the `README` sheet of the workbook with full citations and URLs. Headline list:

| # | Source | What it provides |
|---|---|---|
| 1 | Pew Research Center (June 2025) | US adult ChatGPT use by age and education |
| 2 | Pew Research Center (Dec 2025) | US teen AI use by household income |
| 3 | Pew Research Center (Oct 2025) | US workers using AI on the job |
| 4 | Brookings / Real-Time Population Survey (2024) | US adult generative AI demographic patterns |
| 5 | Stanford HAI 2025 AI Index, Public Opinion chapter | Country-level AI optimism |
| 6 | Ipsos AI Monitor 2024 | 32-country attitudes & understanding |
| 7 | Microsoft AI Diffusion Report 2025 H2 | Country-level adoption % (top 30 + 4 large economies) |
| 8 | Eurostat (Dec 2025) | EU GenAI use by country and age |
| 9 | Anthropic Economic Index (Sept 2025) | Per-capita Claude usage by country |
| 10 | Visual Capitalist / GPO-AI 2024 | Daily/weekly ChatGPT use across 21 countries |
| 11 | World Bank / IMF (2024) | GDP per capita PPP, internet penetration |

---

## Reproduce the Outputs

### 1. Install dependencies

```bash
python3 -m pip install -r requirements.txt
npm install   # only needed if you want to rebuild the Word report
```

### 2. Run the notebook (recommended path)

```bash
jupyter notebook notebooks/01_exploration.ipynb
```

The notebook loads the merged `Master_Country` sheet, computes the correlations, renders the key figures inline, and walks through the interpretation, limitations, and refined hypothesis.

### 3. Or rebuild everything from source

```bash
python3 build_spreadsheet.py        # rebuilds data/AI_Adoption_Research_Data.xlsx
python3 explore.py                  # recomputes correlations + regenerates figures/
node build_report.js                # regenerates reports/AI_Adoption_Research_Report.docx
```

The workbook contains an additional verification step: the `Correlations` sheet computes the same Pearson coefficients live using Excel's `CORREL()` formula, so opening the workbook in Excel/LibreOffice should produce the same numbers as the Python notebook.

---

## Methods

- **Country-level**: Pearson correlation between AI adoption (Microsoft Diffusion 2025 H2) and three macro indicators (GDP per capita PPP from IMF/World Bank 2024, internet penetration from ITU/World Bank 2024, tertiary education completion rate from World Bank). Same approach repeated for the optimism measure (Stanford AI Index 2025 / Ipsos 2024).
- **US demographic**: descriptive comparison of adoption rates by age, education, and household income from Pew Research Center (2025).
- **Verification**: correlations are computed two ways — once in Python (notebook) and once with Excel's `CORREL()` formula (workbook) — and required to match.

---

## Limitations and Scope of Inference

1. **Source comparability.** Different surveys define "adoption" differently (have-ever-used vs. used-in-last-3-months vs. daily). Cross-source comparisons should be read as indicative.
2. **Online sampling bias.** Most country-level surveys are administered online, biasing samples toward already-connected populations.
3. **Correlational only.** Findings are cross-sectional Pearson correlations. Causation cannot be established.
4. **Small n for optimism analysis.** The optimism × adoption relationship rests on only 16 countries with both measures; larger sampling needed to firmly establish the paradox.
5. **Coverage skews OECD.** The 34-country adoption sample under-represents Sub-Saharan Africa and parts of Central Asia.

**Generalization** is supported for OECD and middle-income economies. Lower-income contexts are under-represented in the underlying data and inferences should not be extended to them without additional evidence.
