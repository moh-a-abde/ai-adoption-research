# AI Adoption Research

This project studies whether generative AI adoption is reinforcing the digital divide. It compares adoption across U.S. demographic groups and across countries, then relates adoption to indicators such as GDP per capita, internet penetration, tertiary education, and public optimism about AI. The data comes from public sources compiled into the project workbook, including Pew Research Center, Brookings, Microsoft AI Diffusion reporting, Stanford HAI/Ipsos, Eurostat, Anthropic, Visual Capitalist, and World Bank/IMF indicators. Repository: https://github.com/moh-a-abde/ai-adoption-research

## Repository Contents

- `data/AI_Adoption_Research_Data.xlsx` contains the compiled workbook and source notes.
- `figures/` contains each chart as its own PNG file.
- `reports/` contains the generated Word and PDF reports.
- `build_spreadsheet.py` rebuilds the workbook.
- `explore.py` computes summary correlations and regenerates the figures.
- `build_report.js` regenerates the Word report from the figures.

## Reproduce The Outputs

Install Python dependencies:

```bash
python3 -m pip install -r requirements.txt
```

Install Node dependencies:

```bash
npm install
```

Regenerate the workbook, figures, and Word report:

```bash
python3 build_spreadsheet.py
python3 explore.py
npm run build:report
```
