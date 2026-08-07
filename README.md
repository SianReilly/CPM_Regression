# Common Progress Measures — Regression Analysis Dashboard

**Built by Sian Reilly, Data & Intelligence Analyst**
Strategy & Intelligence, Westminster City Council

---

## What this is

This dashboard goes beyond the existing CPM ward scoring to answer three questions:

1. **Which CPM indicators actually matter for life expectancy?** The current scoring treats all indicators equally. This tool uses LASSO regression to identify which ones genuinely predict LE outcomes and which are just duplicating information from others.

2. **Which wards have unexpectedly high or low LE?** The current scoring tells you "Church Street scores low on anxiety." But it doesn't tell you whether that's *surprising* given Church Street's overall profile. Residual analysis flags where the gap between expected and actual performance is largest — pointing to specific areas for targeted intervention.

3. **Is the picture stable across years?** Running the same analysis on 2023-24 and 2024-25 data shows which findings are consistent and which might be noise.

The analysis runs for **male, female, and overall** (average of both) life expectancy simultaneously, and cross-references results with the Index of Multiple Deprivation (IMD 2025).

---

## How to run it

```bash
pip install -r requirements.txt
streamlit run cpm_regression_app.py
```

Upload `normalised_outputs.xlsx` when the app opens.

---

## What each tab shows

| Tab | What you'll find |
|-----|-----------------|
| **How to Read This** | Plain-English explanations of LASSO regression, residual analysis, and R² — start here if you're new to the techniques |
| **Data Overview (EDA)** | Profiles the dataset before modelling: how many indicators, missing data, which indicators have the most variation across wards, and how the CPM pillars correlate with each other |
| **What Explains LE?** | The LASSO results: which indicators were selected for male, female, and overall LE, how strong the model is (R²), and which indicators appear across all three |
| **Unexpected Wards** | Residual analysis: which wards have LE that's higher or lower than their CPM profile would predict, with drill-down into what's driving the gap |
| **Year-on-Year** | Feature stability (which indicators appear in both years), residual shifts, and individual indicator movement per ward |
| **Deprivation & LE** | IMD 2025 profiles + scatter plots crossing deprivation with LE levels and LE *change* — are deprived wards falling further behind? |
| **Explore the Data** | Interactive scatter plots for any indicator against LE, plus the raw data tables |

---

## Key concepts explained

### What is LASSO regression?

Imagine you have 29 CPM indicators and want to know which ones actually matter for life expectancy. You could look at each one individually, but many indicators move together — wards with high poverty also tend to have high worklessness, poor housing, and worse health. So which indicator is actually doing the explaining, and which are just along for the ride?

LASSO regression looks at all indicators simultaneously and **automatically drops** the ones that don't add anything beyond what the others already capture. The output is a short list of indicators that genuinely matter, with everything else filtered out.

This approach comes from the regression techniques covered in my Data Analytics Principles (LD7155) module, adapted here for the CPM context. The EDA pipeline follows the same profile → identify issues → visualise → model structure from the module's workshop notebooks.

### What is residual analysis?

A **residual** is just: **actual value minus predicted value**.

If a ward's actual LE is much lower than what the model predicts given its CPM scores, that's a large negative residual — something specific may be driving poor outcomes beyond the general pattern. Conversely, a large positive residual means something is going *right* that the indicators don't fully capture.

This shifts the conversation from "which wards are worst" (which we already know) to "where are the specific, unexpected gaps where intervention could make a difference."

### What does R² mean?

R² is a number between 0 and 1 that answers: **what share of the variation in life expectancy can our CPM indicators explain?**

- R² of 0.75 = the indicators explain 75% of why LE differs across wards. Strong.
- R² of 0.30 = they only capture 30%. We're missing something important.

### What is "Overall LE"?

The average of the male and female normalised LE scores for each ward. It's not a separate data source — just a combined view that smooths out sex-specific patterns.

---

## Important caveats

- **18 wards is a small sample.** Results are directional signals, not definitive answers. Adding or removing a single ward can shift the results.
- **Life expectancy is lagging.** It's measured over 5-year periods and may not reflect recent CPM changes. Complementary shorter-term health outcomes would strengthen the analysis.
- **Air quality has reverse polarity** (as Damian flagged). Pollution is worst in the wealthiest areas (West End, Marylebone). The normalisation handles scoring direction, but the underlying confound means air quality indicators should be interpreted with extra care.
- **Correlation isn't causation.** These indicators co-occur with LE differences; they don't necessarily cause them.
- **Year-on-year stability matters.** Indicators appearing in only one year should be treated cautiously — only those stable across both years are strong signals.

---

## Data structure

The app expects an Excel file (`normalised_outputs.xlsx`) with these sheets:

| Sheet | What's in it |
|-------|-------------|
| `CPM_2024_2025_Scores` | Ward × indicator matrix with raw values and min-max normalised scores for 2024-25 |
| `CPM_2023_2024_Scores` | Same structure for 2023-24 |
| `Dataset_Overview` | Indicator metadata: names, pillars, polarity, sources |
| `IMD_2025_Ward` | IMD 2025 scores, ranks, and deciles by domain for each ward |

The app reads the **min-max normalised scores** (not raw values) because polarity is already handled in the normalisation (higher score always = better). LASSO applies z-scoring on top of this to ensure fair penalty across indicators — this doesn't distort anything, it's a standard preprocessing step.

---

## Results summary (from test data)

| Year | Target | R² | Indicators selected |
|------|--------|-----|---------------------|
| 2024-25 | Male LE | 0.64 | 4 out of 29 |
| 2024-25 | Female LE | 0.46 | 4 out of 29 |
| 2024-25 | Overall LE | 0.51 | 6 out of 29 |
| 2023-24 | Male LE | 0.50 | 13 out of 32 |
| 2023-24 | Female LE | 0.38 | 18 out of 32 |
| 2023-24 | Overall LE | 0.33 | 3 out of 32 |

IMD × Male LE correlation: **−0.74** (strong negative — more deprived wards have lower LE, as expected).

---

## For stakeholders

Every chart has a **⬇ Download as slide** button that exports it as a single-slide PPTX — ready to paste into a presentation.

---

*Questions or feedback: Sian Reilly (sreilly1@westminster.gov.uk)*
