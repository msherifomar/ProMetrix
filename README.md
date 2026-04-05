# ProMetrix Analytics

**Statistical Analysis & Publication-Quality Visualization for Dental and Biomedical Research**

> *Courtesy of Dr. M S Omar BDS MSc — For Doers, Researchers and Innovators*

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.XXXXXXX.svg)](https://doi.org/10.5281/zenodo.XXXXXXX)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Python 3.8+](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/)
[![Version](https://img.shields.io/badge/version-2.2.0-green.svg)](CHANGELOG.md)

---

## Overview

ProMetrix Analytics is an open-source desktop application for rigorous statistical analysis and publication-ready figure generation. It is designed for researchers in prosthodontics, restorative dentistry, and related biomedical fields who require validated statistical workflows without programming expertise.

The software implements an adaptive analysis pipeline that automatically selects parametric or nonparametric tests based on normality and variance homogeneity testing, generates manuscript-ready results text, and exports a complete Word report in a single click.

![ProMetrix Screenshot](docs/screenshot.png)

---

## Features

- **Adaptive statistical pipeline** — Shapiro-Wilk normality + Levene variance testing → automatic ANOVA or Kruskal-Wallis selection
- **Sensitivity analysis** — Welch ANOVA run in parallel for all outcomes
- **Pairwise comparisons** — Mann-Whitney U with Bonferroni correction
- **Effect sizes** — Hedges' g with bias correction and 95% bootstrap confidence intervals (2,000 resamples); rank-biserial correlation
- **Power analysis** — Minimum detectable effect (MDE) at 80% power for every pairwise comparison
- **Publication figures** — Raincloud plots, bar plots (mean ± SD), scatter with regression, correlation heatmap
- **Tables** — Descriptive statistics, pairwise comparisons, Pearson correlations, coefficient of variation (CV%)
- **Manuscript text generator** — Auto-generates Methods and Results sections formatted for journal submission
- **Full report export** — Single `.docx` file containing all statistics, tables, and figures (TNR 12pt double-spaced body; TNR 10pt single-spaced tables)
- **Flexible data input** — Excel (multi-sheet) or CSV with group column
- **Color presets** — Default, Clinical, Pastel, Grayscale, Journal (colorblind-friendly)

---

## Installation

### Prerequisites

- Python 3.8 or higher
- pip

### Install dependencies

```bash
pip install pandas numpy scipy matplotlib seaborn openpyxl PyQt5
pip install python-docx        # optional — required only for Word export
```

### Run

**In Spyder:**
```python
%run ProMetrix.py
```

**Standalone:**
```bash
python ProMetrix.py
```

---

## Quick Start

1. Launch ProMetrix
2. Click **Load Demo Data (Crown Study)** to explore with built-in data
3. Select an outcome variable from the dropdown
4. Click **Run Full Statistics** to view the full analysis
5. Generate figures using the plot buttons (Raincloud, Bar, Scatter, Heatmap)
6. Click **EXPORT FULL REPORT (.docx)** to generate a complete Word document

---

## Data Format

### Excel (recommended)
One sheet per group. Each sheet contains numeric outcome columns. Column names must match across all sheets.

| Design Time (s) | Volume Removed (mm³) | Surface Deviation (µm) |
|---|---|---|
| 144 | 18.18 | 50.5 |
| 138 | 8.54 | 34.0 |

### CSV
A single file with a `Group` column (also recognized: `Groups`, `Platform`, `Method`, `Type`, `Category`) followed by numeric outcome columns.

| Group | Design Time (s) | Volume Removed (mm³) |
|---|---|---|
| Human Expert | 144 | 18.18 |
| Human Novice | 155 | 7.05 |

A sample dataset is provided: [`sample_data/crown_study_demo.xlsx`](sample_data/crown_study_demo.xlsx)

---

## Statistical Methods

The following statistical pipeline is implemented:

| Step | Method |
|---|---|
| Normality testing | Shapiro-Wilk test (per group) |
| Variance homogeneity | Levene's test |
| Omnibus test (nonparametric) | Kruskal-Wallis H test |
| Omnibus test (parametric) | One-way ANOVA |
| Sensitivity analysis | Welch's ANOVA |
| Pairwise comparisons | Mann-Whitney U (Bonferroni correction) |
| Effect size | Hedges' g (bias-corrected) with 95% bootstrap CI |
| Nonparametric effect size | Rank-biserial correlation |
| Omnibus effect size | Epsilon-squared (KW); Omega-squared (ANOVA) |
| Power analysis | Minimum detectable effect at α = 0.05, power = 0.80 |
| Correlation | Pearson r (pooled across groups) |

---

## Citing ProMetrix

If you use ProMetrix Analytics in your research, please cite:

> Omar, M. S. (2025). *ProMetrix Analytics: Statistical Analysis & Publication-Quality Visualization* (Version 2.2.0) [Software]. Zenodo. https://doi.org/10.5281/zenodo.XXXXXXX

**BibTeX:**
```bibtex
@software{omar_2025_prometrix,
  author       = {Omar, M S},
  title        = {ProMetrix Analytics: Statistical Analysis \& Publication-Quality Visualization},
  version      = {2.2.0},
  year         = {2025},
  publisher    = {Zenodo},
  doi          = {10.5281/zenodo.XXXXXXX},
  url          = {https://doi.org/10.5281/zenodo.XXXXXXX}
}
```

**In-text (Methods section):**
> Statistical analyses were performed using ProMetrix Analytics (v2.2; Omar, 2025). The software implements adaptive omnibus testing (Kruskal-Wallis or one-way ANOVA based on Shapiro-Wilk and Levene testing), Welch ANOVA sensitivity analysis, Mann-Whitney U pairwise comparisons with Bonferroni correction, and Hedges' g effect sizes with 95% bootstrap confidence intervals.

---

## License

MIT License — see [LICENSE](LICENSE) for details.

---

## Author

**Dr. M S Omar BDS MSc**  
Clinical Assistant Professor of Prosthodontics  
Director, Digital Innovation Laboratory  
Indiana University School of Dentistry
