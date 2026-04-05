ProMetrix Analytics
Statistical Analysis & Publication-Quality Visualization for Dental and Biomedical Research

Courtesy of Dr. M S Omar BDS MSc — For Doers, Researchers and Innovators

Show Image
Show Image
Show Image
Show Image
Overview
ProMetrix Analytics is an open-source desktop application for rigorous statistical analysis and publication-ready figure generation. It is designed for researchers in prosthodontics, restorative dentistry, and related biomedical fields who require validated statistical workflows without programming expertise.
The software implements an adaptive analysis pipeline that automatically selects parametric or nonparametric tests based on normality and variance homogeneity testing, generates manuscript-ready results text, and exports a complete Word report in a single click.
Features

Adaptive statistical pipeline — Shapiro-Wilk normality + Levene variance testing → automatic ANOVA or Kruskal-Wallis selection
Sensitivity analysis — Welch ANOVA run in parallel for all outcomes
Pairwise comparisons — Mann-Whitney U with Bonferroni correction
Effect sizes — Hedges' g with bias correction and 95% bootstrap confidence intervals (2,000 resamples); rank-biserial correlation
Power analysis — Minimum detectable effect (MDE) at 80% power for every pairwise comparison
Publication figures — Raincloud plots, bar plots (mean ± SD), scatter with regression, correlation heatmap
Tables — Descriptive statistics, pairwise comparisons, Pearson correlations, coefficient of variation (CV%)
Manuscript text generator — Auto-generates Methods and Results sections formatted for journal submission
Full report export — Single .docx file containing all statistics, tables, and figures (TNR 12pt double-spaced body; TNR 10pt single-spaced tables)
Flexible data input — Excel (multi-sheet) or CSV with group column
Color presets — Default, Clinical, Pastel, Grayscale, Journal (colorblind-friendly)

Installation
pip install pandas numpy scipy matplotlib seaborn openpyxl PyQt5
pip install python-docx
Run
In Spyder: %run ProMetrix.py
Standalone: python ProMetrix.py
Quick Start

Launch ProMetrix
Click Load Demo Data (Crown Study) to explore with built-in data
Select an outcome variable from the dropdown
Click Run Full Statistics to view the full analysis
Generate figures using the plot buttons (Raincloud, Bar, Scatter, Heatmap)
Click EXPORT FULL REPORT (.docx) to generate a complete Word document

Data Format
Excel: One sheet per group. Column names must match across all sheets.
CSV: A single file with a Group column followed by numeric outcome columns.
A blank data entry template is provided: ProMetrix_Data_Template.xlsx
Statistical Methods
Normality testing: Shapiro-Wilk test per group
Variance homogeneity: Levene's test
Omnibus nonparametric: Kruskal-Wallis H test
Omnibus parametric: One-way ANOVA
Sensitivity analysis: Welch's ANOVA
Pairwise comparisons: Mann-Whitney U with Bonferroni correction
Effect size: Hedges' g bias-corrected with 95% bootstrap CI
Nonparametric effect size: Rank-biserial correlation
Omnibus effect size: Epsilon-squared (KW) and Omega-squared (ANOVA)
Power analysis: Minimum detectable effect at alpha 0.05 power 0.80
Correlation: Pearson r pooled across groups
Citing ProMetrix
Omar, M. S. (2026). ProMetrix Analytics: Statistical Analysis & Publication-Quality Visualization (Version 2.2.0) [Software]. Zenodo. https://doi.org/10.5281/zenodo.19429680
BibTeX:
@software{omar_2026_prometrix, author = {Omar, M S}, title = {ProMetrix Analytics: Statistical Analysis & Publication-Quality Visualization}, version = {2.2.0}, year = {2026}, publisher = {Zenodo}, doi = {10.5281/zenodo.19429680}, url = {https://doi.org/10.5281/zenodo.19429680} }
In-text Methods section:
Statistical analyses were performed using ProMetrix Analytics (v2.2; Omar, 2026). The software implements adaptive omnibus testing based on Shapiro-Wilk normality and Levene variance assessments, with Kruskal-Wallis H tests and Mann-Whitney U post-hoc comparisons (Bonferroni correction) when assumptions were violated. Effect sizes are reported as Hedges' g with bias correction and 95% bootstrap confidence intervals (2,000 resamples). Welch's ANOVA was conducted as a sensitivity analysis.
License
MIT License
Author
Dr. M S Omar BDS MSc
Clinical Assistant Professor of Prosthodontics
Director, Digital Innovation Laboratory
Indiana University School of Dentistry
