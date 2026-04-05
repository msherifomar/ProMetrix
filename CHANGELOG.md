# Changelog

All notable changes to ProMetrix Analytics are documented here.  
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [2.2.0] — 2025 (Initial Public Release)

### Added
- Full desktop GUI built with PyQt5
- Adaptive statistical pipeline: Shapiro-Wilk normality + Levene variance testing with automatic ANOVA / Kruskal-Wallis selection
- Welch ANOVA sensitivity analysis for all outcomes
- Mann-Whitney U pairwise comparisons with Bonferroni correction
- Hedges' g effect size with bias correction and 95% bootstrap confidence intervals (2,000 resamples)
- Rank-biserial correlation as nonparametric effect size
- Epsilon-squared (Kruskal-Wallis) and omega-squared (ANOVA) omnibus effect sizes
- Minimum detectable effect (MDE) calculation at 80% power for all pairwise comparisons
- Pearson correlation matrix (pooled across groups)
- Coefficient of variation (CV%) table with automatic flagging of CV > 50%
- Raincloud plot (half-violin + boxplot + jitter + mean diamond)
- Bar plot (mean ± SD with individual data overlay)
- Scatter plot with linear regression, 95% CI ribbon, and per-group Pearson r
- Correlation heatmap (lower triangle, Pearson r with significance annotations)
- Auto-generated manuscript Methods and Results text
- Full `.docx` report export (TNR 12pt double-spaced body; TNR 10pt single-spaced tables)
- Excel (multi-sheet) and CSV data loading
- Color presets: Default, Clinical, Pastel, Grayscale, Journal (colorblind-friendly)
- Per-group color customization via color picker
- PNG, PDF, TIFF, SVG figure export at 300 DPI
- CSV statistics export
- Built-in crown study demo dataset (5 groups, 4 outcomes)

---

## Planned for v2.3.0

- QQ plots for per-group normality visualization
- Forest plot of all pairwise Hedges' g with confidence intervals
- Spearman correlation option
- Benjamini-Hochberg FDR correction as alternative to Bonferroni
- Paired test support (Wilcoxon signed-rank) for within-subject designs
- Sample size calculator (reverse MDE: required n given expected effect size)
