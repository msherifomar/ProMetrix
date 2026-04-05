# ProMetrix — GitHub & Zenodo Publication Guide
## Step-by-Step Instructions for Dr. M S Omar

---

## PART 1 — PREPARE YOUR FILES LOCALLY

Before touching GitHub, make sure your folder contains exactly these files:

```
ProMetrix/
├── ProMetrix.py              ← your main script (fixed version)
├── README.md                 ← drafted for you
├── requirements.txt          ← drafted for you
├── LICENSE                   ← drafted for you (MIT)
├── CHANGELOG.md              ← drafted for you
├── CITATION.cff              ← drafted for you
└── sample_data/
    └── crown_study_demo.xlsx ← generated for you
```

---

## PART 2 — GITHUB (click by click)

### Step 1 — Create your GitHub account (skip if you already have one)
1. Go to https://github.com
2. Click **Sign up**
3. Enter your email, create a password, choose a username (e.g., `msomar-iu`)
4. Verify your email

---

### Step 2 — Create a new repository
1. Log in to GitHub
2. Click the **+** button in the top-right corner
3. Click **New repository**
4. Fill in:
   - **Repository name:** `ProMetrix`
   - **Description:** `Statistical Analysis & Publication-Quality Visualization for Dental and Biomedical Research`
   - **Public** ← must be Public for Zenodo to index it
   - ✅ Check **Add a README file** — NO, leave it unchecked (you already have one)
   - **License:** leave as None (you already have a LICENSE file)
5. Click **Create repository**

---

### Step 3 — Upload your files
1. You are now on your new empty repository page
2. Click **uploading an existing file** (the link in the middle of the page)
3. **Drag and drop ALL your files** into the upload area:
   - `ProMetrix.py`
   - `README.md`
   - `requirements.txt`
   - `LICENSE`
   - `CHANGELOG.md`
   - `CITATION.cff`
4. For the `sample_data` folder: you cannot drag a folder directly.
   - Click **choose your files**
   - Navigate into your `sample_data` folder
   - Select `crown_study_demo.xlsx`
   - GitHub will ask for the path — type: `sample_data/crown_study_demo.xlsx`
5. Scroll down to **Commit changes**
6. In the first box type: `Initial release v2.2.0`
7. Click **Commit changes**

---

### Step 4 — Create a Release (this is what Zenodo archives)
1. On your repository page, look at the right sidebar
2. Click **Releases** (or **Create a new release**)
3. Click **Draft a new release**
4. Click **Choose a tag**
5. Type `v2.2.0` in the box and click **Create new tag: v2.2.0**
6. **Release title:** `ProMetrix Analytics v2.2.0 — Initial Public Release`
7. In the description box, paste this:

```
## ProMetrix Analytics v2.2.0 — Initial Public Release

Statistical Analysis & Publication-Quality Visualization for Dental and Biomedical Research

### What's included
- Adaptive statistical pipeline (Shapiro-Wilk, Levene, ANOVA/Kruskal-Wallis, Welch ANOVA)
- Pairwise comparisons with Bonferroni correction and Hedges' g effect sizes
- Publication-quality figures: Raincloud, Bar, Scatter, Heatmap
- Full Word report export (.docx)
- Auto-generated manuscript Methods and Results text
- Sample crown study dataset included

### Requirements
pip install pandas numpy scipy matplotlib seaborn openpyxl PyQt5
pip install python-docx  # for Word export
```

8. Click **Publish release**

✅ Your software is now live on GitHub.

---

## PART 3 — ZENODO (click by click)

### Step 1 — Create a Zenodo account
1. Go to https://zenodo.org
2. Click **Log in** in the top right
3. Click **Log in with GitHub** ← use this option, it connects automatically
4. Authorize Zenodo to access your GitHub account

---

### Step 2 — Connect your GitHub repository to Zenodo
1. After logging in, click your name in the top right
2. Click **GitHub** in the dropdown menu
3. You will see a list of all your GitHub repositories
4. Find **ProMetrix** in the list
5. Toggle the switch next to it to **ON** (it turns blue)

---

### Step 3 — Get your DOI
1. Now go back to GitHub
2. Go to your ProMetrix repository
3. Click **Releases** in the right sidebar
4. Click **Draft a new release** (or if you already published v2.2.0, this triggers automatically)

   **If you already published the release in Part 2:**
   - Zenodo may have already archived it automatically
   - Go back to Zenodo → click your name → click **Upload** or **My records**
   - Look for ProMetrix — it should appear within a few minutes

   **If it did not appear automatically:**
   - Go back to GitHub → your ProMetrix repo → Releases
   - Click your v2.2.0 release
   - Click **Edit** (pencil icon)
   - Click **Update release** (no changes needed — this re-triggers Zenodo)

5. On Zenodo, click **My records** or **My uploads**
6. Click on **ProMetrix Analytics**
7. Your DOI will be shown — it looks like: `10.5281/zenodo.XXXXXXX`

---

### Step 4 — Fill in the Zenodo metadata (important for citations)
1. On the Zenodo record page, click **Edit**
2. Fill in these fields:
   - **Title:** `ProMetrix Analytics: Statistical Analysis & Publication-Quality Visualization`
   - **Authors:** `Omar, M S` — Affiliation: `Indiana University School of Dentistry`
   - **Description:** copy from the README Overview section
   - **Resource type:** Software
   - **Version:** 2.2.0
   - **Keywords:** statistics, prosthodontics, dental research, effect size, Python, open source, nonparametric
   - **License:** MIT
3. Click **Save** then **Publish**

---

### Step 5 — Update your files with the real DOI
Once you have your DOI (e.g., `10.5281/zenodo.1234567`):

1. Open `README.md` — replace `XXXXXXX` with your real number in two places
2. Open `CITATION.cff` — replace `XXXXXXX` with your real number
3. Go to GitHub → ProMetrix repository
4. Click on `README.md`
5. Click the **pencil icon** (Edit this file)
6. Make the replacement
7. Scroll down → click **Commit changes**
8. Repeat for `CITATION.cff`

---

## PART 4 — HOW TO CITE IT IN YOUR PAPERS

### In your Methods section:
> Statistical analyses were performed using ProMetrix Analytics
> (v2.2; Omar, 2025). The software implements adaptive omnibus
> testing based on Shapiro-Wilk normality and Levene variance
> assessments, with Kruskal-Wallis H tests and Mann-Whitney U
> post-hoc comparisons (Bonferroni correction) when assumptions
> were violated. Effect sizes are reported as Hedges' g with
> bias correction and 95% bootstrap confidence intervals
> (2,000 resamples). Welch's ANOVA was conducted as a
> sensitivity analysis.

### Reference entry:
> Omar MS. ProMetrix Analytics: Statistical Analysis &
> Publication-Quality Visualization (Version 2.2.0) [Software].
> 2025. Available from: https://doi.org/10.5281/zenodo.XXXXXXX

---

## PART 5 — CHECKLIST BEFORE PUBLISHING

- [ ] ProMetrix.py is the fixed version (n<3 guard, statsmodels removed)
- [ ] README.md is in the folder
- [ ] requirements.txt is in the folder
- [ ] LICENSE is in the folder
- [ ] CHANGELOG.md is in the folder
- [ ] CITATION.cff is in the folder
- [ ] sample_data/crown_study_demo.xlsx is in the folder
- [ ] GitHub repository is set to PUBLIC
- [ ] Zenodo toggle is ON for the ProMetrix repo
- [ ] Release v2.2.0 is published on GitHub
- [ ] Zenodo DOI is confirmed
- [ ] README and CITATION.cff updated with real DOI

---

*Guide prepared by Claude for Dr. M S Omar — Indiana University School of Dentistry*
