# Data Projection: Clinical Dataset Simulation

This project simulates synthetic patient data based on real clinical datasets provided in an Excel file. The goal is to project the characteristics of a sample of 50 patients to a hypothetical cohort of 120 patients, preserving the statistical properties of the original data.

## Project Overview

- Input: An Excel file with 4 sheets, each representing a different diagnostic protocol (e.g., CT head, abdomen-pelvis, chest, mammography).
- Output: A new Excel file with 120 simulated patients per sheet, and QQ plots showing the distribution checks.

## Features

- Reads multiple sheets from an Excel file using `pandas`.
- Performs normality tests (Shapiro-Wilk + QQ plot) on continuous variables (`Age`, `Dose`).
- Generates synthetic data:
  - If normal: random sampling from a normal distribution (`numpy.random.normal`)
  - If not normal: bootstrap sampling with replacement
- Categorical variables (`Sex`, `Month`) are simulated using proportional sampling.
- Saves:
  - Simulated datasets in `/output/simulazione_120_pazienti_python.xlsx`
  - QQ plots in `/plots/` as PNG images

## How to Run

1. Install required libraries (if not already installed):

   ```bash
   pip install pandas numpy scipy matplotlib openpyxl
