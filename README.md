# Healthcare Analytics: Cleaning, EDA & Billing Prediction

**Author:** Karan Pratap

## Overview
End-to-end analytics project on a hospital admissions dataset:
1. **Data cleaning**: duplicates, invalid (negative) billing, name formatting, date parsing, feature engineering
2. **EDA**: 11 charts covering demographics, conditions, billing, insurance, admission trends and correlations
3. **Machine learning**:
   - Regression (Linear Regression, Decision Tree, Random Forest) to predict `Billing Amount`
   - Random Forest classification to predict `Test Results`
4. **Recommendations** based on the findings (see project report)

## Dataset
Healthcare Dataset (Kaggle): https://www.kaggle.com/datasets/prasad22/healthcare-dataset
Download `healthcare_dataset.csv` and place it in the project root (same folder as the notebook).
Note: the dataset is synthetic, so low model scores are expected and are discussed in the report.

## Tech Stack
Python 3.9+, pandas, NumPy, Matplotlib, Seaborn, scikit-learn, Jupyter, python-docx

## Project Structure
```
.
├── KaranPratap_HealthcareAnalytics.ipynb   # main notebook
├── build_report.py                         # generates the Word report from notebook results
├── KaranPratap_ProjectReport.docx          # project report
├── requirements.txt
├── README.md
├── healthcare_dataset.csv                  # (download from Kaggle)
└── figures/                                # charts saved by the notebook
```

## How to Run
```bash
git clone <your-repo-link>
cd <repo-folder>
pip install -r requirements.txt
jupyter notebook KaranPratap_HealthcareAnalytics.ipynb   # Run All cells
python build_report.py                                   # creates KaranPratap_ProjectReport.docx
```
Running the notebook creates the `figures/` folder and `results.json`, which `build_report.py` uses to fill the report.

## Key Findings
See the notebook's final section and the project report for results, model metrics and recommendations.
