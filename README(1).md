# Data Cleaning & Reporting Automation

## 1. Project Overview

This project automates a complete retail-sales data cleaning, validation, analysis, and reporting workflow using Python.

The workflow takes a raw retail sales CSV file, cleans inconsistent and missing data, validates the resulting dataset, performs business analysis, and generates both cleaned data files and an Excel report with visual summaries.

## 2. Task Requirements Covered

This project addresses the internship requirements:

- Use Python, Excel, or Power BI for automation
- Handle missing values
- Handle duplicates
- Handle inconsistent data
- Generate automated reports
- Generate visual summaries
- Improve reporting efficiency through a reusable workflow

## 3. Dataset

Input file:

`retail_store_sales.csv`

Dataset size:

- Rows: 12,575
- Columns: 11
- Date range: 2022-01-01 to 2025-01-18
- Customers: 25
- Categories: 8
- Locations: 2

Original columns:

- Transaction ID
- Customer ID
- Category
- Item
- Price Per Unit
- Quantity
- Total Spent
- Payment Method
- Location
- Transaction Date
- Discount Applied

## 4. Data Quality Findings

The raw dataset contained:

- 7,229 missing cells
- 4,199 missing Discount Applied values
- 1,213 missing Item values
- 609 missing Price Per Unit values
- 604 missing Quantity values
- 604 missing Total Spent values
- 0 duplicate rows
- 0 duplicate Transaction IDs

## 5. Cleaning Approach

The automation performs the following steps:

1. Standardises text fields and null-like values.
2. Removes exact duplicate rows.
3. Removes duplicate Transaction IDs while keeping the first occurrence.
4. Converts dates and numeric fields to appropriate data types.
5. Replaces missing Discount Applied values with `Unknown` rather than making an unsupported assumption.
6. Recovers Item values only when a unique Category + Price mapping is available.
7. Labels Item values as `Unknown Item` when they cannot be safely recovered.
8. Recovers missing Price Per Unit from Total Spent / Quantity when mathematically available.
9. Recovers Quantity from Total Spent / Price when mathematically determined.
10. Imputes remaining missing Quantity values using the category median.
11. Recalculates missing Total Spent as Price Per Unit × Quantity.
12. Records recovery/imputation actions in a separate audit file.

Estimated quantity rows are explicitly flagged rather than hidden.

## 6. Validation Results

Final validation:

| Check | Result |
|---|---:|
| Raw rows | 12,575 |
| Clean rows | 12,575 |
| Raw missing cells | 7,229 |
| Clean missing cells | 0 |
| Duplicate rows | 0 |
| Duplicate Transaction IDs | 0 |
| Invalid dates | 0 |
| Negative price/quantity/total | 0 |
| Total != Price × Quantity | 0 |
| Estimated quantity rows | 604 |

The cleaning process therefore preserves all 12,575 records while producing a complete validated business dataset.

## 7. Automation Outputs

The workflow generates:

### Clean_Retail_Store_Sales.csv

The final cleaned business dataset containing the original 11 business columns.

### Cleaning_Audit_Flags.csv

A provenance file showing which records had values recovered, imputed, or recalculated.

Audit fields include:

- Price_Recovered
- Item_Recovered
- Quantity_Imputed
- Total_Recalculated
- Is_Estimated

### Retail_Data_Automated_Report.xlsx

An automated Excel report containing:

- Executive Summary
- Data Quality
- Cleaning Actions
- Category Analysis
- Payment Analysis
- Location Analysis
- Customer Analysis
- Monthly Trend
- Clean Data
- Audit Flags

The report also contains automated charts for category revenue, payment-method revenue share, location revenue, and monthly revenue trend.

## 8. Business Summary

The cleaned dataset contains:

- Total Revenue: 1,635,694.50
- Verified Revenue excluding estimated rows: 1,552,071.00
- Estimated Revenue Share: 5.1%
- Total Transactions: 12,575
- Unique Customers: 25
- Average Transaction Value: 130.08
- Total Quantity Sold: 69,828

## 9. Project Structure

```text
Data-Cleaning-Reporting-Automation/
│
├── task4_automation_final_v2.py
├── retail_store_sales.csv
├── Clean_Retail_Store_Sales.csv
├── Cleaning_Audit_Flags.csv
├── Retail_Data_Automated_Report.xlsx
├── README.md
└── Task4_Data_Cleaning_Automation.ipynb
```

## 10. How to Run

Install the required Python packages:

```bash
pip install pandas openpyxl
```

Run the automation:

```bash
python task4_automation_final_v2.py
```

The script reads `retail_store_sales.csv` and generates the cleaned CSV, audit flags CSV, and automated Excel report.

## 11. Limitations

- The dataset contains one retail-sales data source and does not include external factors such as promotions, weather, or supplier information.
- Some missing quantities cannot be known exactly, so category-median imputation is used and those rows are flagged as estimated.
- Missing discount status is labelled `Unknown` rather than assuming whether a discount was applied.
- The report is designed for reproducible data cleaning and reporting rather than real-time production deployment.

## 12. Conclusion

This project demonstrates a reusable Python workflow for data preprocessing, quality validation, business analysis, and automated reporting. The process preserves the original record count, removes data-quality issues, documents estimation decisions, and produces ready-to-use CSV and Excel outputs.
