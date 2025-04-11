# Invoice Accrual Automation

## Overview

This project automates the monthly processing of PeopleSoft invoice exports for Medicare accruals. It classifies invoices, applies business logic, flags outliers, and generates fully formatted, audit-ready Excel reports.

Additionally, it calculates MMP (Medicare-Medicaid Plan) reclass allocations from a reference table, since MMP costs are embedded in Medicare invoices and not explicitly separated.

## Features

- **Automated Data Processing**: Finds and processes the most recent invoice data file
- **Business Logic Implementation**: Applies complex categorization rules to invoice transactions
- **Anomaly Detection**: Identifies potential duplicates and outliers in invoice data
- **Cost Summarization**: Aggregates and visualizes invoice costs by category
- **MMP Reallocation**: Calculates reclass allocations based on reference percentages
- **Rich Excel Output**: Generates professionally formatted Excel reports with multiple sheets

## System Components

### 1. Invoice Processor

The main script that performs all processing operations:

- Reads the latest PeopleSoft export file from the raw_data directory
- Normalizes and categorizes invoice data
- Identifies statistical outliers and duplicates
- Generates formatted summary reports with visualizations

### 2. MMP Allocation Calculator

A specialized component that handles MMP reclass calculations:

- Takes "Charts & Coding" totals from the invoice processing
- Applies allocation percentages from the reference file
- Calculates adjusted and subset allocations
- Creates a formatted allocation report with visualizations

## Directory Structure

```
/
├── peoplesoft_invoice_processor.py    # Main script
├── PeopleSoft_Invoice_Reports/        # Data directories (created automatically)
│   ├── raw_data/                      # Place raw invoice exports here
│   ├── processed_reports/             # Output will be saved here (by month)
│   └── MMP_Reclass_Ref/               # Reference data
│       └── MMP_Reclass_Ref.xlsx       # Allocation percentages reference
└── invoice_processor.log              # Log file (created when script runs)
```

## Required Data Format

### PeopleSoft Export Files

- Excel files exported from PeopleSoft
- Must include the following fields:
  - Journal Date
  - Invoice identifier
  - Source (AP2 or COR)
  - Contract (1111 or 2222)
  - Line Descr (description of the invoice line)
  - Amount
  - AP Amount

### MMP Reference File

- Located at `PeopleSoft_Invoice_Reports/MMP_Reclass_Ref/MMP_Reclass_Ref.xlsx`
- Must contain columns for `State`, `Contract`, and `% of Payments`
- Must include a row with `State` value of "Total"
- Must include a row with `Contract` value of "Subset"

## PeopleSoft Query Instructions

Use the following parameters to pull the monthly data:

- UNIT: XXXXX
- YEAR: [Enter Year]
- BEG PERIOD: [Enter Month Number]
- END PERIOD: [Enter Month Number]
- ACCOUNT: XXXX
- DEPT: XXX
- CONTRACT: %
- PRODUCT: %

Then save the file as Excel and drop it into the `raw_data/` folder. The script will process the latest file automatically.

## Output Details

### Invoice Report (`Invoice_Report_YYYY_MM.xlsx`)

- **Summary Sheet**: 
  - Aggregated totals by invoice category with professional formatting
  - Pie chart visualization of invoice distribution
  - Color-coded rows for different categories
  - Grand total calculation

- **Full Data Sheet**: 
  - Complete processed dataset with categorizations
  - Column filtering for easy data analysis
  - Proper currency formatting

- **Flags Sheet**: 
  - Potential issues - duplicates and statistical outliers
  - Special header with description
  - Color-coded for immediate attention

### MMP Allocation Report (`MMP_Reclass_Allocations_YYYY_MM.xlsx`)

- **MMP Allocation Sheet**: 
  - Professional title and formatting
  - Calculated allocations based on reference percentages
  - Pie chart showing allocation distribution by state
  - Legend explaining color coding
  - Summary information section

## Invoice Categorization Logic

The script categorizes invoices based on the following rules:

| Source | Contract | Amount     | Category Label        |
|--------|----------|------------|----------------------|
| AP2    | 1111     | Any        | Charts & Coding      |
| AP2    | 2222     | Any        | Misc. exp.           |
| COR    | 1111     | Negative   | 1111 Coupa Reversal  |
| COR    | 1111     | Positive   | 1111 Coupa Pending   |
| COR    | 2222     | Negative   | 2222 Coupa Reversal  |
| COR    | 2222     | Positive   | 2222 Coupa Pending   |

## Usage Instructions

```bash
python Run_Invoice_Report.py
```

This will:
- Find the most recent invoice file in the raw_data directory
- Process and categorize all invoice data
- Generate formatted reports in the processed_reports directory
- Show a summary of processed data

## Sample Data

This repository includes sample data files to demonstrate the pipeline:

- `sample_data_invoice.xlsx`: Contains 5 fake invoice lines across contracts and sources
- `sample_data_mmp_ref.xlsx`: A mock MMP reference table with allocation percentages

These can be used to test the script and understand the expected input/output formats without requiring real data.

## Dependencies

- Python 3.7+
- pandas
- xlsxwriter

## Disclaimer

This project contains **sample data only** and does not expose any real vendors, contracts, or PHI. All logic and structure are generic and anonymized for demonstration purposes.
