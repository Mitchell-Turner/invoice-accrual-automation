# Invoice Accrual Automation

This project automates the monthly processing of PeopleSoft invoice exports for Medicare accruals. It classifies invoices, applies business logic, flags outliers, and generates fully formatted, audit-ready Excel reports.

Additionally, it calculates MMP (Medicare-Medicaid Plan) reclass allocations from a reference table, since MMP costs are embedded in Medicare invoices and not explicitly separated.

---

## 🔧 Features

- ✅ Loads the latest PeopleSoft export automatically
- ✅ Filters and labels invoice rows based on source, contract, and amount
- ✅ Flags duplicate invoices and high outliers
- ✅ Summarizes invoice costs by type
- ✅ Allocates MMP reclass values from a reference file
- ✅ Outputs formatted Excel files with multiple sheets

---

## 📁 Folder Structure

PeopleSoft_Invoice_Reports/ ├── raw_data/ # Drop monthly PeopleSoft export here ├── processed_reports/ # Output reports are saved here ├── MMP_Reclass_Ref/ │ └── MMP_Reclass_Ref.xlsx # Contract allocation reference file


---

## 🧾 PeopleSoft Query Instructions

Use the following parameters to pull the monthly data:

UNIT: XXXXX
YEAR: [Enter Year]
BEG PERIOD: [Enter Month Number]
END PERIOD: [Enter Month Number]
ACCOUNT: XXXX
DEPT: XXX
CONTRACT: %
PRODUCT: %


Then save the file as Excel and drop it into the `raw_data/` folder. The script will process the latest file automatically.

---

## 📦 Outputs

Each run produces:

- `Invoice_Report_YYYY_MM.xlsx`  
  - Summary tab (total by label)  
  - Full Data tab  
  - Flags tab (duplicates & outliers)

- `MMP_Reclass_Allocations_YYYY_MM.xlsx`  
  - Filled allocation sheet with formatting and totals

---

## 🛠️ Technologies Used

- Python 3
- pandas
- xlsxwriter

---

## 🧪 Sample Data

This repo includes two sample Excel files to demonstrate the pipeline:
- `sample_data_invoice.xlsx`: Contains 5 fake invoice lines across contracts and sources
- `sample_data_mmp_ref.xlsx`: A mock MMP reference table with allocation percentages

These can be used to test the script and understand the expected format.



## ⚠️ Disclaimer

This project contains **sample data only** and does not expose any real vendors, contracts, or PHI. All logic and structure are generic and anonymized for demonstration purposes.

---
