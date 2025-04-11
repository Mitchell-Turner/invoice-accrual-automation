# Invoice Accrual Automation

<img src="https://via.placeholder.com/1200x300?text=Invoice+Accrual+Automation" alt="Project Banner" width="100%">

A robust Python solution that automates the monthly processing of PeopleSoft invoice exports for Medicare accruals. This tool classifies invoices, applies sophisticated business logic, flags outliers, and generates fully formatted, audit-ready Excel reports with visualizations.

Additionally, it calculates MMP (Medicare-Medicaid Plan) reclass allocations from a reference table, since MMP costs are embedded in Medicare invoices and not explicitly separated in the source data.

## ✨ Features

| Feature | Description |
|---------|-------------|
| 🔍 **Automatic Detection** | Locates and processes the latest PeopleSoft export file |
| 🏷️ **Smart Classification** | Labels invoice rows based on source, contract, and amount |
| ⚠️ **Quality Control** | Identifies duplicate invoices and statistical outliers |
| 📊 **Cost Analysis** | Summarizes invoice costs by category with visualizations |
| 💰 **MMP Allocation** | Calculates reclass values from reference percentages |
| 📑 **Professional Reports** | Generates formatted Excel workbooks with multiple sheets |

## 📂 Project Structure

```
PeopleSoft_Invoice_Reports/
├── raw_data/                  # Place monthly PeopleSoft exports here
├── processed_reports/         # Output reports organized by month
│   └── YYYY_MM/               # Month-specific output folders
│       ├── Invoice_Report_YYYY_MM.xlsx
│       └── MMP_Reclass_Allocations_YYYY_MM.xlsx
├── MMP_Reclass_Ref/
│   └── MMP_Reclass_Ref.xlsx   # Contract allocation reference file
└── peoplesoft_invoice_processor.py  # Main processing script
```

## 🚀 Getting Started

### Prerequisites

- Python 3.7+
- Required packages: pandas, xlsxwriter

### Installation

```bash
# Clone this repository
git clone https://github.com/yourusername/invoice-accrual-automation.git

# Navigate to the project directory
cd invoice-accrual-automation

# Install required packages
pip install pandas xlsxwriter
```

### PeopleSoft Query Instructions

Use the following parameters to pull the monthly data:

- **UNIT**: XXXXX
- **YEAR**: [Enter Year]
- **BEG PERIOD**: [Enter Month Number]
- **END PERIOD**: [Enter Month Number]
- **ACCOUNT**: XXXX
- **DEPT**: XXX
- **CONTRACT**: %
- **PRODUCT**: %

Save the file as Excel and place it in the `raw_data/` folder. The script will automatically process the most recent file.

### Running the Process

```bash
python peoplesoft_invoice_processor.py
```

## 📊 Output Reports

Each monthly processing run generates two professional Excel reports:

### 1. Invoice Report (`Invoice_Report_YYYY_MM.xlsx`)

<img src="https://via.placeholder.com/800x400?text=Invoice+Report+Screenshot" alt="Invoice Report" width="90%">

- **Summary Sheet**: 
  - Aggregated totals by invoice category with professional formatting
  - Pie chart visualization of invoice distribution
  - Color-coded rows for different categories
  - Grand total calculation

- **Full Data Sheet**: 
  - Complete processed dataset with categorizations
  - Column filtering for easy data analysis
  - Proper currency formatting
  - Frozen header row for easier navigation

- **Flags Sheet**: 
  - Potential issues - duplicates and statistical outliers
  - Special header with description
  - Color-coded for immediate attention

### 2. MMP Allocation Report (`MMP_Reclass_Allocations_YYYY_MM.xlsx`)

<img src="https://via.placeholder.com/800x400?text=MMP+Allocation+Report+Screenshot" alt="MMP Allocation Report" width="90%">

- **MMP Allocation Sheet**: 
  - Professional title and formatting
  - Calculated allocations based on reference percentages
  - Pie chart showing allocation distribution by state
  - Legend explaining color coding
  - Summary information section

## 🧪 Sample Data

This repository includes sample data files to demonstrate the pipeline:

- `sample_data_invoice.xlsx`: Contains 5 fake invoice lines across contracts and sources
- `sample_data_mmp_ref.xlsx`: A mock MMP reference table with allocation percentages

These can be used to test the script and understand the expected input/output formats without requiring real data.

### How to Use Sample Data

1. Copy the sample files to their respective directories:
   ```bash
   cp sample_data_invoice.xlsx PeopleSoft_Invoice_Reports/raw_data/
   cp sample_data_mmp_ref.xlsx PeopleSoft_Invoice_Reports/MMP_Reclass_Ref/MMP_Reclass_Ref.xlsx
   ```

2. Run the processor:
   ```bash
   python peoplesoft_invoice_processor.py
   ```

3. Check the `PeopleSoft_Invoice_Reports/processed_reports/` directory for the generated output files.

## 🛠️ Technical Implementation

- **Python 3**: Core programming language
- **pandas**: Data manipulation and analysis
- **xlsxwriter**: Advanced Excel report generation
- **Object-oriented design**: Modular, maintainable code structure

## ⚠️ Disclaimer

This project contains **sample data only** and does not expose any real vendors, contracts, or PHI. All logic and structure are generic and anonymized for demonstration purposes.

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 👤 Author

Your Name
