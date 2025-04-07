import os
import pandas as pd
from datetime import datetime

# === STEP 1: Paths ===
raw_data_dir = "PeopleSoft_Invoice_Reports/raw_data"
processed_root = "PeopleSoft_Invoice_Reports/processed_reports"
mmp_ref_path = "PeopleSoft_Invoice_Reports/MMP_Reclass_Ref/MMP_Reclass_Ref.xlsx"

# === STEP 2: Load the latest file from raw_data ===
try:
    excel_files = [f for f in os.listdir(raw_data_dir) if f.endswith('.xlsx')]
    if not excel_files:
        raise FileNotFoundError("❌ No Excel files found in raw_data!")

    latest_file = max(excel_files, key=lambda f: os.path.getmtime(os.path.join(raw_data_dir, f)))
    latest_file_path = os.path.join(raw_data_dir, latest_file)
    print(f"📥 Found latest raw file: {latest_file}")
except Exception as e:
    print(f"❌ Error finding latest file: {str(e)}")
    raise

# === STEP 3: Load invoice data (skip row 1) ===
try:
    df = pd.read_excel(latest_file_path, skiprows=1)
except Exception as e:
    print(f"❌ Error reading invoice data: {str(e)}")
    raise

# === STEP 4: Extract month for folder/output naming ===
report_date = pd.to_datetime(df['Journal Date'].iloc[0])
report_folder = report_date.strftime("%Y_%m")
report_filename = f"Invoice_Report_{report_folder}.xlsx"
mmp_output_filename = f"MMP_Reclass_Allocations_{report_folder}.xlsx"

# === STEP 5: Setup output folders ===
processed_month_dir = os.path.join(processed_root, report_folder)
os.makedirs(processed_month_dir, exist_ok=True)
report_path = os.path.join(processed_month_dir, report_filename)
mmp_output_path = os.path.join(processed_month_dir, mmp_output_filename)

# === STEP 6: Clean + Label ===
df = df[df['Contract'].isin([1111, 2222])]
df = df[~df['Line Descr'].isin(["MSG Chart Expense", "MSG Misc Chart Expense"])]

label_conditions = {
    lambda row: row['Source'] == 'AP2' and row['Contract'] == 1111: "Charts & Coding",
    lambda row: row['Source'] == 'AP2' and row['Contract'] == 2222: "Misc. exp.",
    lambda row: row['Source'] == 'COR' and row['Contract'] == 1111 and row['Amount'] < 0: "1111 Coupa Reversal",
    lambda row: row['Source'] == 'COR' and row['Contract'] == 1111 and row['Amount'] > 0: "1111 Coupa Pending",
    lambda row: row['Source'] == 'COR' and row['Contract'] == 2222 and row['Amount'] < 0: "2222 Coupa Reversal",
    lambda row: row['Source'] == 'COR' and row['Contract'] == 2222 and row['Amount'] > 0: "2222 Coupa Pending"
}

def label_invoice(row):
    for condition, label in label_conditions.items():
        if condition(row):
            return label
    return "Unlabeled"

df['Label'] = df.apply(label_invoice, axis=1)
df['Value Used'] = df['Amount'].where(df['Source'] == 'COR', df['AP Amount'])

# === STEP 7: Summary Sheet ===
summary_df = df.groupby('Label')['Value Used'].sum().reset_index()
summary_df.columns = ['Label', 'Total']

# === STEP 8: Flags Sheet ===
abs_value_threshold = df['Value Used'].abs().quantile(0.99)
duplicates = df[df.duplicated(subset=['Invoice'], keep=False)]
outliers = df[df['Value Used'].abs() > abs_value_threshold]
flags_df = pd.concat([duplicates, outliers]).drop_duplicates()

# === STEP 9: MMP Reclass Allocation ===
print("📄 Reading MMP Reclass reference...")
try:
    ref_df = pd.read_excel(mmp_ref_path, converters={'% of Payments': float})
except FileNotFoundError:
    print(f"❌ MMP Reclass reference not found at: {mmp_ref_path}")
    raise
except Exception as e:
    print(f"❌ Error reading MMP Reclass reference: {str(e)}")
    raise

# Get 'Charts & Coding' total from summary
charts_total = summary_df.loc[summary_df['Label'] == 'Charts & Coding', 'Total'].values[0]
print(f"📊 Charts & Coding Total: ${charts_total:,.2f}")

ref_df['Payment Allocation'] = ref_df['% of Payments'] * charts_total

subset_alloc = ref_df.loc[ref_df['Contract'] == 'Subset', 'Payment Allocation'].values[0]
summary_df.loc[len(summary_df.index)] = ['Total MMP Reclass', subset_alloc]
print(f"✅ Reclass (Subset) Allocation: ${subset_alloc:,.2f}")

total_alloc = ref_df.loc[ref_df['State'] == 'Total', 'Payment Allocation'].sum()
adjusted_value = total_alloc - subset_alloc
ref_df.loc[ref_df['State'] == 'Adjusted', 'Payment Allocation'] = adjusted_value
print(f"✅ Adjusted Allocation = Total - Subset = ${adjusted_value:,.2f}")

# === STEP 10: Save output files ===
writer_args = [
    (mmp_output_path, "MMP Allocation"),
    (report_path, ["Summary", "Full Data", "Flags"])
]

for file_path, sheet_names in writer_args:
    try:
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            if isinstance(sheet_names, str):
                ref_df.to_excel(writer, sheet_name=sheet_names, index=False)
                workbook = writer.book
                worksheet = writer.sheets[sheet_names]

                # Define formats
                percent_fmt = workbook.add_format({'num_format': '0.00%', 'align': 'center'})
                currency_fmt = workbook.add_format({'num_format': '$#,##0', 'align': 'right'})
                gray_fmt = workbook.add_format({'num_format': '$#,##0', 'align': 'right', 'bg_color': '#D9D9D9'})
                gray_pct_fmt = workbook.add_format({'num_format': '0.00%', 'align': 'center', 'bg_color': '#D9D9D9'})
                yellow_fmt = workbook.add_format({'num_format': '$#,##0', 'align': 'right', 'bg_color': '#FFFACD'})
                yellow_pct_fmt = workbook.add_format({'num_format': '0.00%', 'align': 'center', 'bg_color': '#FFFACD'})

                # Get column indexes
                percent_col = ref_df.columns.get_loc('% of Payments')
                alloc_col = ref_df.columns.get_loc('Payment Allocation')

                worksheet.set_column(percent_col, percent_col, None, percent_fmt)
                worksheet.set_column(alloc_col, alloc_col, None, currency_fmt)

                for row_idx, row in ref_df.iterrows():
                    if str(row['State']).strip().lower() == 'total':
                        worksheet.write(row_idx + 1, alloc_col, row['Payment Allocation'], gray_fmt)
                        worksheet.write(row_idx + 1, percent_col, row['% of Payments'], gray_pct_fmt)
                    if str(row['Contract']).strip().lower() == 'subset':
                        worksheet.write(row_idx + 1, alloc_col, row['Payment Allocation'], yellow_fmt)
                        worksheet.write(row_idx + 1, percent_col, row['% of Payments'], yellow_pct_fmt)
            else:
                summary_df.to_excel(writer, sheet_name=sheet_names[0], index=False)
                summary_ws = writer.sheets[sheet_names[0]]
                total_col_idx = summary_df.columns.get_loc('Total')
                currency_fmt_summary = writer.book.add_format({'num_format': '$#,##0.00'})
                summary_ws.set_column(total_col_idx, total_col_idx, None, currency_fmt_summary)
                df.to_excel(writer, sheet_name=sheet_names[1], index=False)
                flags_df.to_excel(writer, sheet_name=sheet_names[2], index=False)

        print(f"✅ Saved: {file_path}")
    except Exception as e:
        print(f"❌ Error saving {file_path}: {str(e)}")
        raise

