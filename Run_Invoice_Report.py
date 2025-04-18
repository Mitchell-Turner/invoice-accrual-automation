#!/usr/bin/env python3
"""
PeopleSoft Invoice Report Processor

This script processes PeopleSoft invoice data, applies business logic for classification,
identifies anomalies, and generates formatted reports with allocations for MMP reclassification.

Author: Mitchell Turner
Date: April 2025
"""

import os
import pandas as pd
from datetime import datetime
import logging
from pathlib import Path
import sys

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('invoice_processor.log', encoding='utf-8')
    ]
)
logger = logging.getLogger("invoice_processor")


# ANSI Color codes for console output
class Colors:
    PINK = '\033[38;5;219m'  # Pastel Pink
    BLUE = '\033[38;5;153m'  # Pastel Blue
    GREEN = '\033[38;5;157m'  # Pastel Green
    YELLOW = '\033[38;5;229m'  # Pastel Yellow
    PURPLE = '\033[38;5;183m'  # Pastel Purple
    CYAN = '\033[38;5;159m'  # Pastel Cyan
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    END = '\033[0m'  # Reset


# Constants for configuration
REQUIRED_CONTRACTS = [1111, 2222]
EXCLUDED_LINE_DESCRIPTIONS = ["MSG Chart Expense", "MSG Misc Chart Expense"]
OUTLIER_PERCENTILE = 0.99


class InvoiceProcessor:
    """
    Handles the processing of PeopleSoft invoice reports, including:
    - Loading and cleaning raw data
    - Categorizing invoices by type
    - Calculating MMP reclass allocations
    - Generating formatted reports
    """

    def __init__(self, raw_data_dir, processed_root, mmp_ref_path):
        """
        Initialize the processor with directory paths.

        Args:
            raw_data_dir (str): Path to directory containing raw invoice files
            processed_root (str): Path to directory for processed output
            mmp_ref_path (str): Path to MMP reclass reference Excel file
        """
        self.raw_data_dir = raw_data_dir
        self.processed_root = processed_root
        self.mmp_ref_path = mmp_ref_path

        # Ensure directories exist
        Path(raw_data_dir).mkdir(exist_ok=True, parents=True)
        Path(processed_root).mkdir(exist_ok=True, parents=True)

        # Will be set during processing
        self.latest_file = None
        self.report_date = None
        self.report_folder = None
        self.invoice_df = None
        self.summary_df = None
        self.flags_df = None
        self.mmp_ref_df = None
        self.charts_total = None

    def find_latest_invoice_file(self):
        """
        Find the most recently modified Excel file in the raw data directory.

        Returns:
            str: Path to the latest invoice file

        Raises:
            FileNotFoundError: If no Excel files exist in the directory
        """
        try:
            excel_files = [f for f in os.listdir(self.raw_data_dir) if f.endswith('.xlsx')]
            if not excel_files:
                raise FileNotFoundError(f"No Excel files found in {self.raw_data_dir}")

            self.latest_file = max(
                excel_files,
                key=lambda f: os.path.getmtime(os.path.join(self.raw_data_dir, f))
            )
            latest_file_path = os.path.join(self.raw_data_dir, self.latest_file)
            logger.info(f"Found latest raw file: {self.latest_file}")
            return latest_file_path
        except Exception as e:
            logger.error(f"Error finding latest file: {str(e)}")
            raise

    def load_invoice_data(self, file_path):
        """
        Load and clean invoice data from Excel file.

        Args:
            file_path (str): Path to the invoice Excel file

        Raises:
            Exception: If file cannot be read or processed
        """
        try:
            # Skip the first row which contains header information
            self.invoice_df = pd.read_excel(file_path, skiprows=1)

            # Extract date information from the first journal entry
            self.report_date = pd.to_datetime(self.invoice_df['Journal Date'].iloc[0])
            self.report_folder = self.report_date.strftime("%Y_%m")

            logger.info(f"Loaded invoice data for {self.report_folder}")
            logger.info(f"Found {len(self.invoice_df)} invoice records")

            # Initial data filtering
            self._filter_invoice_data()
        except Exception as e:
            logger.error(f"Error reading invoice data: {str(e)}")
            raise

    def _filter_invoice_data(self):
        """Apply initial filtering to the invoice data."""
        # Filter for required contracts
        self.invoice_df = self.invoice_df[self.invoice_df['Contract'].isin(REQUIRED_CONTRACTS)]

        # Exclude specific line descriptions
        self.invoice_df = self.invoice_df[~self.invoice_df['Line Descr'].isin(EXCLUDED_LINE_DESCRIPTIONS)]

        logger.info(f"After filtering: {len(self.invoice_df)} invoice records")

    def categorize_invoices(self):
        """
        Categorize invoices based on business rules and add labels.
        """
        label_conditions = {
            lambda row: row['Source'] == 'AP2' and row['Contract'] == 1111: "Charts & Coding",
            lambda row: row['Source'] == 'AP2' and row['Contract'] == 2222: "Misc. exp.",
            lambda row: row['Source'] == 'COR' and row['Contract'] == 1111 and row['Amount'] < 0: "1111 Coupa Reversal",
            lambda row: row['Source'] == 'COR' and row['Contract'] == 1111 and row['Amount'] > 0: "1111 Coupa Pending",
            lambda row: row['Source'] == 'COR' and row['Contract'] == 2222 and row['Amount'] < 0: "2222 Coupa Reversal",
            lambda row: row['Source'] == 'COR' and row['Contract'] == 2222 and row['Amount'] > 0: "2222 Coupa Pending"
        }

        def label_invoice(row):
            """Determine the appropriate label for an invoice row."""
            for condition, label in label_conditions.items():
                if condition(row):
                    return label
            return "Unlabeled"

        self.invoice_df['Label'] = self.invoice_df.apply(label_invoice, axis=1)

        # Use AP Amount for AP2 records, Amount for COR records
        self.invoice_df['Value Used'] = self.invoice_df['Amount'].where(
            self.invoice_df['Source'] == 'COR',
            self.invoice_df['AP Amount']
        )

        # Log the counts of each category
        category_counts = self.invoice_df['Label'].value_counts().sort_index()
        logger.info("\n" + "-" * 50)
        logger.info(f"{Colors.BLUE}INVOICE CATEGORIES SUMMARY:{Colors.END}")
        logger.info("-" * 50)
        logger.info(f"{Colors.YELLOW}{'Category':<25s} {'Count':>10s}{Colors.END}")
        logger.info("-" * 50)
        for category, count in category_counts.items():
            logger.info(f"{category:<25s} {count:>10d}")
        logger.info("-" * 50)

    def create_summary(self):
        """
        Create a summary dataframe with totals by label.
        """
        self.summary_df = self.invoice_df.groupby('Label')['Value Used'].sum().reset_index()
        self.summary_df.columns = ['Label', 'Total']

        # Format the summary for logging
        logger.info("\n" + "-" * 50)
        logger.info(f"{Colors.PINK}SUMMARY TOTALS BY CATEGORY:{Colors.END}")
        logger.info("-" * 50)
        for _, row in self.summary_df.iterrows():
            logger.info(f"{row['Label']:.<30s} {Colors.GREEN}${row['Total']:>15,.2f}{Colors.END}")
        logger.info("-" * 50)

    def identify_flags(self):
        """
        Identify potential issues in the data - duplicates and outliers.
        """
        # Identify duplicates
        duplicates = self.invoice_df[self.invoice_df.duplicated(subset=['Invoice'], keep=False)]

        # Identify outliers - values above the 99th percentile
        abs_value_threshold = self.invoice_df['Value Used'].abs().quantile(OUTLIER_PERCENTILE)
        outliers = self.invoice_df[self.invoice_df['Value Used'].abs() > abs_value_threshold]

        # Combine flagged items
        self.flags_df = pd.concat([duplicates, outliers]).drop_duplicates()

        # Format the flags summary
        logger.info("\n" + "-" * 50)
        logger.info(f"{Colors.YELLOW}FLAGGED ITEMS SUMMARY:{Colors.END}")
        logger.info("-" * 50)
        logger.info(f"Duplicate invoices: {Colors.PURPLE}{len(duplicates)}{Colors.END}")
        logger.info(f"Outliers (>${abs_value_threshold:,.2f}): {Colors.PURPLE}{len(outliers)}{Colors.END}")
        logger.info(f"Total flagged items: {Colors.PURPLE}{len(self.flags_df)}{Colors.END}")
        logger.info("-" * 50)

    def process_mmp_allocation(self):
        """
        Process MMP reclass allocations based on reference data.

        Raises:
            FileNotFoundError: If MMP reference file is not found
            Exception: For other processing errors
        """
        logger.info("\n" + "-" * 50)
        logger.info(f"{Colors.CYAN}PROCESSING MMP RECLASS ALLOCATION:{Colors.END}")
        logger.info("-" * 50)

        try:
            # Check if the summary DataFrame is empty
            if self.summary_df.empty:
                logger.warning(f"{Colors.YELLOW}No data available for MMP allocation - summary is empty{Colors.END}")
                return

            # Check if "Charts & Coding" category exists
            if "Charts & Coding" not in self.summary_df["Label"].values:
                logger.warning(f"{Colors.YELLOW}No \"Charts & Coding\" category found in summary{Colors.END}")
                return

            # Load the MMP reference data
            self.mmp_ref_df = pd.read_excel(self.mmp_ref_path, converters={'% of Payments': float})

            # Get the total for "Charts & Coding" category
            self.charts_total = self.summary_df.loc[
                self.summary_df['Label'] == 'Charts & Coding', 'Total'
            ].values[0]

            # Calculate payment allocations based on percentages
            self.mmp_ref_df['Payment Allocation'] = self.mmp_ref_df['% of Payments'] * self.charts_total

            # Get subset allocation
            subset_alloc = self.mmp_ref_df.loc[
                self.mmp_ref_df['Contract'] == 'Subset', 'Payment Allocation'
            ].values[0]

            # Add to summary
            self.summary_df.loc[len(self.summary_df.index)] = ['Total MMP Reclass', subset_alloc]

            # Calculate adjusted allocation
            total_alloc = self.mmp_ref_df.loc[
                self.mmp_ref_df['State'] == 'Total', 'Payment Allocation'
            ].sum()

            adjusted_value = total_alloc - subset_alloc
            self.mmp_ref_df.loc[
                self.mmp_ref_df['State'] == 'Adjusted', 'Payment Allocation'
            ] = adjusted_value

            # Log the MMP allocation summary
            logger.info(f"{'Charts & Coding Total:':<30s} {Colors.GREEN}${self.charts_total:>15,.2f}{Colors.END}")
            logger.info(f"{'Reclass (Subset) Allocation:':<30s} {Colors.GREEN}${subset_alloc:>15,.2f}{Colors.END}")
            logger.info(f"{'Total Allocation:':<30s} {Colors.GREEN}${total_alloc:>15,.2f}{Colors.END}")
            logger.info(f"{'Adjusted Allocation:':<30s} {Colors.GREEN}${adjusted_value:>15,.2f}{Colors.END}")
            logger.info("-" * 50)

        except FileNotFoundError:
            logger.error(f"{Colors.YELLOW}MMP Reclass reference not found at: {self.mmp_ref_path}{Colors.END}")
            raise
        except Exception as e:
            logger.error(f"{Colors.YELLOW}Error processing MMP allocation: {str(e)}{Colors.END}")
            raise

        except FileNotFoundError:
            logger.error(f"MMP Reclass reference not found at: {self.mmp_ref_path}")
            raise
        except Exception as e:
            logger.error(f"Error processing MMP allocation: {str(e)}")
            raise

    def save_reports(self):
        """
        Save processed data to Excel files with formatting.

        Returns:
            tuple: Paths to the generated report files
        """
        # Create month-specific directory
        processed_month_dir = os.path.join(self.processed_root, self.report_folder)
        os.makedirs(processed_month_dir, exist_ok=True)

        # Define output file paths
        report_filename = f"Invoice_Report_{self.report_folder}.xlsx"
        mmp_output_filename = f"MMP_Reclass_Allocations_{self.report_folder}.xlsx"

        report_path = os.path.join(processed_month_dir, report_filename)
        mmp_output_path = os.path.join(processed_month_dir, mmp_output_filename)

        # Save MMP Allocation file
        self._save_mmp_allocation_file(mmp_output_path)

        # Save main report file
        self._save_main_report_file(report_path)

        return report_path, mmp_output_path

    def _save_mmp_allocation_file(self, file_path):
        """
        Save MMP allocation data with formatting.

        Args:
            file_path (str): Output file path
        """
        try:
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                self.mmp_ref_df.to_excel(writer, sheet_name="MMP Allocation", index=False)
                workbook = writer.book
                worksheet = writer.sheets["MMP Allocation"]

                # Define formats
                header_fmt = workbook.add_format({
                    'bold': True,
                    'fg_color': '#E6E6FA',  # Pastel lavender
                    'align': 'center',
                    'border': 1
                })
                percent_fmt = workbook.add_format({'num_format': '0.00%', 'align': 'center'})
                currency_fmt = workbook.add_format({'num_format': '$#,##0', 'align': 'right'})
                gray_fmt = workbook.add_format({
                    'num_format': '$#,##0',
                    'align': 'right',
                    'bg_color': '#D9D9D9'
                })
                gray_pct_fmt = workbook.add_format({
                    'num_format': '0.00%',
                    'align': 'center',
                    'bg_color': '#D9D9D9'
                })
                yellow_fmt = workbook.add_format({
                    'num_format': '$#,##0',
                    'align': 'right',
                    'bg_color': '#FFFACD'  # Pastel yellow
                })
                yellow_pct_fmt = workbook.add_format({
                    'num_format': '0.00%',
                    'align': 'center',
                    'bg_color': '#FFFACD'  # Pastel yellow
                })

                # Apply header format
                for col_num, value in enumerate(self.mmp_ref_df.columns.values):
                    worksheet.write(0, col_num, value, header_fmt)

                # Auto-adjust column widths based on content
                for i, col in enumerate(self.mmp_ref_df.columns):
                    # Find the maximum length in the column
                    max_len = max(
                        self.mmp_ref_df[col].astype(str).map(len).max(),
                        len(str(col))
                    )
                    # Add a little extra space
                    worksheet.set_column(i, i, max_len + 2)

                # Get column indexes
                percent_col = self.mmp_ref_df.columns.get_loc('% of Payments')
                alloc_col = self.mmp_ref_df.columns.get_loc('Payment Allocation')

                worksheet.set_column(percent_col, percent_col, None, percent_fmt)
                worksheet.set_column(alloc_col, alloc_col, None, currency_fmt)

                # Apply conditional formatting
                for row_idx, row in self.mmp_ref_df.iterrows():
                    if str(row['State']).strip().lower() == 'total':
                        worksheet.write(row_idx + 1, alloc_col, row['Payment Allocation'], gray_fmt)
                        worksheet.write(row_idx + 1, percent_col, row['% of Payments'], gray_pct_fmt)
                    if str(row['Contract']).strip().lower() == 'subset':
                        worksheet.write(row_idx + 1, alloc_col, row['Payment Allocation'], yellow_fmt)
                        worksheet.write(row_idx + 1, percent_col, row['% of Payments'], yellow_pct_fmt)

            logger.info(f"Saved MMP allocation file: {file_path}")
        except Exception as e:
            logger.error(f"Error saving MMP allocation file: {str(e)}")
            raise

    def _save_main_report_file(self, file_path):
        """
        Save main report file with summary, full data, and flags.

        Args:
            file_path (str): Output file path
        """
        try:
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                workbook = writer.book

                # Define formats
                header_fmt = workbook.add_format({
                    'bold': True,
                    'align': 'center',
                    'border': 1
                })
                summary_header_fmt = workbook.add_format({
                    'bold': True,
                    'fg_color': '#FFD1DC',  # Pastel pink
                    'align': 'center',
                    'border': 1
                })
                data_header_fmt = workbook.add_format({
                    'bold': True,
                    'fg_color': '#CCFFCC',  # Pastel green
                    'align': 'center',
                    'border': 1
                })
                flags_header_fmt = workbook.add_format({
                    'bold': True,
                    'fg_color': '#FFFFCC',  # Pastel yellow
                    'align': 'center',
                    'border': 1
                })
                currency_fmt = workbook.add_format({'num_format': '$#,##0.00'})

                # Summary sheet
                self.summary_df.to_excel(writer, sheet_name="Summary", index=False)
                summary_ws = writer.sheets["Summary"]

                # Apply header format to Summary sheet
                for col_num, value in enumerate(self.summary_df.columns.values):
                    summary_ws.write(0, col_num, value, summary_header_fmt)

                # Auto-adjust column widths for Summary sheet
                for i, col in enumerate(self.summary_df.columns):
                    # Find the maximum length in the column
                    max_len = max(
                        self.summary_df[col].astype(str).map(len).max(),
                        len(str(col))
                    )
                    # Add a little extra space
                    summary_ws.set_column(i, i, max_len + 2)

                # Format currency in summary
                total_col_idx = self.summary_df.columns.get_loc('Total')
                summary_ws.set_column(total_col_idx, total_col_idx, None, currency_fmt)

                # Full data sheet
                self.invoice_df.to_excel(writer, sheet_name="Full Data", index=False)
                data_ws = writer.sheets["Full Data"]

                # Apply header format to Full Data sheet
                for col_num, value in enumerate(self.invoice_df.columns.values):
                    data_ws.write(0, col_num, value, data_header_fmt)

                # Auto-adjust column widths for Full Data sheet
                for i, col in enumerate(self.invoice_df.columns):
                    # For large datasets, use a sample to determine width
                    column_data = self.invoice_df[col].astype(str)
                    # Consider only a sample of the data to avoid slow performance on large datasets
                    if len(column_data) > 100:
                        sample = column_data.sample(100).map(len)
                        max_len = max(sample.max(), len(str(col)))
                    else:
                        max_len = max(column_data.map(len).max(), len(str(col)))
                    # Set a reasonable limit to column width
                    col_width = min(max_len + 2, 50)  # Limit to 50 characters
                    data_ws.set_column(i, i, col_width)

                # Flags sheet
                self.flags_df.to_excel(writer, sheet_name="Flags", index=False)
                flags_ws = writer.sheets["Flags"]

                # Apply header format to Flags sheet
                for col_num, value in enumerate(self.flags_df.columns.values):
                    flags_ws.write(0, col_num, value, flags_header_fmt)

                # Auto-adjust column widths for Flags sheet
                for i, col in enumerate(self.flags_df.columns):
                    if not self.flags_df.empty:
                        max_len = max(
                            self.flags_df[col].astype(str).map(len).max(),
                            len(str(col))
                        )
                        col_width = min(max_len + 2, 50)  # Limit to 50 characters
                    else:
                        col_width = len(str(col)) + 2
                    flags_ws.set_column(i, i, col_width)

            logger.info(f"Saved main report file: {file_path}")
        except Exception as e:
            logger.error(f"Error saving main report file: {str(e)}")
            raise

    def process(self):
        """
        Run the complete processing pipeline.

        Returns:
            tuple: Paths to the generated report files
        """
        try:
            # Step 1: Find and load the latest invoice file
            latest_file_path = self.find_latest_invoice_file()
            self.load_invoice_data(latest_file_path)

            # Check if we have data to process
            if self.invoice_df.empty:
                logger.warning("No data available after filtering. Ensure your data contains contracts 1111 and 2222.")
                raise ValueError("No data available after filtering. Check the contract numbers in your data.")

            # Step 2: Process the data
            self.categorize_invoices()
            self.create_summary()
            self.identify_flags()
            self.process_mmp_allocation()

            # Step 3: Save the reports
            report_paths = self.save_reports()

            logger.info("Invoice processing completed successfully")
            return report_paths

        except Exception as e:
            logger.error(f"Processing failed: {str(e)}")
            raise


def main():
    """
    Main execution function for the invoice processor.

    Returns:
        int: 0 for success, 1 for failure
    """
    # Configuration - paths
    base_dir = os.path.dirname(os.path.abspath(__file__))
    raw_data_dir = os.path.join(base_dir, "PeopleSoft_Invoice_Reports", "raw_data")
    processed_root = os.path.join(base_dir, "PeopleSoft_Invoice_Reports", "processed_reports")
    mmp_ref_path = os.path.join(
        base_dir, "PeopleSoft_Invoice_Reports", "MMP_Reclass_Ref", "MMP_Reclass_Ref.xlsx"
    )

    # Print startup information
    print("\n" + "=" * 60)
    print(f"{Colors.BLUE}{Colors.BOLD}{'PEOPLESOFT INVOICE REPORT PROCESSOR':^60}{Colors.END}")
    print("=" * 60)
    print(f"{Colors.PINK}{Colors.BOLD}{'CONFIGURATION':^60}{Colors.END}")
    print("-" * 60)
    print(f"{Colors.YELLOW}{'Raw data directory:':<25}{Colors.END} {raw_data_dir}")
    print(f"{Colors.YELLOW}{'Output directory:':<25}{Colors.END} {processed_root}")
    print(f"{Colors.YELLOW}{'MMP reference file:':<25}{Colors.END} {mmp_ref_path}")
    print("=" * 60 + "\n")

    try:
        # Create and run the processor
        processor = InvoiceProcessor(raw_data_dir, processed_root, mmp_ref_path)
        report_path, mmp_path = processor.process()

        # Print success message
        print("\n" + "=" * 60)
        print(f"{Colors.BLUE}{Colors.BOLD}{'PROCESSING SUMMARY':^60}{Colors.END}")
        print("=" * 60)
        print(f"{Colors.YELLOW}{'STATUS:':<15}{Colors.END} {Colors.GREEN}{'COMPLETED SUCCESSFULLY':>45}{Colors.END}")
        print("-" * 60)
        print(f"{Colors.CYAN}{'Main report:':<15}{Colors.END} {report_path}")
        print(f"{Colors.CYAN}{'MMP allocation:':<15}{Colors.END} {mmp_path}")
        print("=" * 60 + "\n")

        return 0
    except Exception as e:
        # Print error message
        print("\n" + "=" * 60)
        print(f"{Colors.BLUE}{Colors.BOLD}{'PROCESSING SUMMARY':^60}{Colors.END}")
        print("=" * 60)
        print(f"{Colors.YELLOW}{'STATUS:':<15}{Colors.END} {Colors.PINK}{'FAILED':>45}{Colors.END}")
        print("-" * 60)
        print(f"{Colors.PURPLE}{'Error:':<15}{Colors.END} {str(e)}")
        print("=" * 60 + "\n")

        return 1


if __name__ == "__main__":
    sys.exit(main())
