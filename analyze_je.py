import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os
import sys

def analyze_data():
    input_file = 'je_samples (1).xlsx'
    output_dir = 'analysis_output'

    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    print(f"Reading {input_file}...")
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        print(f"Error reading file: {e}")
        sys.exit(1)

    # Data Cleaning
    # Convert Debit to numeric, coercing errors to NaN
    df['Debit'] = pd.to_numeric(df['Debit'], errors='coerce')
    df['Credit'] = pd.to_numeric(df['Credit'], errors='coerce')

    # Fill NaN with 0 for calculations
    df['Debit'] = df['Debit'].fillna(0)
    df['Credit'] = df['Credit'].fillna(0)

    # Clean string columns
    str_cols = df.select_dtypes(include=['object']).columns
    for col in str_cols:
        df[col] = df[col].astype(str).str.strip()

    # Convert dates
    df['EffectiveDate'] = pd.to_datetime(df['EffectiveDate'])
    df['EntryDate'] = pd.to_datetime(df['EntryDate'])

    # Summary Statistics
    row_count = len(df)
    effective_date_min = df['EffectiveDate'].min()
    effective_date_max = df['EffectiveDate'].max()
    entry_date_min = df['EntryDate'].min()
    entry_date_max = df['EntryDate'].max()

    total_debit = df['Debit'].sum()
    total_credit = df['Credit'].sum()

    bu_counts = df['BusinessUnit'].value_counts()
    source_counts = df['Source'].value_counts()

    # Write summary to file
    summary_file = os.path.join(output_dir, 'summary.txt')
    with open(summary_file, 'w') as f:
        f.write("Basic Analysis Summary\n")
        f.write("======================\n\n")
        f.write(f"Total Rows: {row_count}\n")
        f.write(f"Effective Date Range: {effective_date_min} to {effective_date_max}\n")
        f.write(f"Entry Date Range: {entry_date_min} to {entry_date_max}\n")
        f.write(f"Total Debits: {total_debit:,.2f}\n")
        f.write(f"Total Credits: {total_credit:,.2f}\n")
        f.write(f"Net Difference (Debits + Credits): {total_debit + total_credit:,.2f}\n\n")

        f.write("Entries by Business Unit:\n")
        f.write(bu_counts.to_string())
        f.write("\n\n")

        f.write("Entries by Source:\n")
        f.write(source_counts.to_string())
        f.write("\n")

    print(f"Summary written to {summary_file}")

    # Visualization: Entries over time (by Effective Date Month)
    plt.figure(figsize=(10, 6))
    # Group by month
    monthly_counts = df.groupby(df['EffectiveDate'].dt.to_period('M')).size()
    monthly_counts.index = monthly_counts.index.astype(str)

    monthly_counts.plot(kind='bar')
    plt.title('Number of Journal Entries by Effective Month')
    plt.xlabel('Month')
    plt.ylabel('Count')
    plt.tight_layout()

    plot_file = os.path.join(output_dir, 'entries_over_time.png')
    plt.savefig(plot_file)
    print(f"Plot saved to {plot_file}")

if __name__ == "__main__":
    analyze_data()
