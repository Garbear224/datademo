import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import os
import sys
import numpy as np

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

    # --- Benford's Law Analysis ---
    print("Performing Benford's Law Analysis...")
    # Combine Debits and Credits (abs value), filter out 0
    all_amounts = pd.concat([df['Debit'], df['Credit'].abs()])
    all_amounts = all_amounts[all_amounts > 0]

    # Extract first digit
    # Using string conversion for simplicity and robustness against float representation
    first_digits = all_amounts.astype(str).str.lstrip('0.').str[0].astype(int)

    # Calculate observed counts and frequencies
    observed_counts = first_digits.value_counts().sort_index()
    total_count = len(first_digits)
    observed_freq = observed_counts / total_count

    # Expected frequencies (Benford's Law)
    digits = np.arange(1, 10)
    expected_freq = np.log10(1 + 1/digits)
    expected_freq_series = pd.Series(expected_freq, index=digits)

    # DataFrame for comparison
    benford_df = pd.DataFrame({
        'Observed': observed_freq,
        'Expected': expected_freq_series
    }).fillna(0) # In case some digits are missing in observed

    # Plot Benford's Law
    plt.figure(figsize=(10, 6))
    benford_df.plot(kind='bar', width=0.8)
    plt.title("Benford's Law Analysis: Observed vs Expected First Digit Frequencies")
    plt.xlabel('First Digit')
    plt.ylabel('Frequency')
    plt.tight_layout()
    benford_plot_file = os.path.join(output_dir, 'benford_law.png')
    plt.savefig(benford_plot_file)
    print(f"Benford plot saved to {benford_plot_file}")
    plt.close()

    # --- Pie Chart (Source Distribution) ---
    print("Generating Pie Chart...")
    plt.figure(figsize=(10, 8))
    # Limit to top 10 sources if too many
    if len(source_counts) > 10:
         top_sources = source_counts.head(10)
         other_count = source_counts.iloc[10:].sum()
         # Create a new series for plotting to avoid SettingWithCopy warning or index issues
         plot_data = top_sources.copy()
         plot_data['Other'] = other_count
    else:
         plot_data = source_counts

    plot_data.plot(kind='pie', autopct='%1.1f%%', startangle=90)
    plt.title('Distribution of Journal Entries by Source')
    plt.ylabel('') # Hide y-label
    plt.tight_layout()
    pie_plot_file = os.path.join(output_dir, 'source_distribution_pie.png')
    plt.savefig(pie_plot_file)
    print(f"Pie chart saved to {pie_plot_file}")
    plt.close()

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
        f.write("\n\n")

        f.write("Benford's Law Analysis (First Digit Frequencies):\n")
        f.write(f"{'Digit':<6} {'Observed':<10} {'Expected':<10} {'Diff':<10}\n")
        f.write("-" * 40 + "\n")
        for digit in digits:
            obs = benford_df.loc[digit, 'Observed']
            exp = benford_df.loc[digit, 'Expected']
            diff = obs - exp
            f.write(f"{digit:<6} {obs:.4f}     {exp:.4f}     {diff:.4f}\n")

        f.write("\nInterpretation:\n")
        f.write("Significant deviations from the expected Benford's Law frequencies may indicate anomalies or potential fraud.\n")
        f.write("However, certain legitimate accounting patterns (e.g., recurring identical amounts) can also cause deviations.\n")

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
    plt.close()

if __name__ == "__main__":
    analyze_data()
