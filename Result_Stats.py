import pandas as pd


path = "C:\\Users\\newjo\\Downloads\\CM_results_internal.xlsx"
CMdeals = pd.read_excel(path, sheet_name="CM", usecols='A:AL')

# drop rows where Status is NA, these are deals under review
CMdeals_drop = CMdeals.dropna(subset=["Repayment status", "Date"], how='any')

data = CMdeals_drop.assign(
    Date=lambda x: pd.to_datetime(x['Date']),
    Month=lambda x: x['Date'].dt.to_period('M')
)


def aggregate_stats(data, country=None):
    """
    Aggregate statistics by month, optionally filtered by country

    Parameters:
    data (DataFrame): Input data
    country (str, optional): Country to filter by. If None, uses all data

    Returns:
    DataFrame: Aggregated statistics with a total row
    """
    exclude_status = ['Not Funded', 'Completed', 'ongoing']

    # Filter by country if specified
    if country:
        data_filtered = data[data['Country'] == country].copy()
    else:
        data_filtered = data.copy()

    stats = (
        data_filtered.groupby('Month')
        .agg(
            Submission_Count=('Date', 'size'),
            Funded_Count=("Repayment status", lambda x: (x != 'Not Funded').sum()),
            Issue_Count=("Repayment status", lambda x: (~x.isin(exclude_status)).sum()),
            Funded_USD_Amount=("USD Amount",
                               lambda x: x[data_filtered.loc[x.index, "Repayment status"] != 'Not Funded'].sum()),
            Repayment_Issue_Notional=("USD Amount",
                                      lambda x: x[~data_filtered.loc[x.index, "Repayment status"].isin(
                                          exclude_status)].sum()),
            Actual_Loss=("Actual Loss (USD)", lambda x: x[~data_filtered.loc[x.index, "Repayment status"].isin(
                exclude_status)].sum())
        )
        .reset_index()
        .sort_values('Month')
    )

    # Calculate Default Rate
    stats['Default_Rate'] = stats.apply(
        lambda row: row['Actual_Loss'] / row['Funded_USD_Amount']
        if row['Funded_USD_Amount'] != 0 else 0,
        axis=1
    )

    # Reorder columns
    column_order = ['Month', 'Submission_Count', 'Funded_Count', 'Issue_Count', 'Funded_USD_Amount',
                    'Repayment_Issue_Notional', 'Actual_Loss', 'Default_Rate']
    stats = stats[column_order]

    # Add a total row with average default rate
    total_row = pd.DataFrame({
        'Month': ['Total'],
        'Submission_Count': [stats['Submission_Count'].sum()],
        'Funded_Count': [stats['Funded_Count'].sum()],
        'Issue_Count': [stats['Issue_Count'].sum()],
        'Funded_USD_Amount': [stats['Funded_USD_Amount'].sum()],
        'Repayment_Issue_Notional': [stats['Repayment_Issue_Notional'].sum()],
        'Actual_Loss': [stats['Actual_Loss'].sum()],
        'Default_Rate': [stats['Default_Rate'].mean()]  # Average of default rates
    })

    # Concatenate the original stats with the total row
    stats = pd.concat([stats, total_row], ignore_index=True)

    return stats

if __name__ == '__main__':
    stat_SG = aggregate_stats(data,'SG')
    stat_HK = aggregate_stats(data,'HK')
    stat_AU = aggregate_stats(data, 'AU')
    stat_MY = aggregate_stats(data,'MY')
    stat_Agg = aggregate_stats(data)

    output_path = "C:\\Users\\newjo\\Downloads\\stats1.xlsx"
    with pd.ExcelWriter(output_path) as writer:
        data.to_excel(writer, sheet_name='Data')
        stat_SG.to_excel(writer, sheet_name='stat_SG')
        stat_HK.to_excel(writer, sheet_name='stat_HK')
        stat_AU.to_excel(writer, sheet_name='stat_AU')
        stat_MY.to_excel(writer, sheet_name='stat_MY')
        stat_Agg.to_excel(writer, sheet_name='Aggregate')
