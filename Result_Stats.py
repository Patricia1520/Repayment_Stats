import pandas as pd


path = "C:\\Users\\newjo\\Downloads\\CM_results_internal.xlsx"
CMdeals = pd.read_excel(path, sheet_name="CM", usecols='A:AL')

# drop rows where Status is NA, these are deals under review
CMdeals_drop = CMdeals.dropna(subset=["Repayment status", "Date"], how='any')

data = CMdeals_drop.assign(
    Date=lambda x: pd.to_datetime(x['Date']),
    Month=lambda x: x['Date'].dt.to_period('M')
)


def country_stat(country, data):
    data_country = data[data['Country'] == country].copy()
    exclude_status = ['Not Funded', 'Completed', 'ongoing']

    stats = (
        data_country.groupby('Month')
        .agg(
            Submission_Count=('Date', 'size'),
            Funded_Count=("Repayment status", lambda x: (x != 'Not Funded').sum()),
            Issue_Count=("Repayment status", lambda x: (~x.isin(exclude_status)).sum()),  # New column
            Funded_USD_Amount=("USD Amount",
                               lambda x: x[data_country.loc[x.index, "Repayment status"] != 'Not Funded'].sum()),
            Repayment_Issue_Notional=("USD Amount",
                                    lambda x: x[~data_country.loc[x.index, "Repayment status"].isin(
                                        exclude_status)].sum()),
            Actual_Loss = ("Actual Loss (USD)", lambda x: x[~data_country.loc[x.index, "Repayment status"].isin(
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

    # Reorder columns if needed (optional)
    column_order = ['Month', 'Submission_Count', 'Funded_Count', 'Issue_Count', 'Funded_USD_Amount',
                    'Repayment_Issue_Amount', 'Default_Rate']
    stats = stats[column_order]

    return stats


stat_SG = country_stat('SG', data)
stat_HK = country_stat('HK', data)
stat_AU = country_stat('AU', data)
stat_MY = country_stat('MY', data)
print("SG2: ", stat_SG)

output_path = "C:\\Users\\newjo\\Downloads\\stats3.xlsx"
with pd.ExcelWriter(output_path) as writer:
    data.to_excel(writer, sheet_name='Data')
    stat_SG.to_excel(writer, sheet_name='stat_SG')
    stat_HK.to_excel(writer, sheet_name='stat_HK')
    stat_AU.to_excel(writer, sheet_name='stat_AU')
    stat_MY.to_excel(writer, sheet_name='stat_MY')
