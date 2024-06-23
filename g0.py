import pandas as pd


class DataMerger:
    def __init__(self, predictor_file, risk_free_file, sp500_file, bond_index_file):
        self.predictor_file = predictor_file
        self.risk_free_file = risk_free_file
        self.sp500_file = sp500_file
        self.bond_index_file = bond_index_file

    def load_data(self):
        self.predictor_data = pd.read_excel(self.predictor_file,
                                            usecols=['Dates', 'E12', 'b/m', 'tbl', 'ntis', 'infl'])
        self.risk_free_data = pd.read_excel(self.risk_free_file)
        self.sp500_data = pd.read_excel(self.sp500_file)
        self.bond_index_data = pd.read_excel(self.bond_index_file)

    def normalize_dates(self):
        self.predictor_data['Dates'] = pd.to_datetime(self.predictor_data['Dates'], format='%Y%m').dt.strftime('%m-%Y')
        self.risk_free_data['Date'] = pd.to_datetime(self.risk_free_data['Date'], unit='d', origin='1899-12-30').dt.strftime('%m-%Y')
        self.sp500_data['Dates'] = pd.to_datetime(self.sp500_data['Dates']).dt.strftime('%m-%Y')
        self.bond_index_data['Dates'] = pd.to_datetime(self.bond_index_data['Dates']).dt.strftime('%m-%Y')

    def rename_columns(self):
        self.predictor_data.rename(columns={'Dates': 'Date'}, inplace=True)
        self.risk_free_data.rename(columns={'Date': 'Date', 'Risk free rate of return ': 'Risk_Free_Rate'}, inplace=True)
        self.sp500_data.rename(columns={'Dates': 'Date', 'SP500 index price': 'SP500'}, inplace=True)
        self.bond_index_data.rename(columns={'Dates': 'Date', 'LBUSTRUU Index price': 'LBUSTRUU'}, inplace=True)

    def merge_data(self):
        merged_data = pd.merge(self.predictor_data, self.sp500_data, on='Date', how='inner')
        merged_data = pd.merge(merged_data, self.risk_free_data, on='Date', how = 'left')
        merged_data = pd.merge(merged_data, self.bond_index_data, on='Date', how='inner')              
        return merged_data

    def calculate_excess_returns(self, merged_data):
        merged_data['SP500_Return'] = merged_data['SP500'].pct_change() 
        merged_data['LBUSTRUU_Return'] = merged_data['LBUSTRUU'].pct_change() 
        merged_data['Excess_Return_Stocks'] = merged_data['SP500_Return'] - merged_data['Risk_Free_Rate']
        merged_data['Excess_Return_Bonds'] = merged_data['LBUSTRUU_Return'] - merged_data['Risk_Free_Rate']
        return merged_data

def g0_demo():
    predictor_file = 'PredictorData2022.xlsx'
    risk_free_file = 'Risk-free_rate_of_return.xlsx'
    sp500_file = 'S&P_500_stock_index.xlsx'
    bond_index_file = 'US_Aggregate_Bond_index.xlsx'

    data_merger = DataMerger(predictor_file, risk_free_file, sp500_file, bond_index_file)
    data_merger.load_data()
    data_merger.normalize_dates()
    data_merger.rename_columns()
    merged_data = data_merger.merge_data()
    final_data = data_merger.calculate_excess_returns(merged_data)

    # Print a preview of the merged data
    print(f'Columns in final data.xlsx: \n{final_data.columns.tolist()}')
    print('-'*30)
    print(f'Number of rows: {len(final_data)}')
    print('-' * 30)
    print(f'data.xlsx "data" sheet preview: \n{final_data.head()}')
    
    # Save the merged data to a new Excel file
    final_data.to_excel('data.xlsx', index=False, sheet_name='data')


if __name__ == '__main__':
    g0_demo()