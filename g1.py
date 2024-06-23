import pandas as pd
import numpy as np
from scipy.stats import skew, kurtosis
from openpyxl import load_workbook

class AssetStatistics:
    def __init__(self, data):
        self.data = data.dropna(subset=['Excess_Return'])

    def annualized_mean(self):
        return self.data['Excess_Return'].mean() * 12

    def annualized_volatility(self):
        return self.data['Excess_Return'].std() * np.sqrt(12)

    def annualized_sharpe_ratio(self, risk_free_rate):
        return (self.annualized_mean()) / self.annualized_volatility()

    def skewness(self):
        return skew(self.data['Excess_Return'])

    def kurtosis(self):
        return kurtosis(self.data['Excess_Return'], fisher=False)

    def summary_statistics(self, risk_free_rate):
        return {
            'Annualized Mean': self.annualized_mean(),
            'Annualized Volatility': self.annualized_volatility(),
            'Sharpe Ratio': self.annualized_sharpe_ratio(risk_free_rate),
            'Skewness': self.skewness(),
            'Kurtosis': self.kurtosis()
        }

def save_to_excel(data, file_path, sheet_name):
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
        data.to_excel(writer, sheet_name=sheet_name, index=True)

def g1_demo():
    file_path = 'data.xlsx'
    data = pd.read_excel(file_path)

    risk_free_rate = data['Risk_Free_Rate'].mean() * 12

    sp500_data = data[['Date', 'Excess_Return_Stocks']].rename(columns={'Excess_Return_Stocks': 'Excess_Return'})
    bonds_data = data[['Date', 'Excess_Return_Bonds']].rename(columns={'Excess_Return_Bonds': 'Excess_Return'})
       
    sp500_stats = AssetStatistics(sp500_data)
    bonds_stats = AssetStatistics(bonds_data)

    sp500_stats_dict = sp500_stats.summary_statistics(risk_free_rate)
    bonds_stats_dict = bonds_stats.summary_statistics(risk_free_rate)

    # Formatting the results into a DataFrame
    stats_df = pd.DataFrame({'S&P 500': sp500_stats_dict, 'Bonds': bonds_stats_dict})
    print(stats_df)

    # Saving the DataFrame to a new sheet in 'data.xlsx'
    save_to_excel(stats_df, file_path, 'g1.statistics')


if __name__ == '__main__':
    g1_demo()