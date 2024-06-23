import pandas as pd
import numpy as np

class AlternativePortfolioStatisticsCalculator:
    def __init__(self, file_path, weights_sheet, data_sheet):
        self.file_path = file_path
        self.weights_sheet = weights_sheet
        self.data_sheet = data_sheet
        self.weights = None
        self.data = None

    def load_data(self):
        """
        Load necessary data from Excel file.
        """
        self.weights = pd.read_excel(self.file_path, sheet_name=self.weights_sheet)
        self.data = pd.read_excel(self.file_path, sheet_name=self.data_sheet)

    def calculate_annualized_statistics(self):
        """
        Calculate mean, volatility, and Sharpe ratio for each set of portfolio weights.
        """
        model_names = ['Forecast_E12', 'Forecast_b/m', 'Forecast_tbl', 'Forecast_ntis', 'Forecast_infl', 'Combined_Forecast']
        stats_results = []

        for model in model_names:
            # Extracting weights for Stocks and Bonds
            stocks_weight = self.weights.loc[self.weights['Unnamed: 0'] == 'Stocks', model].iloc[0]
            bonds_weight = self.weights.loc[self.weights['Unnamed: 0'] == 'Bonds', model].iloc[0]
            portfolio_weights = np.array([stocks_weight, bonds_weight]).reshape(1, -1)

            # Calculate monthly portfolio returns
            monthly_returns = (self.data[['Excess_Return_Stocks', 'Excess_Return_Bonds']] @ portfolio_weights.T).sum(axis=1)

            # Annualize the statistics
            mean_return = monthly_returns.mean() * 12  # Assuming 12 months in a year
            volatility = monthly_returns.std() * np.sqrt(12)  # Annualizing standard deviation
            sharpe_ratio = mean_return / volatility  # Assuming risk-free rate is 0 or negligible

            # Append the results
            stats_results.append(pd.DataFrame({
                'Model': [model],
                'Mean Return': [mean_return],
                'Volatility': [volatility],
                'Sharpe Ratio': [sharpe_ratio]
            }))

        return pd.concat(stats_results).reset_index(drop=True)

def g5_4_demo():
    calculator = AlternativePortfolioStatisticsCalculator(
        'data.xlsx',
        'g5.3.Alt_Opt_Tan_Portfolio_W',
        'data')
    calculator.load_data()
    statistics = calculator.calculate_annualized_statistics()

    # Saving the results to the Excel file
    with pd.ExcelWriter('data.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        statistics.to_excel(writer, sheet_name='g5.4.Alt_Opt_Portfolio_Stats')

    return statistics

def g6_5_4_demo():
    print('Task 6.5.4:')
    calculator = AlternativePortfolioStatisticsCalculator(
        'data.xlsx',
        'g6.5.3.Alt_Opt_Tan_Portfolio_W',
        'data')
    calculator.load_data()
    statistics = calculator.calculate_annualized_statistics()

    # # Saving the results to the Excel file
    # with pd.ExcelWriter('data.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    #     statistics.to_excel(writer, sheet_name='g6.5.4.Alt_Opt_Portfolio_Stats')

    return statistics

if __name__ == '__main__':
    # print(g5_4_demo())
    print(g6_5_4_demo())