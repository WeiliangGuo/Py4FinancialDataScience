import pandas as pd
import numpy as np

class AlternativeOptimalPortfolioCalculator:
    def __init__(self, file_path):
        self.file_path = file_path
        self.stock_forecasts = None
        self.bond_forecasts = None
        self.cov_matrices = None

    def load_data(self,
                  s1='g3.3_Forecast_Excess_R_Stocks',
                  s2='g3.3_Forecast_Excess_R_Bonds',
                  s3='g4.all_cov_matrices'):
        """
        Load necessary data from Excel file.
        """
        self.stock_forecasts = pd.read_excel(self.file_path, sheet_name=s1)
        self.bond_forecasts = pd.read_excel(self.file_path, sheet_name=s2)
        self.cov_matrices = pd.read_excel(self.file_path, sheet_name=s3)

    def calculate_optimal_weights(self):
        """
        Calculate optimal weights for each predictive model.
        """
        optimal_weights = pd.DataFrame()

        # Use the last row of cov_matrices as the most recent covariance matrix
        recent_cov_matrix = self.cov_matrices.iloc[-1][['Variance - Stocks', 'Covariance - Stocks/Bonds',
                                                        'Covariance - Bonds/Stocks', 'Variance - Bonds']]
        recent_cov_matrix = recent_cov_matrix.astype(float).values.reshape(2, 2)
       
        for model in self.stock_forecasts.columns[2:]:  # Exclude 'Unnamed: 0' and 'Date' columns
            mean_returns = np.array([self.stock_forecasts[model].iloc[-1],
                                     self.bond_forecasts[model].iloc[-1]])

            # Inverse of covariance matrix
            inv_cov_matrix = np.linalg.inv(recent_cov_matrix)

            # Calculating tangency portfolio weights
            weights = inv_cov_matrix @ mean_returns
            weights /= np.sum(weights)
            optimal_weights[model] = weights

        optimal_weights.index = ['Stocks', 'Bonds']
        return optimal_weights

def g5_3_demo():
    calculator = AlternativeOptimalPortfolioCalculator('data.xlsx')
    calculator.load_data()
    optimal_weights = calculator.calculate_optimal_weights()

    # Saving the results to the Excel file
    with pd.ExcelWriter('data.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        optimal_weights.to_excel(writer, sheet_name='g5.3.Alt_Opt_Tan_Portfolio_W')
    return optimal_weights

def g6_5_3_demo():
    print('Task 6.5.3:')
    calculator = AlternativeOptimalPortfolioCalculator('data.xlsx')
    s1 = 'g6.3.3_Forecast_Excess_R_Stocks'
    s2 = 'g6.3.3_Forecast_Excess_R_Bonds'
    s3 = 'g6.4.all_cov_matrices'
    calculator.load_data(s1, s2, s3)
    optimal_weights = calculator.calculate_optimal_weights()
    # # Saving the results to the Excel file
    # with pd.ExcelWriter('data.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    #     optimal_weights.to_excel(writer, sheet_name='g6.5.3.Alt_Opt_Tan_Portfolio_W')
    return optimal_weights


if __name__ == '__main__':
    # print(g5_3_demo())
    print(g6_5_3_demo())