import pandas as pd
import numpy as np

class OptimalTangencyPortfolio:
    def __init__(self, file_path):
        self.file_path = file_path
        self.stock_forecasts = None
        self.bond_forecasts = None
        self.cov_matrices = None

    def load_data(self,
                  s1='g2.SP500_Monthly_Mean_Forecast',
                  s2='g2.Bonds_Monthly_Mean_Forecast',
                  s3='g4.all_cov_matrices'):
        """
        Load necessary data from Excel file.
        """
        self.stock_forecasts = pd.read_excel(self.file_path, sheet_name=s1)
        self.bond_forecasts = pd.read_excel(self.file_path, sheet_name=s2)
        self.cov_matrices = pd.read_excel(self.file_path, sheet_name=s3)

    def calculate_optimal_weights(self):
        """
        Calculate the optimal tangency portfolio weights for stocks and bonds.
        """
        # Assuming the last row of cov_matrices is the most recent covariance matrix
        recent_cov_matrix = self.cov_matrices.iloc[-1][['Variance - Stocks', 'Covariance - Stocks/Bonds',
                                                        'Covariance - Bonds/Stocks', 'Variance - Bonds']]
        # Convert to numeric type
        recent_cov_matrix = recent_cov_matrix.astype(float).values.reshape(2, 2)
        mean_returns = np.array([self.stock_forecasts['Mean_Forecast'].iloc[-1],
                                 self.bond_forecasts['Mean_Forecast'].iloc[-1]])

        # Inverse of covariance matrix
        inv_cov_matrix = np.linalg.inv(recent_cov_matrix)

        # Calculating tangency portfolio weights
        weights = inv_cov_matrix @ mean_returns
        weights /= np.sum(weights)
        return pd.DataFrame(weights, index=['Stocks', 'Bonds'], columns=['Weights'])

def g5_1_demo():
    calculator = OptimalTangencyPortfolio('data.xlsx')
    calculator.load_data()
    optimal_weights = calculator.calculate_optimal_weights()
    # Saving the results to the Excel file
    with pd.ExcelWriter('data.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        optimal_weights.to_excel(writer, sheet_name='g5.1.Opt_Tan_Portfolio_W')
    return optimal_weights

def g6_5_1_demo():
    print('Task 6.5.1:')
    calculator = OptimalTangencyPortfolio('data.xlsx')
    s1 = 'g6.2.SP500_Monthly_Mu_Forecast'
    s2 = 'g6.2.Bonds_Monthly_Mu_Forecast'
    s3 = 'g6.4.all_cov_matrices'
    calculator.load_data(s1, s2, s3)
    optimal_weights = calculator.calculate_optimal_weights()
    # # Saving the results to the Excel file
    # with pd.ExcelWriter('data.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    #     optimal_weights.to_excel(writer, sheet_name='g6.5.1.Opt_Tan_Portfolio_W')
    return optimal_weights


if __name__ == '__main__':
    # print(g5_1_demo())
    print(g6_5_1_demo())