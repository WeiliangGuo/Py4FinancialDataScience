import pandas as pd

class MonthlyRecursiveVarianceCovarianceMatrixCalculator:
    def __init__(self, file_path):
        self.file_path = file_path
        self.data = None

    def load_data(self):
        """
        Load necessary data from Excel file and adjust the scale of returns.
        """
        self.data = pd.read_excel(self.file_path, sheet_name='data')
        # Adjusting the returns 
        self.data['Excess_Return_Stocks'] 
        self.data['Excess_Return_Bonds'] 

    def calculate_monthly_recursive_covariance_matrix(self, initial_window=242):
        """
        Calculate the monthly recursive 2x2 variance-covariance matrix for stocks and bonds.
        """
        stocks_returns = self.data['Excess_Return_Stocks']
        bonds_returns = self.data['Excess_Return_Bonds']
        returns = pd.DataFrame({'Stocks': stocks_returns, 'Bonds': bonds_returns})

        # DataFrame to store all matrices
        all_matrices = pd.DataFrame()

        # Calculate the covariance matrix for each month
        for end_month in range(initial_window, len(returns)+1):
            monthly_returns = returns.iloc[1:end_month]
            cov_matrix = monthly_returns.cov()
            matrix_flat = cov_matrix.unstack().to_frame().T
            matrix_flat.index = [self.data['Date'].iloc[end_month-1]]
            all_matrices = pd.concat([all_matrices, matrix_flat])
            
        # Renaming the columns for clarity
        all_matrices.columns = ['Variance - Stocks', 'Covariance - Stocks/Bonds', 'Covariance - Bonds/Stocks', 'Variance - Bonds']

        return all_matrices


def g4_demo():
    calculator = MonthlyRecursiveVarianceCovarianceMatrixCalculator('data.xlsx')
    calculator.load_data()
    all_cov_matrices = calculator.calculate_monthly_recursive_covariance_matrix()
    # Save the results to the same Excel file under a new sheet
    with pd.ExcelWriter('data.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        all_cov_matrices.to_excel(writer, sheet_name='g4.all_cov_matrices')
    return all_cov_matrices


class MonthlyRollingWindowVarCovMatrixCalculator:
    def __init__(self, file_path):
        self.file_path = file_path
        self.data = None

    def load_data(self):
        """
        Load necessary data from Excel file.
        """
        self.data = pd.read_excel(self.file_path, sheet_name='data')

    def calculate_monthly_rolling_window_covariance_matrix(self, initial_window=242):
        """
        Calculate the monthly rolling window 2x2 variance-covariance matrix for stocks and bonds.
        """
        stocks_returns = self.data['Excess_Return_Stocks']
        bonds_returns = self.data['Excess_Return_Bonds']
        returns = pd.DataFrame({'Stocks': stocks_returns, 'Bonds': bonds_returns})

        # DataFrame to store all matrices
        all_matrices = pd.DataFrame()

        # Calculate the covariance matrix for each month using rolling window
        for end_month in range(initial_window, len(returns)+1):
            window_returns = returns.iloc[end_month-initial_window:end_month]
            cov_matrix = window_returns.cov()
            matrix_flat = cov_matrix.unstack().to_frame().T
            matrix_flat.index = [self.data['Date'].iloc[end_month-1]]
            all_matrices = pd.concat([all_matrices, matrix_flat])

        # Renaming the columns for clarity
        all_matrices.columns = ['Variance - Stocks', 'Covariance - Stocks/Bonds', 'Covariance - Bonds/Stocks', 'Variance - Bonds']

        return all_matrices

def g6_4_demo():
    calculator = MonthlyRollingWindowVarCovMatrixCalculator('data.xlsx')
    calculator.load_data()
    all_cov_matrices = calculator.calculate_monthly_rolling_window_covariance_matrix()
    # # Save the results to the same Excel file under a new sheet
    # with pd.ExcelWriter('data.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    #     all_cov_matrices.to_excel(writer, sheet_name='g6.4.all_cov_matrices', index=False)
    return all_cov_matrices.head()


if __name__ == '__main__':
    # g4_demo()
    print(g6_4_demo())