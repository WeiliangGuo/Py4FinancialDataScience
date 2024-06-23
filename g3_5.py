import pandas as pd
from scipy import stats
import numpy as np

class DmTestCalculator:
    def __init__(self, file_path, out_sample_start, out_sample_end):
        self.file_path = file_path
        self.data = None
        self.dm_test_stocks_results = pd.DataFrame(columns=['p-value', 'conclusion'])
        self.dm_test_bonds_results = pd.DataFrame(columns=['p-value', 'conclusion'])
        self.out_sample_start = out_sample_start
        self.out_sample_end = out_sample_end

    def load_data(self):
        self.data = pd.read_excel(self.file_path, sheet_name=None)

    def newey_west_se(self, errors, lag=4):
        """
        Calculate the Newey-West standard error of the mean forecast error.
        """
        T = len(errors)
        covariances = np.zeros(lag + 1)
        for k in range(lag + 1):
            cov = np.sum((errors[:T - k] - errors.mean()) * (errors[k:] - errors.mean())) / T
            covariances[k] = cov
        weights = 1 - np.arange(lag + 1) / (lag + 1)
        return np.sqrt(2 * np.dot(weights, covariances) / T)

    def perform_dm_test(self, actuals, forecast1, forecast2, lag=4):
        e1 = actuals - forecast1
        e2 = actuals - forecast2
        
        d = np.square(e1) - np.square(e2)
        mean_d = np.mean(d)
        se_d = self.newey_west_se(d, lag)
        dm_stat = mean_d / se_d
        
        #Bilateral-test
        p_value = stats.norm.cdf(-np.abs(dm_stat))
        
        #Bilateral 95% confidence level
        conclusion = 'Reject H0' if p_value < 0.025 else 'Failed to reject H0'
        return p_value, conclusion

    def run_tests(self,
                  sheet_prefix1='g2',
                  sheet_prefix2='g3.3'):

        # Initialize an empty list to store recursive mean estimates for SP500
        self.data['data']['Date'] = pd.to_datetime(self.data['data']['Date']) 
        
        #Index to filter out-of-sample periods
        out_sample_start_index = self.data['data'][(self.data['data']['Date'] == self.out_sample_start)].index  
        out_sample_end_index = self.data['data'][(self.data['data']['Date'] == self.out_sample_end)].index
        out_sample_start_index = out_sample_start_index[0]
        out_sample_end_index  = out_sample_end_index[0]
         
        #Filter out-of-sample periods
        actuals_stocks = self.data['data']['Excess_Return_Stocks'].iloc[out_sample_start_index : out_sample_end_index+1]
        actuals_bonds = self.data['data']['Excess_Return_Bonds'].iloc[out_sample_start_index : out_sample_end_index+1]

        if sheet_prefix1 == 'g6.2':
            mu = 'Mu'
        else:
            mu = 'Mean'
        benchmark_stocks = self.data[f'{sheet_prefix1}.SP500_Monthly_{mu}_Forecast']['Mean_Forecast']
        benchmark_bonds = self.data[f'{sheet_prefix1}.Bonds_Monthly_{mu}_Forecast']['Mean_Forecast']
        
        #Transforming to array
        actuals_stocks = actuals_stocks.array
        actuals_bonds = actuals_bonds.array
        benchmark_stocks = benchmark_stocks.array
        benchmark_bonds = benchmark_bonds.array
              
        
        for sheet in [f'{sheet_prefix2}_Forecast_Excess_R_Stocks',
                      f'{sheet_prefix2}_Forecast_Excess_R_Bonds']:
            for column in self.data[sheet].columns:
                if column not in ['Date', 'Unnamed: 0']:
                    forecasts = self.data[sheet][column]
                    forecasts = forecasts.array
                    
                    if 'Stocks' in sheet:
                        actuals = actuals_stocks
                        benchmark = benchmark_stocks
                        p_value, conclusion = self.perform_dm_test(actuals, forecasts, benchmark)
                        self.dm_test_stocks_results.loc[column] = [p_value, conclusion]
                        
                    else:
                        benchmark = benchmark_bonds
                        p_value, conclusion = self.perform_dm_test(actuals_bonds, forecasts, benchmark)
                        self.dm_test_bonds_results.loc[column] = [p_value, conclusion]                   
                    

def g3_5_demo():
     # Set the boundaries for in-sample and out-of-sample periods
    out_sample_start = "2000-01-01"
    out_sample_end = "2021-12-01" #This is equivalent to 2021-12-31 (last observation)
    
    file_path = 'data.xlsx'
    dm_test_calculator = DmTestCalculator(file_path, out_sample_start,out_sample_end)
    dm_test_calculator.load_data()
    dm_test_calculator.run_tests()
    print('\nStocks:\n', dm_test_calculator.dm_test_stocks_results)
    print('-'*50)
    print('\nBonds:\n', dm_test_calculator.dm_test_bonds_results)


def g6_3_5_demo():
    print('Task 6.3.5:')
    # Set the boundaries for in-sample and out-of-sample periods
    out_sample_start = "2000-01-01"
    out_sample_end = "2021-12-01"  # This is equivalent to 2021-12-31 (last observation)

    file_path = 'data.xlsx'
    dm_test_calculator = DmTestCalculator(file_path, out_sample_start, out_sample_end)
    dm_test_calculator.load_data()

    sheet_prefix1 = 'g6.2'
    sheet_prefix2 = 'g6.3.3'
    dm_test_calculator.run_tests(sheet_prefix1, sheet_prefix2)
    print('\nStocks:\n', dm_test_calculator.dm_test_stocks_results)
    print('-' * 50)
    print('\nBonds:\n', dm_test_calculator.dm_test_bonds_results)


if __name__ == '__main__':
    # g3_5_demo()
    g6_3_5_demo()