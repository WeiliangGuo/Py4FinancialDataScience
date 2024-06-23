import pandas as pd

class MSFECalculator:
    """
    A class to calculate Mean Squared Forecast Error (MSFE) and MSFE ratio for various forecasting models.
    """
    def __init__(self, file_path):
        self.file_path = file_path
        self.data = None
        self.msfe_results = {}
        self.msfe_ratios = {}

    def load_data(self):
        """
        Load necessary data from Excel file.
        """
        self.data = pd.read_excel(self.file_path, sheet_name=None)

    def calculate_msfe(self, actuals, forecasts):
        """
        Calculate Mean Squared Forecast Error (MSFE) between actuals and forecasts.
        """
        return ((actuals - forecasts) ** 2).mean()

    def calculate_msfe_ratios(self, benchmark_name):
        """
        Calculate MSFE ratios for each model relative to the benchmark.
        """
        benchmark_msfe = self.msfe_results[benchmark_name]

        for key, value in self.msfe_results.items():
            if key != benchmark_name:
                self.msfe_ratios[f'{key} Ratio'] = value / benchmark_msfe

    def process_forecasts(self,
                          sheet_prefix1='g2',
                          sheet_prefix2='g3.3'):
        """
        Process each forecast and calculate its MSFE.
        """      
        actuals_stocks = self.data['data']['Excess_Return_Stocks']
        actuals_bonds = self.data['data']['Excess_Return_Bonds']

        # Mapping of sheet names and columns to meaningful result names
        if sheet_prefix1 == 'g6.2':
            mu = 'Mu'
        else:
            mu = 'Mean'
        meaningful_names = {
            f'{sheet_prefix1}.SP500_Monthly_{mu}_Forecast': {'Mean_Forecast': 'SP500 Mean Forecast MSFE'},
            f'{sheet_prefix2}_Forecast_Excess_R_Stocks': {
                'Forecast_E12': 'Stocks E12 Forecast MSFE',
                'Forecast_b/m': 'Stocks Book-to-Market Ratio Forecast MSFE',
                'Forecast_tbl': 'Stocks Treasury Bill Rate Forecast MSFE',
                'Forecast_ntis': 'Stocks Net Equity Expansion Forecast MSFE',
                'Forecast_infl': 'Stocks Inflation Forecast MSFE',
                'Combined_Forecast': 'Stocks Combined Forecast MSFE'
            },
            f'{sheet_prefix1}.Bonds_Monthly_{mu}_Forecast': {'Mean_Forecast': 'Bonds Mean Forecast MSFE'},
            f'{sheet_prefix2}_Forecast_Excess_R_Bonds': {
                'Forecast_E12': 'Bonds E12 Forecast MSFE',
                'Forecast_b/m': 'Bonds Book-to-Market Ratio Forecast MSFE',
                'Forecast_tbl': 'Bonds Treasury Bill Rate Forecast MSFE',
                'Forecast_ntis': 'Bonds Net Equity Expansion Forecast MSFE',
                'Forecast_infl': 'Bonds Inflation Forecast MSFE',
                'Combined_Forecast': 'Bonds Combined Forecast MSFE'
            }
        }

        # Calculate MSFE for each model
        for sheet, name_mappings in meaningful_names.items():
            for column, result_name in name_mappings.items():
                forecasts = self.data[sheet][column]               
                
                #Index for filter
                out_sample_start_index = len(actuals_stocks) - len(forecasts) 
                out_sample_end_index = len(actuals_stocks)-1                     
                
                actuals = actuals_stocks.iloc[out_sample_start_index : out_sample_end_index+1] if result_name.__contains__('Stocks') or result_name.__contains__('SP500') else actuals_bonds.iloc[out_sample_start_index : out_sample_end_index+1]
                
                
                if result_name.__contains__('Stocks') or result_name.__contains__('SP500'):
                    benchmark = self.data[f'{sheet_prefix1}.SP500_Monthly_{mu}_Forecast']['Mean_Forecast']
                    benchmark = benchmark.array
                    self.msfe_results['SP500 Mean Forecast MSFE'] = self.calculate_msfe(actuals, benchmark)
                    benchmark_mean = self.msfe_results['SP500 Mean Forecast MSFE']
                    
                                      
                else: 
                    benchmark = self.data[f'{sheet_prefix1}.Bonds_Monthly_{mu}_Forecast']['Mean_Forecast']
                    benchmark = benchmark.array
                    self.msfe_results['Bonds Mean Forecast MSFE'] = self.calculate_msfe(actuals, benchmark)
                    benchmark_mean = self.msfe_results['Bonds Mean Forecast MSFE']
                    
                    #Transform to array
                actuals = actuals.array
                forecasts = forecasts.array
                    
                    #Caclculate results
                self.msfe_results[result_name] = self.calculate_msfe(actuals, forecasts)
                self.msfe_results[result_name + ' Ratio'] = self.msfe_results[result_name]/benchmark_mean

        # Calculate MSFE ratios
        self.calculate_msfe_ratios('SP500 Mean Forecast MSFE')
        self.calculate_msfe_ratios('Bonds Mean Forecast MSFE')
               
        
    def save_results(self,
                     s1='g.3.4.MSFE_Results',
                     s2='g.3.4.MSFE_Ratios'):
        """
        Save MSFE results and ratios to a new sheet in the Excel file.
        """
        results_df1 = pd.DataFrame.from_dict({**self.msfe_results}, orient='index', columns=['Value'])
        results_df2 = pd.DataFrame.from_dict({**self.msfe_ratios}, orient='index', columns=['Value'])
        with pd.ExcelWriter(self.file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
            results_df1.to_excel(writer, sheet_name=s1)
            results_df2.to_excel(writer, sheet_name=s2)

def g3_4_demo():
    """
    Run the MSFE calculation process.
    """
    calculator = MSFECalculator('data.xlsx')
    calculator.load_data()
    calculator.process_forecasts()
    # calculator.save_results()
    print("MSFE Results:")

    # Convert the dictionary to a DataFrame
    msfe_results = pd.DataFrame(list(calculator.msfe_results.items()), columns=['MSFE Type', 'Value'])
    print(msfe_results)
    print('-'*50)
    print("MSFE Ratios:")
    msfe_ratios = pd.DataFrame(list(calculator.msfe_ratios.items()), columns=['MSFE Type', 'Value'])
    print(msfe_ratios)

def g6_3_4_demo():
    """
    Run the MSFE calculation process.
    """
    print('Task 6.3.4:')
    calculator = MSFECalculator('data.xlsx')
    calculator.load_data()
    sheet_prefix1 = 'g6.2'
    sheet_prefix2 = 'g6.3.3'
    calculator.process_forecasts(sheet_prefix1, sheet_prefix2)
    s1 = 'g6.3.4.MSFE_Results'
    s2 = 'g6.3.4.MSFE_Ratios'
    # calculator.save_results(s1, s2)
    print("MSFE Results:")

    # Convert the dictionary to a DataFrame
    msfe_results = pd.DataFrame(list(calculator.msfe_results.items()), columns=['MSFE Type', 'Value'])
    print(msfe_results)
    print('-'*50)
    print("MSFE Ratios:")
    msfe_ratios = pd.DataFrame(list(calculator.msfe_ratios.items()), columns=['MSFE Type', 'Value'])
    print(msfe_ratios)


# Main execution block
if __name__ == '__main__':
    # g3_4_demo()
    g6_3_4_demo()