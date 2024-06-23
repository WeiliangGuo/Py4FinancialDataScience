import pandas as pd

class ComparativeAnalysis:
    def __init__(self, file_path, benchmark_sheet, alternative_sheet):
        self.file_path = file_path
        self.benchmark_sheet = benchmark_sheet
        self.alternative_sheet = alternative_sheet
        self.benchmark_data = None
        self.alternative_data = None

    def load_data(self):
        """ Load the benchmark and alternative portfolio statistics from the Excel file. """
        self.benchmark_data = pd.read_excel(self.file_path, sheet_name=self.benchmark_sheet)
        self.alternative_data = pd.read_excel(self.file_path, sheet_name=self.alternative_sheet)

    def perform_comparative_analysis(self):
        """ Compare the summary statistics of the optimal portfolio based on the mean benchmark
        forecast with those based on predictive model forecasts. """
        # Creating a separate benchmark DataFrame with a 'Model' column
        benchmark_stats = pd.DataFrame(self.benchmark_data.iloc[0]).T
        benchmark_stats['Model'] = 'Benchmark'
        benchmark_stats = benchmark_stats[['Model', 'Mean Return', 'Volatility',
                                           'Sharpe Ratio']]

        # Concatenating the benchmark data with the alternative data
        comparison = pd.concat([self.alternative_data, benchmark_stats],
                               ignore_index=True)
        comparison.drop('Unnamed: 0', axis=1, inplace=True)
        return comparison


def g5_5_demo():
    analyst = ComparativeAnalysis('data.xlsx',
                                  'g5.2.Opt_Portfolio_Stats',
                                  'g5.4.Alt_Opt_Portfolio_Stats')
    analyst.load_data()
    comparison_results = analyst.perform_comparative_analysis()
    #
    # Saving the results to the Excel file
    with pd.ExcelWriter('data.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        comparison_results.to_excel(writer, sheet_name='g5.5.Comparative_Analysis', index=False)
    return comparison_results

def g6_5_5_demo():
    print('Task 6.5.5:')
    analyst = ComparativeAnalysis('data.xlsx',
                                  'g6.5.2.Opt_Portfolio_Stats',
                                  'g6.5.4.Alt_Opt_Portfolio_Stats')
    analyst.load_data()
    comparison_results = analyst.perform_comparative_analysis()

    # # Saving the results to the Excel file
    # with pd.ExcelWriter('data.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    #     comparison_results.to_excel(writer, sheet_name='g6.5.5.Comparative_Analysis', index=False)
    return comparison_results


if __name__ == '__main__':
    # print(g5_5_demo())
    print(g6_5_5_demo())