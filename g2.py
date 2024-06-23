import pandas as pd

class RecursiveEstimator:
    def __init__(self, data, in_sample_end, out_sample_start, column_name):
        self.data = data.dropna(subset=[column_name]).copy()
        self.in_sample_end = pd.to_datetime(in_sample_end)
        self.out_sample_start = pd.to_datetime(out_sample_start)
        self.column_name = column_name

    def generate_forecasts(self):
        self.data['Date'] = pd.to_datetime(self.data['Date']).dt.to_period('M').dt.to_timestamp()
        forecasts_list = []
        
        for current_date in self.data[self.data['Date'] >= self.out_sample_start]['Date']:
            current_in_sample = self.data[(self.data['Date'] < current_date)]
            mean_forecast = current_in_sample[self.column_name].mean() 
            forecasts_list.append({'Date': current_date.strftime('%m-%Y'), 'Mean_Forecast': mean_forecast})

        return pd.DataFrame(forecasts_list)

def g2_demo():
    file_path = 'data.xlsx'
    data = pd.read_excel(file_path)

    in_sample_end = '1999-12-31'
    out_sample_start = '2000-01-01'

    sp500_estimator = RecursiveEstimator(data, in_sample_end, out_sample_start, 'Excess_Return_Stocks')
    bonds_estimator = RecursiveEstimator(data, in_sample_end, out_sample_start, 'Excess_Return_Bonds')

    sp500_forecasts = sp500_estimator.generate_forecasts()
    bonds_forecasts = bonds_estimator.generate_forecasts()

    print("S&P 500 Forecasts Sample:")
    print(f'# of rows in sp500_forecasts: {len(sp500_forecasts)}')
    print(sp500_forecasts.head())
    print("\nBond Forecasts Sample:")
    print(f'# of rows in bonds_forecasts: {len(bonds_forecasts)}')
    print(bonds_forecasts.head())

    # # Saving the forecasts to the data.xlsx file
    # with pd.ExcelWriter(file_path, mode='a', engine='openpyxl') as writer:
    #     sp500_forecasts.to_excel(writer, sheet_name='g2.SP500_Monthly_Mean_Forecast', index=False)
    #     bonds_forecasts.to_excel(writer, sheet_name='g2.Bonds_Monthly_Mean_Forecast', index=False)


class RollingWindowEstimator:
    def __init__(self, data, in_sample_end, out_sample_start, column_name, window_size):
        self.data = data.dropna(subset=[column_name]).copy()
        self.in_sample_end = pd.to_datetime(in_sample_end)
        self.out_sample_start = pd.to_datetime(out_sample_start)
        self.column_name = column_name
        self.window_size = window_size

    def generate_forecasts(self):
        self.data['Date'] = pd.to_datetime(self.data['Date']).dt.to_period('M').dt.to_timestamp()
        forecasts_list = []

        for current_date in self.data[self.data['Date'] >= self.out_sample_start]['Date']:
            current_in_sample = self.data[(self.data['Date'] < current_date)]
            if len(current_in_sample) > self.window_size:
                current_in_sample = current_in_sample.iloc[-self.window_size:]
            mean_forecast = current_in_sample[self.column_name].mean()
            forecasts_list.append({'Date': current_date.strftime('%m-%Y'), 'Mean_Forecast': mean_forecast})

        return pd.DataFrame(forecasts_list)


def g6_2_demo():
    print('Task 2 using rolling window estimator:')
    file_path = 'data.xlsx'
    data = pd.read_excel(file_path)
    window_size = 12

    in_sample_end = '1999-12-31'
    out_sample_start = '2000-01-01'

    sp500_estimator = RollingWindowEstimator(data, in_sample_end, out_sample_start, 'Excess_Return_Stocks', window_size)
    bonds_estimator = RollingWindowEstimator(data, in_sample_end, out_sample_start, 'Excess_Return_Bonds', window_size)

    sp500_forecasts = sp500_estimator.generate_forecasts()
    bonds_forecasts = bonds_estimator.generate_forecasts()

    print("S&P 500 Forecasts Sample:")
    print(f'# of rows in sp500_forecasts: {len(sp500_forecasts)}')
    print(sp500_forecasts.head())
    print("\nBond Forecasts Sample:")
    print(f'# of rows in bonds_forecasts: {len(bonds_forecasts)}')
    print(bonds_forecasts.head())

    # # Saving the forecasts to the data.xlsx file
    # with pd.ExcelWriter(file_path, mode='a', engine='openpyxl') as writer:
    #     sp500_forecasts.to_excel(writer, sheet_name='g6.2.SP500_Monthly_Mu_Forecast', index=False)
    #     bonds_forecasts.to_excel(writer, sheet_name='g6.2.Bonds_Monthly_Mu_Forecast', index=False)


if __name__ == '__main__':
    g2_demo()
    print('-'*40)
    g6_2_demo()