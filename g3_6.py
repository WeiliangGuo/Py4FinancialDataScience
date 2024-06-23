import pandas as pd
import numpy as np
from sklearn.linear_model import Lasso, Ridge
import matplotlib.pyplot as plt


class PenalizedModelVisualizer:
    def __init__(self, data_file, in_sample_start, in_sample_end, out_sample_start, out_sample_end):
        self.data_file = data_file
        self.data = None
        self.models = {}
        self.in_sample_start = in_sample_start
        self.in_sample_end = in_sample_end
        self.out_sample_start = out_sample_start
        self.out_sample_end = out_sample_end  

    def load_data(self):
        """
        Load the necessary data from the Excel file.
        """
        self.data = pd.read_excel(self.data_file, sheet_name=None)
        
        # Initialize an empty list to store recursive mean estimates for SP500
        self.data['data']['Date'] = pd.to_datetime(self.data['data']['Date']) 
        
    def fit_models(self):
        """
        Fit Lasso and Ridge regression models for each predictor, handling NaN values.
        """
        
         #Index to filter out-of-sample periods
        self.in_sample_start_index = self.data['data'][(self.data['data']['Date'] == self.in_sample_start)].index  
        self.in_sample_end_index = self.data['data'][(self.data['data']['Date'] == self.in_sample_end)].index
        self.in_sample_start_index = self.in_sample_start_index[0]
        self.in_sample_end_index  = self.in_sample_end_index[0]
        self.out_sample_start_index = self.data['data'][(self.data['data']['Date'] == self.out_sample_start)].index  
        self.out_sample_end_index = self.data['data'][(self.data['data']['Date'] == self.out_sample_end)].index
        self.out_sample_start_index = self.out_sample_start_index[0]
        self.out_sample_end_index  = self.out_sample_end_index[0]
        
        
        self.predictors = ['E12', 'b/m', 'tbl', 'ntis', 'infl']
        self.models = {'stocks': {'lasso': {}, 'ridge': {}}, 'bonds': {'lasso': {}, 'ridge': {}}}

        for asset_class in ['stocks', 'bonds']:        
            target = self.data['data'][f'Excess_Return_{asset_class.capitalize()}'].iloc[self.in_sample_start_index : self.in_sample_end_index+1]      
            target_nonan = target.dropna()
            X = self.data['data'][self.predictors].iloc[self.in_sample_start_index : self.in_sample_end_index+1].values
            y = target_nonan.values
                        
            lasso_model = Lasso(alpha=0.1).fit(X, y)
            self.models[asset_class]['lasso'] = lasso_model

                # Ridge Model
            ridge_model = Ridge(alpha=0.1).fit(X, y)
            self.models[asset_class]['ridge'] = ridge_model
    

    def generate_forecasts(self):
        """
        Generate forecasts using the fitted models.
        """
        forecasts = {'stocks': {}, 'bonds': {}}
        
        X_test = self.data['data'][self.predictors].iloc[self.out_sample_start_index : self.out_sample_end_index+1].values
                
        for asset_class in ['stocks', 'bonds']:
            for model_type in ['lasso', 'ridge']:                
                forecast = self.models[asset_class][model_type].predict(X_test)
                forecasts[asset_class][f'{model_type}'] = forecast
                
        return forecasts

    def plot_forecasts(self, forecasts,
                       sheet_prefix1='g2',
                       sheet_prefix2='g3.3'):
        if sheet_prefix1 == 'g6.2':
            mu = 'Mu'
        else:
            mu = 'Mean'
        """
        Create plots of the forecast time-series.
        """
        #Figure for Stocks
        plt.figure(figsize=(10, 5))
                
        #Define time for x-axis
        time = pd.to_datetime(self.data[f'{sheet_prefix1}.Bonds_Monthly_{mu}_Forecast']['Date'])

        #Plot Lasso and Ridge Stocks
        for model_name, forecast in forecasts['stocks'].items():
             plt.plot(time, forecast, label=model_name)
        
        # Plot benchmark Stocks
        plt.plot(time, self.data[f'{sheet_prefix1}.SP500_Monthly_{mu}_Forecast']['Mean_Forecast'], label='Benchmark', color='green')
        
        #Plot combined Stocks
        plt.plot(time, self.data[f'{sheet_prefix2}_Forecast_Excess_R_Stocks']['Combined_Forecast'], label='Combined Forecast', color='purple')
        plt.title('Stocks Forecast Comparison')
        
        plt.legend()
        plt.grid(True)
        plt.xlabel('Years')
        plt.ylabel('Monthly Excess Return')
        plt.tight_layout()
        plt.show()

# Plot for Bonds - out-of-sample
        plt.figure(figsize=(10, 5))    

        #Plot Lasso and Ridge Bonds
        for model_name, forecast in forecasts['bonds'].items():
            plt.plot(time, forecast, label=model_name)
        
        # Plot benchmark Bonds
        plt.plot(time, self.data['g2.Bonds_Monthly_Mean_Forecast']['Mean_Forecast'], label='Benchmark', color='green')
        
        #Plot combined Bonds
        plt.plot(time, self.data['g3.3_Forecast_Excess_R_Bonds']['Combined_Forecast'], label='Combined Forecast', color='purple')
        
        plt.title('Bonds Forecast Comparison')
        plt.legend()
        plt.grid(True)
        plt.xlabel('Years')
        plt.ylabel('Excess Return')

        plt.tight_layout()
        plt.show()
        

    def run(self,
            sheet_prefix1='g2',
            sheet_prefix2='g3.3'):
        """
        Execute the model fitting, forecasting, and visualization process.
        """
        self.load_data()
        self.fit_models()
        forecasts = self.generate_forecasts()
        self.plot_forecasts(forecasts, sheet_prefix1, sheet_prefix2)
        print("Visualization complete.")

def g3_6_demo():
    
    # Set the boundaries for in-sample and out-of-sample periods
    in_sample_start = "1980-01-01"
    in_sample_end = "1999-12-01"
    out_sample_start = "2000-01-01"
    out_sample_end = "2021-12-01" #This is equivalent to 2021-12-31 (last observation)
    
    visualizer = PenalizedModelVisualizer('data.xlsx', in_sample_start, in_sample_end, out_sample_start, out_sample_end)
    visualizer.run()


def g6_3_6_demo():
    print('Task 6.3.6:')
    # Set the boundaries for in-sample and out-of-sample periods
    in_sample_start = "1980-01-01"
    in_sample_end = "1999-12-01"
    out_sample_start = "2000-01-01"
    out_sample_end = "2021-12-01"  # This is equivalent to 2021-12-31 (last observation)

    visualizer = PenalizedModelVisualizer('data.xlsx', in_sample_start, in_sample_end, out_sample_start, out_sample_end)
    sheet_prefix1 = 'g6.2'
    sheet_prefix2 = 'g6.3.3'
    visualizer.run(sheet_prefix1, sheet_prefix2)


if __name__ == '__main__':
    g3_6_demo()
