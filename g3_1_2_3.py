import os
import pandas as pd
from sklearn.linear_model import LinearRegression
from sklearn.model_selection import train_test_split, cross_val_score
import numpy as np
import pickle
import statsmodels.api as sm

class OLSModeler:
    def __init__(self, data, predictors, target, in_sample_start, in_sample_end, out_sample_start, out_sample_end):
        self.data = data
        self.predictors = predictors
        self.target = target
        self.in_sample_start = in_sample_start
        self.in_sample_end = in_sample_end
        self.out_sample_start = out_sample_start
        self.out_sample_end = out_sample_end      

    def train_and_evaluate_models(self):

        # Initialize an empty list to store recursive mean estimates for SP500
        self.data['Date'] = pd.to_datetime(self.data['Date'])         

        # Filter data for in-sample and out-of-sample periods
        in_sample_end_dyn = self.data[(self.data['Date'] >= self.out_sample_start) & (self.data['Date'] <= self.out_sample_end)].copy()['Date']        
        out_sample_data = self.data[(self.data['Date'] >= self.out_sample_start) & (self.data['Date'] <= self.out_sample_end)].copy()
        
        #Initialized the forecast_df
        self.forecast_df = in_sample_end_dyn.reset_index()
                
        #Drop colum of old index
        self.forecast_df = self.forecast_df.drop(columns=['index'])
        
        for predictor in self.predictors:
            recursive_means = []            
        #Does a recursive approach of re-evaluating the OLS model for each new monthly observation
            for i in in_sample_end_dyn:
                in_sample_data = self.data[(self.data['Date'] >= self.in_sample_start) & (self.data['Date'] < i)]
                out_sample_point = self.data[(self.data['Date'] == i)].copy()
            #x and y for OLS regression
                x = in_sample_data[predictor]
                y = in_sample_data[self.target]            
            # Calculate the OLS for the in-sample period and get predicted "a + bx" results for the entire data
                x = sm.add_constant(x)
                model = sm.OLS(endog = y, exog = x)
                results = model.fit()
            # This initial prediction (in-sample and out-sample) will be used for the recursive estimation
                ypred = results.predict((1, float(out_sample_point[predictor])))    
            #It's necessary to align prediction dimensions
                ypred = ypred[0]
            #Append the estimate
                recursive_means.append(ypred)
                
        #Consolidate all forecasts in one column
            self.recursive_means_df = pd.DataFrame(recursive_means, columns = ['Forecast_'+ predictor])      
        #For each predictor, concatenate it
            self.forecast_df = pd.merge(self.forecast_df, self.recursive_means_df, left_index=True, right_index=True, how = 'left')
            
    def combined_forecast(self):
        #Combined forecast(simple average)        
        self.forecast_df['Combined_Forecast'] = self.forecast_df.mean(axis= 1, numeric_only = True)      
        self.forecast_df['Date'] = self.forecast_df['Date'].dt.strftime('%m-%Y')
                
    def save_model(self, task_prefix='g3.3'):
        forecast_sheet_name = f'{task_prefix}_Forecasts_{self.target}'
        print(forecast_sheet_name)
        
        if 'Stocks' in forecast_sheet_name:
            forecast_sheet_name = f'{task_prefix}_Forecast_Excess_R_Stocks'
        else:
            forecast_sheet_name = f'{task_prefix}_Forecast_Excess_R_Bonds'
        
        print(self.forecast_df.head())
        
        # with pd.ExcelWriter('data.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        #     self.forecast_df.to_excel(writer, sheet_name=forecast_sheet_name)
        # print(f"Forecasts saved in sheet: {forecast_sheet_name}")
        # print('-'*50)

def g3_1_2_3_demo():
    file_path = 'data.xlsx'
    data = pd.read_excel(file_path, sheet_name='data')

    predictors = ['E12', 'b/m', 'tbl', 'ntis', 'infl']
    targets = ['Excess_Return_Stocks', 'Excess_Return_Bonds']
    
    # Set the boundaries for in-sample and out-of-sample periods
    in_sample_start = "1980-01-01"
    in_sample_end = "1999-12-01"
    out_sample_start = "2000-01-01"
    out_sample_end = "2021-12-01" #This is equivalent to 2021-12-31 (last observation)
    
    for target in targets:
        #Now receives data and return 'Forecast_ + predictor' saved in a sheet for target
        modeler = OLSModeler(data, predictors, target, in_sample_start, in_sample_end, out_sample_start, out_sample_end)
        modeler.train_and_evaluate_models()
        modeler.combined_forecast()
        modeler.save_model()


def g6_3_1_2_3_demo():
    print('Task 6.3.1-3:')
    file_path = 'data.xlsx'
    data = pd.read_excel(file_path, sheet_name='data')

    predictors = ['E12', 'b/m', 'tbl', 'ntis', 'infl']
    targets = ['Excess_Return_Stocks', 'Excess_Return_Bonds']

    # Set the boundaries for in-sample and out-of-sample periods
    in_sample_start = "1980-01-01"
    in_sample_end = "1999-12-01"
    out_sample_start = "2000-01-01"
    out_sample_end = "2021-12-01"  # This is equivalent to 2021-12-31 (last observation)

    for target in targets:
        # Now receives data and return 'Forecast_ + predictor' saved in a sheet for target
        modeler = OLSModeler(data, predictors, target, in_sample_start, in_sample_end, out_sample_start, out_sample_end)
        modeler.train_and_evaluate_models()
        modeler.combined_forecast()
        modeler.save_model(task_prefix='g6.3.3')


if __name__ == '__main__':
    # g3_1_2_3_demo()
    g6_3_1_2_3_demo()
    
