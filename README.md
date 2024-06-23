### Instructions for Analysis and Forecasting of U.S. Stock and Bond Returns

#### 1. Data Collection
- **Stock Market Index Data**: Download the monthly prices of the S&P 500 (SP500) from the U.S. Federal Reserve Economic Data (https://fred.stlouisfed.org/) or Bloomberg for the period from December 1979 to December 2021.
- **Bond Index Data**: Download the monthly prices of the Bloomberg Barclays U.S. Aggregate Bond Index (LBUSTRUU) from Bloomberg for the same period.
- **Risk-Free Rate Data**: Download the monthly data on the risk-free rate of return from Professor Kenneth French’s data library (http://mba.tuck.dartmouth.edu/pages/faculty/ken.french/data_library.html) for the same period.

#### 2. Compute Statistics for U.S. Stock and Bond Excess Returns
Write a Python function to compute the following statistics for simple excess returns of U.S. stocks and bonds:
- (a) Annualized mean/average
- (b) Annualized volatility
- (c) Annualized Sharpe ratio
- (d) Skewness
- (e) Kurtosis

Report these summary statistics in a table and provide an explanation of the results.

#### 3. Forecasting Excess Returns
- **Sample Division**: Split the total sample into two periods:
  - In-sample period: January 1980 - December 1999
  - Out-of-sample period: January 2000 - December 2021

- **Recursive Estimation**:
  - Generate a time series of monthly out-of-sample constant expected (mean) excess return forecasts for each asset class using a recursive estimation approach. Write a function to perform this computation.

#### 4. Predictor Variables and Forecasting Models
- **Predictor Variables**: Select five plausible predictors of asset class excess returns based on literature (e.g., Rapach, Ringgenberg, and Zhou, 2016 for stocks, Lin, Wu, and Zhou, 2017).
- **Predictive Models**:
  - For each predictor, use an OLS predictive regression model to generate monthly out-of-sample excess return forecasts for each asset class.
  - Create combination forecasts by averaging the forecasts based on the five predictors from the OLS model.
  
- **Model Performance Evaluation**:
  - Compute the mean squared forecast error (MSFE) for the benchmark forecast.
  - Calculate the ratio of MSFEs for the predictive models relative to the benchmark MSFE.
  - Use the Diebold and Mariano (1995) test to compare the predictive ability of the models relative to the benchmark forecast. Write a function to perform this test, clearly stating the null hypothesis and providing a discussion of the results.
  
- **Visualization**:
  - Create a figure showing the time series of the mean benchmark, combination, and two penalized linear regression excess return forecasts for each asset class.

#### 5. Forecasting the Variance-Covariance Matrix
- **Out-of-Sample Forecasts**: Generate the out-of-sample forecasts of the (2-by-2) sample variance-covariance matrix for a portfolio of the two asset classes using the recursive estimation window approach.

#### 6. Optimal Portfolio Construction
- **Optimal Tangency Portfolio**: Using the mean benchmark excess return forecasts and sample variance-covariance matrix forecasts, construct the out-of-sample optimal tangency portfolio weights for a mean-variance investor.
- **Summary Statistics**:
  - Compute the annualized mean, volatility, and Sharpe ratio for the optimal portfolio’s excess return.
  - Repeat the exercise using the 8 predictive model excess return forecasts and compute the summary statistics for the alternative optimal portfolios.
  
- **Reporting and Visualization**:
  - Report the summary statistics in a table and explain the results, highlighting any evidence of out-of-sample forecasting performance translating into economic gains/significance.
  - Create a figure showing the time series of portfolio weights and cumulative excess returns for the optimal portfolio based on the mean benchmark forecast and those based on the combination and penalized linear regression excess return forecasts.

#### 7. Rolling Window Estimation
- Repeat all the above tasks using a rolling window estimation approach and provide commentary on the results.

This structured approach ensures comprehensive analysis and robust forecasting of U.S. stock and bond returns, providing valuable insights into their performance and predictability.
