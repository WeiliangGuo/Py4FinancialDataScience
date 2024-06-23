import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

class PortfolioVisualizer:
    def __init__(self, file_path):
        self.file_path = file_path

    def create_visualization(self,
                             s1='g5.1.Opt_Tan_Portfolio_W',
                             s2='g5.3.Alt_Opt_Tan_Portfolio_W'):
        # Load data
        returns_data = pd.read_excel(self.file_path, sheet_name='data')
        tan_weights_data = pd.read_excel(self.file_path, sheet_name=s1, index_col=0)
        alt_weights_data = pd.read_excel(self.file_path, sheet_name=s2, index_col=0)

        # Convert 'Date' column to datetime for better x-axis formatting
        returns_data['Date'] = pd.to_datetime(returns_data['Date'], format='%m-%Y')

        # Calculate cumulative returns
        returns_data['Cumulative Return Stocks'] = returns_data['Excess_Return_Stocks'].cumsum()
        returns_data['Cumulative Return Bonds'] = returns_data['Excess_Return_Bonds'].cumsum()

        # Create subplots
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 12), gridspec_kw={'height_ratios': [2, 1]})

        # Subplot for Cumulative Excess Returns
        ax1.plot(returns_data['Date'], returns_data['Cumulative Return Stocks'], label='Cumulative Return Stocks',
                 color='red')
        ax1.plot(returns_data['Date'], returns_data['Cumulative Return Bonds'], label='Cumulative Return Bonds',
                 color='blue')
        ax1.set_title('Cumulative Excess Returns Over Time')
        ax1.set_ylabel('Cumulative Returns')
        ax1.legend()

        # Subplot for Portfolio Weights Allocation
        portfolio_labels = ['Optimal'] + list(alt_weights_data.columns)  # Include all columns
        portfolio_numbers = list(range(1, len(portfolio_labels) + 1))
        weights_stocks = [tan_weights_data.loc['Stocks', 'Weights']] + alt_weights_data.loc['Stocks', :].tolist()
        weights_bonds = [tan_weights_data.loc['Bonds', 'Weights']] + alt_weights_data.loc['Bonds', :].tolist()

        bar_width = 0.35
        indices = np.arange(len(portfolio_numbers))
        ax2.bar(indices - bar_width/2, weights_stocks, bar_width, label='Stocks', color='green')
        ax2.bar(indices + bar_width/2, weights_bonds, bar_width, label='Bonds', color='orange')
        ax2.set_title('Portfolio Weights Allocation')
        ax2.set_xlabel('Portfolio Type')
        ax2.set_xticks(indices)
        ax2.set_xticklabels(portfolio_numbers)
        ax2.set_ylabel('Weights')
        ax2.legend()

        # Adding a legend for portfolio types
        portfolio_legend = {number: label for number, label in zip(portfolio_numbers, portfolio_labels)}
        ax2.text(1.02, 0.5, '\n'.join([f"{num}: {name}" for num, name in portfolio_legend.items()]),
                 transform=ax2.transAxes, verticalalignment='center')

        return fig

def g5_6_demo():
    # Usage
    file_path = 'data.xlsx'
    portfolio_visualizer = PortfolioVisualizer(file_path)
    fig = portfolio_visualizer.create_visualization()

    # Display the figure
    plt.show()

def g6_5_6_demo():
    # Usage
    print('Task 6.5.6:')
    file_path = 'data.xlsx'
    portfolio_visualizer = PortfolioVisualizer(file_path)
    s1 = 'g6.5.1.Opt_Tan_Portfolio_W'
    s2 = 'g6.5.3.Alt_Opt_Tan_Portfolio_W'
    fig = portfolio_visualizer.create_visualization(s1, s2)

    # Display the figure
    plt.show()


if __name__ == '__main__':
    # g5_6_demo()
    g6_5_6_demo()