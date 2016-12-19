# -*- coding: utf-8 -*-
"""
Created on Sat Dec 17 23:30:17 2016

@author: Raymond
"""

import pandas as pd
import pandas_datareader.data as web
import numpy as np
import datetime

'''
The comparison of Equal weight, Volatility weight and adjusted Risk Parity's performance.
The strategies of each methodology don't contain the portfolio choice problem,
the portfolios always have the same components, the difference are the weighting methodologies.
'''


           
def Concat_data(Tickers,start = '2000-01-01'):
    df = pd.DataFrame(columns = Tickers, index = pd.date_range(start = start, end = datetime.datetime.today()))    
    
    for each in Tickers:
        df[each] = web.DataReader(str(each),'yahoo',start)['Adj Close']
    
    return df.dropna()


def Equal_Weight(Tickers):
    Weight = 1./len(Tickers)
    
    Price_df = pd.DataFrame()    
    
    for Ticker in Tickers:
        Price_df[Ticker] = web.DataReader(Ticker,'yahoo','2000-01-01')['Adj Close']
        
    Price_df = Price_df.resample('D', how='last')
    Price_df = Price_df.fillna(method='ffill')
    
    Price_df = Price_df.dropna()

    Return_df = Price_df.pct_change() * Weight
    
    Benchmark_Return = Return_df.sum(axis = 1) + 1.
    Benchmark_CR = Benchmark_Return.cumprod()
    return Benchmark_CR



def Volatility_Weight(Tickers,lookback):
    df = pd.DataFrame()

    for Ticker in Tickers:
        df[Ticker] = web.DataReader(Ticker,'yahoo','2000-01-01')['Adj Close']

    Return_df = df.dropna().pct_change()

    Volatility_df = 1. / pd.rolling_std(Return_df, lookback) * (250 ** 0.5)
    Volatility_df['Sum'] = Volatility_df.sum(axis=1)
    #temp = pd.rolling_std(Return_df, 60) * (250 ** 0.5)

    Weight_df = pd.DataFrame()

    for Ticker in Tickers:
        Weight_df[Ticker] = Volatility_df[Ticker] / Volatility_df['Sum']

    Weight_df = Weight_df.resample('M', how='mean').resample('D', how='last')
    Weight_df = Weight_df.fillna(method='ffill').dropna()

    Holding_Return = (Weight_df * Return_df).dropna()

    Fund_Return = Holding_Return.sum(axis=1) + 1.
    CR = Fund_Return.cumprod()

    return CR

def Global_Volatility_Portfolio_Algorithm(Tickers, lookback):
    
    Process = ['Price', 'Daily_Return', 'Momentum', 'Volatility']

    Price_df = pd.DataFrame()

    for Ticker in Tickers:
        Price_df[Ticker] = web.DataReader(Ticker,'yahoo','2000-01-01')['Adj Close']
    Price_df = Price_df.resample('D', how='last')
    Price_df = Price_df.fillna(method='ffill')

    Price_df = Price_df.dropna()
    Panel = pd.Panel(items=Process, major_axis=Price_df.index, minor_axis=Tickers)
    Panel['Price'] = Price_df
    Panel['Daily_Return'] = Panel['Price'].pct_change()

    Panel['Momentum'] = Panel['Price'].pct_change(lookback).resample('M', how='mean')
    Panel['Momentum'] = Panel['Momentum'].fillna(method='ffill')

    Panel['Volatility'] = pd.rolling_std(Panel['Daily_Return'], lookback).resample('M', how='mean')
    Panel['Volatility'] = Panel['Volatility'].fillna(method='ffill')

    Panel['Volatility_Inverse'] = 1. / Panel['Volatility']

    Panel['Momentum_Rank'] = Panel['Momentum'].rank(axis=1, ascending=True)

    Panel['Selected_asset'] = np.where(Panel['Momentum_Rank'] >= len(Tickers)/2., 1, 0)
    # Panel['Selected_asset'] = np.where(Panel['Momentum_Rank']<=len(Tickers)/3,1,0)
    Panel['Select_Volatility'] = Panel['Volatility_Inverse'] * Panel['Selected_asset']

    Volatility_weight = Panel['Select_Volatility'].replace(0, np.nan)
    Volatility_sum = Volatility_weight.sum(axis=1).replace(0, np.nan)

    Weight_df = pd.DataFrame()
    for Ticker in Volatility_weight.columns:
        Weight_df[Ticker] = Volatility_weight[Ticker] / Volatility_sum

    Panel['Volatility_weight'] = Weight_df

    Portfolio_Return_temp = Panel['Volatility_weight'] * Panel['Daily_Return']

    Portfolio_Return = Portfolio_Return_temp.sum(axis=1) + 1.

    Portfolio_CR = Portfolio_Return.cumprod()

    #Portfolio_CR = Portfolio_Return.cumsum().apply(np.exp)
    # Output:Last weight, Cumulative Return
    Turnover_temp = Panel['Volatility_weight'].diff().abs()
    Total_Turnover = Turnover_temp.abs().sum().sum()

    # print Panel['Volatility_weight'].ix[-1].fillna(0)

    return Panel['Volatility_weight'].ix[-1].fillna(0), Portfolio_CR, Total_Turnover


def Var_Cov_Weight(Tickers,lookback):
    df = pd.DataFrame()

    for Ticker in Tickers:
        df[Ticker] = web.DataReader(Ticker,'yahoo','2000-01-01')['Adj Close']

    Return_df = df.dropna().pct_change()

    Panel = pd.rolling_cov(Return_df,lookback)
    
    Var_Cov_df = pd.DataFrame(index = Panel.items, columns = Tickers)
    
    
    for Date in Panel.items:
        Var_Cov_df.ix[Date] = Panel[Date].sum().tolist()

    return Var_Cov_df


#==============================================================================
Tickers = ['SPY','IWM','VGK','EWJ','EEM','SHY','IEF' ,'TLT' ,'TIP' ,'AGG' ,'HYG' ,
           'EMB' ,'VNQ' ,'RWX' ,'PFF' ,'GLD' ,'USO' ,'DBA' ]
'''
Equal_Weight = Equal_Weight(Tickers = Tickers)
Volatility_Weight = Volatility_Weight(Tickers = Tickers,lookback = 30)
Volatility_Select = Global_Volatility_Portfolio_Algorithm(Tickers = Tickers, lookback = 30)[1]
Comparison = pd.concat([Equal_Weight,Volatility_Weight,Volatility_Select],axis = 1)
Comparison.columns = ['Equal Weight','Volatility Weight','Volatility Select and Weighting']
Comparison = Comparison.dropna()
Comparison = Comparison/Comparison.ix[0]

Comparison.plot()
'''

Var_Cov =  Var_Cov_Weight(Tickers = Tickers,lookback = 30)
Var_Cov = Var_Cov.replace(0,np.nan).dropna()
Var_Cov = 1./Var_Cov
Var_Cov['Total'] = Var_Cov.sum(axis = 1)

for each in Var_Cov.columns:
    Var_Cov[each] = Var_Cov[each]/Var_Cov['Total']

del Var_Cov['Total']    

print Var_Cov
#Negative problem!!






