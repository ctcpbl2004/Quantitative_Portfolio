# -*- coding: utf-8 -*-
"""
Created on Sat Dec 17 13:21:52 2016

@author: Raymond
"""

import pandas as pd
import numpy as np
import pandas_datareader.data as web
import matplotlib.pyplot as plt
import matplotlib
import datetime

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

def Equal_Weight_Benchmark(Tickers):
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
    
    
def Performance_Comparision(df,Period):
    
    def Performance_Calculate(Series,Period):
        Period_Close = Series.resample(Period,how = 'last')
        
        Period_Return = (Period_Close/Period_Close.shift(1) - 1).ix[:-1]
        Period_Return.name = 'Return'
        
        Period_Volatility = ((Series/Series.shift(1) - 1).resample(Period,how = np.std)*250.**0.5).ix[:-1]
        Period_Volatility.name = 'Volatility'
        
        Mean_Variance_df = pd.concat([Period_Return,Period_Volatility],axis = 1)

        return Mean_Variance_df
    
    def Equity_Curve(df):
        Cumulative_Return = df/df.ix[0] - 1
        Cumulative_Return.plot()

    if Period == 'Monthly'or Period == 'monthly' or Period == 'Month' or Period == 'month' or Period == 'M' or Period == 'm':
        Settlement_Period = 'M'
    elif Period == 'Quarterly' or Period == 'quarterly' or Period == 'Quarter' or Period == 'quarter' or Period == 'Q' or Period == 'q':
        Settlement_Period = 'Q'
    else:
        print 'Error !! You got the wrong period input.'
   
    color_list = ['blue','red','green','yellow','black','cyan','magenta']    
    
    
    count = 0
    for each in df.columns:
        Performance = Performance_Calculate(Series = df[each],Period = Settlement_Period) * 100.
        plt.scatter(x = Performance['Volatility'], y = Performance['Return'], color = color_list[count],label = each, lw = 2)
        count = count + 1
        

    plt.xlabel('Volatility(%)', fontsize = 20)

    plt.ylabel('Return(%)', fontsize = 20)
    plt.legend()
    plt.show()    
    Equity_Curve(df)

def Concat_data(Tickers,start = '2000-01-01'):
    df = pd.DataFrame(columns = Tickers, index = pd.date_range(start = start, end = datetime.datetime.today()))    
    
    for each in Tickers:
        df[each] = web.DataReader(str(each),'yahoo',start)['Adj Close']
    
    return df.dropna()


    



#=================================================================================================================================
Tickers = ['SPY','IWM','VGK','EWJ','EEM','SHY','IEF' ,'TLT' ,'TIP' ,'AGG' ,'HYG' ,'EMB' ,'VNQ' ,'RWX' ,'PFF' ,'GLD' ,'USO' ,'DBA' ]
'''
Portfolio =  Global_Volatility_Portfolio_Algorithm(Tickers = Tickers, lookback = 30)[1]
Benchmark = Equal_Weight_Benchmark(Tickers = Tickers)

Compare = pd.concat([Portfolio,Benchmark],axis = 1)

Compare = Compare[Compare.index >= '2008-06-30']
Compare = Compare/Compare.ix[0]
Compare.columns = ['Momentum','Benchmark']

Compare_Return = np.std(Compare.pct_change()) * (365**0.5)
print Compare_Return


#Compare.plot()
Performance_Comparision(Compare,'M')
'''

Data = Concat_data(Tickers)
print Data

Return_data = Data.pct_change().dropna()


