# -*- coding: utf-8 -*-
"""
Created on Wed Jun 22 09:56:22 2016

@author: Raymond
"""

import sqlite3
import pandas as pd
import xlsxwriter
'''
#==============================================================================
#Update_table
#==============================================================================
'''
def Fetch_All_Data():
    connection = sqlite3.connect('db/Taishin_Database.db')
    cursor = connection.cursor()
    cursor.execute("DELETE FROM Update_table")
    cursor.execute("SELECT Ticker, Name, COUNT(Ticker), MIN(Date), MAX(Date) FROM Time_Series GROUP BY Ticker")
    
    data = cursor.fetchall()
    
    for each in data:
        cursor.execute("INSERT OR REPLACE INTO Update_table VALUES (?,?,?,?,?)",(each[0],each[1],each[2],each[3],each[4]))
    connection.commit()
    connection.close()


'''
#==============================================================================
#Delete data
#==============================================================================
'''
def Delete_data(Ticker,Name):
    connection = sqlite3.connect('db/Taishin_Database.db')
    cursor = connection.cursor()
    cursor.execute("DELETE FROM Time_Series WHERE Ticker=? AND Name=?",(Ticker,Name))
    connection.commit()
    connection.close()
    Fetch_All_Data()


'''
#==============================================================================
#Generate Update.xlsx =====> Bloomberg_Update(Tickers,Start,rows)
#==============================================================================
'''
def Fetch_All_Tickers():
    connection = sqlite3.connect('db/Taishin_Database.db')
    cursor = connection.cursor()
    cursor.execute("SELECT Ticker, Name FROM Time_Series GROUP BY Ticker")
    
    data = cursor.fetchall()
    connection.close()
    
    Ticker_list = []
    Name_list = []
    for each in data:
        Ticker_list.append(each[0])
        Name_list.append(each[1])
    return Ticker_list,Name_list



def Word_code(n):
    Word_dict = {1:'A',2:'B',3:'C',4:'D',5:'E',6:'F',7:'G',8:'H',9:'I',10:'J',
                 11:'K',12:'L',13:'M',14:'N',15:'O',16:'P',17:'Q',18:'R',19:'S',20:'T',
                 21:'U',22:'V',23:'W',24:'X',25:'Y',26:'Z'}    
    
    error_list = range(0,500,26)
    error_list[0] = 1
    
    
    
    if (n <= 26.):
        return Word_dict[n]
    
    elif (n>26) & (n<=26**2):
        try:
            second_digit = int(n/26)
            first_digit = n % 26
            return Word_dict[second_digit]+Word_dict[first_digit]
        except:
            first_digit = 26
            second_digit = second_digit -1
            return Word_dict[second_digit]+Word_dict[first_digit]


def Bloomberg_Update(Tickers,Start,rows):
    Start = Start.replace('-','/')
    workbook = xlsxwriter.Workbook('Update.xlsx')
    worksheet = workbook.add_worksheet()  
    
    NumberOfTicker = len(Tickers)
    
    #Write Tickers
    for i in range(NumberOfTicker):
        worksheet.write(0, i+1, Tickers[i])
    
    for j in range(NumberOfTicker-1):
        Cell = str(Word_code(j+3))+'1'
        Bloomberg_formula = '=BDH('+Cell+',"PX_LAST","'+Start+'","","Dir=V","Dts=H","Sort=D","Quote=C","QtTyp=Y","Days=A","Per=cd","DtFmt=D","Fill=B","UseDPDF=N","CshAdjNormal=Y","CshAdjAbnormal=Y","CapChg=Y","cols=1;rows='+str(rows)+'")'
        worksheet.write_formula(1, 2+j, Bloomberg_formula)
    
    
    
    Bloomberg_formula_date = '=BDH(B1,"PX_LAST","'+Start+'","","Dir=V","Dts=S","Sort=D","Quote=C","QtTyp=Y","Days=A","Per=cd","DtFmt=D","Fill=B","UseDPDF=N","CshAdjNormal=Y","CshAdjAbnormal=Y","CapChg=Y","cols=2;rows='+str(rows)+'")'
    worksheet.write_formula(1, 0, Bloomberg_formula_date)
    workbook.close()



def Create_Update_xlsx(Start):
    Tickers = Fetch_All_Tickers()[0]
    Bloomberg_Update(Tickers,Start,rows=500)

'''
#==============================================================================
#Read Historical data from Update.xlsx into DB ====> Data_to_db(File_name='Update.xlsx')
#==============================================================================
Don't have Update_list in db, I have to figure out the way to remain the same
function without creating new table.
'''

def Find_Name(Ticker):
    connection = sqlite3.connect('db/Taishin_Database.db')
    cursor = connection.cursor()
    cursor.execute("SELECT Name FROM Update_table WHERE Ticker =?",(Ticker,))
    data = cursor.fetchall()
    connection.close()
    return data[0][0]


def Series_to_Dataframe(Ticker,Name,Series):
    df = pd.DataFrame(columns=['Ticker','Name','Date','Value'])
    df['Date'] = Series.index
    df['Value'] = Series.tolist()
    df['Ticker'] = Ticker
    df['Name'] = Name
    return df


def Dataframe_to_db(df):
    connection = sqlite3.connect('db/Taishin_Database.db')
    cursor = connection.cursor()
    NumberOfData = len(df)
    
    for i in range(NumberOfData):
        Ticker = df.ix[1][0]
        Name = df.ix[1][1]
        Date = str(df.ix[i][2])[:10]
        Value = float(df.ix[i][3])
        cursor.execute("INSERT OR REPLACE INTO Time_Series VALUES (?, ?, ?, ?)",(Ticker, Name, Date, Value))
        print Date,Value
    connection.commit()
    connection.close()

def Data_to_db(File_name):
    df = pd.read_excel(File_name).ix[1:]
    df = df.sort_index()
    Commodities_list = df.columns
    #print df
    
    for Ticker in Commodities_list:
        try:    
            Name = Find_Name(Ticker)
            print Name
            Data = df[str(Ticker)].dropna()
            Data_df = Series_to_Dataframe(Ticker,Name,Series=Data)
            Dataframe_to_db(Data_df)
        except:
            pass


'''
===============================================================================
Fetch Single Ticker Data
===============================================================================
'''
def Fetch(Ticker,Start=None,End=None):
    connection = sqlite3.connect('db/Taishin_Database.db')
    cursor = connection.cursor()
    
    if (Start==None) & (End==None):
        cursor.execute("SELECT Date,Value FROM Time_Series WHERE Ticker = ?",(Ticker,))
    elif End == None:
        cursor.execute("SELECT Date,Value FROM Time_Series WHERE Ticker = ? AND Date > ?",(Ticker,Start))
    elif Start == None:
        cursor.execute("SELECT Date,Value FROM Time_Series WHERE Ticker = ? AND Date < ?",(Ticker,End))
    elif (Start!=None) & (End!=None):
        cursor.execute("SELECT Date,Value FROM Time_Series WHERE Ticker = ? AND Date BETWEEN ? AND ? ",(Ticker,Start,End))
    else:
        print 'TIMEZONE ERROR!!!!'

    data = cursor.fetchall()
    connection.close()
        
    Date_list = []
    Value_list = []        
    
    for row in data:
        Date_list.append(row[0])
        Value_list.append(row[1])
        
    df = pd.DataFrame(Value_list,index=Date_list,columns = [str(Ticker)])
    df.index = df.index.to_datetime()
    df.sort_index()
    
    return df

'''
===============================================================================
Add New Ticker and Data
===============================================================================
'''
def Add_Index(File_name,Ticker,Name):
    Data = pd.read_excel(File_name).ix[1:].dropna()
    print Data
    df = Series_to_Dataframe(Ticker,Name,Series=Data[str(Data.columns[0])])
    Dataframe_to_db(df)


#Add_Index(File_name='Test.xlsx',Ticker='AUD Curncy',Name='AUD')

    
def Add_Index_to_Strategy_table(Ticker,Name):
    connection = sqlite3.connect('db/Taishin_Database.db')
    cursor = connection.cursor()
    cursor.execute("INSERT OR REPLACE INTO Strategy_table VALUES (?, ?, ?, ?, ?, ?, ?, ?)",(Ticker, Name, '' ,'' ,'','','',''))
    connection.commit()
    connection.close()

#Delete_data(Ticker = 'AUD',Name = 'AUDUSD')

def Delete_Index_to_Strategy_table(Ticker,Name):
    connection = sqlite3.connect('db/Taishin_Database.db')
    cursor = connection.cursor()
    cursor.execute("DELETE FROM Strategy_table WHERE Ticker=? AND Name=?",(Ticker,Name))
    connection.commit()
    connection.close()
    Fetch_All_Data()


