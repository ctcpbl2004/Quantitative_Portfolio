# -*- coding: utf-8 -*-
"""
Created on Wed Jun 29 11:48:33 2016

@author: Raymond
"""
import warnings
import Tkinter as tk
import ttk
#import Database_Functions
import sqlite3
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import numpy as np
import pandas as pd
import datetime
import xlsxwriter

class Database_Functions(object):
    '''
    #==============================================================================
    #Update_table
    #==============================================================================
    '''
    @staticmethod
    def Fetch_All_Data():
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
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
    @staticmethod
    def Delete_data(Ticker,Name):
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
        cursor = connection.cursor()
        cursor.execute("DELETE FROM Time_Series WHERE Ticker=? AND Name=?",(Ticker,Name))
        connection.commit()
        connection.close()
        Database_Functions.Fetch_All_Data()
    
    
    '''
    #==============================================================================
    #Generate Update.xlsx =====> Bloomberg_Update(Tickers,Start,rows)
    #==============================================================================
    '''
    @staticmethod
    def Fetch_All_Tickers():
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
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
    
    
    @staticmethod
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
    
    @staticmethod
    def Bloomberg_Update(Tickers,Start,rows):
        Start = Start.replace('-','/')
        workbook = xlsxwriter.Workbook('Update.xlsx')
        worksheet = workbook.add_worksheet()  
        
        NumberOfTicker = len(Tickers)
        
        #Write Tickers
        for i in range(NumberOfTicker):
            worksheet.write(0, i+1, Tickers[i])
        
        for j in range(NumberOfTicker-1):
            Cell = str(Database_Functions.Word_code(j+3))+'1'
            Bloomberg_formula = '=BDH('+Cell+',"PX_LAST","'+Start+'","","Dir=V","Dts=H","Sort=D","Quote=C","QtTyp=Y","Days=A","Per=cd","DtFmt=D","Fill=B","UseDPDF=N","CshAdjNormal=Y","CshAdjAbnormal=Y","CapChg=Y","cols=1;rows='+str(rows)+'")'
            worksheet.write_formula(1, 2+j, Bloomberg_formula)
        
        
        
        Bloomberg_formula_date = '=BDH(B1,"PX_LAST","'+Start+'","","Dir=V","Dts=S","Sort=D","Quote=C","QtTyp=Y","Days=A","Per=cd","DtFmt=D","Fill=B","UseDPDF=N","CshAdjNormal=Y","CshAdjAbnormal=Y","CapChg=Y","cols=2;rows='+str(rows)+'")'
        worksheet.write_formula(1, 0, Bloomberg_formula_date)
        workbook.close()
    
    
    @staticmethod
    def Create_Update_xlsx(Start):
        Tickers = Database_Functions.Fetch_All_Tickers()[0]
        Database_Functions.Bloomberg_Update(Tickers,Start,rows=500)
    
    '''
    #==============================================================================
    #Read Historical data from Update.xlsx into DB ====> Data_to_db(File_name='Update.xlsx')
    #==============================================================================
    Don't have Update_list in db, I have to figure out the way to remain the same
    function without creating new table.
    '''
    @staticmethod
    def Find_Name(Ticker):
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
        cursor = connection.cursor()
        cursor.execute("SELECT Name FROM Update_table WHERE Ticker =?",(Ticker,))
        data = cursor.fetchall()
        connection.close()
        return data[0][0]
    
    @staticmethod
    def Series_to_Dataframe(Ticker,Name,Series):
        df = pd.DataFrame(columns=['Ticker','Name','Date','Value'])
        df['Date'] = Series.index
        df['Value'] = Series.tolist()
        df['Ticker'] = Ticker
        df['Name'] = Name
        return df
    
    @staticmethod
    def Dataframe_to_db(df):
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
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
    @staticmethod
    def Data_to_db(File_name):
        df = pd.read_excel(File_name).ix[1:]
        df = df.sort_index()
        Commodities_list = df.columns
        #print df
        
        for Ticker in Commodities_list:
            try:    
                Name = Database_Functions.Find_Name(Ticker)
                print Name
                Data = df[str(Ticker)].dropna()
                Data_df = Database_Functions.Series_to_Dataframe(Ticker,Name,Series=Data)
                Database_Functions.Dataframe_to_db(Data_df)
            except:
                pass
    
    
    '''
    ===============================================================================
    Fetch Single Ticker Data
    ===============================================================================
    '''
    @staticmethod
    def Fetch(Ticker,Start=None,End=None):
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
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
    @staticmethod
    def Add_Index(File_name,Ticker,Name):
        Data = pd.read_excel(File_name).ix[1:].dropna()
        print Data
        df = Database_Functions.Series_to_Dataframe(Ticker,Name,Series=Data[str(Data.columns[0])])
        Database_Functions.Dataframe_to_db(df)
    
    
    #Add_Index(File_name='Test.xlsx',Ticker='AUD Curncy',Name='AUD')
    
    @staticmethod
    def Add_Index_to_Strategy_table(Ticker,Name):
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
        cursor = connection.cursor()
        cursor.execute("INSERT OR REPLACE INTO Strategy_table VALUES (?, ?, ?, ?, ?, ?, ?, ?)",(Ticker, Name, '' ,'' ,'','','',''))
        connection.commit()
        connection.close()
    
    #Delete_data(Ticker = 'AUD',Name = 'AUDUSD')
    @staticmethod
    def Delete_Index_to_Strategy_table(Ticker,Name):
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
        cursor = connection.cursor()
        cursor.execute("DELETE FROM Strategy_table WHERE Ticker=? AND Name=?",(Ticker,Name))
        connection.commit()
        connection.close()
        Database_Functions.Fetch_All_Data()


class Market_Timing(tk.Tk):
    def __init__(self):
        self.root = tk.Tk()
        self.root.wm_title('Market Timing 1.0')
        self.root.geometry('1350x760')
        self.root.iconbitmap('D:/Taishin_Platform/pics/Data_Manager.ico')
        self.root._offsetx = 0
        self.root._offsety = 0

        self.Frame_Top=tk.Frame(self.root, width=2000, height=70, background="#282828")
        self.Frame_Top.place(x=0, y=0)
        self.Frame_Down=tk.Frame(self.root, width=1300, height=700, background="#F0F0F0")
        self.Frame_Down.place(x=0, y=70)
        self.Frame_Top.bind('<Button-1>',self.clickwin)
        self.Frame_Top.bind('<B1-Motion>',self.dragwin)
        self.Frame_Down.bind('<Button-1>',self.clickwin)
        self.Frame_Down.bind('<B1-Motion>',self.dragwin)
        
        
        self.Output_Button = ttk.Button(self.Frame_Top,text='Output as Excel(xlsx)',command = self.Output_xls)
        self.Output_Button.place(x=20, y=5)
        
        self.Backtest_icon = tk.PhotoImage(file = 'pics/Excel.gif')
        self.Output_Button.config(image=self.Backtest_icon,compound='left')
        self.Backtest_icon_Adj = self.Backtest_icon.subsample(3,3)
        self.Output_Button.config(image=self.Backtest_icon_Adj)


        
        self.Notebook1 = ttk.Notebook(self.Frame_Down,height=620,width = 1250, style="TNotebook")
        self.Notebook1.place(x=20,y=10)
        
        
        self.Strategy_frame = tk.Frame(self.Notebook1, background="white")
        self.Score_frame = tk.Frame(self.Notebook1, background="white")
        self.Parameter_frame = tk.Frame(self.Notebook1, background="white")
        
        
        self.Note_label = tk.Label(self.Strategy_frame,text = '* : Highest Win(%);     ** : Highest Expected Return',font=('Arial',10),fg='black',background="#FFFFFF")
        self.Note_label.place(x =1190,y = 10 ,anchor='ne')
        
        self.Notebook1.add(self.Strategy_frame, text="  Market Timing Table  ")
        self.Notebook1.add(self.Score_frame, text="  Historical Score  ")
        self.Notebook1.add(self.Parameter_frame, text="  Parameters Table  ")


        self.Strategy_Table = ttk.Treeview(self.Strategy_frame,height="27")

        self.Strategy_Table["columns"]=("column1","column2",'column3','column4','column5','column6','column7','column8','column9')
        self.Strategy_Table.column("#0",width=40, anchor='e')
        self.Strategy_Table.column("column1", width=200, anchor='w' )
        self.Strategy_Table.column("column2", width=380, anchor='w')
        self.Strategy_Table.column("column3", width=80 , anchor='center')
        self.Strategy_Table.column("column4", width=80 , anchor='center')
        self.Strategy_Table.column("column5", width=80 , anchor='center')
        self.Strategy_Table.column("column6", width=80 , anchor='center')
        self.Strategy_Table.column("column7", width=80 , anchor='center')
        self.Strategy_Table.column("column8", width=80 , anchor='center')
        self.Strategy_Table.column("column9", width=80 , anchor='center')
        
        self.Strategy_Table.heading('#0', text='#',command= lambda : self.treeview_sort_column(self.Strategy_Table, "column1", False))
        self.Strategy_Table.heading("column1", text="Bloomberg Ticker",command= lambda : self.treeview_string_sort_column(self.Strategy_Table, "column1", False))
        self.Strategy_Table.heading("column2", text="Name",command= lambda : self.treeview_string_sort_column(self.Strategy_Table, "column2", False))
        self.Strategy_Table.heading("column3", text="RSI",command= lambda : self.treeview_string_sort_column(self.Strategy_Table, "column3", False))
        self.Strategy_Table.heading("column4", text="KD",command= lambda : self.treeview_string_sort_column(self.Strategy_Table, "column4", False))
        self.Strategy_Table.heading("column5", text="EMA",command= lambda : self.treeview_string_sort_column(self.Strategy_Table, "column5", False))
        self.Strategy_Table.heading("column6", text="Break",command= lambda : self.treeview_string_sort_column(self.Strategy_Table, "column6", False))
        self.Strategy_Table.heading("column7", text="MTM",command= lambda : self.treeview_string_sort_column(self.Strategy_Table, "column7", False))
        self.Strategy_Table.heading("column8", text="ATR",command= lambda : self.treeview_string_sort_column(self.Strategy_Table, "column8", False))
        self.Strategy_Table.heading("column9", text="Score",command= lambda : self.treeview_sort_column(self.Strategy_Table, "column9", False))
        

        self.Strategy_Table.place(x=10, y=30)
        self.Strategy_Table.bind("<Double-1>", self.Double_Click_Selection)
        #self.Strategy_Table.tag_configure('ttk', background='yellow')
        #self.Strategy_Table.tag_bind('ttk', '<1>', itemClicked)
        #======================================================================        
        self.Score_Table = ttk.Treeview(self.Score_frame,height="27")

        self.Score_Table["columns"]=("column1","column2",'column3','column4','column5','column6','column7','column8','column9')
        self.Score_Table.column("#0",width=40, anchor='e')
        self.Score_Table.column("column1", width=200, anchor='w' )
        self.Score_Table.column("column2", width=380, anchor='w')
        self.Score_Table.column("column3", width=80 , anchor='center')
        self.Score_Table.column("column4", width=80 , anchor='center')
        self.Score_Table.column("column5", width=80 , anchor='center')
        self.Score_Table.column("column6", width=80 , anchor='center')
        self.Score_Table.column("column7", width=80 , anchor='center')
        self.Score_Table.column("column8", width=80 , anchor='center')
        self.Score_Table.column("column9", width=80 , anchor='center')
        
        self.Score_Table.heading('#0', text='#',command= lambda : self.treeview_sort_column(self.Score_Table, "column1", False))
        self.Score_Table.heading("column1", text="Bloomberg Ticker",command= lambda : self.treeview_string_sort_column(self.Score_Table, "column1", False))
        self.Score_Table.heading("column2", text="Name",command= lambda : self.treeview_string_sort_column(self.Score_Table, "column2", False))
        self.Score_Table.heading("column3", text="Current",command= lambda : self.treeview_string_sort_column(self.Score_Table, "column3", False))
        self.Score_Table.heading("column4", text="5 Days",command= lambda : self.treeview_string_sort_column(self.Score_Table, "column4", False))
        self.Score_Table.heading("column5", text="10 Days",command= lambda : self.treeview_string_sort_column(self.Score_Table, "column5", False))
        self.Score_Table.heading("column6", text="20 Days",command= lambda : self.treeview_string_sort_column(self.Score_Table, "column6", False))
        self.Score_Table.heading("column7", text="30 Days",command= lambda : self.treeview_string_sort_column(self.Score_Table, "column7", False))
        self.Score_Table.heading("column8", text="40 Days",command= lambda : self.treeview_string_sort_column(self.Score_Table, "column8", False))
        self.Score_Table.heading("column9", text="50 Days",command= lambda : self.treeview_sort_column(self.Score_Table, "column9", False))
        
        self.Score_Table.place(x=10, y=30)
        self.Score_Table.bind("<Double-1>", self.Index_Performance)
        self.Fetch_Signal()
        
        #======================================================================        
        self.Parameter_Table = ttk.Treeview(self.Parameter_frame,height="27")

        self.Parameter_Table["columns"]=("column1","column2",'column3','column4','column5','column6','column7','column8')
        self.Parameter_Table.column("#0",width=40, anchor='e')
        self.Parameter_Table.column("column1", width=200, anchor='w' )
        self.Parameter_Table.column("column2", width=380, anchor='w')
        self.Parameter_Table.column("column3", width=80 , anchor='center')
        self.Parameter_Table.column("column4", width=80 , anchor='center')
        self.Parameter_Table.column("column5", width=80 , anchor='center')
        self.Parameter_Table.column("column6", width=80 , anchor='center')
        self.Parameter_Table.column("column7", width=80 , anchor='center')
        self.Parameter_Table.column("column8", width=80 , anchor='center')
        
        
        self.Parameter_Table.heading('#0', text='#',command= lambda : self.treeview_sort_column(self.Parameter_Table, "column1", False))
        self.Parameter_Table.heading("column1", text="Bloomberg Ticker",command= lambda : self.treeview_string_sort_column(self.Parameter_Table, "column1", False))
        self.Parameter_Table.heading("column2", text="Name",command= lambda : self.treeview_string_sort_column(self.Parameter_Table, "column2", False))
        self.Parameter_Table.heading("column3", text="RSI")
        self.Parameter_Table.heading("column4", text="KD")
        self.Parameter_Table.heading("column5", text="EMA")
        self.Parameter_Table.heading("column6", text="Break")
        self.Parameter_Table.heading("column7", text="MTM")
        self.Parameter_Table.heading("column8", text="ATR")
        
        self.Parameter_Table.place(x=10, y=30)
    
        self.Fetch_Strategy_Parameters()

        self.Save_Button = ttk.Button(self.Parameter_frame,text='  Reload  ',command = self.Reload)
        self.Save_Button.place(x=1120, y=570)

        self.Save_Score_to_excel_Button = ttk.Button(self.Score_frame,text='Save as Excel file',command=self.Output_Score_to_excel)
        self.Save_Score_to_excel_Button.place(x=1090,y=2)

        self.Parameter_Table.bind("<Double-1>", self.Double_Click_Revise)
        
        
        
        
        
        self.root.mainloop()
        
        
    def dragwin(self,event):
        x = self.root.winfo_pointerx() - self.root._offsetx
        y = self.root.winfo_pointery() - self.root._offsety
        self.root.geometry('+{x}+{y}'.format(x=x,y=y))

    def clickwin(self,event):
        self.root._offsetx = event.x
        self.root._offsety = event.y


    def treeview_sort_column(self,tv, col, reverse):
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        l.sort(key=lambda t: float(t[0]), reverse=reverse)
        #      ^^^^^^^^^^^^^^^^^^^^^^^
    
        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)
    
        tv.heading(col,
                  command=lambda :self.treeview_sort_column(tv, col, not reverse))
    
    def treeview_string_sort_column(self,tv, col, reverse):
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        l.sort(key=lambda t: str(t[0]), reverse=reverse)
        #      ^^^^^^^^^^^^^^^^^^^^^^^
    
        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)
    
        tv.heading(col,
                  command=lambda :self.treeview_string_sort_column(tv, col, not reverse))






    def Fetch_Strategy_Parameters(self):
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
        cursor = connection.cursor()
        Database_Functions.Fetch_All_Data()
        cursor.execute("SELECT * FROM Strategy_table")
        
        data = cursor.fetchall()

        self.Parameter_Table.delete(*self.Parameter_Table.get_children())    
        
        i = 0
        for row in data:
            i = i +1
            #print i
            self.Parameter_Table.insert("",i,text=str(i),values=(row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7]))
        
    def Reload(self):
        self.Strategy_Table.delete(*self.Strategy_Table.get_children())    
        self.Fetch_Signal()
        self.Done_Messenger('Done !!')


        
    def Done_Messenger(self,info):
        Message_root = tk.Tk()
        Message_root.title('Info!')
        Message_root.geometry('300x200')
        Message_root.iconbitmap('D:/Taishin_Platform/pics/Data_Manager.ico')
        Message_Label = tk.Label(Message_root,text=info,font=('Arial',12,'bold'),fg='black',background="#F0F0F0")
        Message_Label.place(x=40,y=65)
        
        
        OK_Button=ttk.Button(Message_root,width=15, text =u"  OK  ",command = Message_root.destroy)
        OK_Button.place(x=130, y=120)
      
        
        Message_root.config(background="#F0F0F0")
        tk.mainloop()
    
    
    
    
    
    def Fetch_Signal(self):
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
        cursor = connection.cursor()
        Database_Functions.Fetch_All_Data()
        cursor.execute("SELECT * FROM Strategy_table")
        
        data = cursor.fetchall()

        self.Strategy_Table.delete(*self.Strategy_Table.get_children())    
        #workbook = xlsxwriter.Workbook('Spot Score.xlsx')
        #worksheet = workbook.add_worksheet()  
        global Historical_Score_df
        Historical_Score_df = pd.DataFrame()

        global All_Historical_Score_df
        All_Historical_Score_df = pd.DataFrame()


        i = 0
        for row in data:
            Ticker = row[0]
            Data = Database_Functions.Fetch(Ticker = Ticker).dropna()
            #print Data
            Single_df = pd.DataFrame()

            Score = 0
            i = i +1
            RSI_paras = row[2]
            try:
                RSI_para1 = float(RSI_paras.split(',')[0])
                RSI_para2 = float(RSI_paras.split(',')[1])
                #RSI_Signal_function
                #RSI_Signal = 'Long'
                RSI_info = self.RSI_Signal(Data,Para1=RSI_para1,Para2=RSI_para2)
                RSI_Signal = RSI_info[0]
                RSI_Win = RSI_info[1]
                RSI_Expect = RSI_info[2]
                if RSI_Signal == 'Long':
                    Score = Score + 1
                
                RSI_5 = RSI_info[3]
                RSI_10 = RSI_info[4]
                RSI_20 = RSI_info[5]
                RSI_30 = RSI_info[6]
                RSI_40 = RSI_info[7]
                RSI_50 = RSI_info[8]
                Single_df['RSI'] = RSI_info[9]
            except:
                RSI_Signal = ''
                RSI_Win = np.nan
                RSI_Expect = np.nan
                RSI_5 = 0
                RSI_10 = 0
                RSI_20 = 0
                RSI_30 = 0
                RSI_40 = 0
                RSI_50 = 0


            
            KD_paras = row[3]
            try:
                KD_para1 = float(KD_paras.split(',')[0])
                KD_para2 = float(KD_paras.split(',')[1])
                #KD_Signal_function 
                #KD_Signal = 'Long'
                KD_info = self.KD_Signal(Data,Para1=KD_para1,Para2=KD_para2)
                KD_Signal = KD_info[0]
                KD_Win = KD_info[1]
                KD_Expect = KD_info[2]
                if KD_Signal == 'Long':
                    Score = Score + 1
                KD_5 = KD_info[3]

                KD_10 = KD_info[4]
                KD_20 = KD_info[5]
                KD_30 = KD_info[6]
                KD_40 = KD_info[7]
                KD_50 = KD_info[8]
                Single_df['KD'] = KD_info[9]
            except:
                KD_Signal = ''
                KD_Win = np.nan
                KD_Expect = np.nan
                KD_5 = 0
                KD_10 = 0
                KD_20 = 0
                KD_30 = 0
                KD_40 = 0
                KD_50 = 0

            
            try:            
                EMA_paras = row[4]
                EMA_para1 = float(EMA_paras.split(',')[0])
                EMA_para2 = float(EMA_paras.split(',')[1])
                #EMA_Signal_function
                #EMA_Signal = 'Short'
                EMA_info = self.EMA_Signal(Data,Para1=EMA_para1,Para2=EMA_para2)
                EMA_Signal = EMA_info[0]
                EMA_Win = EMA_info[1]
                EMA_Expect = EMA_info[2]
                if EMA_Signal == 'Long':
                    Score = Score + 1
                EMA_5 = EMA_info[3]
                EMA_10 = EMA_info[4]
                EMA_20 = EMA_info[5]
                EMA_30 = EMA_info[6]
                EMA_40 = EMA_info[7]
                EMA_50 = EMA_info[8]
                Single_df['EMA'] = EMA_info[9]
            except:
                EMA_Signal = ''
                EMA_Win = np.nan
                EMA_Expect = np.nan
                EMA_5 = 0
                EMA_10 = 0
                EMA_20 = 0
                EMA_30 = 0
                EMA_40 = 0
                EMA_50 = 0

            try:
            #if row[5] != 'None':
                Break_paras = row[5]
                #print Break_paras
                Break_para1 = float(Break_paras.split(',')[0])
                Break_para2 = float(Break_paras.split(',')[1])
                #print Break_para1,Break_para2
                #Break_Signal_function
                #Break_Signal = 'Long'
                #print Break_para1,Break_para2
                Break_info = self.Break_Signal(Data,Para1=int(Break_para1),Para2=int(Break_para2))
                Break_Signal = Break_info[0]
                Break_Win = Break_info[1]
                Break_Expect = Break_info[2]
                #Break_Signal = ''
                if Break_Signal == 'Long':
                    Score = Score + 1
                Break_5 = Break_info[3]
                Break_10 = Break_info[4]
                Break_20 = Break_info[5]
                Break_30 = Break_info[6]
                Break_40 = Break_info[7]
                Break_50 = Break_info[8]
                Single_df['Break'] = Break_info[9]
            except:
                Break_Signal = ''
                Break_Win = np.nan
                Break_Expect = np.nan
                Break_5 = 0
                Break_10 = 0
                Break_20 = 0
                Break_30 = 0
                Break_40 = 0
                Break_50 = 0

            
            try:
                MTM_paras = row[6]
                MTM_para1 = float(MTM_paras)
                
                #MTM_Signal_function
                #print MTM_para1
                #MTM_Signal = 'Long'
                MTM_info = self.MTM_Signal(Data,Para1=MTM_para1)
                MTM_Signal = MTM_info[0]
                MTM_Win = MTM_info[1]
                MTM_Expect = MTM_info[2]
                if MTM_Signal == 'Long':
                    Score = Score + 1
                MTM_5 = MTM_info[3]
                MTM_10 = MTM_info[4]
                MTM_20 = MTM_info[5]
                MTM_30 = MTM_info[6]
                MTM_40 = MTM_info[7]
                MTM_50 = MTM_info[8]
                Single_df['MTM'] = MTM_info[9]
            except:
                MTM_Signal = ''
                MTM_Win = np.nan
                MTM_Expect = np.nan
                MTM_5 = 0
                MTM_10 = 0
                MTM_20 = 0
                MTM_30 = 0
                MTM_40 = 0
                MTM_50 = 0

            try:
                ATR_paras = row[7]
                ATR_para1 = float(ATR_paras.split(',')[0])
                ATR_para2 = float(ATR_paras.split(',')[1])
                #ATR_Signal_function
                #ATR_Signal = 'Long'
                ATR_info = self.ATR_Signal(Data,Para1=ATR_para1,Para2=ATR_para2)
                ATR_Signal = ATR_info[0]
                ATR_Win = ATR_info[1]
                ATR_Expect = ATR_info[2]
                
                if ATR_Signal == 'Long':
                    Score = Score + 1
                ATR_5 = ATR_info[3]
                ATR_10 = ATR_info[4]
                ATR_20 = ATR_info[5]
                ATR_30 = ATR_info[6]
                ATR_40 = ATR_info[7]
                ATR_50 = ATR_info[8]
                Single_df['ATR'] = ATR_info[9]
            except:
                ATR_Signal = ''
                ATR_Win = np.nan
                ATR_Expect = np.nan
                ATR_5 = 0
                ATR_10 = 0
                ATR_20 = 0
                ATR_30 = 0
                ATR_40 = 0
                ATR_50 = 0
            
            Win_list = [RSI_Win,KD_Win,EMA_Win,Break_Win,MTM_Win,ATR_Win]
            Expect_list = [RSI_Expect,KD_Expect,EMA_Expect,Break_Expect,MTM_Expect,ATR_Expect]
            Signal_list = [RSI_Signal,KD_Signal,EMA_Signal,Break_Signal,MTM_Signal,ATR_Signal]
            
            Signal_process_df = pd.DataFrame(index=['RSI','KD','EMA','Break','MTM','ATR'],columns=['Signal','Win','Win rank','Expect','Expect rank'])
            Signal_process_df['Signal'] = Signal_list
            Signal_process_df['Win'] = Win_list
            Signal_process_df['Win rank'] = Signal_process_df['Win'].rank(ascending=False)
            Signal_process_df['Expect'] = Expect_list
            Signal_process_df['Expect rank'] = Signal_process_df['Expect'].rank(ascending=False)
            Signal_process_df['Signal'] = np.where(Signal_process_df['Win rank'] == 1,Signal_process_df['Signal']+'*',Signal_process_df['Signal'])
            Signal_process_df['Signal'] = np.where(Signal_process_df['Expect rank'] == 1,Signal_process_df['Signal']+'**',Signal_process_df['Signal'])
            #print Signal_process_df
            
            self.Strategy_Table.insert("",i,text=str(i),values=(row[0],row[1],Signal_process_df['Signal'][0],Signal_process_df['Signal'][1],Signal_process_df['Signal'][2],Signal_process_df['Signal'][3],Signal_process_df['Signal'][4],Signal_process_df['Signal'][5],Score))
            self.Score_Table.insert("",i,text=str(i),values=(row[0],row[1],Score,int(sum([RSI_5,KD_5,EMA_5,Break_5,MTM_5,ATR_5])),int(sum([RSI_10,KD_10,EMA_10,Break_10,MTM_10,ATR_10])),int(sum([RSI_20,KD_20,EMA_20,Break_20,MTM_20,ATR_20])),int(sum([RSI_30,KD_30,EMA_30,Break_30,MTM_30,ATR_30])),int(sum([RSI_40,KD_40,EMA_40,Break_40,MTM_40,ATR_40])),int(sum([RSI_50,KD_50,EMA_50,Break_50,MTM_50,ATR_50]))))
            #self.Score_Table.insert("",i,text=str(i),values=(row[0],row[1],Score,1,1,1,1,1,1,1))
            #List = [row[0],row[1],Signal_process_df['Signal'][0],Signal_process_df['Signal'][1],Signal_process_df['Signal'][2],Signal_process_df['Signal'][3],Signal_process_df['Signal'][4],Signal_process_df['Signal'][5],Score]
            #Historical_Score_df[Ticker] = Historical_Score_df[Ticker].dropna()
            #print Single_df.dropna().sum(axis=1)
            Historical_Score_df[Ticker] = Single_df.dropna().sum(axis=1)

            #All_Historical_Score_df[Ticker] = Historical_Score_df[Ticker].resample('M',how='last')

        #Historical_Score_df = Historical_Score_df.resample('W-WED')
        #print Historical_Score_df.sort_index(ascending=False).T
        global Weekly_Score
        Weekly_Score = Historical_Score_df.resample('W-FRI',how='last').sort_index(ascending=False).T

        global Daily_Score
        Daily_Score = Historical_Score_df.sort_index(ascending=False).T


        Historical_Score_df = Historical_Score_df.resample('W-WED',how='last')
        #print Historical_Score_df
        #print All_Historical_Score_df
        
        #writer = pd.ExcelWriter('Historical Scores.xlsx')
        
        #All_Historical_Score_df.to_excel(writer)
        
        #writer.save()

    
    
    
    
    
    
    
    
    
    
            #worksheet.write(i, 0,List[0])
            #worksheet.write(i, 1,List[1])
            #worksheet.write(i, 2,List[2])
            #worksheet.write(i, 3,List[3])
            #worksheet.write(i, 4,List[4])
            #worksheet.write(i, 5,List[5])
            #worksheet.write(i, 6,List[6])
            #worksheet.write(i, 7,List[7])
            #worksheet.write(i, 8,List[8])
        #workbook.close()
        
    def Output_Score_to_excel(self):
        #print Weekly_Score
        #print Daily_Score

        writer = pd.ExcelWriter('D:/Taishin_Platform/Market_Timing_Score/Score.xlsx')
        Daily_Score.to_excel(writer, 'Daily Score')
        Weekly_Score.to_excel(writer, 'Weekly Score')

        writer.save()
        self.Done_Messenger(info='Done')

    def Output_xls(self):
        def Week_Return():
            df = pd.DataFrame()
            
            for Ticker in Output_table['Ticker']:
                df[Ticker] = Database_Functions.Fetch(Ticker)[Ticker].dropna()    


            df = df.resample('W-WED',how='last').pct_change()
            #print df['MXWD Index']
            return df.ix[-1].tolist(),str(df.index[-1].strftime('%Y-%m-%d'))
            
            
            
            
            
        Strategy_list = self.Strategy_Table.get_children()
        Output_table = pd.DataFrame(columns = ['Ticker','Name','Score'])
        Output_tickers = []
        Output_names = []
        Output_scores = []
        for i in Strategy_list:
            Output_tickers.append(self.Strategy_Table.item(i)['values'][0])
            Output_names.append(self.Strategy_Table.item(i)['values'][1])
            Output_scores.append(self.Strategy_Table.item(i)['values'][8])
            
        Output_table['Ticker'] = Output_tickers
        Output_table['Name'] = Output_names
        Weekly_info = Week_Return()        
        
        
        Output_table['Score'] = Output_scores
        Output_table.index = Output_table.index+1
        #print Output_table

        Last_signal = str(Historical_Score_df.index[-2].strftime('%Y-%m-%d'))
        #print Last_signal
        Output_table['Score at '+str(Last_signal)] = Historical_Score_df.ix[-2].tolist()
        Output_table['One Week Return after '+str(Last_signal)] = Weekly_info[0]
        
        #print Output_table
        
        Strong_market = Output_table[Output_table.Score>=5].sort('Score',ascending=0)
        Weak_market = Output_table[Output_table.Score<=2].sort('Score',ascending=1)
        
        Strategy_list = self.Score_Table.get_children()
        Score_df = pd.DataFrame(columns = ['Ticker','Name','Score'])
        Score_tickers = []
        Score_names = []
        Score_current = []
        Score_before = []
        
        for i in Strategy_list:
            Score_tickers.append(self.Score_Table.item(i)['values'][0])
            Score_names.append(self.Score_Table.item(i)['values'][1])
            Score_current.append(self.Score_Table.item(i)['values'][2])
            Score_before.append(self.Score_Table.item(i)['values'][3])
            
        Score_df['Ticker'] = Score_tickers
        Score_df['Name'] = Score_names
        Score_df['Current Score'] = Score_current
        Score_df['Before Score'] = Historical_Score_df.ix[-2].tolist()
        Score_df['Score_change'] = Score_df['Current Score'] - Score_df['Before Score']
        
        #print Historical_Score_df
        
        Score_df.index = Score_df.index+1
        Weak_to_strong = Score_df[['Ticker','Name','Score_change']][Score_df.Score_change>0]
        Stong_to_weak = Score_df[['Ticker','Name','Score_change']][Score_df.Score_change<0]
        
        
        Today = datetime.datetime.today().strftime('%Y-%m-%d')
        
        
        writer = pd.ExcelWriter('D:/Taishin_Platform/Market_Timing_Score/'+str(Historical_Score_df.index[-2].strftime('%Y-%m-%d'))+' Score.xlsx')
        Strong_market.to_excel(writer,'Strong Market')
        Weak_market.to_excel(writer,'Weak Market')
        Weak_to_strong.to_excel(writer,'Weak to Strong')
        Stong_to_weak.to_excel(writer,'Strong to Weak')
        Output_table.to_excel(writer,'All')
        
        
        writer.save()
        self.Done_Messenger(info='Done')
        
        
        
        
        

    def RSI_Signal(self,Data,Para1,Para2):
        def RSI(prices,n):
            n = int(n)
            deltas = np.diff(prices)
            seed = deltas[:n+1]
            up = seed[seed>=0].sum()/n
            down = -seed[seed<0].sum()/n
            rs = up/down
            rsi = np.zeros_like(prices)
            rsi[:n]=100. - 100./(1.+rs)
                
            for i in range(n,len(prices)):
                delta = deltas[i-1]    #cause the diff is 1 shorter
                    
                if delta>0:
                    upval = delta
                    downval = 0.
                else:
                    upval = 0.
                    downval = -delta
                        
                up = (up*(n-1) + upval)/n
                down = (down*(n-1) + downval)/n
                    
                rs = up/down
                rsi[i] = 100. - 100./(1.+rs)
                
            return rsi
        
        Data['RSI1'] = RSI(Data[Data.columns[0]],Para1)
        Data['RSI2'] = RSI(Data[Data.columns[0]],Para2)

        Data['RSI_Signal'] = np.where(Data['RSI1']>Data['RSI2'],1,0)
        
        Current_Signal = Data['RSI_Signal'][-1]

        if Current_Signal == 1:
            Current_Signal = 'Long'
        else:
            Current_Signal = 'Cash'
            
        Data['RSI_Buy_Price'] = np.where((Data['RSI_Signal'] == 1)&(Data['RSI_Signal'].shift(1) == 0),Data[Data.columns[0]],np.nan)
        Data['RSI_Sell_Price'] = np.where((Data['RSI_Signal'] == 0)&(Data['RSI_Signal'].shift(1) == 1),Data[Data.columns[0]],np.nan)
        
        Trading_book = pd.concat([Data['RSI_Buy_Price'].dropna(),Data['RSI_Sell_Price'].dropna()],axis=1)
        Trading_book['RSI_Sell_Price'] = Trading_book['RSI_Sell_Price'].shift(-1)
        Trading_book['RSI_Profit'] = np.log(Trading_book['RSI_Sell_Price']/Trading_book['RSI_Buy_Price'])
        
        
        Sign = np.where(Trading_book['RSI_Profit'].dropna()>0.,1.,0.)
        
        Win_Percentage = round(sum(Sign)/len(Sign)*100.,2)
        Expect =  np.mean(Trading_book['RSI_Profit'])    

            
        return Current_Signal,Win_Percentage,Expect,Data['RSI_Signal'][-5],Data['RSI_Signal'][-10],Data['RSI_Signal'][-20],Data['RSI_Signal'][-30],Data['RSI_Signal'][-40],Data['RSI_Signal'][-50],Data['RSI_Signal']
    
    def KD_Signal(self,Data,Para1,Para2):
        def RSV(price,n):
            RSV = (price - pd.rolling_min(price,int(n)))/(pd.rolling_max(price,int(n))-pd.rolling_min(price,int(n)))
            return RSV*100.
        
        def K(Series,n):
            Series = RSV(Series,n)
            Series = pd.ewma(Series,5)
            return Series
        
        def D(Series,n):
            Series = pd.ewma(Series,n)
            return Series
        
        Data['K'] = K(Data[Data.columns[0]],Para1)
        Data['D'] = D(Data['K'],Para2)
        
        Data['KD_Signal'] = np.where(Data['K']>Data['D'],1,0)
        
        Current_Signal = Data['KD_Signal'][-1]

        if Current_Signal == 1:
            Current_Signal = 'Long'
        else:
            Current_Signal = 'Cash'
        
        
        Data['KD_Buy_Price'] = np.where((Data['KD_Signal'] == 1)&(Data['KD_Signal'].shift(1) == 0),Data[Data.columns[0]],np.nan)
        Data['KD_Sell_Price'] = np.where((Data['KD_Signal'] == 0)&(Data['KD_Signal'].shift(1) == 1),Data[Data.columns[0]],np.nan)
        
        Trading_book = pd.concat([Data['KD_Buy_Price'].dropna(),Data['KD_Sell_Price'].dropna()],axis=1)
        Trading_book['KD_Sell_Price'] = Trading_book['KD_Sell_Price'].shift(-1)
        Trading_book['KD_Profit'] = np.log(Trading_book['KD_Sell_Price']/Trading_book['KD_Buy_Price'])
        
        
        Sign = np.where(Trading_book['KD_Profit'].dropna()>0.,1.,0.)
        
        Win_Percentage = round(sum(Sign)/len(Sign)*100.,2)
        Expect =  np.mean(Trading_book['KD_Profit'])    
            
        
        
        
        
        return Current_Signal,Win_Percentage,Expect,Data['KD_Signal'][-5],Data['KD_Signal'][-10],Data['KD_Signal'][-20],Data['KD_Signal'][-30],Data['KD_Signal'][-40],Data['KD_Signal'][-50],Data['KD_Signal']
    
    def EMA_Signal(self,Data,Para1,Para2):
        Data['EMA1'] = pd.ewma(Data[Data.columns[0]],Para1)
        Data['EMA2'] = pd.ewma(Data[Data.columns[0]],Para2)
        Data['EMA_Signal'] = np.where(Data['EMA1']>Data['EMA2'],1,0)
        
        Current_Signal = Data['EMA_Signal'][-1]
        
        if Current_Signal == 1:
            Current_Signal = 'Long'
        else:
            Current_Signal = 'Cash'
        
        Data['EMA_Buy_Price'] = np.where((Data['EMA_Signal'] == 1)&(Data['EMA_Signal'].shift(1) == 0),Data[Data.columns[0]],np.nan)
        Data['EMA_Sell_Price'] = np.where((Data['EMA_Signal'] == 0)&(Data['EMA_Signal'].shift(1) == 1),Data[Data.columns[0]],np.nan)
        
        Trading_book = pd.concat([Data['EMA_Buy_Price'].dropna(),Data['EMA_Sell_Price'].dropna()],axis=1)
        Trading_book['EMA_Sell_Price'] = Trading_book['EMA_Sell_Price'].shift(-1)
        Trading_book['EMA_Profit'] = np.log(Trading_book['EMA_Sell_Price']/Trading_book['EMA_Buy_Price'])
        
        
        Sign = np.where(Trading_book['EMA_Profit'].dropna()>0.,1.,0.)
        
        Win_Percentage = round(sum(Sign)/len(Sign)*100.,2)
        Expect =  np.mean(Trading_book['EMA_Profit'])    
            
        
        
        
        
        return Current_Signal,Win_Percentage,Expect,Data['EMA_Signal'][-5],Data['EMA_Signal'][-10],Data['EMA_Signal'][-20],Data['EMA_Signal'][-30],Data['EMA_Signal'][-40],Data['EMA_Signal'][-50],Data['EMA_Signal']
    
    def Break_Signal(self,Data,Para1,Para2):

        Data['Return'] = Data[Data.columns[0]].pct_change()
        
        Data['High'] = pd.rolling_max(Data[Data.columns[0]].shift(1),int(Para1))
        Data['Low'] = pd.rolling_min(Data[Data.columns[0]].shift(1),int(Para2))
        Data['Break_Signal_temp'] = np.where(Data[Data.columns[0]]>Data['High'],1,np.where(Data[Data.columns[0]]<Data['Low'],0,np.nan))
        #print Data['Break_Signal_temp']
        Data['Break_Signal'] = Data['Break_Signal_temp'].fillna(method='ffill')

        Current_Signal = Data['Break_Signal'][-1]
        #print Current_Signal
        if Current_Signal == 1:
            Current_Signal = 'Long'
        else:
            Current_Signal = 'Cash'
        
        Data['Break_Buy_Price'] = np.where((Data['Break_Signal'] == 1)&(Data['Break_Signal'].shift(1) == 0),Data[Data.columns[0]],np.nan)
        Data['Break_Sell_Price'] = np.where((Data['Break_Signal'] == 0)&(Data['Break_Signal'].shift(1) == 1),Data[Data.columns[0]],np.nan)

        Trading_book = pd.concat([Data['Break_Buy_Price'].dropna(),Data['Break_Sell_Price'].dropna()],axis=1)
        #print Trading_book
        Trading_book['Break_Sell_Price'] = Trading_book['Break_Sell_Price'].shift(-1)
        Trading_book['Break_Profit'] = np.log(Trading_book['Break_Sell_Price']/Trading_book['Break_Buy_Price'])

        
        Sign = np.where(Trading_book['Break_Profit'].dropna()>0.,1.,0.)
        Win_Percentage = round(sum(Sign)/len(Sign)*100.,2)
        Expect =  np.mean(Trading_book['Break_Profit'])    
            
        
        
        
        return Current_Signal,Win_Percentage,Expect,Data['Break_Signal'][-5],Data['Break_Signal'][-10],Data['EMA_Signal'][-20],Data['Break_Signal'][-30],Data['Break_Signal'][-40],Data['Break_Signal'][-50],Data['Break_Signal']
    
    def MTM_Signal(self,Data,Para1):
        Data['MTM'] =Data[Data.columns[0]].pct_change(Para1)
        Data['MTM_Signal'] = np.where(Data['MTM']>0.,1,0)
        Current_Signal = Data['MTM_Signal'][-1]
        
        if Current_Signal == 1:
            Current_Signal = 'Long'
        else:
            Current_Signal = 'Cash'
        
        Data['MTM_Buy_Price'] = np.where((Data['MTM_Signal'] == 1)&(Data['MTM_Signal'].shift(1) == 0),Data[Data.columns[0]],np.nan)
        Data['MTM_Sell_Price'] = np.where((Data['MTM_Signal'] == 0)&(Data['MTM_Signal'].shift(1) == 1),Data[Data.columns[0]],np.nan)
        
        Trading_book = pd.concat([Data['MTM_Buy_Price'].dropna(),Data['MTM_Sell_Price'].dropna()],axis=1)
        Trading_book['MTM_Sell_Price'] = Trading_book['MTM_Sell_Price'].shift(-1)
        Trading_book['MTM_Profit'] = np.log(Trading_book['MTM_Sell_Price']/Trading_book['MTM_Buy_Price'])
        

        Sign = np.where(Trading_book['MTM_Profit'].dropna()>0.,1.,0.)
        
        Win_Percentage = round(sum(Sign)/len(Sign)*100.,2)
        Expect =  np.mean(Trading_book['MTM_Profit'])    
            
        
        
        return Current_Signal,Win_Percentage,Expect,Data['MTM_Signal'][-5],Data['MTM_Signal'][-10],Data['MTM_Signal'][-20],Data['MTM_Signal'][-30],Data['MTM_Signal'][-40],Data['MTM_Signal'][-50],Data['MTM_Signal']
    
    def ATR_Signal(self,Data,Para1,Para2):
        Data['EMA'] = pd.ewma(Data[Data.columns[0]],Para1)
        Data['ATR'] = pd.ewma(abs(Data[Data.columns[0]].diff()),Para1)
    
        Data['UpBound'] = Data['EMA']+Para2*Data['ATR']
        #Data['LowBound'] = Data['EMA']-Para2*Data['ATR']
        
        Data['ATR_Signal'] = np.where(Data[Data.columns[0]]>Data['UpBound'],1,0)
        Current_Signal = Data['ATR_Signal'][-1]
        
        if Current_Signal == 1:
            Current_Signal = 'Long'
        else:
            Current_Signal = 'Cash'
            
        Data['ATR_Buy_Price'] = np.where((Data['ATR_Signal'] == 1)&(Data['ATR_Signal'].shift(1) == 0),Data[Data.columns[0]],np.nan)
        Data['ATR_Sell_Price'] = np.where((Data['ATR_Signal'] == 0)&(Data['ATR_Signal'].shift(1) == 1),Data[Data.columns[0]],np.nan)
        
        Trading_book = pd.concat([Data['ATR_Buy_Price'].dropna(),Data['ATR_Sell_Price'].dropna()],axis=1)
        Trading_book['ATR_Sell_Price'] = Trading_book['ATR_Sell_Price'].shift(-1)
        Trading_book['ATR_Profit'] = np.log(Trading_book['ATR_Sell_Price']/Trading_book['ATR_Buy_Price'])
        
        
        Sign = np.where(Trading_book['ATR_Profit'].dropna()>0.,1.,0.)
        
        Win_Percentage = round(sum(Sign)/len(Sign)*100.,2)
        Expect =  np.mean(Trading_book['ATR_Profit'])    
            
        
        
        return Current_Signal,Win_Percentage,Expect,Data['ATR_Signal'][-5],Data['ATR_Signal'][-10],Data['ATR_Signal'][-20],Data['ATR_Signal'][-30],Data['ATR_Signal'][-40],Data['ATR_Signal'][-50],Data['ATR_Signal']
    
    
    
    def Double_Click_Revise(self,event):
        def RSI_Optimize():
            
            def RSI(prices,n):
                n = int(n)
                deltas = np.diff(prices)
                seed = deltas[:n+1]
                up = seed[seed>=0].sum()/n
                down = -seed[seed<0].sum()/n
                rs = up/down
                rsi = np.zeros_like(prices)
                rsi[:n]=100. - 100./(1.+rs)
                    
                for i in range(n,len(prices)):
                    delta = deltas[i-1]    #cause the diff is 1 shorter
                        
                    if delta>0:
                        upval = delta
                        downval = 0.
                    else:
                        upval = 0.
                        downval = -delta
                            
                    up = (up*(n-1) + upval)/n
                    down = (down*(n-1) + downval)/n
                        
                    rs = up/down
                    rsi[i] = 100. - 100./(1.+rs)
                    
                return rsi
            
            Data = Database_Functions.Fetch(Ticker).dropna()
            Data['Return'] = Data[str(Ticker)].pct_change()
            
            para_range = np.arange(0,200,5.)
            para_range[0] = 1.
            
            Turnover_limit = 12.
            Turnover_floor = 2.
            Max_Return = -100.
            for para1 in para_range:
                Data['EMA_1'] = RSI(Data[str(Ticker)],para1)
                for para2 in para_range:
                    Data['EMA_2'] = RSI(Data[str(Ticker)],para2)
                    
                    Data['diff'] = Data['EMA_1'] - Data['EMA_2']
                    Data['diff'] = np.where(Data['diff']==0,Data['diff'].shift(1)*(-1),Data['diff'])
                    
                    Data['Signal'] = np.where(Data['diff'] > 0.,1,0)
                    Data['Equity_Return'] = Data['Return']*Data['Signal'].shift(1)
                    Data['Equity_Curve'] = Data['Equity_Return'].cumsum().apply(np.exp)
                    #Data['Buy_Hold'] = Data['Return'].cumsum().apply(np.exp)
                    
                    Horizon = (Data.index[-1] - Data.index[0]).days/365.
                    Total_Turnover = abs(Data['Signal'].diff()).sum()
                    Annual_Turnover = Total_Turnover/Horizon
                    
                    Total_Return = Data['Equity_Curve'][-1]
                    print para1,para2,Total_Return,Annual_Turnover
                    if Total_Return != np.inf:
                       
                        if (Annual_Turnover < Turnover_limit) & (Annual_Turnover>Turnover_floor):
                            if Total_Return > Max_Return:

                                Max_Return = Total_Return
                                
                                Optimal_para1 = para1
                                Optimal_para2 = para2
                                print Optimal_para1,Optimal_para2
                    else:
                        pass
            
            RSI1_Entry.delete(0,'end')
            RSI2_Entry.delete(0,'end')
            
            RSI1_Entry.insert(0,int(Optimal_para1))
            RSI2_Entry.insert(0,int(Optimal_para2))
        
        def KD_Optimize():
            def RSV(price,n):
                RSV = (price - pd.rolling_min(price,int(n)))/(pd.rolling_max(price,int(n))-pd.rolling_min(price,int(n)))
                return RSV*100.
            
            def K(Series,n):
                Series = RSV(Series,n)
                Series = pd.ewma(Series,5)
                return Series
            
            def D(Series,n):
                Series = pd.ewma(Series,n)
                return Series
            
            Data = Database_Functions.Fetch(Ticker).dropna()
            Data['Return'] = Data[str(Ticker)].pct_change()
            
            para_range = np.arange(0,200,5.)
            para_range[0] = 1.
            
            Turnover_limit = 12.
            Turnover_floor = 2.
            Max_Return = -100.
            for para1 in para_range:
                Data['EMA_1'] = K(Data[str(Ticker)],para1)
                for para2 in para_range:
                    Data['EMA_2'] = D(Data['EMA_1'],para2)
                    
                    Data['diff'] = Data['EMA_1'] - Data['EMA_2']
                    Data['diff'] = np.where(Data['diff']==0,Data['diff'].shift(1)*(-1),Data['diff'])
                    
                    Data['Signal'] = np.where(Data['diff'] > 0.,1,0)
                    Data['Equity_Return'] = Data['Return']*Data['Signal'].shift(1)
                    Data['Equity_Curve'] = Data['Equity_Return'].cumsum().apply(np.exp)
                    #Data['Buy_Hold'] = Data['Return'].cumsum().apply(np.exp)
                    
                    Horizon = (Data.index[-1] - Data.index[0]).days/365.
                    Total_Turnover = abs(Data['Signal'].diff()).sum()
                    Annual_Turnover = Total_Turnover/Horizon
                    
                    Total_Return = Data['Equity_Curve'][-1]
                    print para1,para2,Total_Return,Annual_Turnover
                    if Total_Return != np.inf:
                       
                        if (Annual_Turnover < Turnover_limit) & (Annual_Turnover>Turnover_floor):
                            if Total_Return > Max_Return:

                                Max_Return = Total_Return
                                
                                Optimal_para1 = para1
                                Optimal_para2 = para2
                    else:
                        pass
            
            KD1_Entry.delete(0,'end')
            KD2_Entry.delete(0,'end')
            
            KD1_Entry.insert(0,int(Optimal_para1))
            KD2_Entry.insert(0,int(Optimal_para2))
        
        def EMA_Optimize():
            Data = Database_Functions.Fetch(Ticker).dropna()
            Data['Return'] = Data[str(Ticker)].pct_change()
            
            para_range = np.arange(0,200,5.)
            para_range[0] = 1.
            
            Turnover_limit = 12.
            Turnover_floor = 2.
            Max_Return = -100.
            for para1 in para_range:
                Data['EMA_1'] = pd.ewma(Data[str(Ticker)],para1)
                for para2 in para_range:
                    if para1 < para2:
                        Data['EMA_2'] = pd.ewma(Data[str(Ticker)],para2)
                        
                        Data['diff'] = Data['EMA_1'] - Data['EMA_2']
                        Data['diff'] = np.where(Data['diff']==0,Data['diff'].shift(1)*(-1),Data['diff'])
                        
                        Data['Signal'] = np.where(Data['diff'] > 0.,1,0)
                        Data['Equity_Return'] = Data['Return']*Data['Signal'].shift(1)
                        Data['Equity_Curve'] = Data['Equity_Return'].cumsum().apply(np.exp)
                        #Data['Buy_Hold'] = Data['Return'].cumsum().apply(np.exp)

                        Horizon = (Data.index[-1] - Data.index[0]).days/365.
                        Total_Turnover = abs(Data['Signal'].diff()).sum()
                        Annual_Turnover = Total_Turnover/Horizon
                        
                        Total_Return = Data['Equity_Curve'][-1]
                        #print para1,para2,Total_Return
                        if Total_Return != np.inf:
                           
                            if (Annual_Turnover < Turnover_limit) & (Annual_Turnover>Turnover_floor):
                                if Total_Return > Max_Return:

                                    Max_Return = Total_Return
                                    
                                    Optimal_para1 = para1
                                    Optimal_para2 = para2
                        else:
                            pass
            
            EMA1_Entry.delete(0,'end')
            EMA2_Entry.delete(0,'end')
            
            EMA1_Entry.insert(0,int(Optimal_para1))
            EMA2_Entry.insert(0,int(Optimal_para2))
        def Break_Optimize():
            def Fill_data2(Series):
                Number = len(Series)
                Current_data = 0    
                
                for i in range(Number):
                    if pd.isnull(Series[i]):
                        Series[i] = Current_data
                    else:
                        Current_data = Series[i]
                
                return Series

            
            Data = Database_Functions.Fetch(Ticker).dropna()
            Data['Return'] = Data[str(Ticker)].pct_change()
            
            para_range = np.arange(0,30,5.)
            para_range[0] = 1.
            
            Turnover_limit = 12.
            Turnover_floor = 2.
            Max_Return = -100.
            for para1 in para_range:
                Data['High'] = pd.rolling_max(Data[str(Ticker)],int(para1)).shift(1)
                for para2 in para_range:
                    Data['Low'] = pd.rolling_min(Data[str(Ticker)],int(para2)).shift(1)
                    
                    Data['Signal'] = np.where(Data[str(Ticker)]>Data['High'],1,np.where(Data[str(Ticker)]<Data['Low'],0,np.nan))
                    Data['Signal'] = Data['Signal'].fillna(method='ffill')
                    Data['Equity_Return'] = Data['Return']*Data['Signal'].shift(1)
                    Data['Equity_Curve'] = Data['Equity_Return'].cumsum().apply(np.exp)
                    #Data['Buy_Hold'] = Data['Return'].cumsum().apply(np.exp)
                    
                    Horizon = (Data.index[-1] - Data.index[0]).days/365.
                    Total_Turnover = abs(Data['Signal'].diff()).sum()
                    Annual_Turnover = Total_Turnover/Horizon
                    
                    Total_Return = Data['Equity_Curve'][-1]
                    print para1,para2,Total_Return,Annual_Turnover
                    if Total_Return != np.inf:
                       
                        if (Annual_Turnover < Turnover_limit) & (Annual_Turnover>Turnover_floor):
                            if Total_Return > Max_Return:

                                Max_Return = Total_Return
                                
                                Optimal_para1 = para1
                                Optimal_para2 = para2
                    else:
                        pass
            
            Break1_Entry.delete(0,'end')
            Break2_Entry.delete(0,'end')
            
            Break1_Entry.insert(0,int(Optimal_para1))
            Break2_Entry.insert(0,int(Optimal_para2))
        
        def MTM_Optimize():
            Data = Database_Functions.Fetch(Ticker).dropna()
            Data['Return'] = Data[str(Ticker)].pct_change()
            
            para_range = np.arange(0,200,5.)
            para_range[0] = 1.
            
            Turnover_limit = 12.
            Turnover_floor = 2.
            Max_Return = -100.
            for para1 in para_range:
                Data['MTM'] = pd.ewma(Data[str(Ticker)],para1)
                
                Data['diff'] = Data[Ticker] - Data['MTM']
                Data['diff'] = np.where(Data['diff']==0,Data['diff'].shift(1)*(-1),Data['diff'])
                
                Data['Signal'] = np.where(Data['diff'] > 0.,1,0)
                Data['Equity_Return'] = Data['Return']*Data['Signal'].shift(1)
                Data['Equity_Curve'] = Data['Equity_Return'].cumsum().apply(np.exp)
                #Data['Buy_Hold'] = Data['Return'].cumsum().apply(np.exp)
                
                Horizon = (Data.index[-1] - Data.index[0]).days/365.
                Total_Turnover = abs(Data['Signal'].diff()).sum()
                Annual_Turnover = Total_Turnover/Horizon
                
                Total_Return = Data['Equity_Curve'][-1]
                print para1,Total_Return
                if Total_Return != np.inf:
                   
                    if (Annual_Turnover < Turnover_limit) & (Annual_Turnover>Turnover_floor):
                        if Total_Return > Max_Return:
                            print Max_Return,Total_Return
                            Max_Return = Total_Return
                            Optimal_para1 = para1
                            print Optimal_para1
                else:
                    pass
            MTM1_Entry.delete(0,'end')
            MTM1_Entry.insert(0,int(Optimal_para1))
        
        def ATR_Optimize():
            Data = Database_Functions.Fetch(Ticker).dropna()
            Data['Return'] = Data[str(Ticker)].pct_change()
            
            para_range = np.arange(0,200,5.)
            para_range[0] = 1.
            para2_range = np.arange(0,5,0.1)
            
            Turnover_limit = 12.
            Turnover_floor = 2.
            Max_Return = -100.
            
            for para1 in para_range:
                Data['EMA'] = pd.ewma(Data[str(Ticker)],para1)
                Data['ATR'] = pd.ewma(abs(Data[Ticker].diff()),para1)
                for para2 in para2_range:
                    Data['UpBound'] = Data['EMA']+para2*Data['ATR']
                    Data['LowBound'] = Data['EMA']-para2*Data['ATR']
                    
                    Data['Signal'] = np.where(Data[Ticker]>Data['UpBound'],1,0)
                    Data['Equity_Return'] = Data['Return']*Data['Signal'].shift(1)
                    Data['Equity_Curve'] = Data['Equity_Return'].cumsum().apply(np.exp)
                    #Data['Buy_Hold'] = Data['Return'].cumsum().apply(np.exp)
                    
                    Horizon = (Data.index[-1] - Data.index[0]).days/365.
                    Total_Turnover = abs(Data['Signal'].diff()).sum()
                    Annual_Turnover = Total_Turnover/Horizon
                    
                    Total_Return = Data['Equity_Curve'][-1]
                    print para1,para2,Total_Return
                    
                    if Total_Return != np.inf:
                       
                        if (Annual_Turnover < Turnover_limit) & (Annual_Turnover>Turnover_floor):
                            if Total_Return > Max_Return:
                                Max_Return = Total_Return
                                Optimal_para1 = para1
                                Optimal_para2 = para2
                    else:
                        pass
            
            ATR1_Entry.delete(0,'end')
            ATR2_Entry.delete(0,'end')
            
            ATR1_Entry.insert(0,int(Optimal_para1))
            ATR2_Entry.insert(0,round(float(Optimal_para2),1))
        
        def Save_Parameters_to_db():
            RSI1 = str(RSI1_Entry.get())
            RSI2 = str(RSI2_Entry.get())
            
            KD1 = str(KD1_Entry.get())
            KD2 = str(KD2_Entry.get())
            
            EMA1 = str(EMA1_Entry.get())
            EMA2 = str(EMA2_Entry.get())
        
            Break1 = str(Break1_Entry.get())        
            Break2 = str(Break2_Entry.get())    
            
            MTM = str(MTM1_Entry.get())
        
            ATR1 = str(ATR1_Entry.get())
            ATR2 = str(ATR2_Entry.get())
            
        
            RSI_set =  RSI1+','+RSI2
            KD_set = KD1+','+KD2
            EMA_set = EMA1+','+EMA2
            Break_set = Break1+','+Break2
            MTM_set = MTM
            ATR_set = ATR1+','+ATR2
            
            if RSI_set == ',':
                RSI_set = ''
            if KD_set == ',':
                KD_set = ''
            if EMA_set == ',':
                EMA_set = ''
            if Break_set == ',':
                Break_set = ''
            if MTM_set == ',':
                MTM_set = ''
            if ATR_set == ',':
                ATR_set = ''
            
            connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
            cursor = connection.cursor()
            cursor.execute("INSERT OR REPLACE INTO Strategy_table VALUES (?, ?, ?, ?, ?, ?, ?, ?)",(Ticker, Name, RSI_set, KD_set,EMA_set,Break_set,MTM_set,ATR_set))
            connection.commit()
            connection.close()
            #print Ticker,Name,RSI_set,KD_set,EMA_set,Break_set,MTM_set,ATR_set
            self.Fetch_Strategy_Parameters()
            Parameter_root.destroy()
        
        curItem = self.Parameter_Table.focus()
        Index_info = self.Parameter_Table.item(curItem)['values']
        Ticker = Index_info[0]
        Name = Index_info[1]

        try:
            RSI_para1 = Index_info[2].split(',')[0]
            if RSI_para1 == 'None':
                RSI_para1 = ''
            
        except:
            RSI_para1 = ''
        
        try:
            RSI_para2 = Index_info[2].split(',')[1]
        except:
            RSI_para2 = ''
        
        try:
            KD_para1 = Index_info[3].split(',')[0]
            if KD_para1 == 'None':
                KD_para1 = ''

        except:
            KD_para1 = ''
        
        try:
            KD_para2 = Index_info[3].split(',')[1]

        except:
            KD_para2 = ''
        
        try:
            EMA_para1 = Index_info[4].split(',')[0]
            if EMA_para1 == 'None':
                EMA_para1 = ''

        except:
            EMA_para1 = ''


        try:
            EMA_para2 = Index_info[4].split(',')[1]
        except:
            EMA_para2 = ''

        try:
            Break_para1 = Index_info[5].split(',')[0]
            if Break_para1 == 'None':
                Break_para1 = ''
        except:
            Break_para1 = ''

        try:
            Break_para2 = Index_info[5].split(',')[1]
        except:
            Break_para2 = ''

        try:
            MTM_para1 = Index_info[6]
            if MTM_para1 == 'None':
                MTM_para1 = ''
        except:
            MTM_para1 = ''

        try:
            ATR_para1 = Index_info[7].split(',')[0]
            if ATR_para1 == 'None':
                ATR_para1 = ''
        except:
            ATR_para1 = ''

        try:
            ATR_para2 = Index_info[7].split(',')[1]
            #print ATR_para2
        except:
            ATR_para2 = ''




        
        
        
        #======================================================================
        Parameter_root = tk.Tk()
        Parameter_root.wm_title('Parameters Revise')
        Parameter_root.geometry('500x500')
        Parameter_root.iconbitmap('D:/Taishin_Platform/pics/Data_Manager.ico')
        
        Name_Label = tk.Label(Parameter_root,text=Ticker,font=('Arial',14,'bold'),fg='black',background="#F0F0F0")
        Name_Label.place(x=10,y=10)
       
        RSI_Label = tk.Label(Parameter_root,text='RSI',font=('Arial',12),fg='black',background="#F0F0F0")
        RSI_Label.place(x=10,y=100)
        
        KD_Label = tk.Label(Parameter_root,text='KD',font=('Arial',12),fg='black',background="#F0F0F0")
        KD_Label.place(x=10,y=140)

        EMA_Label = tk.Label(Parameter_root,text='EMA',font=('Arial',12),fg='black',background="#F0F0F0")
        EMA_Label.place(x=10,y=180)

        Break_Label = tk.Label(Parameter_root,text='Break',font=('Arial',12),fg='black',background="#F0F0F0")
        Break_Label.place(x=10,y=220)

        MTM_Label = tk.Label(Parameter_root,text='MTM',font=('Arial',12),fg='black',background="#F0F0F0")
        MTM_Label.place(x=10,y=260)

        ATR_Label = tk.Label(Parameter_root,text='ATR',font=('Arial',12),fg='black',background="#F0F0F0")
        ATR_Label.place(x=10,y=300)


        Para1_Label = tk.Label(Parameter_root,text='Parameter 1',font=('Arial',12),fg='black',background="#F0F0F0")
        Para1_Label.place(x=100,y=50)

        Para2_Label = tk.Label(Parameter_root,text='Parameter 2',font=('Arial',12),fg='black',background="#F0F0F0")
        Para2_Label.place(x=220,y=50)

        #Entry=================================================================

        RSI1_Entry = tk.Entry(Parameter_root, width = 10,justify="left",font=('Arial',12),background = 'white',fg = 'black',bd=1)
        RSI1_Entry.place(x=100,y=100)
        RSI1_Entry.insert(0,RSI_para1)
        
        KD1_Entry = tk.Entry(Parameter_root, width = 10,justify="left",font=('Arial',12),background = 'white',fg = 'black',bd=1)
        KD1_Entry.place(x=100,y=140)
        KD1_Entry.insert(0,KD_para1)

        EMA1_Entry = tk.Entry(Parameter_root, width = 10,justify="left",font=('Arial',12),background = 'white',fg = 'black',bd=1)
        EMA1_Entry.place(x=100,y=180)
        EMA1_Entry.insert(0,EMA_para1)

        Break1_Entry = tk.Entry(Parameter_root, width = 10,justify="left",font=('Arial',12),background = 'white',fg = 'black',bd=1)
        Break1_Entry.place(x=100,y=220)
        Break1_Entry.insert(0,Break_para1)

        MTM1_Entry = tk.Entry(Parameter_root, width = 10,justify="left",font=('Arial',12),background = 'white',fg = 'black',bd=1)
        MTM1_Entry.place(x=100,y=260)
        MTM1_Entry.insert(0,MTM_para1)
        
        ATR1_Entry = tk.Entry(Parameter_root, width = 10,justify="left",font=('Arial',12),background = 'white',fg = 'black',bd=1)
        ATR1_Entry.place(x=100,y=300)
        ATR1_Entry.insert(0,ATR_para1)



        RSI2_Entry = tk.Entry(Parameter_root, width = 10,justify="left",font=('Arial',12),background = 'white',fg = 'black',bd=1)
        RSI2_Entry.place(x=220,y=100)
        RSI2_Entry.insert(0,RSI_para2)
        
        KD2_Entry = tk.Entry(Parameter_root, width = 10,justify="left",font=('Arial',12),background = 'white',fg = 'black',bd=1)
        KD2_Entry.place(x=220,y=140)
        KD2_Entry.insert(0,KD_para2)

        EMA2_Entry = tk.Entry(Parameter_root, width = 10,justify="left",font=('Arial',12),background = 'white',fg = 'black',bd=1)
        EMA2_Entry.place(x=220,y=180)
        EMA2_Entry.insert(0,EMA_para2)

        Break2_Entry = tk.Entry(Parameter_root, width = 10,justify="left",font=('Arial',12),background = 'white',fg = 'black',bd=1)
        Break2_Entry.place(x=220,y=220)
        Break2_Entry.insert(0,Break_para2)

        ATR2_Entry = tk.Entry(Parameter_root, width = 10,justify="left",font=('Arial',12),background = 'white',fg = 'black',bd=1)
        ATR2_Entry.place(x=220,y=300)
        ATR2_Entry.insert(0,ATR_para2)

        RSI_Optimize_Button = ttk.Button(Parameter_root,text='Parameters Optimize',command=RSI_Optimize)
        RSI_Optimize_Button.place(x=335, y=97)

        KD_Optimize_Button = ttk.Button(Parameter_root,text='Parameters Optimize',command=KD_Optimize)
        KD_Optimize_Button.place(x=335, y=137)

        EMA_Optimize_Button = ttk.Button(Parameter_root,text='Parameters Optimize',command=EMA_Optimize)
        EMA_Optimize_Button.place(x=335, y=177)

        Break_Optimize_Button = ttk.Button(Parameter_root,text='Parameters Optimize',command=Break_Optimize)
        Break_Optimize_Button.place(x=335, y=217)

        MTM_Optimize_Button = ttk.Button(Parameter_root,text='Parameters Optimize',command=MTM_Optimize)
        MTM_Optimize_Button.place(x=335, y=257)

        ATR_Optimize_Button = ttk.Button(Parameter_root,text='Parameters Optimize',command=ATR_Optimize)
        ATR_Optimize_Button.place(x=335, y=297)





        Save_Button = ttk.Button(Parameter_root,text='  Save  ',command = Save_Parameters_to_db)
        Save_Button.place(x=220, y=400)

        Cancel_Button = ttk.Button(Parameter_root,text='  Cancel  ',command=Parameter_root.destroy)
        Cancel_Button.place(x=300, y=400)

        
        Parameter_root.mainloop()
    
    def Find_Parameters(self,Ticker,Name):
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
        cursor = connection.cursor()
        #Database_Functions.Fetch_All_Data()
        cursor.execute("SELECT * FROM Strategy_table WHERE Ticker = ? AND Name = ?",(Ticker,Name))
        data = cursor.fetchall()
        
        #print data[0] 
        RSI_Para1 = int(data[0][2].split(',')[0])
        RSI_Para2 = int(data[0][2].split(',')[1])

        KD_Para1 = int(data[0][3].split(',')[0])
        KD_Para2 = int(data[0][3].split(',')[1])

        EMA_Para1 = int(data[0][4].split(',')[0])
        EMA_Para2 = int(data[0][4].split(',')[1])

        Break_Para1 = int(data[0][5].split(',')[0])
        Break_Para2 = int(data[0][5].split(',')[1])

        MTM_Para1 = int(data[0][6])

        ATR_Para1 = int(data[0][7].split(',')[0])
        ATR_Para2 = float(data[0][7].split(',')[1])

        return RSI_Para1,RSI_Para2,KD_Para1,KD_Para2,EMA_Para1,EMA_Para2,Break_Para1,Break_Para2,MTM_Para1,ATR_Para1,ATR_Para2
    
    def Index_Performance(self,event):
        
        curItem = self.Score_Table.focus()
        Index_info = self.Score_Table.item(curItem)['values']
        Ticker = Index_info[0]
        Name = Index_info[1]
        
        
        
        #Get Data and Parameters===============================================
        Data = Database_Functions.Fetch(Ticker).dropna()
        Data['Return'] = Data[Ticker].pct_change()
        Data['Cumulative_Return'] = Data['Return'].cumsum()#.apply(np.exp)
        Data['Cumulative_Return1'] = Data['Return'].cumsum().apply(np.exp)

        Horizon = (Data.index[-1] - Data.index[0]).days/365.
        Annual_Return = Data['Cumulative_Return1'][-1]**(1./Horizon) - 1
        Annual_Return_str = str(round(Annual_Return*100.,2))+'%'
        
        Volatility = np.std(Data['Return'])*(250**0.5)
        Volatility_str = str(round(Volatility*100.,2))+'%'
        
        Sharpe = Annual_Return / Volatility
        Sharpe_str = str(round(Sharpe,2))

        
        
        
        
        
        
        
        Backtest_root = tk.Tk()
        Backtest_root.wm_title(str(Name)+' Backtest Report')
        Backtest_root.geometry('1000x580')
        Backtest_root.iconbitmap('D:/Taishin_Platform/pics/Data_Manager.ico')
        Backtest_root._offsetx = 0
        Backtest_root._offsety = 0

        Backtest_Frame_Top=tk.Frame(Backtest_root, width=1000, height=80, background="#282828")
        Backtest_Frame_Top.place(x=0, y=0)
        

        Name_Label = tk.Label(Backtest_Frame_Top,text=str(Name),font=('Arial',24,'bold'),fg='white',background="#282828")
        Name_Label.place(x=10,y=15)
        
        Return_Label = tk.Label(Backtest_Frame_Top,text='Annual Return = '+Annual_Return_str,font=('Arial',12,'bold'),fg='white',background="#282828")
        Return_Label.place(x=10,y=55)

        Volatility_Label = tk.Label(Backtest_Frame_Top,text='Volatility = '+Volatility_str,font=('Arial',12,'bold'),fg='white',background="#282828")
        Volatility_Label.place(x=210,y=55)
        
        Sharpe_Label = tk.Label(Backtest_Frame_Top,text='Sharpe = '+Sharpe_str,font=('Arial',12,'bold'),fg='white',background="#282828")
        Sharpe_Label.place(x=355,y=55)
        
        
        Backtest_Frame_Down=tk.Frame(Backtest_root,relief=tk.GROOVE,width=1000,height=100,bd=0)
        Backtest_Frame_Down.place(x=0,y=80)
        
        #canvas=tk.Canvas(Backtest_Frame_Down)
        #frame=tk.Frame(canvas)

        #style = ttk.Style()
        #style.configure('CustomScroll.Vertical.TScrollbar')

        #myscrollbar=ttk.Scrollbar(Backtest_Frame_Down,orient="vertical",command=canvas.yview)
        #myscrollbar['style'] = 'CustomScroll.Vertical.TScrollbar'

        #canvas.configure(yscrollcommand=myscrollbar.set)
        '''
        myscrollbar.pack(side="right",fill="y")
        canvas.pack(side="left")
        canvas.create_window((0,0),window=frame,anchor='nw')
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        frame.bind("<Configure>",myfunction)
        '''
        RSI_Frame=tk.Frame(Backtest_Frame_Down, width=1000, height=500, background="#F0F0F0")
        RSI_Frame.grid(row=0,column=1)

        #======================================================================
        fig_RSI = Figure(figsize=(8,4), dpi=120)
        fig_RSI.set_tight_layout(True)
        fig_RSI.patch.set_facecolor('#F0F0F0')
        
        RSI_Chart = fig_RSI.add_subplot(111,axisbg='#F0F0F0')
        Chart2 = RSI_Chart.twinx()
        
        RSI_Chart.tick_params(axis='both', which='major', labelsize=8)
        RSI_Chart.tick_params(axis='both', which='major',colors='black', labelsize=6)
        RSI_Chart.spines['bottom'].set_color('black')
        RSI_Chart.spines['top'].set_color('black')
        RSI_Chart.spines['left'].set_color('black')
        RSI_Chart.spines['right'].set_color('black')
        RSI_Chart.xaxis.label.set_color('black')
        RSI_Chart.yaxis.label.set_color('black')
        
        Chart2.tick_params(axis='both', which='major', labelsize=8)
        Chart2.tick_params(axis='both', which='major',colors='black', labelsize=6)
        Chart2.spines['bottom'].set_color('black')
        Chart2.spines['top'].set_color('black')
        Chart2.spines['left'].set_color('black')
        Chart2.spines['right'].set_color('black')
        Chart2.xaxis.label.set_color('black')
        Chart2.yaxis.label.set_color('black')





        
        canvas_RSI = FigureCanvasTkAgg(fig_RSI, master=RSI_Frame)
        canvas_RSI.get_tk_widget().place(x=0,y=0)
        canvas_RSI.get_tk_widget().configure(background='#F0F0F0',  highlightcolor='#F0F0F0', highlightbackground='#F0F0F0')

        #======================================================================
        
        RSI_Chart.clear()
        Chart2.clear()
        #Buy & Hold Chart======================================================
        #print Historical_Score_df[Ticker][-400:]
        #RSI_Chart.plot( Data.index[-60:], Data['Cumulative_Return'][-60:],lw=1,color='#01485E',alpha=0.5, label='Buy&Hold')
        #RSI_Chart.plot( Historical_Score_df[Ticker].index[-60:], Historical_Score_df[Ticker][-60:],lw=1,color='#01485E',alpha=0.5, label='Buy&Hold')
        print Historical_Score_df[Ticker][-400:]
        Chart2.bar( Historical_Score_df[Ticker].index[-400:], Historical_Score_df[Ticker][-400:],lw=0.5,color='#00A3DC',alpha=0.5)
        RSI_Chart.plot( Data.index[-2000:], Data['Cumulative_Return'][-2000:],lw=1,color='red',alpha=1, label='Buy&Hold')
        #00A3DC
        #01485E
    
        Backtest_root.mainloop()
    
    def Double_Click_Selection(self,event):
        
        def myfunction(event):
            canvas.configure(background="#181818",scrollregion=canvas.bbox("all"),width=980,height=660)
        
        def _on_mousewheel(event):
            canvas.yview_scroll(-1*(event.delta/120), "units")
        
        def EMA_Performance(Data,para1,para2):
            Data['EMA1'] = pd.ewma(Data[Ticker],para1)
            Data['EMA2'] = pd.ewma(Data[Ticker],para2)
            Data['EMA_Signal'] = np.where(Data['EMA1']>Data['EMA2'],1,0)
            Data['Return'] = Data[Ticker].pct_change()
            Data['EMA_Return'] = Data['Return']*Data['EMA_Signal'].shift(1)
            Data['EMA_EquityCurve'] = Data['EMA_Return'].cumsum().apply(np.exp)
            
            Horizon = (Data.index[-1] - Data.index[0]).days/365.
            Annual_Return = Data['EMA_EquityCurve'][-1]**(1./Horizon) - 1
            Volatility = np.std(Data['EMA_Return'])*(250**0.5)
            Sharpe = Annual_Return / Volatility

            Total_Turnover = abs(Data['EMA_Signal'].diff()).sum()
            Annual_Turnover = Total_Turnover/Horizon


            Data['EMA_Buy_Price'] = np.where((Data['EMA_Signal'] == 1)&(Data['EMA_Signal'].shift(1) == 0),Data[str(Ticker)],np.nan)
            Data['EMA_Sell_Price'] = np.where((Data['EMA_Signal'] == 0)&(Data['EMA_Signal'].shift(1) == 1),Data[str(Ticker)],np.nan)
            
            Trading_book = pd.concat([Data['EMA_Buy_Price'].dropna(),Data['EMA_Sell_Price'].dropna()],axis=1)
            Trading_book['EMA_Sell_Price'] = Trading_book['EMA_Sell_Price'].shift(-1)
            Trading_book['EMA_Profit'] = np.log(Trading_book['EMA_Sell_Price']/Trading_book['EMA_Buy_Price'])
            
            Sign = np.where(Trading_book['EMA_Profit'].dropna()>0.,1.,0.)
            
            Win_Percentage = str(round(sum(Sign)/len(Sign)*100.,2))+'%'
            Expect =  np.mean(Trading_book['EMA_Profit'])    
            
            Buy = pd.DataFrame(Data['EMA_Buy_Price'].dropna())
            Buy['Sign'] = 'Buy'
            Sell = pd.DataFrame(Data['EMA_Sell_Price'].dropna())
            Sell['Sign'] = 'Sell'
            
            df = pd.concat([Buy,Sell]).sort_index()
            df = df.replace(np.nan,0)
            df['Total'] = df['EMA_Buy_Price']+df['EMA_Sell_Price']
            Trading_detail = df[['Sign','Total']].tail()            
            
            
            return str(round(Annual_Return*100.,2))+'%',str(round(Volatility*100.,2))+'%',str(round(Sharpe,2)),Win_Percentage,str(round(Annual_Turnover,2)),str(round(Expect*100.,2))+'%',Trading_detail

        def RSI_func(prices,n):
            n = int(n)
            deltas = np.diff(prices)
            seed = deltas[:n+1]
            up = seed[seed>=0].sum()/n
            down = -seed[seed<0].sum()/n
            rs = up/down
            rsi = np.zeros_like(prices)
            rsi[:n]=100. - 100./(1.+rs)
                
            for i in range(n,len(prices)):
                delta = deltas[i-1]    #cause the diff is 1 shorter
                    
                if delta>0:
                    upval = delta
                    downval = 0.
                else:
                    upval = 0.
                    downval = -delta
                        
                up = (up*(n-1) + upval)/n
                down = (down*(n-1) + downval)/n
                    
                rs = up/down
                rsi[i] = 100. - 100./(1.+rs)
                
            return rsi

        def RSI_Performance(Data,para1,para2):
            Data['RSI1'] = RSI_func(Data[Ticker],para1)
            Data['RSI2'] = RSI_func(Data[Ticker],para2)
            Data['RSI_Signal'] = np.where(Data['RSI1']>Data['RSI2'],1,0)
            Data['Return'] = Data[Ticker].pct_change()
            Data['RSI_Return'] = Data['Return']*Data['RSI_Signal'].shift(1)
            Data['RSI_EquityCurve'] = Data['RSI_Return'].cumsum().apply(np.exp)
            
            Horizon = (Data.index[-1] - Data.index[0]).days/365.
            Annual_Return = Data['RSI_EquityCurve'][-1]**(1./Horizon) - 1
            Volatility = np.std(Data['RSI_Return'])*(250**0.5)
            Sharpe = Annual_Return / Volatility

            Total_Turnover = abs(Data['RSI_Signal'].diff()).sum()
            Annual_Turnover = Total_Turnover/Horizon

            
            Data['RSI_Buy_Price'] = np.where((Data['RSI_Signal'] == 1)&(Data['RSI_Signal'].shift(1) == 0),Data[str(Ticker)],np.nan)
            Data['RSI_Sell_Price'] = np.where((Data['RSI_Signal'] == 0)&(Data['RSI_Signal'].shift(1) == 1),Data[str(Ticker)],np.nan)
            
            Trading_book = pd.concat([Data['RSI_Buy_Price'].dropna(),Data['RSI_Sell_Price'].dropna()],axis=1)
            Trading_book['RSI_Sell_Price'] = Trading_book['RSI_Sell_Price'].shift(-1)
            Trading_book['RSI_Profit'] = np.log(Trading_book['RSI_Sell_Price']/Trading_book['RSI_Buy_Price'])
            
            
            Sign = np.where(Trading_book['RSI_Profit'].dropna()>0.,1.,0.)
            
            Win_Percentage = str(round(sum(Sign)/len(Sign)*100.,2))+'%'
            Expect =  np.mean(Trading_book['RSI_Profit'])    
            
            
            Buy = pd.DataFrame(Data['RSI_Buy_Price'].dropna())
            Buy['Sign'] = 'Buy'
            Sell = pd.DataFrame(Data['RSI_Sell_Price'].dropna())
            Sell['Sign'] = 'Sell'
            
            df = pd.concat([Buy,Sell]).sort_index()
            df = df.replace(np.nan,0)
            df['Total'] = df['RSI_Buy_Price']+df['RSI_Sell_Price']
            Trading_detail = df[['Sign','Total']].tail()            
            
            
            
            return str(round(Annual_Return*100.,2))+'%',str(round(Volatility*100.,2))+'%',str(round(Sharpe,2)),Win_Percentage,str(round(Annual_Turnover,2)),str(round(Expect*100.,2))+'%',Trading_detail

        def RSV(price,n):
            RSV = (price - pd.rolling_min(price,int(n)))/(pd.rolling_max(price,int(n))-pd.rolling_min(price,int(n)))
            return RSV*100.
        
        def K(Series,n):
            Series = RSV(Series,n)
            Series = pd.ewma(Series,5)
            return Series
        
        def D(Series,n):
            Series = pd.ewma(Series,n)
            return Series


        def KD_Performance(Data,para1,para2):
            Data['KD1'] = K(Data[Ticker],para1)
            Data['KD2'] = D(Data['KD1'],para2)
            Data['KD_Signal'] = np.where(Data['KD1']>Data['KD2'],1,0)
            Data['Return'] = Data[Ticker].pct_change()
            Data['KD_Return'] = Data['Return']*Data['KD_Signal'].shift(1)
            Data['KD_EquityCurve'] = Data['KD_Return'].cumsum().apply(np.exp)

            Horizon = (Data.index[-1] - Data.index[0]).days/365.
            Annual_Return = Data['KD_EquityCurve'][-1]**(1./Horizon) - 1
            Volatility = np.std(Data['KD_Return'])*(250**0.5)
            Sharpe = Annual_Return / Volatility

            Total_Turnover = abs(Data['KD_Signal'].diff()).sum()
            Annual_Turnover = Total_Turnover/Horizon


            Data['KD_Buy_Price'] = np.where((Data['KD_Signal'] == 1)&(Data['KD_Signal'].shift(1) == 0),Data[str(Ticker)],np.nan)
            Data['KD_Sell_Price'] = np.where((Data['KD_Signal'] == 0)&(Data['KD_Signal'].shift(1) == 1),Data[str(Ticker)],np.nan)
            
            Trading_book = pd.concat([Data['KD_Buy_Price'].dropna(),Data['KD_Sell_Price'].dropna()],axis=1)
            Trading_book['KD_Sell_Price'] = Trading_book['KD_Sell_Price'].shift(-1)
            Trading_book['KD_Profit'] = np.log(Trading_book['KD_Sell_Price']/Trading_book['KD_Buy_Price'])
            
            
            Sign = np.where(Trading_book['KD_Profit'].dropna()>0.,1.,0.)
            
            Win_Percentage = str(round(sum(Sign)/len(Sign)*100.,2))+'%'
            Expect =  np.mean(Trading_book['KD_Profit'])    

            Buy = pd.DataFrame(Data['KD_Buy_Price'].dropna())
            Buy['Sign'] = 'Buy'
            Sell = pd.DataFrame(Data['KD_Sell_Price'].dropna())
            Sell['Sign'] = 'Sell'
            
            df = pd.concat([Buy,Sell]).sort_index()
            df = df.replace(np.nan,0)
            df['Total'] = df['KD_Buy_Price']+df['KD_Sell_Price']
            Trading_detail = df[['Sign','Total']].tail()            




            return str(round(Annual_Return*100.,2))+'%',str(round(Volatility*100.,2))+'%',str(round(Sharpe,2)),Win_Percentage,str(round(Annual_Turnover,2)),str(round(Expect*100.,2))+'%',Trading_detail

        def Break_Performance(Data,para1,para2):
            Data['High'] = pd.rolling_max(Data[Data.columns[0]].shift(1),int(para1))
            Data['Low'] = pd.rolling_min(Data[Data.columns[0]].shift(1),int(para2))
            Data['Break_Signal_temp'] = np.where(Data[Data.columns[0]]>Data['High'],1,np.where(Data[Data.columns[0]]<Data['Low'],0,np.nan))
            #print Data['Break_Signal_temp']
            Data['Break_Signal'] = Data['Break_Signal_temp'].fillna(method='ffill')
            Data['Return'] = Data[Ticker].pct_change()
            Data['Break_Return'] = Data['Return']*Data['Break_Signal'].shift(1)
            Data['Break_EquityCurve'] = Data['Break_Return'].cumsum().apply(np.exp)

            Horizon = (Data.index[-1] - Data.index[0]).days/365.
            Annual_Return = Data['Break_EquityCurve'][-1]**(1./Horizon) - 1
            Volatility = np.std(Data['Break_Return'])*(250**0.5)
            Sharpe = Annual_Return / Volatility

            Total_Turnover = abs(Data['Break_Signal'].diff()).sum()
            Annual_Turnover = Total_Turnover/Horizon

            
            Data['Break_Buy_Price'] = np.where((Data['Break_Signal'] == 1)&(Data['Break_Signal'].shift(1) == 0),Data[str(Ticker)],np.nan)
            Data['Break_Sell_Price'] = np.where((Data['Break_Signal'] == 0)&(Data['Break_Signal'].shift(1) == 1),Data[str(Ticker)],np.nan)
            
            Trading_book = pd.concat([Data['Break_Buy_Price'].dropna(),Data['Break_Sell_Price'].dropna()],axis=1)
            Trading_book['Break_Sell_Price'] = Trading_book['Break_Sell_Price'].shift(-1)
            Trading_book['Break_Profit'] = np.log(Trading_book['Break_Sell_Price']/Trading_book['Break_Buy_Price'])
            
            
            Sign = np.where(Trading_book['Break_Profit'].dropna()>0.,1.,0.)
            
            Win_Percentage = str(round(sum(Sign)/len(Sign)*100.,2))+'%'
            Expect =  np.mean(Trading_book['Break_Profit'])    
            
            Buy = pd.DataFrame(Data['Break_Buy_Price'].dropna())
            Buy['Sign'] = 'Buy'
            Sell = pd.DataFrame(Data['Break_Sell_Price'].dropna())
            Sell['Sign'] = 'Sell'
            
            df = pd.concat([Buy,Sell]).sort_index()
            df = df.replace(np.nan,0)
            df['Total'] = df['Break_Buy_Price']+df['Break_Sell_Price']
            Trading_detail = df[['Sign','Total']].tail()            

            
            return str(round(Annual_Return*100.,2))+'%',str(round(Volatility*100.,2))+'%',str(round(Sharpe,2)),Win_Percentage,str(round(Annual_Turnover,2)),str(round(Expect*100.,2))+'%',Trading_detail

        def MTM_Performance(Data,para1):
            Data['MTM'] =Data[Data.columns[0]].pct_change(para1)
            Data['MTM_Signal'] = np.where(Data['MTM']>0.,1,0)
            Data['Return'] = Data[Ticker].pct_change()
            Data['MTM_Return'] = Data['Return']*Data['MTM_Signal'].shift(1)
            Data['MTM_EquityCurve'] = Data['MTM_Return'].cumsum().apply(np.exp)

            Horizon = (Data.index[-1] - Data.index[0]).days/365.
            Annual_Return = Data['MTM_EquityCurve'][-1]**(1./Horizon) - 1
            Volatility = np.std(Data['MTM_Return'])*(250**0.5)
            Sharpe = Annual_Return / Volatility

            Total_Turnover = abs(Data['MTM_Signal'].diff()).sum()
            Annual_Turnover = Total_Turnover/Horizon

            
            Data['MTM_Buy_Price'] = np.where((Data['MTM_Signal'] == 1)&(Data['MTM_Signal'].shift(1) == 0),Data[str(Ticker)],np.nan)
            Data['MTM_Sell_Price'] = np.where((Data['MTM_Signal'] == 0)&(Data['MTM_Signal'].shift(1) == 1),Data[str(Ticker)],np.nan)
            
            Trading_book = pd.concat([Data['MTM_Buy_Price'].dropna(),Data['MTM_Sell_Price'].dropna()],axis=1)
            Trading_book['MTM_Sell_Price'] = Trading_book['MTM_Sell_Price'].shift(-1)
            Trading_book['MTM_Profit'] = np.log(Trading_book['MTM_Sell_Price']/Trading_book['MTM_Buy_Price'])
            
            
            Sign = np.where(Trading_book['MTM_Profit'].dropna()>0.,1.,0.)
            
            Win_Percentage = str(round(sum(Sign)/len(Sign)*100.,2))+'%'
            Expect =  np.mean(Trading_book['MTM_Profit'])    
            
            Buy = pd.DataFrame(Data['MTM_Buy_Price'].dropna())
            Buy['Sign'] = 'Buy'
            Sell = pd.DataFrame(Data['MTM_Sell_Price'].dropna())
            Sell['Sign'] = 'Sell'
            
            df = pd.concat([Buy,Sell]).sort_index()
            df = df.replace(np.nan,0)
            df['Total'] = df['MTM_Buy_Price']+df['MTM_Sell_Price']
            Trading_detail = df[['Sign','Total']].tail()            
            
            return str(round(Annual_Return*100.,2))+'%',str(round(Volatility*100.,2))+'%',str(round(Sharpe,2)),Win_Percentage,str(round(Annual_Turnover,2)),str(round(Expect*100.,2))+'%',Trading_detail

        def ATR_Performance(Data,para1,para2):
            Data['EMA_ATR'] = pd.ewma(Data[Data.columns[0]],para1)
            Data['ATR'] = pd.ewma(abs(Data[Data.columns[0]].diff()),para1)
        
            Data['UpBound'] = Data['EMA_ATR']+para2*Data['ATR']
            #Data['LowBound'] = Data['EMA']-Para2*Data['ATR']
            
            Data['ATR_Signal'] = np.where(Data[Data.columns[0]]>Data['UpBound'],1,0)
            Data['Return'] = Data[Ticker].pct_change()
            Data['ATR_Return'] = Data['Return']*Data['ATR_Signal'].shift(1)
            Data['ATR_EquityCurve'] = Data['ATR_Return'].cumsum().apply(np.exp)
            
            Horizon = (Data.index[-1] - Data.index[0]).days/365.
            Annual_Return = Data['ATR_EquityCurve'][-1]**(1./Horizon) - 1
            Volatility = np.std(Data['ATR_Return'])*(250**0.5)
            Sharpe = Annual_Return / Volatility

            Total_Turnover = abs(Data['ATR_Signal'].diff()).sum()
            Annual_Turnover = Total_Turnover/Horizon

            
            Data['ATR_Buy_Price'] = np.where((Data['ATR_Signal'] == 1)&(Data['ATR_Signal'].shift(1) == 0),Data[str(Ticker)],np.nan)
            Data['ATR_Sell_Price'] = np.where((Data['ATR_Signal'] == 0)&(Data['ATR_Signal'].shift(1) == 1),Data[str(Ticker)],np.nan)
            
            Trading_book = pd.concat([Data['ATR_Buy_Price'].dropna(),Data['ATR_Sell_Price'].dropna()],axis=1)
            Trading_book['ATR_Sell_Price'] = Trading_book['ATR_Sell_Price'].shift(-1)
            Trading_book['ATR_Profit'] = np.log(Trading_book['ATR_Sell_Price']/Trading_book['ATR_Buy_Price'])
            
            
            Sign = np.where(Trading_book['ATR_Profit'].dropna()>0.,1.,0.)
            
            Win_Percentage = str(round(sum(Sign)/len(Sign)*100.,2))+'%'
            Expect =  np.mean(Trading_book['ATR_Profit'])    

            Buy = pd.DataFrame(Data['ATR_Buy_Price'].dropna())
            Buy['Sign'] = 'Buy'
            Sell = pd.DataFrame(Data['ATR_Sell_Price'].dropna())
            Sell['Sign'] = 'Sell'
            
            df = pd.concat([Buy,Sell]).sort_index()
            df = df.replace(np.nan,0)
            df['Total'] = df['ATR_Buy_Price']+df['ATR_Sell_Price']
            Trading_detail = df[['Sign','Total']].tail()            
            
            return str(round(Annual_Return*100.,2))+'%',str(round(Volatility*100.,2))+'%',str(round(Sharpe,2)),Win_Percentage,str(round(Annual_Turnover,2)),str(round(Expect*100.,2))+'%',Trading_detail


        
        curItem = self.Strategy_Table.focus()
        Index_info = self.Strategy_Table.item(curItem)['values']
        Ticker = Index_info[0]
        Name = Index_info[1]
        Parameters = self.Find_Parameters(Ticker,Name)
        
        
        
        #Get Data and Parameters===============================================
        Data = Database_Functions.Fetch(Ticker).dropna()
        Data['Return'] = Data[Ticker].pct_change()
        Data['Cumulative_Return'] = Data['Return'].cumsum()#.apply(np.exp)
        Data['Cumulative_Return1'] = Data['Return'].cumsum().apply(np.exp)

        Horizon = (Data.index[-1] - Data.index[0]).days/365.
        Annual_Return = Data['Cumulative_Return1'][-1]**(1./Horizon) - 1
        Annual_Return_str = str(round(Annual_Return*100.,2))+'%'
        
        Volatility = np.std(Data['Return'])*(250**0.5)
        Volatility_str = str(round(Volatility*100.,2))+'%'
        
        Sharpe = Annual_Return / Volatility
        Sharpe_str = str(round(Sharpe,2))
        
        
        
        
        #print Parameters
        RSI = [Parameters[0],Parameters[1]]
        KD = [Parameters[2],Parameters[3]]
        EMA = [Parameters[4],Parameters[5]]
        Break = [Parameters[6],Parameters[7]]
        MTM = [Parameters[8]]
        ATR = [Parameters[9],Parameters[10]]
        #======================================================================
        
        
        
        Backtest_root = tk.Tk()
        Backtest_root.wm_title(str(Name)+' Backtest Report')
        Backtest_root.geometry('1000x760')
        Backtest_root.iconbitmap('D:/Taishin_Platform/pics/Data_Manager.ico')
        Backtest_root._offsetx = 0
        Backtest_root._offsety = 0

        Backtest_Frame_Top=tk.Frame(Backtest_root, width=1000, height=80, background="#282828")
        Backtest_Frame_Top.place(x=0, y=0)
        

        Name_Label = tk.Label(Backtest_Frame_Top,text=str(Name),font=('Arial',24,'bold'),fg='white',background="#282828")
        Name_Label.place(x=10,y=15)
        
        Return_Label = tk.Label(Backtest_Frame_Top,text='Annual Return = '+Annual_Return_str,font=('Arial',12,'bold'),fg='white',background="#282828")
        Return_Label.place(x=10,y=55)

        Volatility_Label = tk.Label(Backtest_Frame_Top,text='Volatility = '+Volatility_str,font=('Arial',12,'bold'),fg='white',background="#282828")
        Volatility_Label.place(x=210,y=55)
        
        Sharpe_Label = tk.Label(Backtest_Frame_Top,text='Sharpe = '+Sharpe_str,font=('Arial',12,'bold'),fg='white',background="#282828")
        Sharpe_Label.place(x=355,y=55)
        
        
        Backtest_Frame_Down=tk.Frame(Backtest_root,relief=tk.GROOVE,width=1000,height=100,bd=0)
        Backtest_Frame_Down.place(x=0,y=80)
        
        canvas=tk.Canvas(Backtest_Frame_Down)
        frame=tk.Frame(canvas)

        style = ttk.Style()
        style.configure('CustomScroll.Vertical.TScrollbar')




        myscrollbar=ttk.Scrollbar(Backtest_Frame_Down,orient="vertical",command=canvas.yview)
        myscrollbar['style'] = 'CustomScroll.Vertical.TScrollbar'

        canvas.configure(yscrollcommand=myscrollbar.set)
        
        myscrollbar.pack(side="right",fill="y")
        canvas.pack(side="left")
        canvas.create_window((0,0),window=frame,anchor='nw')
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        frame.bind("<Configure>",myfunction)
        #======================================================================
        RSI_Frame=tk.Frame(frame, width=1000, height=500, background="#F0F0F0")
        RSI_Frame.grid(row=0,column=1)
        
        RSI_Label = tk.Label(RSI_Frame,text='RSI',font=('Arial',18,'bold'),fg='black',background="#F0F0F0")
        RSI_Label.place(x=750,y=15)
        
        RSI_text = str(RSI[0])+','+str(RSI[1])
        RSI_Parameter_Label = tk.Label(RSI_Frame,text=RSI_text,font=('Arial',18,'bold'),fg='black',background="#F0F0F0")
        RSI_Parameter_Label.place(x=980,y=15,anchor='ne')
        
        ttk.Separator(RSI_Frame,orient=tk.HORIZONTAL).place(x=750,y=50,width=300)

        RSI_Win_Label = tk.Label(RSI_Frame,text='Win(%)',font=('Arial',14),fg='black',background="#F0F0F0")
        RSI_Win_Label.place(x=750,y=60)

        RSI_Return_Label = tk.Label(RSI_Frame,text='Annual Return',font=('Arial',14),fg='black',background="#F0F0F0")
        RSI_Return_Label.place(x=750,y=100)

        RSI_Volatility_Label = tk.Label(RSI_Frame,text='Volatility',font=('Arial',14),fg='black',background="#F0F0F0")
        RSI_Volatility_Label.place(x=750,y=140)

        RSI_Sharpe_Label = tk.Label(RSI_Frame,text='Sharpe',font=('Arial',14),fg='black',background="#F0F0F0")
        RSI_Sharpe_Label.place(x=750,y=180)

        RSI_Turnover_Label = tk.Label(RSI_Frame,text='Turnover',font=('Arial',14),fg='black',background="#F0F0F0")
        RSI_Turnover_Label.place(x=750,y=220)

        RSI_Expect_Label = tk.Label(RSI_Frame,text='Expected Return',font=('Arial',14),fg='black',background="#F0F0F0")
        RSI_Expect_Label.place(x=750,y=260)

        
        #Performance
        RSI_Performance = RSI_Performance(Data = Data,para1 = RSI[0],para2 = RSI[1])
        
        RSI_Win_Performance = tk.Label(RSI_Frame,text=RSI_Performance[3],font=('Arial',14),fg='black',background="#F0F0F0")
        RSI_Win_Performance.place(x=980,y=60,anchor='ne')

        RSI_Return_Performance = tk.Label(RSI_Frame,text=RSI_Performance[0],font=('Arial',14),fg='black',background="#F0F0F0")
        RSI_Return_Performance.place(x=980,y=100,anchor='ne')

        RSI_Volatility_Performance = tk.Label(RSI_Frame,text=RSI_Performance[1],font=('Arial',14),fg='black',background="#F0F0F0")
        RSI_Volatility_Performance.place(x=980,y=140,anchor='ne')

        RSI_Sharpe_Performance = tk.Label(RSI_Frame,text=RSI_Performance[2],font=('Arial',14),fg='black',background="#F0F0F0")
        RSI_Sharpe_Performance.place(x=980,y=180,anchor='ne')

        RSI_Turnover_Performance = tk.Label(RSI_Frame,text=RSI_Performance[4],font=('Arial',14),fg='black',background="#F0F0F0")
        RSI_Turnover_Performance.place(x=980,y=220,anchor='ne')

        RSI_Expect_Performance = tk.Label(RSI_Frame,text=RSI_Performance[5],font=('Arial',14),fg='black',background="#F0F0F0")
        RSI_Expect_Performance.place(x=980,y=260,anchor='ne')


        RSI_Trading = ttk.Treeview(RSI_Frame,height="5")

        RSI_Trading["columns"]=("column1","column2",'column3')
        RSI_Trading.column("#0",width=10, anchor='e')
        RSI_Trading.column("column1", width=80, anchor='center' )
        RSI_Trading.column("column2", width=60, anchor='center')
        RSI_Trading.column("column3", width=80 , anchor='center')
        
        
        RSI_Trading.heading('#0', text='')
        RSI_Trading.heading("column1", text="Date")
        RSI_Trading.heading("column2", text="Signal")
        RSI_Trading.heading("column3", text="Price")
        
        RSI_Trading.place(x=750, y=320)


        RSI_Trading.delete(*RSI_Trading.get_children())    
        
        RSI_table = RSI_Performance[6]
        
        for i in range(5):
            RSI_Trading.insert("",i,text=str(i),values=(str(RSI_table.index[i])[:10],RSI_table['Sign'][i],RSI_table['Total'][i]))









        #======================================================================
        KD_Frame=tk.Frame(frame, width=1000, height=500, background="#F0F0F0")
        KD_Frame.grid(row=1,column=1)
        
        KD_Label = tk.Label(KD_Frame,text='KD',font=('Arial',18,'bold'),fg='black',background="#F0F0F0")
        KD_Label.place(x=750,y=15)

        KD_text = str(KD[0])+','+str(KD[1])
        KD_Parameter_Label = tk.Label(KD_Frame,text=KD_text,font=('Arial',18,'bold'),fg='black',background="#F0F0F0")
        KD_Parameter_Label.place(x=980,y=15,anchor='ne')


        ttk.Separator(KD_Frame,orient=tk.HORIZONTAL).place(x=750,y=50,width=300)


        KD_Win_Label = tk.Label(KD_Frame,text='Win(%)',font=('Arial',14),fg='black',background="#F0F0F0")
        KD_Win_Label.place(x=750,y=60)

        KD_Return_Label = tk.Label(KD_Frame,text='Annual Return',font=('Arial',14),fg='black',background="#F0F0F0")
        KD_Return_Label.place(x=750,y=100)

        KD_Volatility_Label = tk.Label(KD_Frame,text='Volatility',font=('Arial',14),fg='black',background="#F0F0F0")
        KD_Volatility_Label.place(x=750,y=140)

        KD_Sharpe_Label = tk.Label(KD_Frame,text='Sharpe',font=('Arial',14),fg='black',background="#F0F0F0")
        KD_Sharpe_Label.place(x=750,y=180)

        KD_Turnover_Label = tk.Label(KD_Frame,text='Turnover',font=('Arial',14),fg='black',background="#F0F0F0")
        KD_Turnover_Label.place(x=750,y=220)

        KD_Expect_Label = tk.Label(KD_Frame,text='Expected Return',font=('Arial',14),fg='black',background="#F0F0F0")
        KD_Expect_Label.place(x=750,y=260)



        #Performance
        KD_Performance = KD_Performance(Data = Data,para1 = KD[0],para2 = KD[1])
        
        
        KD_Win_Performance = tk.Label(KD_Frame,text=KD_Performance[3],font=('Arial',14),fg='black',background="#F0F0F0")
        KD_Win_Performance.place(x=980,y=60,anchor='ne')

        KD_Return_Performance = tk.Label(KD_Frame,text=KD_Performance[0],font=('Arial',14),fg='black',background="#F0F0F0")
        KD_Return_Performance.place(x=980,y=100,anchor='ne')

        KD_Volatility_Performance = tk.Label(KD_Frame,text=KD_Performance[1],font=('Arial',14),fg='black',background="#F0F0F0")
        KD_Volatility_Performance.place(x=980,y=140,anchor='ne')

        KD_Sharpe_Performance = tk.Label(KD_Frame,text=KD_Performance[2],font=('Arial',14),fg='black',background="#F0F0F0")
        KD_Sharpe_Performance.place(x=980,y=180,anchor='ne')

        KD_Turnover_Performance = tk.Label(KD_Frame,text=KD_Performance[4],font=('Arial',14),fg='black',background="#F0F0F0")
        KD_Turnover_Performance.place(x=980,y=220,anchor='ne')

        KD_Expect_Performance = tk.Label(KD_Frame,text=KD_Performance[5],font=('Arial',14),fg='black',background="#F0F0F0")
        KD_Expect_Performance.place(x=980,y=260,anchor='ne')

        KD_Trading = ttk.Treeview(KD_Frame,height="5")

        KD_Trading["columns"]=("column1","column2",'column3')
        KD_Trading.column("#0",width=10, anchor='e')
        KD_Trading.column("column1", width=80, anchor='center' )
        KD_Trading.column("column2", width=60, anchor='center')
        KD_Trading.column("column3", width=80 , anchor='center')
        
        
        KD_Trading.heading('#0', text='')
        KD_Trading.heading("column1", text="Date")
        KD_Trading.heading("column2", text="Signal")
        KD_Trading.heading("column3", text="Price")
        
        KD_Trading.place(x=750, y=320)


        KD_Trading.delete(*KD_Trading.get_children())    
        
        KD_table = KD_Performance[6]
        
        for i in range(5):
            KD_Trading.insert("",i,text=str(i),values=(str(KD_table.index[i])[:10],KD_table['Sign'][i],KD_table['Total'][i]))



        #======================================================================
        EMA_Frame=tk.Frame(frame, width=1000, height=500, background="#F0F0F0")
        EMA_Frame.grid(row=2,column=1)
        
        EMA_Label = tk.Label(EMA_Frame,text='EMA',font=('Arial',18,'bold'),fg='black',background="#F0F0F0")
        EMA_Label.place(x=750,y=15)
        
        EMA_text = str(EMA[0])+','+str(EMA[1])
        EMA_Parameter_Label = tk.Label(EMA_Frame,text=EMA_text,font=('Arial',18,'bold'),fg='black',background="#F0F0F0")
        EMA_Parameter_Label.place(x=980,y=15,anchor='ne')



        ttk.Separator(EMA_Frame,orient=tk.HORIZONTAL).place(x=750,y=50,width=300)



        
        EMA_Win_Label = tk.Label(EMA_Frame,text='Win(%)',font=('Arial',14),fg='black',background="#F0F0F0")
        EMA_Win_Label.place(x=750,y=60)


        EMA_Return_Label = tk.Label(EMA_Frame,text='Annual Return',font=('Arial',14),fg='black',background="#F0F0F0")
        EMA_Return_Label.place(x=750,y=100)
 
        EMA_Volatility_Label = tk.Label(EMA_Frame,text='Volatility',font=('Arial',14),fg='black',background="#F0F0F0")
        EMA_Volatility_Label.place(x=750,y=140)

        EMA_Sharpe_Label = tk.Label(EMA_Frame,text='Sharpe',font=('Arial',14),fg='black',background="#F0F0F0")
        EMA_Sharpe_Label.place(x=750,y=180)

        EMA_Turnover_Label = tk.Label(EMA_Frame,text='Turnover',font=('Arial',14),fg='black',background="#F0F0F0")
        EMA_Turnover_Label.place(x=750,y=220)

        EMA_Expect_Label = tk.Label(EMA_Frame,text='Expected Return',font=('Arial',14),fg='black',background="#F0F0F0")
        EMA_Expect_Label.place(x=750,y=260)


        #Performance
        Performance_EMA = EMA_Performance(Data = Data,para1 = EMA[0],para2 = EMA[1])
        
        EMA_Win_Performance = tk.Label(EMA_Frame,text=Performance_EMA[3],font=('Arial',14),fg='black',background="#F0F0F0")
        EMA_Win_Performance.place(x=980,y=60,anchor='ne')

        EMA_Return_Performance = tk.Label(EMA_Frame,text=Performance_EMA[0],font=('Arial',14),fg='black',background="#F0F0F0")
        EMA_Return_Performance.place(x=980,y=100,anchor='ne')

        EMA_Volatility_Performance = tk.Label(EMA_Frame,text=Performance_EMA[1],font=('Arial',14),fg='black',background="#F0F0F0")
        EMA_Volatility_Performance.place(x=980,y=140,anchor='ne')

        EMA_Sharpe_Performance = tk.Label(EMA_Frame,text=Performance_EMA[2],font=('Arial',14),fg='black',background="#F0F0F0")
        EMA_Sharpe_Performance.place(x=980,y=180,anchor='ne')

        EMA_Turnover_Performance = tk.Label(EMA_Frame,text=Performance_EMA[4],font=('Arial',14),fg='black',background="#F0F0F0")
        EMA_Turnover_Performance.place(x=980,y=220,anchor='ne')

        EMA_Expect_Performance = tk.Label(EMA_Frame,text=Performance_EMA[5],font=('Arial',14),fg='black',background="#F0F0F0")
        EMA_Expect_Performance.place(x=980,y=260,anchor='ne')

        EMA_Trading = ttk.Treeview(EMA_Frame,height="5")

        EMA_Trading["columns"]=("column1","column2",'column3')
        EMA_Trading.column("#0",width=10, anchor='e')
        EMA_Trading.column("column1", width=80, anchor='center' )
        EMA_Trading.column("column2", width=60, anchor='center')
        EMA_Trading.column("column3", width=80 , anchor='center')
        
        
        EMA_Trading.heading('#0', text='')
        EMA_Trading.heading("column1", text="Date")
        EMA_Trading.heading("column2", text="Signal")
        EMA_Trading.heading("column3", text="Price")
        
        EMA_Trading.place(x=750, y=320)


        EMA_Trading.delete(*EMA_Trading.get_children())    
        
        EMA_table = Performance_EMA[6]
        #print EMA_table
        for i in range(5):
            EMA_Trading.insert("",i,text=str(i),values=(str(EMA_table.index[i])[:10],EMA_table['Sign'][i],EMA_table['Total'][i]))




        #======================================================================
        Break_Frame=tk.Frame(frame, width=1000, height=500, background="#F0F0F0")
        Break_Frame.grid(row=3,column=1)
        
        Break_Label = tk.Label(Break_Frame,text='Break',font=('Arial',18,'bold'),fg='black',background="#F0F0F0")
        Break_Label.place(x=750,y=15)

        Break_text = str(Break[0])+','+str(Break[1])
        Break_Parameter_Label = tk.Label(Break_Frame,text=Break_text,font=('Arial',18,'bold'),fg='black',background="#F0F0F0")
        Break_Parameter_Label.place(x=980,y=15,anchor='ne')


        ttk.Separator(Break_Frame,orient=tk.HORIZONTAL).place(x=750,y=50,width=300)



        Break_Win_Label = tk.Label(Break_Frame,text='Win(%)',font=('Arial',14),fg='black',background="#F0F0F0")
        Break_Win_Label.place(x=750,y=60)

        Break_Return_Label = tk.Label(Break_Frame,text='Annual Return',font=('Arial',14),fg='black',background="#F0F0F0")
        Break_Return_Label.place(x=750,y=100)

        Break_Volatility_Label = tk.Label(Break_Frame,text='Volatility',font=('Arial',14),fg='black',background="#F0F0F0")
        Break_Volatility_Label.place(x=750,y=140)

        Break_Sharpe_Label = tk.Label(Break_Frame,text='Sharpe',font=('Arial',14),fg='black',background="#F0F0F0")
        Break_Sharpe_Label.place(x=750,y=180)

        Break_Turnover_Label = tk.Label(Break_Frame,text='Turnover',font=('Arial',14),fg='black',background="#F0F0F0")
        Break_Turnover_Label.place(x=750,y=220)

        Break_Expect_Label = tk.Label(Break_Frame,text='Expected Return',font=('Arial',14),fg='black',background="#F0F0F0")
        Break_Expect_Label.place(x=750,y=260)


        #Performance
        Break_Performance = Break_Performance(Data = Data,para1 = Break[0],para2 = Break[1])

        Break_Win_Performance = tk.Label(Break_Frame,text=Break_Performance[3],font=('Arial',14),fg='black',background="#F0F0F0")
        Break_Win_Performance.place(x=980,y=60,anchor='ne')

        Break_Return_Performance = tk.Label(Break_Frame,text=Break_Performance[0],font=('Arial',14),fg='black',background="#F0F0F0")
        Break_Return_Performance.place(x=980,y=100,anchor='ne')

        Break_Volatility_Performance = tk.Label(Break_Frame,text=Break_Performance[1],font=('Arial',14),fg='black',background="#F0F0F0")
        Break_Volatility_Performance.place(x=980,y=140,anchor='ne')

        Break_Sharpe_Performance = tk.Label(Break_Frame,text=Break_Performance[2],font=('Arial',14),fg='black',background="#F0F0F0")
        Break_Sharpe_Performance.place(x=980,y=180,anchor='ne')

        Break_Turnover_Performance = tk.Label(Break_Frame,text=Break_Performance[4],font=('Arial',14),fg='black',background="#F0F0F0")
        Break_Turnover_Performance.place(x=980,y=220,anchor='ne')

        Break_Expect_Performance = tk.Label(Break_Frame,text=Break_Performance[5],font=('Arial',14),fg='black',background="#F0F0F0")
        Break_Expect_Performance.place(x=980,y=260,anchor='ne')

        Break_Trading = ttk.Treeview(Break_Frame,height="5")

        Break_Trading["columns"]=("column1","column2",'column3')
        Break_Trading.column("#0",width=10, anchor='e')
        Break_Trading.column("column1", width=80, anchor='center' )
        Break_Trading.column("column2", width=60, anchor='center')
        Break_Trading.column("column3", width=80 , anchor='center')
        
        
        Break_Trading.heading('#0', text='')
        Break_Trading.heading("column1", text="Date")
        Break_Trading.heading("column2", text="Signal")
        Break_Trading.heading("column3", text="Price")
        
        Break_Trading.place(x=750, y=320)


        Break_Trading.delete(*Break_Trading.get_children())    
        
        Break_table = Break_Performance[6]
        
        for i in range(5):
            Break_Trading.insert("",i,text=str(i),values=(str(Break_table.index[i])[:10],Break_table['Sign'][i],Break_table['Total'][i]))





        #======================================================================
        MTM_Frame=tk.Frame(frame, width=1000, height=500, background="#F0F0F0")
        MTM_Frame.grid(row=4,column=1)
        
        MTM_Label = tk.Label(MTM_Frame,text='MTM',font=('Arial',18,'bold'),fg='black',background="#F0F0F0")
        MTM_Label.place(x=750,y=15)

        MTM_text = str(MTM[0])
        MTM_Parameter_Label = tk.Label(MTM_Frame,text=MTM_text,font=('Arial',18,'bold'),fg='black',background="#F0F0F0")
        MTM_Parameter_Label.place(x=980,y=15,anchor='ne')


        ttk.Separator(MTM_Frame,orient=tk.HORIZONTAL).place(x=750,y=50,width=300)





        MTM_Win_Label = tk.Label(MTM_Frame,text='Win(%)',font=('Arial',14),fg='black',background="#F0F0F0")
        MTM_Win_Label.place(x=750,y=60)

        MTM_Return_Label = tk.Label(MTM_Frame,text='Annual Return',font=('Arial',14),fg='black',background="#F0F0F0")
        MTM_Return_Label.place(x=750,y=100)

        MTM_Volatility_Label = tk.Label(MTM_Frame,text='Volatility',font=('Arial',14),fg='black',background="#F0F0F0")
        MTM_Volatility_Label.place(x=750,y=140)

        MTM_Sharpe_Label = tk.Label(MTM_Frame,text='Sharpe',font=('Arial',14),fg='black',background="#F0F0F0")
        MTM_Sharpe_Label.place(x=750,y=180)

        MTM_Turnover_Label = tk.Label(MTM_Frame,text='Turnover',font=('Arial',14),fg='black',background="#F0F0F0")
        MTM_Turnover_Label.place(x=750,y=220)

        MTM_Expect_Label = tk.Label(MTM_Frame,text='Expected Return',font=('Arial',14),fg='black',background="#F0F0F0")
        MTM_Expect_Label.place(x=750,y=260)


        #Performance
        MTM_Performance = MTM_Performance(Data = Data,para1 = MTM[0])
        
        MTM_Win_Performance = tk.Label(MTM_Frame,text=MTM_Performance[3],font=('Arial',14),fg='black',background="#F0F0F0")
        MTM_Win_Performance.place(x=980,y=60,anchor='ne')

        MTM_Return_Performance = tk.Label(MTM_Frame,text=MTM_Performance[0],font=('Arial',14),fg='black',background="#F0F0F0")
        MTM_Return_Performance.place(x=980,y=100,anchor='ne')

        MTM_Volatility_Performance = tk.Label(MTM_Frame,text=MTM_Performance[1],font=('Arial',14),fg='black',background="#F0F0F0")
        MTM_Volatility_Performance.place(x=980,y=140,anchor='ne')

        MTM_Sharpe_Performance = tk.Label(MTM_Frame,text=MTM_Performance[2],font=('Arial',14),fg='black',background="#F0F0F0")
        MTM_Sharpe_Performance.place(x=980,y=180,anchor='ne')

        MTM_Turnover_Performance = tk.Label(MTM_Frame,text=MTM_Performance[4],font=('Arial',14),fg='black',background="#F0F0F0")
        MTM_Turnover_Performance.place(x=980,y=220,anchor='ne')

        MTM_Expect_Performance = tk.Label(MTM_Frame,text=MTM_Performance[5],font=('Arial',14),fg='black',background="#F0F0F0")
        MTM_Expect_Performance.place(x=980,y=260,anchor='ne')

        MTM_Trading = ttk.Treeview(MTM_Frame,height="5")

        MTM_Trading["columns"]=("column1","column2",'column3')
        MTM_Trading.column("#0",width=10, anchor='e')
        MTM_Trading.column("column1", width=80, anchor='center' )
        MTM_Trading.column("column2", width=60, anchor='center')
        MTM_Trading.column("column3", width=80 , anchor='center')
        
        
        MTM_Trading.heading('#0', text='')
        MTM_Trading.heading("column1", text="Date")
        MTM_Trading.heading("column2", text="Signal")
        MTM_Trading.heading("column3", text="Price")
        
        MTM_Trading.place(x=750, y=320)


        MTM_Trading.delete(*MTM_Trading.get_children())    
        
        MTM_table = MTM_Performance[6]
        
        for i in range(5):
            MTM_Trading.insert("",i,text=str(i),values=(str(MTM_table.index[i])[:10],MTM_table['Sign'][i],MTM_table['Total'][i]))




        #======================================================================
        ATR_Frame=tk.Frame(frame, width=1000, height=500, background="#F0F0F0")
        ATR_Frame.grid(row=5,column=1)
        
        ATR_Label = tk.Label(ATR_Frame,text='ATR',font=('Arial',18,'bold'),fg='black',background="#F0F0F0")
        ATR_Label.place(x=750,y=15)

        ATR_text = str(ATR[0])+','+str(ATR[1])
        ATR_Parameter_Label = tk.Label(ATR_Frame,text=ATR_text,font=('Arial',18,'bold'),fg='black',background="#F0F0F0")
        ATR_Parameter_Label.place(x=980,y=15,anchor='ne')


        ttk.Separator(ATR_Frame,orient=tk.HORIZONTAL).place(x=750,y=50,width=300)


        ATR_Win_Label = tk.Label(ATR_Frame,text='Win(%)',font=('Arial',14),fg='black',background="#F0F0F0")
        ATR_Win_Label.place(x=750,y=60)

        ATR_Return_Label = tk.Label(ATR_Frame,text='Annual Return',font=('Arial',14),fg='black',background="#F0F0F0")
        ATR_Return_Label.place(x=750,y=100)

        ATR_Volatility_Label = tk.Label(ATR_Frame,text='Volatility',font=('Arial',14),fg='black',background="#F0F0F0")
        ATR_Volatility_Label.place(x=750,y=140)

        ATR_Sharpe_Label = tk.Label(ATR_Frame,text='Sharpe',font=('Arial',14),fg='black',background="#F0F0F0")
        ATR_Sharpe_Label.place(x=750,y=180)

        ATR_Turnover_Label = tk.Label(ATR_Frame,text='Turnover',font=('Arial',14),fg='black',background="#F0F0F0")
        ATR_Turnover_Label.place(x=750,y=220)

        ATR_Expect_Label = tk.Label(ATR_Frame,text='Expected Return',font=('Arial',14),fg='black',background="#F0F0F0")
        ATR_Expect_Label.place(x=750,y=260)


        #Performance
        ATR_Performance = ATR_Performance(Data = Data,para1 = ATR[0],para2 = ATR[1])
        
        ATR_Win_Performance = tk.Label(ATR_Frame,text=ATR_Performance[3],font=('Arial',14),fg='black',background="#F0F0F0")
        ATR_Win_Performance.place(x=980,y=60,anchor='ne')

        ATR_Return_Performance = tk.Label(ATR_Frame,text=ATR_Performance[0],font=('Arial',14),fg='black',background="#F0F0F0")
        ATR_Return_Performance.place(x=980,y=100,anchor='ne')

        ATR_Volatility_Performance = tk.Label(ATR_Frame,text=ATR_Performance[1],font=('Arial',14),fg='black',background="#F0F0F0")
        ATR_Volatility_Performance.place(x=980,y=140,anchor='ne')

        ATR_Sharpe_Performance = tk.Label(ATR_Frame,text=ATR_Performance[2],font=('Arial',14),fg='black',background="#F0F0F0")
        ATR_Sharpe_Performance.place(x=980,y=180,anchor='ne')

        ATR_Turnover_Performance = tk.Label(ATR_Frame,text=ATR_Performance[4],font=('Arial',14),fg='black',background="#F0F0F0")
        ATR_Turnover_Performance.place(x=980,y=220,anchor='ne')

        ATR_Expect_Performance = tk.Label(ATR_Frame,text=ATR_Performance[5],font=('Arial',14),fg='black',background="#F0F0F0")
        ATR_Expect_Performance.place(x=980,y=260,anchor='ne')

        ATR_Trading = ttk.Treeview(ATR_Frame,height="5")

        ATR_Trading["columns"]=("column1","column2",'column3')
        ATR_Trading.column("#0",width=10, anchor='e')
        ATR_Trading.column("column1", width=80, anchor='center' )
        ATR_Trading.column("column2", width=60, anchor='center')
        ATR_Trading.column("column3", width=80 , anchor='center')
        
        
        ATR_Trading.heading('#0', text='')
        ATR_Trading.heading("column1", text="Date")
        ATR_Trading.heading("column2", text="Signal")
        ATR_Trading.heading("column3", text="Price")
        
        ATR_Trading.place(x=750, y=320)


        ATR_Trading.delete(*ATR_Trading.get_children())    
        
        ATR_table = ATR_Performance[6]
        
        for i in range(5):
            ATR_Trading.insert("",i,text=str(i),values=(str(ATR_table.index[i])[:10],ATR_table['Sign'][i],ATR_table['Total'][i]))






        #RSI===================================================================
        fig_RSI = Figure(figsize=(6,4), dpi=120)
        fig_RSI.set_tight_layout(True)
        fig_RSI.patch.set_facecolor('#F0F0F0')
        
        RSI_Chart = fig_RSI.add_subplot(111,axisbg='#F0F0F0')
        RSI_Chart.tick_params(axis='both', which='major', labelsize=8)
        
        RSI_Chart.tick_params(axis='both', which='major',colors='black', labelsize=6)
        RSI_Chart.spines['bottom'].set_color('black')
        RSI_Chart.spines['top'].set_color('black')
        RSI_Chart.spines['left'].set_color('black')
        RSI_Chart.spines['right'].set_color('black')
        RSI_Chart.xaxis.label.set_color('black')
        RSI_Chart.yaxis.label.set_color('black')
        
        
        canvas_RSI = FigureCanvasTkAgg(fig_RSI, master=RSI_Frame)
        canvas_RSI.get_tk_widget().place(x=0,y=0)
        canvas_RSI.get_tk_widget().configure(background='#F0F0F0',  highlightcolor='#F0F0F0', highlightbackground='#F0F0F0')

        '''
        Stock_Chart.clear()
        Stock_Chart.plot( STOCK.index, STOCK[Ticker],lw=1,color='#01485E',alpha=0.5)
        canvas1.show()

        '''
        
        #KD====================================================================
        fig_KD = Figure(figsize=(6,4), dpi=120)
        fig_KD.set_tight_layout(True)
        fig_KD.patch.set_facecolor('#F0F0F0')
        
        KD_Chart = fig_KD.add_subplot(111,axisbg='#F0F0F0')
        KD_Chart.tick_params(axis='both', which='major', labelsize=8)
        
        KD_Chart.tick_params(axis='both', which='major',colors='black', labelsize=6)
        KD_Chart.spines['bottom'].set_color('black')
        KD_Chart.spines['top'].set_color('black')
        KD_Chart.spines['left'].set_color('black')
        KD_Chart.spines['right'].set_color('black')
        KD_Chart.xaxis.label.set_color('black')
        KD_Chart.yaxis.label.set_color('black')
        
        
        canvas_KD = FigureCanvasTkAgg(fig_KD, master=KD_Frame)
        canvas_KD.get_tk_widget().place(x=0,y=0)
        canvas_KD.get_tk_widget().configure(background='#F0F0F0',  highlightcolor='#F0F0F0', highlightbackground='#F0F0F0')


        #EMA====================================================================
        fig_EMA = Figure(figsize=(6,4), dpi=120)
        fig_EMA.set_tight_layout(True)
        fig_EMA.patch.set_facecolor('#F0F0F0')
        
        EMA_Chart = fig_EMA.add_subplot(111,axisbg='#F0F0F0')
        EMA_Chart.tick_params(axis='both', which='major', labelsize=8)
        
        EMA_Chart.tick_params(axis='both', which='major',colors='black', labelsize=6)
        EMA_Chart.spines['bottom'].set_color('black')
        EMA_Chart.spines['top'].set_color('black')
        EMA_Chart.spines['left'].set_color('black')
        EMA_Chart.spines['right'].set_color('black')
        EMA_Chart.xaxis.label.set_color('black')
        EMA_Chart.yaxis.label.set_color('black')
        
        canvas_EMA = FigureCanvasTkAgg(fig_EMA, master=EMA_Frame)
        canvas_EMA.get_tk_widget().place(x=0,y=0)
        canvas_EMA.get_tk_widget().configure(background='#F0F0F0',  highlightcolor='#F0F0F0', highlightbackground='#F0F0F0')


        #Break=================================================================
        fig_Break = Figure(figsize=(6,4), dpi=120)
        fig_Break.set_tight_layout(True)
        fig_Break.patch.set_facecolor('#F0F0F0')
        
        Break_Chart = fig_Break.add_subplot(111,axisbg='#F0F0F0')
        Break_Chart.tick_params(axis='both', which='major', labelsize=8)
        
        Break_Chart.tick_params(axis='both', which='major',colors='black', labelsize=6)
        Break_Chart.spines['bottom'].set_color('black')
        Break_Chart.spines['top'].set_color('black')
        Break_Chart.spines['left'].set_color('black')
        Break_Chart.spines['right'].set_color('black')
        Break_Chart.xaxis.label.set_color('black')
        Break_Chart.yaxis.label.set_color('black')
        
        canvas_Break = FigureCanvasTkAgg(fig_Break, master=Break_Frame)
        canvas_Break.get_tk_widget().place(x=0,y=0)
        canvas_Break.get_tk_widget().configure(background='#F0F0F0',  highlightcolor='#F0F0F0', highlightbackground='#F0F0F0')

        #MTM===================================================================
        fig_MTM = Figure(figsize=(6,4), dpi=120)
        fig_MTM.set_tight_layout(True)
        fig_MTM.patch.set_facecolor('#F0F0F0')
        
        MTM_Chart = fig_MTM.add_subplot(111,axisbg='#F0F0F0')
        MTM_Chart.tick_params(axis='both', which='major', labelsize=8)
        
        MTM_Chart.tick_params(axis='both', which='major',colors='black', labelsize=6)
        MTM_Chart.spines['bottom'].set_color('black')
        MTM_Chart.spines['top'].set_color('black')
        MTM_Chart.spines['left'].set_color('black')
        MTM_Chart.spines['right'].set_color('black')
        MTM_Chart.xaxis.label.set_color('black')
        MTM_Chart.yaxis.label.set_color('black')
        
        canvas_MTM = FigureCanvasTkAgg(fig_MTM, master=MTM_Frame)
        canvas_MTM.get_tk_widget().place(x=0,y=0)
        canvas_MTM.get_tk_widget().configure(background='#F0F0F0',  highlightcolor='#F0F0F0', highlightbackground='#F0F0F0')

        #ATR===================================================================
        fig_ATR = Figure(figsize=(6,4), dpi=120)
        fig_ATR.set_tight_layout(True)
        fig_ATR.patch.set_facecolor('#F0F0F0')
        
        ATR_Chart = fig_ATR.add_subplot(111,axisbg='#F0F0F0')
        ATR_Chart.tick_params(axis='both', which='major', labelsize=8)
        
        ATR_Chart.tick_params(axis='both', which='major',colors='black', labelsize=6)
        ATR_Chart.spines['bottom'].set_color('black')
        ATR_Chart.spines['top'].set_color('black')
        ATR_Chart.spines['left'].set_color('black')
        ATR_Chart.spines['right'].set_color('black')
        ATR_Chart.xaxis.label.set_color('black')
        ATR_Chart.yaxis.label.set_color('black')
        
        canvas_ATR = FigureCanvasTkAgg(fig_ATR, master=ATR_Frame)
        canvas_ATR.get_tk_widget().place(x=0,y=0)
        canvas_ATR.get_tk_widget().configure(background='#F0F0F0',  highlightcolor='#F0F0F0', highlightbackground='#F0F0F0')

        #======================================================================
        
        RSI_Chart.clear()
        KD_Chart.clear()
        EMA_Chart.clear()
        Break_Chart.clear()
        MTM_Chart.clear()
        ATR_Chart.clear()
        
        #Buy & Hold Chart======================================================
        
        RSI_Chart.plot( Data.index, Data['Cumulative_Return'],lw=1,color='#01485E',alpha=0.5, label='Buy&Hold')
        KD_Chart.plot( Data.index, Data['Cumulative_Return'],lw=1,color='#01485E',alpha=0.5, label='Buy&Hold')
        EMA_Chart.plot( Data.index, Data['Cumulative_Return'],lw=1,color='#01485E',alpha=0.5, label='Buy&Hold')
        Break_Chart.plot( Data.index, Data['Cumulative_Return'],lw=1,color='#01485E',alpha=0.5, label='Buy&Hold')
        MTM_Chart.plot( Data.index, Data['Cumulative_Return'],lw=1,color='#01485E',alpha=0.5, label='Buy&Hold')
        ATR_Chart.plot( Data.index, Data['Cumulative_Return'],lw=1,color='#01485E',alpha=0.5, label='Buy&Hold')
        
        def RSI_calculate(prices,n):
            n = int(n)
            deltas = np.diff(prices)
            seed = deltas[:n+1]
            up = seed[seed>=0].sum()/n
            down = -seed[seed<0].sum()/n
            rs = up/down
            rsi = np.zeros_like(prices)
            rsi[:n]=100. - 100./(1.+rs)
                
            for i in range(n,len(prices)):
                delta = deltas[i-1]    #cause the diff is 1 shorter
                    
                if delta>0:
                    upval = delta
                    downval = 0.
                else:
                    upval = 0.
                    downval = -delta
                        
                up = (up*(n-1) + upval)/n
                down = (down*(n-1) + downval)/n
                    
                rs = up/down
                rsi[i] = 100. - 100./(1.+rs)
                
            return rsi
        

        #Strategy==============================================================
        #RSI Equity Curve
        Data['RSI_1'] = RSI_calculate(Data[str(Ticker)],RSI[0])
        Data['RSI_2'] = RSI_calculate(Data[str(Ticker)],RSI[1])
        
        Data['diff'] = Data['RSI_1'] - Data['RSI_2']
        Data['diff'] = np.where(Data['diff']==0,Data['diff'].shift(1)*(-1),Data['diff'])
        
        Data['RSI_Signal'] = np.where(Data['diff'] > 0.,1,0)
        Data['RSI_Return'] = Data['Return']*Data['RSI_Signal'].shift(1)
        Data['RSI_Equity_Curve'] = Data['RSI_Return'].cumsum()

        RSI_Chart.plot( Data.index, Data['RSI_Equity_Curve'],lw=2,color='#00A3DC',alpha=0.5, label='Equity Curve')
        #KD Equity Curve
        Data['K'] = K(Data[str(Ticker)],KD[0])
        Data['D'] = D(Data['K'],KD[1])
        
        Data['diff'] = Data['K'] - Data['D']
        Data['diff'] = np.where(Data['diff']==0,Data['diff'].shift(1)*(-1),Data['diff'])
        
        Data['KD_Signal'] = np.where(Data['diff'] > 0.,1,0)
        Data['KD_Return'] = Data['Return']*Data['KD_Signal'].shift(1)
        Data['KD_Equity_Curve'] = Data['KD_Return'].cumsum()
        KD_Chart.plot( Data.index, Data['KD_Equity_Curve'],lw=2,color='#00A3DC',alpha=0.5, label='Equity Curve')




        #EMA Equity Curve
        Data['EMA_1'] = pd.ewma(Data[Ticker],EMA[0])
        Data['EMA_2'] = pd.ewma(Data[Ticker],EMA[1])
        
        Data['EMA_Signal'] = np.where(Data['EMA_1'] - Data['EMA_2']>0.,1,0)
        Data['EMA_Return'] = Data['Return']*Data['EMA_Signal']
        Data['EMA_EquityCurve'] = Data['EMA_Return'].cumsum()#.apply(np.exp)
        
        EMA_Chart.plot( Data.index, Data['EMA_EquityCurve'],lw=2,color='#00A3DC',alpha=0.5, label='Equity Curve')
        
        
        
        #Break Equity Curve
        Data['High'] = pd.rolling_max(Data[str(Ticker)],int(Break[0])).shift(1)
        Data['Low'] = pd.rolling_min(Data[str(Ticker)],int(Break[1])).shift(1)
        
        Data['Break_Signal'] = np.where(Data[str(Ticker)]>Data['High'],1,np.where(Data[str(Ticker)]<Data['Low'],0,np.nan))
        Data['Break_Signal'] = Data['Break_Signal'].fillna(method='ffill')
        Data['Break_Return'] = Data['Return']*Data['Break_Signal'].shift(1)
        Data['Break_Equity_Curve'] = Data['Break_Return'].cumsum()#.apply(np.exp)
        
        Break_Chart.plot( Data.index, Data['Break_Equity_Curve'],lw=2,color='#00A3DC',alpha=0.5, label='Equity Curve')
        
        #MTM Equity Curve
        Data['MTM'] = pd.ewma(Data[str(Ticker)],MTM[0])
        
        Data['MTM_diff'] = Data[Ticker] - Data['MTM']
        Data['MTM_diff'] = np.where(Data['MTM_diff']==0,Data['MTM_diff'].shift(1)*(-1),Data['MTM_diff'])
        
        Data['MTM_Signal'] = np.where(Data['MTM_diff'] > 0.,1,0)
        Data['MTM_Return'] = Data['Return']*Data['MTM_Signal'].shift(1)
        Data['MTM_Equity_Curve'] = Data['MTM_Return'].cumsum()#.apply(np.exp)
        MTM_Chart.plot( Data.index, Data['MTM_Equity_Curve'],lw=2,color='#00A3DC',alpha=0.5, label='Equity Curve')
        
        #ATR Equity Curve
        Data['ATR_EMA'] = pd.ewma(Data[str(Ticker)],ATR[0])
        Data['ATR'] = pd.ewma(abs(Data[Ticker].diff()),ATR[0])
        Data['UpBound'] = Data['ATR_EMA']+ATR[1]*Data['ATR']
        #Data['LowBound'] = Data['ATR_EMA']-para2*Data['ATR']
        
        Data['ATR_Signal'] = np.where(Data[Ticker]>Data['UpBound'],1,0)
        Data['ATR_Return'] = Data['Return']*Data['ATR_Signal'].shift(1)
        Data['ATR_Equity_Curve'] = Data['ATR_Return'].cumsum()#.apply(np.exp)
        
        ATR_Chart.plot( Data.index, Data['ATR_Equity_Curve'],lw=2,color='#00A3DC',alpha=0.5, label='Equity Curve')
    
        
        
        #======================================================================
        
        RSI_Chart.legend(loc=2, shadow=True,fontsize=8).get_frame().set_facecolor('#F0F0F0')
        KD_Chart.legend(loc=2, shadow=True,fontsize=8).get_frame().set_facecolor('#F0F0F0')
        EMA_Chart.legend(loc=2, shadow=True,fontsize=8).get_frame().set_facecolor('#F0F0F0')
        Break_Chart.legend(loc=2, shadow=True,fontsize=8).get_frame().set_facecolor('#F0F0F0')
        MTM_Chart.legend(loc=2, shadow=True,fontsize=8).get_frame().set_facecolor('#F0F0F0')
        ATR_Chart.legend(loc=2, shadow=True,fontsize=8).get_frame().set_facecolor('#F0F0F0')
        
        RSI_Chart.set_ylabel('Log Scale Return', color='black', fontsize=8)
        KD_Chart.set_ylabel('Log Scale Return', color='black', fontsize=8)
        EMA_Chart.set_ylabel('Log Scale Return', color='black', fontsize=8)
        Break_Chart.set_ylabel('Log Scale Return', color='black', fontsize=8)
        MTM_Chart.set_ylabel('Log Scale Return', color='black', fontsize=8)
        ATR_Chart.set_ylabel('Log Scale Return', color='black', fontsize=8)
        
        
        
        canvas_RSI.show()
        canvas_KD.show()
        canvas_EMA.show()
        canvas_Break.show()
        canvas_MTM.show()
        canvas_ATR.show()

        
        Backtest_root.mainloop()

        





if __name__ == "__main__":
    warnings.simplefilter(action = "ignore", category = FutureWarning)
    app = Market_Timing()
    #app.mainloop()
    #print 'Fucking'
