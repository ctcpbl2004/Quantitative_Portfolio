# -*- coding: utf-8 -*-
"""
Created on Mon Jul 25 11:23:54 2016

@author: T105041
"""

import Tkinter as tk
import ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib import cm
import numpy as np
import re
# import Database_Functions
import pandas as pd
import sqlite3
import datetime
import xlsxwriter
import warnings
import matplotlib.pyplot as plt


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
   ''





   
        for each in data:
            cursor.execute("INSERT OR REPLACE INTO Update_table VALUES (?,?,?,?,?)",
                           (each[0], each[1], each[2], each[3], each[4]))
        connection.commit()
        connection.close()

    '''
    #==============================================================================
    #Delete data
    #==============================================================================
    '''

    @staticmethod
    def Delete_data(Ticker, Name):
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
        cursor = connection.cursor()
        cursor.execute("DELETE FROM Time_Series WHERE Ticker=? AND Name=?", (Ticker, Name))
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
        return Ticker_list, Name_list

    @staticmethod
    def Word_code(n):
        Word_dict = {1: 'A', 2: 'B', 3: 'C', 4: 'D', 5: 'E', 6: 'F', 7: 'G', 8: 'H', 9: 'I', 10: 'J',
                     11: 'K', 12: 'L', 13: 'M', 14: 'N', 15: 'O', 16: 'P', 17: 'Q', 18: 'R', 19: 'S', 20: 'T',
                     21: 'U', 22: 'V', 23: 'W', 24: 'X', 25: 'Y', 26: 'Z'}

        error_list = range(0, 500, 26)
        error_list[0] = 1

        if (n <= 26.):
            return Word_dict[n]

        elif (n > 26) & (n <= 26 ** 2):
            try:
                second_digit = int(n / 26)
                first_digit = n % 26
                return Word_dict[second_digit] + Word_dict[first_digit]
            except:
                first_digit = 26
                second_digit = second_digit - 1
                return Word_dict[second_digit] + Word_dict[first_digit]

    @staticmethod
    def Bloomberg_Update(Tickers, Start, rows):
        Start = Start.replace('-', '/')
        workbook = xlsxwriter.Workbook('Update.xlsx')
        worksheet = workbook.add_worksheet()

        NumberOfTicker = len(Tickers)

        # Write Tickers
        for i in range(NumberOfTicker):
            worksheet.write(0, i + 1, Tickers[i])

        for j in range(NumberOfTicker - 1):
            Cell = str(Database_Functions.Word_code(j + 3)) + '1'
            Bloomberg_formula = '=BDH(' + Cell + ',"PX_LAST","' + Start + '","","Dir=V","Dts=H","Sort=D","Quote=C","QtTyp=Y","Days=A","Per=cd","DtFmt=D","Fill=B","UseDPDF=N","CshAdjNormal=Y","CshAdjAbnormal=Y","CapChg=Y","cols=1;rows=' + str(
                rows) + '")'
            worksheet.write_formula(1, 2 + j, Bloomberg_formula)

        Bloomberg_formula_date = '=BDH(B1,"PX_LAST","' + Start + '","","Dir=V","Dts=S","Sort=D","Quote=C","QtTyp=Y","Days=A","Per=cd","DtFmt=D","Fill=B","UseDPDF=N","CshAdjNormal=Y","CshAdjAbnormal=Y","CapChg=Y","cols=2;rows=' + str(
            rows) + '")'
        worksheet.write_formula(1, 0, Bloomberg_formula_date)
        workbook.close()

    @staticmethod
    def Create_Update_xlsx(Start):
        Tickers = Database_Functions.Fetch_All_Tickers()[0]
        Database_Functions.Bloomberg_Update(Tickers, Start, rows=500)

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
        cursor.execute("SELECT Name FROM Update_table WHERE Ticker =?", (Ticker,))
        data = cursor.fetchall()
        connection.close()
        return data[0][0]

    @staticmethod
    def Series_to_Dataframe(Ticker, Name, Series):
        df = pd.DataFrame(columns=['Ticker', 'Name', 'Date', 'Value'])
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
            cursor.execute("INSERT OR REPLACE INTO Time_Series VALUES (?, ?, ?, ?)", (Ticker, Name, Date, Value))
            print Date, Value
        connection.commit()
        connection.close()

    @staticmethod
    def Data_to_db(File_name):
        df = pd.read_excel(File_name).ix[1:]
        df = df.sort_index()
        Commodities_list = df.columns
        # print df

        for Ticker in Commodities_list:
            try:
                Name = Database_Functions.Find_Name(Ticker)
                print Name
                Data = df[str(Ticker)].dropna()
                Data_df = Database_Functions.Series_to_Dataframe(Ticker, Name, Series=Data)
                Database_Functions.Dataframe_to_db(Data_df)
            except:
                pass

    '''
    ===============================================================================
    Fetch Single Ticker Data
    ===============================================================================
    '''

    @staticmethod
    def Fetch(Ticker, Start=None, End=None):
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
        cursor = connection.cursor()

        if (Start == None) & (End == None):
            cursor.execute("SELECT Date,Value FROM Time_Series WHERE Ticker = ?", (Ticker,))
        elif End == None:
            cursor.execute("SELECT Date,Value FROM Time_Series WHERE Ticker = ? AND Date > ?", (Ticker, Start))
        elif Start == None:
            cursor.execute("SELECT Date,Value FROM Time_Series WHERE Ticker = ? AND Date < ?", (Ticker, End))
        elif (Start != None) & (End != None):
            cursor.execute("SELECT Date,Value FROM Time_Series WHERE Ticker = ? AND Date BETWEEN ? AND ? ",
                           (Ticker, Start, End))
        else:
            print 'TIMEZONE ERROR!!!!'

        data = cursor.fetchall()
        connection.close()

        Date_list = []
        Value_list = []

        for row in data:
            Date_list.append(row[0])
            Value_list.append(row[1])

        df = pd.DataFrame(Value_list, index=Date_list, columns=[str(Ticker)])
        df.index = df.index.to_datetime()
        df.sort_index()

        return df

    '''
    ===============================================================================
    Add New Ticker and Data
    ===============================================================================
    '''

    @staticmethod
    def Add_Index(File_name, Ticker, Name):
        Data = pd.read_excel(File_name).ix[1:].dropna()
        print Data
        df = Database_Functions.Series_to_Dataframe(Ticker, Name, Series=Data[str(Data.columns[0])])
        Database_Functions.Dataframe_to_db(df)

    # Add_Index(File_name='Test.xlsx',Ticker='AUD Curncy',Name='AUD')

    @staticmethod
    def Add_Index_to_Strategy_table(Ticker, Name):
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
        cursor = connection.cursor()
        cursor.execute("INSERT OR REPLACE INTO Strategy_table VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                       (Ticker, Name, '', '', '', '', '', ''))
        connection.commit()
        connection.close()

    # Delete_data(Ticker = 'AUD',Name = 'AUDUSD')
    @staticmethod
    def Delete_Index_to_Strategy_table(Ticker, Name):
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
        cursor = connection.cursor()
        cursor.execute("DELETE FROM Strategy_table WHERE Ticker=? AND Name=?", (Ticker, Name))
        connection.commit()
        connection.close()
        Database_Functions.Fetch_All_Data()


class AutocompleteEntry(tk.Entry):
    def __init__(self, lista, *args, **kwargs):

        tk.Entry.__init__(self, *args, **kwargs)
        self.lista = lista
        self.var = self["textvariable"]
        if self.var == '':
            self.var = self["textvariable"] = tk.StringVar()

        self.var.trace('w', self.changed)
        self.bind("<Right>", self.selection)
        self.bind("<Up>", self.up)
        self.bind("<Down>", self.down)

        self.lb_up = False

    def changed(self, name, index, mode):

        if self.var.get() == '':
            self.lb.destroy()
            self.lb_up = False
        else:
            words = self.comparison()
            if words:
                if not self.lb_up:
                    self.lb = tk.Listbox(font=('Arial', 10), width=55, background="#181818", fg="white",
                                         selectforeground='black',
                                         selectbackground="#FF9C29", highlightcolor="#181818", activestyle=tk.NONE)
                    self.lb.bind("<Double-Button-1>", self.selection)
                    self.lb.bind("<Right>", self.selection)
                    self.lb.place(x=self.winfo_x(), y=self.winfo_y() + self.winfo_height() + 70)
                    self.lb_up = True

                self.lb.delete(0, tk.END)
                for w in words:
                    self.lb.insert(tk.END, w)
            else:
                if self.lb_up:
                    self.lb.destroy()
                    self.lb_up = False

    def selection(self, event):

        if self.lb_up:
            self.var.set(self.lb.get(tk.ACTIVE))
            self.lb.destroy()
            self.lb_up = False
            self.icursor(tk.END)

    def up(self, event):

        if self.lb_up:
            if self.lb.curselection() == ():
                index = '0'
            else:
                index = self.lb.curselection()[0]
            if index != '0':
                self.lb.selection_clear(first=index)
                index = str(int(index) - 1)
                self.lb.selection_set(first=index)
                self.lb.activate(index)

    def down(self, event):

        if self.lb_up:
            if self.lb.curselection() == ():
                index = '0'
            else:
                index = self.lb.curselection()[0]
            if index != tk.END:
                self.lb.selection_clear(first=index)
                index = str(int(index) + 1)
                self.lb.selection_set(first=index)
                self.lb.activate(index)

    def comparison(self):
        pattern = re.compile('.*' + self.var.get() + '.*')
        return [w for w in self.lista if re.match(pattern, w)]


class GUI(tk.Tk):
    def __init__(self):

        self.Today = datetime.datetime.today()
        self.Today_str = self.Today.strftime('%Y-%m-%d')

        self.root = tk.Tk()
        self.root.wm_title('Portfolio Manager')
        self.root.geometry('1250x660')#1350x760
        self.root.iconbitmap('D:/Taishin_Platform/pics/chart_diagram_analytics_business_flat_icon-512.ico')
        # ======================================================================
        self.Frame_Top = tk.Frame(self.root, width=2000, height=70, background="#282828")
        self.Frame_Top.place(x=0, y=0)

        self.Frame_Down = tk.Frame(self.root, width=1400, height=700, background="#F0F0F0")
        self.Frame_Down.place(x=0, y=70)

        # ======================================================================
        self.Overview_frame = tk.Frame(self.Frame_Down, width=1400, height=1300, background='#F0F0F0')
        self.Portfolio1_frame = tk.Frame(self.Frame_Down, width=1400, height=1300, background='#F0F0F0')
        self.Portfolio2_frame = tk.Frame(self.Frame_Down, width=1400, height=1300, background='#F0F0F0')
        self.Portfolio3_frame = tk.Frame(self.Frame_Down, width=1400, height=1300, background='#F0F0F0')
        self.Portfolio4_frame = tk.Frame(self.Frame_Down, width=1400, height=1300, background='#F0F0F0')
        self.Portfolio5_frame = tk.Frame(self.Frame_Down, width=1400, height=1300, background='#F0F0F0')
        self.Portfolio6_frame = tk.Frame(self.Frame_Down, width=1400, height=1300, background='#F0F0F0')
        self.Portfolio7_frame = tk.Frame(self.Frame_Down, width=1400, height=1300, background='#F0F0F0')
        self.Portfolio8_frame = tk.Frame(self.Frame_Down, width=1400, height=1300, background='#F0F0F0')
        self.Portfolio9_frame = tk.Frame(self.Frame_Down, width=1400, height=1300, background='#F0F0F0')

        self.Overview_frame.place(x=0, y=0)
        self.Portfolio1_frame.place(x=0, y=0)
        self.Portfolio2_frame.place(x=0, y=0)
        self.Portfolio3_frame.place(x=0, y=0)
        self.Portfolio4_frame.place(x=0, y=0)
        self.Portfolio5_frame.place(x=0, y=0)
        self.Portfolio6_frame.place(x=0, y=0)
        self.Portfolio7_frame.place(x=0, y=0)
        self.Portfolio8_frame.place(x=0, y=0)
        self.Portfolio9_frame.place(x=0, y=0)

        self.Overview_Button = tk.Button(self.Frame_Top, text='Overview', relief='groove', fg='white',
                                         background="#181818", width=20,
                                         command=lambda: self.Raise_frame(self.Overview_frame))
        self.Button1 = tk.Button(self.Frame_Top, text='Aggressive', relief='groove', fg='white', background="#181818",
                                 width=20, command=lambda: self.Raise_frame(self.Portfolio1_frame))
        self.Button2 = tk.Button(self.Frame_Top, text='Moderate', relief='groove', fg='white', background="#181818",
                                 width=20, command=lambda: self.Raise_frame(self.Portfolio2_frame))
        self.Button3 = tk.Button(self.Frame_Top, text='Conservative', relief='groove', fg='white', background="#181818",
                                 width=20, command=lambda: self.Raise_frame(self.Portfolio3_frame))
        # self.Button4 = tk.Button(self.Frame_Top,text='Aggregate',relief='groove',fg='white',background="#181818",width=20, command=lambda:self.Raise_frame(self.Portfolio4_frame))
        self.Button5 = tk.Button(self.Frame_Top, text='Global Strategic Bond', relief='groove', fg='white',
                                 background="#181818", width=20,
                                 command=lambda: self.Raise_frame(self.Portfolio5_frame))
        self.Button6 = tk.Button(self.Frame_Top, text='Global Momentum', relief='groove', fg='white',
                                 background="#181818", width=20,
                                 command=lambda: self.Raise_frame(self.Portfolio6_frame))
        self.Button7 = tk.Button(self.Frame_Top, text='Low Volatility Portfolio (Equity & Fixed Income)', relief='groove', fg='white', background="#181818",
                                 width=63, command=lambda: self.Raise_frame(self.Portfolio7_frame))
        self.Button8 = tk.Button(self.Frame_Top, text='6 Portfolio', relief='groove', fg='white', background="#181818",
                                 width=20, command=lambda: self.Raise_frame(self.Portfolio8_frame))
        self.Button9 = tk.Button(self.Frame_Top, text='All Weather Portfolio', relief='groove', fg='white', background="#181818",
                                 width=20, command=lambda: self.Raise_frame(self.Portfolio9_frame))




        self.Overview_Button.place(x=0, y=44)
        #self.Button1.place(x=150, y=44)
        #self.Button2.place(x=300, y=44)
        #self.Button3.place(x=450, y=44)

        self.Button7.place(x=150, y=44)

        # self.Button4.place(x=600,y=44)
        self.Button5.place(x=750, y=44)
        self.Button6.place(x=600, y=44)
        self.Button8.place(x=1050, y=44)
        self.Button9.place(x=900, y=44)
        self.Raise_frame(self.Overview_frame)
        # ======================================================================

        self.Overview_fig = Figure(figsize=(6, 5), dpi=100)#6,4,110
        self.Overview_fig.set_tight_layout(True)
        self.Overview_fig.patch.set_facecolor('#F0F0F0')

        self.Overview_Chart = self.Overview_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Overview_Chart.tick_params(axis='both', which='major', labelsize=8)
        self.Overview_Chart.set_xlabel('Standard Deviation(%)', fontsize=8)
        self.Overview_Chart.set_ylabel('Annual Return(%)', fontsize=8)

        self.Overview_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Overview_Chart.spines['bottom'].set_color('black')
        self.Overview_Chart.spines['top'].set_color('black')
        self.Overview_Chart.spines['left'].set_color('black')
        self.Overview_Chart.spines['right'].set_color('black')
        self.Overview_Chart.xaxis.label.set_color('black')
        self.Overview_Chart.yaxis.label.set_color('black')

        self.Overview_canvas = FigureCanvasTkAgg(self.Overview_fig, master=self.Overview_frame)
        self.Overview_canvas.get_tk_widget().place(x=0, y=10)
        self.Overview_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                       highlightbackground='#FFFFFF')

        # ----------------------------------------------------------------------

        self.Overview1_fig = Figure(figsize=(6, 5), dpi=100)#6,4,110
        self.Overview1_fig.set_tight_layout(True)
        self.Overview1_fig.patch.set_facecolor('#F0F0F0')

        self.Overview1_Chart = self.Overview1_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Overview1_Chart.tick_params(axis='both', which='major', labelsize=8)
        self.Overview1_Chart.set_xlabel('Date', fontsize=8)
        self.Overview1_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.Overview1_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Overview1_Chart.spines['bottom'].set_color('black')
        self.Overview1_Chart.spines['top'].set_color('black')
        self.Overview1_Chart.spines['left'].set_color('black')
        self.Overview1_Chart.spines['right'].set_color('black')
        self.Overview1_Chart.xaxis.label.set_color('black')
        self.Overview1_Chart.yaxis.label.set_color('black')

        self.Overview1_canvas = FigureCanvasTkAgg(self.Overview1_fig, master=self.Overview_frame)
        self.Overview1_canvas.get_tk_widget().place(x=600, y=10)
        self.Overview1_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                        highlightbackground='#FFFFFF')
        self.Overview_Output_Button = ttk.Button(self.Overview_frame, text='Output as Excel', width=20,command = self.All_Portfolio_Output)
        self.Overview_Output_Button.place(x=550, y=520)

        # ======================================================================
        # NAV Chart
        self.Aggressive_fig = Figure(figsize=(6, 5), dpi=120)
        self.Aggressive_fig.set_tight_layout(True)
        self.Aggressive_fig.patch.set_facecolor('#F0F0F0')

        self.Aggressive_Chart = self.Aggressive_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Aggressive_Chart.tick_params(axis='both', which='major', labelsize=8)
        self.Aggressive_Chart.set_xlabel('Time', fontsize=8)
        self.Aggressive_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.Aggressive_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Aggressive_Chart.spines['bottom'].set_color('black')
        self.Aggressive_Chart.spines['top'].set_color('black')
        self.Aggressive_Chart.spines['left'].set_color('black')
        self.Aggressive_Chart.spines['right'].set_color('black')
        self.Aggressive_Chart.xaxis.label.set_color('black')
        self.Aggressive_Chart.yaxis.label.set_color('black')

        self.Aggressive_canvas = FigureCanvasTkAgg(self.Aggressive_fig, master=self.Portfolio1_frame)
        self.Aggressive_canvas.get_tk_widget().place(x=0, y=0)
        self.Aggressive_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                         highlightbackground='#FFFFFF')

        # Pie Chart
        self.Aggressive_pie_fig = Figure(figsize=(4, 4), dpi=60)
        self.Aggressive_pie_fig.set_tight_layout(True)
        self.Aggressive_pie_fig.patch.set_facecolor('#F0F0F0')

        self.Aggressive_pie_Chart = self.Aggressive_pie_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Aggressive_pie_Chart.tick_params(axis='both', which='major', labelsize=8)
        # self.Aggressive_pie_Chart.set_xlabel('Time', fontsize=8)
        # self.Aggressive_pie_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.Aggressive_pie_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Aggressive_pie_Chart.spines['bottom'].set_color('#F0F0F0')
        self.Aggressive_pie_Chart.spines['top'].set_color('#F0F0F0')
        self.Aggressive_pie_Chart.spines['left'].set_color('#F0F0F0')
        self.Aggressive_pie_Chart.spines['right'].set_color('#F0F0F0')
        self.Aggressive_pie_Chart.xaxis.label.set_color('#F0F0F0')
        self.Aggressive_pie_Chart.yaxis.label.set_color('#F0F0F0')
        self.Aggressive_pie_Chart.tick_params(axis='x', colors='#F0F0F0')
        self.Aggressive_pie_Chart.tick_params(axis='y', colors='black')

        self.Aggressive_pie_canvas = FigureCanvasTkAgg(self.Aggressive_pie_fig, master=self.Portfolio1_frame)
        self.Aggressive_pie_canvas.get_tk_widget().place(x=720, y=80)
        self.Aggressive_pie_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                             highlightbackground='#FFFFFF')

        self.labels = 'Equity', 'Fixed Income'
        self.sizes = [80, 20]

        self.Aggressive_pie_Chart.pie(self.sizes, colors=cm.Blues(np.arange(len(self.sizes)) / float(len(self.sizes))),
                                      labels=self.labels, autopct='%1.1f%%', shadow=True, startangle=150,
                                      labeldistance=0.8)  # ,autopct='%1.1f%%'
        self.Aggressive_pie_canvas.show()

        # Pie Chart
        self.Aggressive_pie2_fig = Figure(figsize=(4, 4), dpi=60)
        self.Aggressive_pie2_fig.set_tight_layout(True)
        self.Aggressive_pie2_fig.patch.set_facecolor('#F0F0F0')

        self.Aggressive_pie2_Chart = self.Aggressive_pie2_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Aggressive_pie2_Chart.tick_params(axis='both', which='major', labelsize=8)
        # self.Aggressive_pie_Chart.set_xlabel('Time', fontsize=8)
        # self.Aggressive_pie_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.Aggressive_pie2_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Aggressive_pie2_Chart.spines['bottom'].set_color('#F0F0F0')
        self.Aggressive_pie2_Chart.spines['top'].set_color('#F0F0F0')
        self.Aggressive_pie2_Chart.spines['left'].set_color('#F0F0F0')
        self.Aggressive_pie2_Chart.spines['right'].set_color('#F0F0F0')
        self.Aggressive_pie2_Chart.xaxis.label.set_color('#F0F0F0')
        self.Aggressive_pie2_Chart.yaxis.label.set_color('#F0F0F0')
        self.Aggressive_pie2_Chart.tick_params(axis='x', colors='#F0F0F0')
        self.Aggressive_pie2_Chart.tick_params(axis='y', colors='black')

        self.Aggressive_pie2_canvas = FigureCanvasTkAgg(self.Aggressive_pie2_fig, master=self.Portfolio1_frame)
        self.Aggressive_pie2_canvas.get_tk_widget().place(x=1000, y=80)
        self.Aggressive_pie2_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                              highlightbackground='#FFFFFF')

        '''
        self.labels = 'S&P500', 'Emerging Market', 'Japan', 'EAFE', 'US Treasury Bond'
        self.sizes = [20,20,20,20, 20]
        self.cs=cm.Set1(np.arange(5)/5.)
        #self.colors = ['yellowgreen', 'gold', 'lightskyblue', 'lightcoral']
        self.explode = (0, 0)  # only "explode" the 2nd slice (i.e. 'Hogs')
        
        self.Aggressive_pie2_Chart.pie(self.sizes,colors=cm.Blues(np.arange(len(self.sizes))/float(len(self.sizes))), labels=self.labels, autopct='%1.1f%%', shadow=True, startangle=150, labeldistance=0.8)#,autopct='%1.1f%%'   
        self.Aggressive_pie2_canvas.show()
        #colors=self.colors,
        '''

        self.Aggressive_Tree_table = ttk.Treeview(self.Portfolio1_frame, height="10")

        self.Aggressive_Tree_table["columns"] = ("column1", "column2", 'column3', 'column4')
        self.Aggressive_Tree_table.column("#0", width=10, anchor='e')
        self.Aggressive_Tree_table.column("column1", width=60, anchor='center')
        self.Aggressive_Tree_table.column("column2", width=80, anchor='w')
        self.Aggressive_Tree_table.column("column3", width=250, anchor='w')
        self.Aggressive_Tree_table.column("column4", width=130, anchor='center')

        self.Aggressive_Tree_table.heading('#0', text='')
        self.Aggressive_Tree_table.heading("column1", text="Category")
        self.Aggressive_Tree_table.heading("column2", text="Ticker")
        self.Aggressive_Tree_table.heading("column3", text="Name")
        self.Aggressive_Tree_table.heading("column4", text="Weight(%)",
                                           command=lambda: self.treeview_sort_column(self.Aggressive_Tree_table,
                                                                                     "column4", False))

        self.Aggressive_Tree_table.place(x=750, y=340)

        self.Aggressive_Return_StrVar = tk.StringVar()
        self.Aggressive_Volatility_StrVar = tk.StringVar()
        self.Aggressive_Sharpe_StrVar = tk.StringVar()

        self.Aggressive_Return = tk.Label(self.Portfolio1_frame, textvariable=self.Aggressive_Return_StrVar)
        self.Aggressive_Return.place(x=50, y=470)

        self.Aggressive_Volatility = tk.Label(self.Portfolio1_frame, textvariable=self.Aggressive_Volatility_StrVar)
        self.Aggressive_Volatility.place(x=50, y=500)

        self.Aggressive_Sharpe = tk.Label(self.Portfolio1_frame, textvariable=self.Aggressive_Sharpe_StrVar)
        self.Aggressive_Sharpe.place(x=50, y=530)

        self.Aggressive_dollar_entry = tk.Entry(self.Portfolio1_frame)
        self.Aggressive_dollar_entry.place(x=750, y=312)

        self.Aggressive_dollar_weight = ttk.Button(self.Portfolio1_frame, text='Portfolio Weight in Dollar',
                                                   command=lambda: self.Portfolio_dollar(
                                                       Entry=self.Aggressive_dollar_entry,
                                                       Tree=self.Aggressive_Tree_table))
        self.Aggressive_dollar_weight.place(x=900, y=310)

        #self.Aggressive_Performance_Output = ttk.Button(self.Portfolio1_frame, text='Performance Output',
                                                        #command=lambda: self.Performance_print(Aggressive_compare,
                                                                                               #'Aggressive_Performance.xlsx',Aggressive_Weight_df,Aggressive_AR,Aggressive_Std,Aggressive_Sharpe))
        #self.Aggressive_Performance_Output.place(x=600, y=530)

        # self.Aggressive_Last_Label = tk.Label(self.Portfolio1_frame,textvariable=self.Last_StrVar)
        # self.Aggressive_Last_Label.place(x=750,y=500)










        # Aggressive frame setting
        self.Aggressive_Title = tk.Label(self.Portfolio1_frame, text='Aggressive Portfolio', font=('Arial', 22, 'bold'))
        self.Aggressive_Title.place(x=750, y=10)

        # self.Aggressive_description = tk.Label(self.Portfolio1_frame,text='80% of Equity, 20% of Fixed income',font=('Arial',14))
        # self.Aggressive_description.place(x=750,y=60)

        ttk.Separator(self.Portfolio1_frame, orient=tk.HORIZONTAL).place(x=750, y=50, width=500)

        # ======================================================================


        self.Moderate_fig = Figure(figsize=(6, 4), dpi=120)
        self.Moderate_fig.set_tight_layout(True)
        self.Moderate_fig.patch.set_facecolor('#F0F0F0')

        self.Moderate_Chart = self.Moderate_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Moderate_Chart.tick_params(axis='both', which='major', labelsize=8)
        self.Moderate_Chart.set_xlabel('Time', fontsize=8)
        self.Moderate_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.Moderate_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Moderate_Chart.spines['bottom'].set_color('black')
        self.Moderate_Chart.spines['top'].set_color('black')
        self.Moderate_Chart.spines['left'].set_color('black')
        self.Moderate_Chart.spines['right'].set_color('black')
        self.Moderate_Chart.xaxis.label.set_color('black')
        self.Moderate_Chart.yaxis.label.set_color('black')

        self.Moderate_canvas = FigureCanvasTkAgg(self.Moderate_fig, master=self.Portfolio2_frame)
        self.Moderate_canvas.get_tk_widget().place(x=0, y=0)
        self.Moderate_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                       highlightbackground='#FFFFFF')

        # Moderate frame setting
        self.Moderate_Title = tk.Label(self.Portfolio2_frame, text='Moderate Portfolio', font=('Arial', 22, 'bold'))
        self.Moderate_Title.place(x=750, y=10)

        # self.Moderate_description = tk.Label(self.Portfolio2_frame,text='60% of Equity, 40% of Fixed income',font=('Arial',14))
        # self.Moderate_description.place(x=750,y=60)

        ttk.Separator(self.Portfolio2_frame, orient=tk.HORIZONTAL).place(x=750, y=50, width=500)

        # Pie Chart
        self.Moderate_pie_fig = Figure(figsize=(4, 4), dpi=60)
        self.Moderate_pie_fig.set_tight_layout(True)
        self.Moderate_pie_fig.patch.set_facecolor('#F0F0F0')

        self.Moderate_pie_Chart = self.Moderate_pie_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Moderate_pie_Chart.tick_params(axis='both', which='major', labelsize=8)
        # self.Aggressive_pie_Chart.set_xlabel('Time', fontsize=8)
        # self.Aggressive_pie_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.Moderate_pie_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Moderate_pie_Chart.spines['bottom'].set_color('#F0F0F0')
        self.Moderate_pie_Chart.spines['top'].set_color('#F0F0F0')
        self.Moderate_pie_Chart.spines['left'].set_color('#F0F0F0')
        self.Moderate_pie_Chart.spines['right'].set_color('#F0F0F0')
        self.Moderate_pie_Chart.xaxis.label.set_color('#F0F0F0')
        self.Moderate_pie_Chart.yaxis.label.set_color('#F0F0F0')
        self.Moderate_pie_Chart.tick_params(axis='x', colors='#F0F0F0')
        self.Moderate_pie_Chart.tick_params(axis='y', colors='black')

        self.Moderate_pie_canvas = FigureCanvasTkAgg(self.Moderate_pie_fig, master=self.Portfolio2_frame)
        self.Moderate_pie_canvas.get_tk_widget().place(x=720, y=80)
        self.Moderate_pie_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                           highlightbackground='#FFFFFF')

        self.labels = 'Equity', 'Fixed Income'
        self.sizes = [60, 40]

        self.Moderate_pie_Chart.pie(self.sizes, colors=cm.Blues(np.arange(len(self.sizes)) / float(len(self.sizes))),
                                    labels=self.labels, autopct='%1.1f%%', shadow=True, startangle=150,
                                    labeldistance=0.8)  # ,autopct='%1.1f%%'
        self.Moderate_pie_canvas.show()

        # Pie Chart
        self.Moderate_pie2_fig = Figure(figsize=(4, 4), dpi=60)
        self.Moderate_pie2_fig.set_tight_layout(True)
        self.Moderate_pie2_fig.patch.set_facecolor('#F0F0F0')

        self.Moderate_pie2_Chart = self.Moderate_pie2_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Moderate_pie2_Chart.tick_params(axis='both', which='major', labelsize=8)
        # self.Moderate_pie_Chart.set_xlabel('Time', fontsize=8)
        # self.Moderate_pie_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.Moderate_pie2_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Moderate_pie2_Chart.spines['bottom'].set_color('#F0F0F0')
        self.Moderate_pie2_Chart.spines['top'].set_color('#F0F0F0')
        self.Moderate_pie2_Chart.spines['left'].set_color('#F0F0F0')
        self.Moderate_pie2_Chart.spines['right'].set_color('#F0F0F0')
        self.Moderate_pie2_Chart.xaxis.label.set_color('#F0F0F0')
        self.Moderate_pie2_Chart.yaxis.label.set_color('#F0F0F0')
        self.Moderate_pie2_Chart.tick_params(axis='x', colors='#F0F0F0')
        self.Moderate_pie2_Chart.tick_params(axis='y', colors='black')

        self.Moderate_pie2_canvas = FigureCanvasTkAgg(self.Moderate_pie2_fig, master=self.Portfolio2_frame)
        self.Moderate_pie2_canvas.get_tk_widget().place(x=1000, y=80)
        self.Moderate_pie2_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                            highlightbackground='#FFFFFF')

        self.Moderate_Tree_table = ttk.Treeview(self.Portfolio2_frame, height="10")

        self.Moderate_Tree_table["columns"] = ("column1", "column2", 'column3', 'column4')
        self.Moderate_Tree_table.column("#0", width=10, anchor='e')
        self.Moderate_Tree_table.column("column1", width=60, anchor='center')
        self.Moderate_Tree_table.column("column2", width=80, anchor='w')
        self.Moderate_Tree_table.column("column3", width=250, anchor='w')
        self.Moderate_Tree_table.column("column4", width=130, anchor='center')

        self.Moderate_Tree_table.heading('#0', text='')
        self.Moderate_Tree_table.heading("column1", text="Category")
        self.Moderate_Tree_table.heading("column2", text="Ticker")
        self.Moderate_Tree_table.heading("column3", text="Name")
        self.Moderate_Tree_table.heading("column4", text="Weight(%)",
                                         command=lambda: self.treeview_sort_column(self.Moderate_Tree_table, "column4",
                                                                                   False))

        self.Moderate_Tree_table.place(x=750, y=340)

        self.Moderate_Return_StrVar = tk.StringVar()
        self.Moderate_Volatility_StrVar = tk.StringVar()
        self.Moderate_Sharpe_StrVar = tk.StringVar()

        self.Moderate_Return = tk.Label(self.Portfolio2_frame, textvariable=self.Moderate_Return_StrVar)
        self.Moderate_Return.place(x=50, y=470)

        self.Moderate_Volatility = tk.Label(self.Portfolio2_frame, textvariable=self.Moderate_Volatility_StrVar)
        self.Moderate_Volatility.place(x=50, y=500)

        self.Moderate_Sharpe = tk.Label(self.Portfolio2_frame, textvariable=self.Moderate_Sharpe_StrVar)
        self.Moderate_Sharpe.place(x=50, y=530)

        self.Moderate_dollar_entry = tk.Entry(self.Portfolio2_frame)
        self.Moderate_dollar_entry.place(x=750, y=312)

        self.Moderate_dollar_weight = ttk.Button(self.Portfolio2_frame, text='Portfolio Weight in Dollar',
                                                 command=lambda: self.Portfolio_dollar(Entry=self.Moderate_dollar_entry,
                                                                                       Tree=self.Moderate_Tree_table))
        self.Moderate_dollar_weight.place(x=900, y=310)

        #self.Moderate_Performance_Output = ttk.Button(self.Portfolio2_frame, text='Performance Output',
                                                      #command=lambda: self.Performance_print(Moderate_compare,
                                                                                             #'Moderate_Performance.xlsx',Moderate_Weight_df,Moderate_AR,Moderate_Std,Moderate_Sharpe))
        #self.Moderate_Performance_Output.place(x=600, y=530)

        # self.Moderate_Last_Label = tk.Label(self.Portfolio2_frame,textvariable=self.Last_StrVar)
        # self.Moderate_Last_Label.place(x=750,y=500)












        # ======================================================================


        self.Conservative_fig = Figure(figsize=(6, 4), dpi=120)
        self.Conservative_fig.set_tight_layout(True)
        self.Conservative_fig.patch.set_facecolor('#F0F0F0')

        self.Conservative_Chart = self.Conservative_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Conservative_Chart.tick_params(axis='both', which='major', labelsize=8)
        self.Conservative_Chart.set_xlabel('Time', fontsize=8)
        self.Conservative_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.Conservative_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Conservative_Chart.spines['bottom'].set_color('black')
        self.Conservative_Chart.spines['top'].set_color('black')
        self.Conservative_Chart.spines['left'].set_color('black')
        self.Conservative_Chart.spines['right'].set_color('black')
        self.Conservative_Chart.xaxis.label.set_color('black')
        self.Conservative_Chart.yaxis.label.set_color('black')

        self.Conservative_canvas = FigureCanvasTkAgg(self.Conservative_fig, master=self.Portfolio3_frame)
        self.Conservative_canvas.get_tk_widget().place(x=0, y=0)
        self.Conservative_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                           highlightbackground='#FFFFFF')

        # Conservative frame setting
        self.Conservative_Title = tk.Label(self.Portfolio3_frame, text='Conservative Portfolio',
                                           font=('Arial', 22, 'bold'))
        self.Conservative_Title.place(x=750, y=10)

        # self.Conservative_description = tk.Label(self.Portfolio3_frame,text='20% of Equity, 80% of Fixed income',font=('Arial',14))
        # self.Conservative_description.place(x=750,y=60)

        ttk.Separator(self.Portfolio3_frame, orient=tk.HORIZONTAL).place(x=750, y=50, width=500)

        # Pie Chart
        self.Conservative_pie_fig = Figure(figsize=(4, 4), dpi=60)
        self.Conservative_pie_fig.set_tight_layout(True)
        self.Conservative_pie_fig.patch.set_facecolor('#F0F0F0')

        self.Conservative_pie_Chart = self.Conservative_pie_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Conservative_pie_Chart.tick_params(axis='both', which='major', labelsize=8)
        # self.Aggressive_pie_Chart.set_xlabel('Time', fontsize=8)
        # self.Aggressive_pie_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.Conservative_pie_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Conservative_pie_Chart.spines['bottom'].set_color('#F0F0F0')
        self.Conservative_pie_Chart.spines['top'].set_color('#F0F0F0')
        self.Conservative_pie_Chart.spines['left'].set_color('#F0F0F0')
        self.Conservative_pie_Chart.spines['right'].set_color('#F0F0F0')
        self.Conservative_pie_Chart.xaxis.label.set_color('#F0F0F0')
        self.Conservative_pie_Chart.yaxis.label.set_color('#F0F0F0')
        self.Conservative_pie_Chart.tick_params(axis='x', colors='#F0F0F0')
        self.Conservative_pie_Chart.tick_params(axis='y', colors='black')

        self.Conservative_pie_canvas = FigureCanvasTkAgg(self.Conservative_pie_fig, master=self.Portfolio3_frame)
        self.Conservative_pie_canvas.get_tk_widget().place(x=720, y=80)
        self.Conservative_pie_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                               highlightbackground='#FFFFFF')

        self.labels = 'Equity', 'Fixed Income'
        self.sizes = [20, 80]

        self.Conservative_pie_Chart.pie(self.sizes,
                                        colors=cm.Blues(np.arange(len(self.sizes)) / float(len(self.sizes))),
                                        labels=self.labels, autopct='%1.1f%%', shadow=True, startangle=150,
                                        labeldistance=0.8)  # ,autopct='%1.1f%%'
        self.Conservative_pie_canvas.show()

        # Pie Chart
        self.Conservative_pie2_fig = Figure(figsize=(4, 4), dpi=60)
        self.Conservative_pie2_fig.set_tight_layout(True)
        self.Conservative_pie2_fig.patch.set_facecolor('#F0F0F0')

        self.Conservative_pie2_Chart = self.Conservative_pie2_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Conservative_pie2_Chart.tick_params(axis='both', which='major', labelsize=8)
        # self.Conservative_pie2_Chart.set_xlabel('Time', fontsize=8)
        # self.Conservative_pie2_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.Conservative_pie2_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Conservative_pie2_Chart.spines['bottom'].set_color('#F0F0F0')
        self.Conservative_pie2_Chart.spines['top'].set_color('#F0F0F0')
        self.Conservative_pie2_Chart.spines['left'].set_color('#F0F0F0')
        self.Conservative_pie2_Chart.spines['right'].set_color('#F0F0F0')
        self.Conservative_pie2_Chart.xaxis.label.set_color('#F0F0F0')
        self.Conservative_pie2_Chart.yaxis.label.set_color('#F0F0F0')
        self.Conservative_pie2_Chart.tick_params(axis='x', colors='#F0F0F0')
        self.Conservative_pie2_Chart.tick_params(axis='y', colors='black')

        self.Conservative_pie2_canvas = FigureCanvasTkAgg(self.Conservative_pie2_fig, master=self.Portfolio3_frame)
        self.Conservative_pie2_canvas.get_tk_widget().place(x=1000, y=80)
        self.Conservative_pie2_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                                highlightbackground='#FFFFFF')

        self.Conservative_Tree_table = ttk.Treeview(self.Portfolio3_frame, height="10")

        self.Conservative_Tree_table["columns"] = ("column1", "column2", 'column3', 'column4')
        self.Conservative_Tree_table.column("#0", width=10, anchor='e')
        self.Conservative_Tree_table.column("column1", width=60, anchor='center')
        self.Conservative_Tree_table.column("column2", width=80, anchor='w')
        self.Conservative_Tree_table.column("column3", width=250, anchor='w')
        self.Conservative_Tree_table.column("column4", width=130, anchor='center')

        self.Conservative_Tree_table.heading('#0', text='')
        self.Conservative_Tree_table.heading("column1", text="Category")
        self.Conservative_Tree_table.heading("column2", text="Ticker")
        self.Conservative_Tree_table.heading("column3", text="Name")
        self.Conservative_Tree_table.heading("column4", text="Weight(%)",
                                             command=lambda: self.treeview_sort_column(self.Conservative_Tree_table,
                                                                                       "column4", False))

        self.Conservative_Tree_table.place(x=750, y=340)

        self.Conservative_Return_StrVar = tk.StringVar()
        self.Conservative_Volatility_StrVar = tk.StringVar()
        self.Conservative_Sharpe_StrVar = tk.StringVar()

        self.Conservative_Return = tk.Label(self.Portfolio3_frame, textvariable=self.Conservative_Return_StrVar)
        self.Conservative_Return.place(x=50, y=470)

        self.Conservative_Volatility = tk.Label(self.Portfolio3_frame, textvariable=self.Conservative_Volatility_StrVar)
        self.Conservative_Volatility.place(x=50, y=500)

        self.Conservative_Sharpe = tk.Label(self.Portfolio3_frame, textvariable=self.Conservative_Sharpe_StrVar)
        self.Conservative_Sharpe.place(x=50, y=530)

        self.Conservative_dollar_entry = tk.Entry(self.Portfolio3_frame)
        self.Conservative_dollar_entry.place(x=750, y=312)

        self.Conservative_dollar_weight = ttk.Button(self.Portfolio3_frame, text='Portfolio Weight in Dollar',
                                                     command=lambda: self.Portfolio_dollar(
                                                         Entry=self.Conservative_dollar_entry,
                                                         Tree=self.Conservative_Tree_table))
        self.Conservative_dollar_weight.place(x=900, y=310)

        #self.Conservative_Performance_Output = ttk.Button(self.Portfolio3_frame, text='Performance Output',
                                                          #command=lambda: self.Performance_print(Conservative_compare,
                                                                                                 #'Conservative_Performance.xlsx',Consevative_Weight_df,Conservative_AR,Conservative_Std,Conservative_Sharpe))
        #self.Conservative_Performance_Output.place(x=600, y=530)

        # ======================================================================

        # ======================================================================


        self.Aggregate_fig = Figure(figsize=(6, 4), dpi=120)
        self.Aggregate_fig.set_tight_layout(True)
        self.Aggregate_fig.patch.set_facecolor('#F0F0F0')

        self.Aggregate_Chart = self.Aggregate_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Aggregate_Chart.tick_params(axis='both', which='major', labelsize=8)
        self.Aggregate_Chart.set_xlabel('Time', fontsize=8)
        self.Aggregate_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.Aggregate_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Aggregate_Chart.spines['bottom'].set_color('black')
        self.Aggregate_Chart.spines['top'].set_color('black')
        self.Aggregate_Chart.spines['left'].set_color('black')
        self.Aggregate_Chart.spines['right'].set_color('black')
        self.Aggregate_Chart.xaxis.label.set_color('black')
        self.Aggregate_Chart.yaxis.label.set_color('black')

        self.Aggregate_canvas = FigureCanvasTkAgg(self.Aggregate_fig, master=self.Portfolio4_frame)
        self.Aggregate_canvas.get_tk_widget().place(x=0, y=0)
        self.Aggregate_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                        highlightbackground='#FFFFFF')

        # Conservative frame setting
        self.Aggregate_Title = tk.Label(self.Portfolio4_frame, text='Aggregate Portfolio', font=('Arial', 22, 'bold'))
        self.Aggregate_Title.place(x=750, y=10)

        # self.Conservative_description = tk.Label(self.Portfolio3_frame,text='20% of Equity, 80% of Fixed income',font=('Arial',14))
        # self.Conservative_description.place(x=750,y=60)

        ttk.Separator(self.Portfolio4_frame, orient=tk.HORIZONTAL).place(x=750, y=50, width=500)

        # Pie Chart
        self.Aggregate_pie_fig = Figure(figsize=(4, 4), dpi=60)
        self.Aggregate_pie_fig.set_tight_layout(True)
        self.Aggregate_pie_fig.patch.set_facecolor('#F0F0F0')

        self.Aggregate_pie_Chart = self.Aggregate_pie_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Aggregate_pie_Chart.tick_params(axis='both', which='major', labelsize=8)
        # self.Aggressive_pie_Chart.set_xlabel('Time', fontsize=8)
        # self.Aggressive_pie_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.Aggregate_pie_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Aggregate_pie_Chart.spines['bottom'].set_color('#F0F0F0')
        self.Aggregate_pie_Chart.spines['top'].set_color('#F0F0F0')
        self.Aggregate_pie_Chart.spines['left'].set_color('#F0F0F0')
        self.Aggregate_pie_Chart.spines['right'].set_color('#F0F0F0')
        self.Aggregate_pie_Chart.xaxis.label.set_color('#F0F0F0')
        self.Aggregate_pie_Chart.yaxis.label.set_color('#F0F0F0')
        self.Aggregate_pie_Chart.tick_params(axis='x', colors='#F0F0F0')
        self.Aggregate_pie_Chart.tick_params(axis='y', colors='black')

        self.Aggregate_pie_canvas = FigureCanvasTkAgg(self.Aggregate_pie_fig, master=self.Portfolio4_frame)
        self.Aggregate_pie_canvas.get_tk_widget().place(x=720, y=80)
        self.Aggregate_pie_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                            highlightbackground='#FFFFFF')

        self.labels = 'Equity', 'Fixed Income'
        self.sizes = [20, 80]

        # self.Aggregate_pie_Chart.pie(self.sizes,colors=cm.Blues(np.arange(len(self.sizes))/float(len(self.sizes))), labels=self.labels, autopct='%1.1f%%', shadow=True, startangle=150, labeldistance=0.8)#,autopct='%1.1f%%'
        # self.Aggregate_pie_canvas.show()



        # Pie Chart
        self.Aggregate_pie2_fig = Figure(figsize=(4, 4), dpi=60)
        self.Aggregate_pie2_fig.set_tight_layout(True)
        self.Aggregate_pie2_fig.patch.set_facecolor('#F0F0F0')

        self.Aggregate_pie2_Chart = self.Aggregate_pie2_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Aggregate_pie2_Chart.tick_params(axis='both', which='major', labelsize=8)
        # self.Conservative_pie2_Chart.set_xlabel('Time', fontsize=8)
        # self.Conservative_pie2_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.Aggregate_pie2_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Aggregate_pie2_Chart.spines['bottom'].set_color('#F0F0F0')
        self.Aggregate_pie2_Chart.spines['top'].set_color('#F0F0F0')
        self.Aggregate_pie2_Chart.spines['left'].set_color('#F0F0F0')
        self.Aggregate_pie2_Chart.spines['right'].set_color('#F0F0F0')
        self.Aggregate_pie2_Chart.xaxis.label.set_color('#F0F0F0')
        self.Aggregate_pie2_Chart.yaxis.label.set_color('#F0F0F0')
        self.Aggregate_pie2_Chart.tick_params(axis='x', colors='#F0F0F0')
        self.Aggregate_pie2_Chart.tick_params(axis='y', colors='black')

        self.Aggregate_pie2_canvas = FigureCanvasTkAgg(self.Aggregate_pie2_fig, master=self.Portfolio4_frame)
        self.Aggregate_pie2_canvas.get_tk_widget().place(x=1000, y=80)
        self.Aggregate_pie2_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                             highlightbackground='#FFFFFF')

        self.Aggregate_Tree_table = ttk.Treeview(self.Portfolio4_frame, height="10")

        self.Aggregate_Tree_table["columns"] = ("column1", "column2", 'column3', 'column4')
        self.Aggregate_Tree_table.column("#0", width=10, anchor='e')
        self.Aggregate_Tree_table.column("column1", width=60, anchor='center')
        self.Aggregate_Tree_table.column("column2", width=80, anchor='w')
        self.Aggregate_Tree_table.column("column3", width=250, anchor='w')
        self.Aggregate_Tree_table.column("column4", width=130, anchor='center')

        self.Aggregate_Tree_table.heading('#0', text='')
        self.Aggregate_Tree_table.heading("column1", text="Category")
        self.Aggregate_Tree_table.heading("column2", text="Ticker")
        self.Aggregate_Tree_table.heading("column3", text="Name")
        self.Aggregate_Tree_table.heading("column4", text="Weight(%)",
                                          command=lambda: self.treeview_sort_column(self.Aggregate_Tree_table,
                                                                                    "column4", False))

        self.Aggregate_Tree_table.place(x=750, y=340)

        self.Aggregate_Return_StrVar = tk.StringVar()
        self.Aggregate_Volatility_StrVar = tk.StringVar()
        self.Aggregate_Sharpe_StrVar = tk.StringVar()

        self.Aggregate_Return = tk.Label(self.Portfolio4_frame, textvariable=self.Aggregate_Return_StrVar)
        self.Aggregate_Return.place(x=50, y=470)

        self.Aggregate_Volatility = tk.Label(self.Portfolio4_frame, textvariable=self.Aggregate_Volatility_StrVar)
        self.Aggregate_Volatility.place(x=50, y=500)

        self.Aggregate_Sharpe = tk.Label(self.Portfolio4_frame, textvariable=self.Aggregate_Sharpe_StrVar)
        self.Aggregate_Sharpe.place(x=50, y=530)

        self.Aggregate_dollar_entry = tk.Entry(self.Portfolio4_frame)
        self.Aggregate_dollar_entry.place(x=750, y=312)

        self.Aggregate_dollar_weight = ttk.Button(self.Portfolio4_frame, text='Portfolio Weight in Dollar',
                                                  command=lambda: self.Portfolio_dollar(
                                                      Entry=self.Aggregate_dollar_entry,
                                                      Tree=self.Aggregate_Tree_table))
        self.Aggregate_dollar_weight.place(x=900, y=310)

        # ======================================================================

        # ======================================================================


        self.Bond_fig = Figure(figsize=(6, 4), dpi=100)
        self.Bond_fig.set_tight_layout(True)
        self.Bond_fig.patch.set_facecolor('#F0F0F0')

        self.Bond_Chart = self.Bond_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Bond_Chart.tick_params(axis='both', which='major', labelsize=8)
        self.Bond_Chart.set_xlabel('Time', fontsize=8)
        self.Bond_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.Bond_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Bond_Chart.spines['bottom'].set_color('black')
        self.Bond_Chart.spines['top'].set_color('black')
        self.Bond_Chart.spines['left'].set_color('black')
        self.Bond_Chart.spines['right'].set_color('black')
        self.Bond_Chart.xaxis.label.set_color('black')
        self.Bond_Chart.yaxis.label.set_color('black')

        self.Bond_canvas = FigureCanvasTkAgg(self.Bond_fig, master=self.Portfolio5_frame)
        self.Bond_canvas.get_tk_widget().place(x=0, y=0)
        self.Bond_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                   highlightbackground='#FFFFFF')

        # Conservative frame setting
        self.Bond_Title = tk.Label(self.Portfolio5_frame, text='Global Strategic Bond (Market Timing)',
                                   font=('Arial', 22, 'bold'))
        self.Bond_Title.place(x=620, y=10)

        # self.Conservative_description = tk.Label(self.Portfolio3_frame,text='20% of Equity, 80% of Fixed income',font=('Arial',14))
        # self.Conservative_description.place(x=750,y=60)

        ttk.Separator(self.Portfolio5_frame, orient=tk.HORIZONTAL).place(x=620, y=50, width=500)

        # Pie Chart
        self.Bond_pie_fig = Figure(figsize=(4, 4), dpi=60)
        self.Bond_pie_fig.set_tight_layout(True)
        self.Bond_pie_fig.patch.set_facecolor('#F0F0F0')

        self.Bond_pie_Chart = self.Bond_pie_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Bond_pie_Chart.tick_params(axis='both', which='major', labelsize=8)
        # self.Aggressive_pie_Chart.set_xlabel('Time', fontsize=8)
        # self.Aggressive_pie_Chart.set_ylabel('Cumulative Return', fontsize=8)
        self.Bond_pie_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Bond_pie_Chart.spines['bottom'].set_color('#F0F0F0')
        self.Bond_pie_Chart.spines['top'].set_color('#F0F0F0')
        self.Bond_pie_Chart.spines['left'].set_color('#F0F0F0')
        self.Bond_pie_Chart.spines['right'].set_color('#F0F0F0')
        self.Bond_pie_Chart.xaxis.label.set_color('#F0F0F0')
        self.Bond_pie_Chart.yaxis.label.set_color('#F0F0F0')
        self.Bond_pie_Chart.tick_params(axis='x', colors='#F0F0F0')
        self.Bond_pie_Chart.tick_params(axis='y', colors='black')

        self.Bond_pie_canvas = FigureCanvasTkAgg(self.Bond_pie_fig, master=self.Portfolio5_frame)
        self.Bond_pie_canvas.get_tk_widget().place(x=600, y=60)
        self.Bond_pie_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                       highlightbackground='#FFFFFF')

        '''
        #Pie Chart
        self.Bond_pie2_fig = Figure(figsize=(4,4), dpi=60)
        self.Bond_pie2_fig.set_tight_layout(True)
        self.Bond_pie2_fig.patch.set_facecolor('#F0F0F0')
        
        self.Bond_pie2_Chart = self.Aggregate_pie2_fig.add_subplot(111,axisbg='#F0F0F0')
        self.Bond_pie2_Chart.tick_params(axis='both', which='major', labelsize=8)
        #self.Conservative_pie2_Chart.set_xlabel('Time', fontsize=8)
        #self.Conservative_pie2_Chart.set_ylabel('Cumulative Return', fontsize=8)
        
        self.Bond_pie2_Chart.tick_params(axis='both', which='major',colors='black', labelsize=6)
        self.Bond_pie2_Chart.spines['bottom'].set_color('#F0F0F0')
        self.Bond_pie2_Chart.spines['top'].set_color('#F0F0F0')
        self.Bond_pie2_Chart.spines['left'].set_color('#F0F0F0')
        self.Bond_pie2_Chart.spines['right'].set_color('#F0F0F0')
        self.Bond_pie2_Chart.xaxis.label.set_color('#F0F0F0')
        self.Bond_pie2_Chart.yaxis.label.set_color('#F0F0F0')
        self.Bond_pie2_Chart.tick_params(axis='x', colors='#F0F0F0')
        self.Bond_pie2_Chart.tick_params(axis='y', colors='black')
        
        
        self.Bond_pie2_canvas = FigureCanvasTkAgg(self.Bond_pie2_fig, master=self.Portfolio5_frame)
        self.Bond_pie2_canvas.get_tk_widget().place(x=1000,y=80)
        self.Bond_pie2_canvas.get_tk_widget().configure(background='#000000',  highlightcolor='#FFFFFF', highlightbackground='#FFFFFF')
        '''

        self.Bond_Tree_table = ttk.Treeview(self.Portfolio5_frame, height="10")

        self.Bond_Tree_table["columns"] = ("column1", "column2", 'column3', 'column4')
        self.Bond_Tree_table.column("#0", width=10, anchor='e')
        self.Bond_Tree_table.column("column1", width=60, anchor='center')
        self.Bond_Tree_table.column("column2", width=80, anchor='w')
        self.Bond_Tree_table.column("column3", width=250, anchor='w')
        self.Bond_Tree_table.column("column4", width=130, anchor='center')

        self.Bond_Tree_table.heading('#0', text='')
        self.Bond_Tree_table.heading("column1", text="Category")
        self.Bond_Tree_table.heading("column2", text="Ticker")
        self.Bond_Tree_table.heading("column3", text="Name")
        self.Bond_Tree_table.heading("column4", text="Weight(%)",
                                     command=lambda: self.treeview_sort_column(self.Bond_Tree_table, "column4", False))

        self.Bond_Tree_table.place(x=630, y=340)

        self.Bond_Return_StrVar = tk.StringVar()
        self.Bond_Volatility_StrVar = tk.StringVar()
        self.Bond_Sharpe_StrVar = tk.StringVar()

        self.Bond_Return = tk.Label(self.Portfolio5_frame, textvariable=self.Bond_Return_StrVar)
        self.Bond_Return.place(x=50, y=470)

        self.Bond_Volatility = tk.Label(self.Portfolio5_frame, textvariable=self.Bond_Volatility_StrVar)
        self.Bond_Volatility.place(x=50, y=500)

        self.Bond_Sharpe = tk.Label(self.Portfolio5_frame, textvariable=self.Bond_Sharpe_StrVar)
        self.Bond_Sharpe.place(x=50, y=530)

        self.Bond_dollar_entry = tk.Entry(self.Portfolio5_frame)
        self.Bond_dollar_entry.place(x=630, y=312)

        self.Bond_dollar_weight = ttk.Button(self.Portfolio5_frame, text='Portfolio Weight in Dollar',
                                             command=lambda: self.Portfolio_dollar(Entry=self.Bond_dollar_entry,
                                                                                   Tree=self.Bond_Tree_table))
        self.Bond_dollar_weight.place(x=750, y=310)

        #self.Bond_Performance_Output = ttk.Button(self.Portfolio5_frame, text='Performance Output',
                                                  #command=lambda: self.Performance_print(Bond_Portfolio_compare,
                                                                                         #'Strategic_Bond_Performance.xlsx',Strategic_Bond_Weight,Bond_AR,Bond_Portfolio_Std,Bond_Sharpe))
        #self.Bond_Performance_Output.place(x=600, y=530)

        # ======================================================================

        # ======================================================================


        self.GMP_fig = Figure(figsize=(6, 4), dpi=100)
        self.GMP_fig.set_tight_layout(True)
        self.GMP_fig.patch.set_facecolor('#F0F0F0')

        self.GMP_Chart = self.GMP_fig.add_subplot(111, axisbg='#F0F0F0')
        self.GMP_Chart.tick_params(axis='both', which='major', labelsize=8)
        self.GMP_Chart.set_xlabel('Time', fontsize=8)
        self.GMP_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.GMP_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.GMP_Chart.spines['bottom'].set_color('black')
        self.GMP_Chart.spines['top'].set_color('black')
        self.GMP_Chart.spines['left'].set_color('black')
        self.GMP_Chart.spines['right'].set_color('black')
        self.GMP_Chart.xaxis.label.set_color('black')
        self.GMP_Chart.yaxis.label.set_color('black')

        self.GMP_canvas = FigureCanvasTkAgg(self.GMP_fig, master=self.Portfolio6_frame)
        self.GMP_canvas.get_tk_widget().place(x=0, y=0)
        self.GMP_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                  highlightbackground='#FFFFFF')

        # Conservative frame setting
        self.GMP_Title = tk.Label(self.Portfolio6_frame, text='Global Momentum Portfolio', font=('Arial', 22, 'bold'))
        self.GMP_Title.place(x=620, y=10)

        # self.Conservative_description = tk.Label(self.Portfolio3_frame,text='20% of Equity, 80% of Fixed income',font=('Arial',14))
        # self.Conservative_description.place(x=750,y=60)

        ttk.Separator(self.Portfolio6_frame, orient=tk.HORIZONTAL).place(x=620, y=50, width=500)

        # Pie Chart
        self.GMP_pie_fig = Figure(figsize=(4, 4), dpi=60)
        self.GMP_pie_fig.set_tight_layout(True)
        self.GMP_pie_fig.patch.set_facecolor('#F0F0F0')

        self.GMP_pie_Chart = self.GMP_pie_fig.add_subplot(111, axisbg='#F0F0F0')
        self.GMP_pie_Chart.tick_params(axis='both', which='major', labelsize=8)
        # self.Aggressive_pie_Chart.set_xlabel('Time', fontsize=8)
        # self.Aggressive_pie_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.GMP_pie_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.GMP_pie_Chart.spines['bottom'].set_color('#F0F0F0')
        self.GMP_pie_Chart.spines['top'].set_color('#F0F0F0')
        self.GMP_pie_Chart.spines['left'].set_color('#F0F0F0')
        self.GMP_pie_Chart.spines['right'].set_color('#F0F0F0')
        self.GMP_pie_Chart.xaxis.label.set_color('#F0F0F0')
        self.GMP_pie_Chart.yaxis.label.set_color('#F0F0F0')
        self.GMP_pie_Chart.tick_params(axis='x', colors='#F0F0F0')
        self.GMP_pie_Chart.tick_params(axis='y', colors='black')

        self.GMP_pie_canvas = FigureCanvasTkAgg(self.GMP_pie_fig, master=self.Portfolio6_frame)
        self.GMP_pie_canvas.get_tk_widget().place(x=600, y=60)
        self.GMP_pie_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                      highlightbackground='#FFFFFF')

        self.labels = 'Equity', 'Fixed Income'
        self.sizes = [20, 80]

        # self.GMP_pie_Chart.pie(self.sizes,colors=cm.Blues(np.arange(len(self.sizes))/float(len(self.sizes))), labels=self.labels, autopct='%1.1f%%', shadow=True, startangle=150, labeldistance=0.8)#,autopct='%1.1f%%'
        # self.GMP_pie_canvas.show()



        # Pie Chart
        self.GMP_pie2_fig = Figure(figsize=(4, 4), dpi=60)
        self.GMP_pie2_fig.set_tight_layout(True)
        self.GMP_pie2_fig.patch.set_facecolor('#F0F0F0')

        self.GMP_pie2_Chart = self.GMP_pie2_fig.add_subplot(111, axisbg='#F0F0F0')
        self.GMP_pie2_Chart.tick_params(axis='both', which='major', labelsize=8)
        # self.Conservative_pie2_Chart.set_xlabel('Time', fontsize=8)
        # self.Conservative_pie2_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.GMP_pie2_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.GMP_pie2_Chart.spines['bottom'].set_color('#F0F0F0')
        self.GMP_pie2_Chart.spines['top'].set_color('#F0F0F0')
        self.GMP_pie2_Chart.spines['left'].set_color('#F0F0F0')
        self.GMP_pie2_Chart.spines['right'].set_color('#F0F0F0')
        self.GMP_pie2_Chart.xaxis.label.set_color('#F0F0F0')
        self.GMP_pie2_Chart.yaxis.label.set_color('#F0F0F0')
        self.GMP_pie2_Chart.tick_params(axis='x', colors='#F0F0F0')
        self.GMP_pie2_Chart.tick_params(axis='y', colors='black')

        self.GMP_pie2_canvas = FigureCanvasTkAgg(self.GMP_pie2_fig, master=self.Portfolio6_frame)
        self.GMP_pie2_canvas.get_tk_widget().place(x=900, y=60)
        self.GMP_pie2_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                       highlightbackground='#FFFFFF')

        self.GMP_Tree_table = ttk.Treeview(self.Portfolio6_frame, height="10")

        self.GMP_Tree_table["columns"] = ("column1", "column2", 'column3', 'column4')
        self.GMP_Tree_table.column("#0", width=10, anchor='e')
        self.GMP_Tree_table.column("column1", width=60, anchor='center')
        self.GMP_Tree_table.column("column2", width=80, anchor='w')
        self.GMP_Tree_table.column("column3", width=250, anchor='w')
        self.GMP_Tree_table.column("column4", width=130, anchor='center')

        self.GMP_Tree_table.heading('#0', text='')
        self.GMP_Tree_table.heading("column1", text="Category")
        self.GMP_Tree_table.heading("column2", text="Ticker")
        self.GMP_Tree_table.heading("column3", text="Name")
        self.GMP_Tree_table.heading("column4", text="Weight(%)",
                                    command=lambda: self.treeview_sort_column(self.GMP_Tree_table, "column4", False))

        self.GMP_Tree_table.place(x=630, y=340)

        self.GMP_Return_StrVar = tk.StringVar()
        self.GMP_Volatility_StrVar = tk.StringVar()
        self.GMP_Sharpe_StrVar = tk.StringVar()

        self.GMP_Return = tk.Label(self.Portfolio6_frame, textvariable=self.GMP_Return_StrVar)
        self.GMP_Return.place(x=50, y=470)

        self.GMP_Volatility = tk.Label(self.Portfolio6_frame, textvariable=self.GMP_Volatility_StrVar)
        self.GMP_Volatility.place(x=50, y=500)

        self.GMP_Sharpe = tk.Label(self.Portfolio6_frame, textvariable=self.GMP_Sharpe_StrVar)
        self.GMP_Sharpe.place(x=50, y=530)

        self.GMP_dollar_entry = tk.Entry(self.Portfolio6_frame)
        self.GMP_dollar_entry.place(x=630, y=312)

        self.GMP_dollar_weight = ttk.Button(self.Portfolio6_frame, text='Portfolio Weight in Dollar',
                                            command=lambda: self.Portfolio_dollar(Entry=self.GMP_dollar_entry,
                                                                                  Tree=self.GMP_Tree_table))
        self.GMP_dollar_weight.place(x=780, y=310)

        #self.GMP_Performance_Output = ttk.Button(self.Portfolio6_frame, text='Performance Output',
                                                 #command=lambda: self.Performance_print(GMP_compare,
                                                                                        #'Global_Momentum_Portfolio_Performance.xlsx',GMP_Weight_df,GMP_AR,GMP_Std,GMP_Sharpe))
        #self.GMP_Performance_Output.place(x=600, y=530)

        # ======================================================================

        # ======================================================================






        # ======================================================================

        self.Equity_Title = tk.Label(self.Portfolio7_frame, text='Equity Holdings', font=('Arial', 18, 'bold'))
        self.Equity_Title.place(x=200, y=0)

        self.Equity_fig = Figure(figsize=(6, 3), dpi=100)
        self.Equity_fig.set_tight_layout(True)
        self.Equity_fig.patch.set_facecolor('#F0F0F0')

        self.Equity_Chart = self.Equity_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Equity_Chart.tick_params(axis='both', which='major', labelsize=8)
        self.Equity_Chart.set_xlabel('Time', fontsize=8)
        self.Equity_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.Equity_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Equity_Chart.spines['bottom'].set_color('black')
        self.Equity_Chart.spines['top'].set_color('black')
        self.Equity_Chart.spines['left'].set_color('black')
        self.Equity_Chart.spines['right'].set_color('black')
        self.Equity_Chart.xaxis.label.set_color('black')
        self.Equity_Chart.yaxis.label.set_color('black')

        self.Equity_canvas = FigureCanvasTkAgg(self.Equity_fig, master=self.Portfolio7_frame)
        self.Equity_canvas.get_tk_widget().place(x=0, y=30)
        self.Equity_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                     highlightbackground='#FFFFFF')

        self.Equity_Tree_table = ttk.Treeview(self.Portfolio7_frame, height="8")

        self.Equity_Tree_table["columns"] = ("column1", "column2", 'column3', 'column4')
        self.Equity_Tree_table.column("#0", width=10, anchor='e')
        self.Equity_Tree_table.column("column1", width=110, anchor='w')
        self.Equity_Tree_table.column("column2", width=230, anchor='w')
        self.Equity_Tree_table.column("column3", width=110, anchor='center')
        self.Equity_Tree_table.column("column4", width=110, anchor='center')

        self.Equity_Tree_table.heading('#0', text='')
        self.Equity_Tree_table.heading("column1", text="Ticker",
                                       command=lambda: self.treeview_string_sort_column(self.Equity_Tree_table,
                                                                                        "column1", False))
        self.Equity_Tree_table.heading("column2", text="Name",
                                       command=lambda: self.treeview_string_sort_column(self.Equity_Tree_table,
                                                                                        "column2", False))
        self.Equity_Tree_table.heading("column3", text="Volatility(%)",
                                       command=lambda: self.treeview_sort_column(self.Equity_Tree_table, "column3",
                                                                                 False))
        self.Equity_Tree_table.heading("column4", text="Weight(%)",
                                       command=lambda: self.treeview_sort_column(self.Equity_Tree_table, "column4",
                                                                                 False))

        self.Equity_Tree_table.place(x=20, y=320)

        self.Tickers_list = Database_Functions.Fetch_All_Tickers()
        self.Combined_list = self.Combined_list(list1=self.Tickers_list[0], list2=self.Tickers_list[1])

        self.Equity_Ticker_Entry = AutocompleteEntry(self.Combined_list, self.Portfolio7_frame, width=55, bg='#FF9C29')
        self.Equity_Ticker_Entry.place(x=20, y=520)

        self.Equity_Add_Button = ttk.Button(self.Portfolio7_frame, text='Add', width=10,
                                            command=lambda: self.Insert_holdings(Tree=self.Equity_Tree_table,
                                                                                 Entry=self.Equity_Ticker_Entry))
        self.Equity_Add_Button.place(x=417, y=517)

        self.Equity_Delete_Button = ttk.Button(self.Portfolio7_frame, text='Delete', width=10,
                                               command=self.Delete_holdings_Equity)
        self.Equity_Delete_Button.place(x=507, y=517)

        self.Fetch_Equity_holdings()

        self.Equity_Return_StrVar = tk.StringVar()
        self.Equity_Volatility_StrVar = tk.StringVar()
        self.Equity_Sharpe_StrVar = tk.StringVar()

        self.Equity_Return_label = tk.Label(self.Portfolio7_frame, textvariable=self.Equity_Return_StrVar)
        self.Equity_Return_label.place(x=55, y=60)
        self.Equity_Volatility_label = tk.Label(self.Portfolio7_frame, textvariable=self.Equity_Volatility_StrVar)
        self.Equity_Volatility_label.place(x=55, y=80)
        self.Equity_Sharpe_label = tk.Label(self.Portfolio7_frame, textvariable=self.Equity_Sharpe_StrVar)
        self.Equity_Sharpe_label.place(x=55, y=100)

        ttk.Separator(self.Portfolio7_frame, orient=tk.VERTICAL).place(x=620, y=10, height=520)

        self.Fixed_Income_Title = tk.Label(self.Portfolio7_frame, text='Fixed Income Holdings',
                                           font=('Arial', 18, 'bold'))
        self.Fixed_Income_Title.place(x=850, y=0)

        self.Fixed_Income_fig = Figure(figsize=(6, 3), dpi=100)
        self.Fixed_Income_fig.set_tight_layout(True)
        self.Fixed_Income_fig.patch.set_facecolor('#F0F0F0')

        self.Fixed_Income_Chart = self.Fixed_Income_fig.add_subplot(111, axisbg='#F0F0F0')
        self.Fixed_Income_Chart.tick_params(axis='both', which='major', labelsize=8)
        self.Fixed_Income_Chart.set_xlabel('Time', fontsize=8)
        self.Fixed_Income_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.Fixed_Income_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.Fixed_Income_Chart.spines['bottom'].set_color('black')
        self.Fixed_Income_Chart.spines['top'].set_color('black')
        self.Fixed_Income_Chart.spines['left'].set_color('black')
        self.Fixed_Income_Chart.spines['right'].set_color('black')
        self.Fixed_Income_Chart.xaxis.label.set_color('black')
        self.Fixed_Income_Chart.yaxis.label.set_color('black')

        self.Fixed_Income_canvas = FigureCanvasTkAgg(self.Fixed_Income_fig, master=self.Portfolio7_frame)
        self.Fixed_Income_canvas.get_tk_widget().place(x=630, y=30)
        self.Fixed_Income_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                           highlightbackground='#FFFFFF')

        self.Fixed_Income_Tree_table = ttk.Treeview(self.Portfolio7_frame, height="8")

        self.Fixed_Income_Tree_table["columns"] = ("column1", "column2", 'column3', 'column4')
        self.Fixed_Income_Tree_table.column("#0", width=10, anchor='e')
        self.Fixed_Income_Tree_table.column("column1", width=110, anchor='w')
        self.Fixed_Income_Tree_table.column("column2", width=230, anchor='w')
        self.Fixed_Income_Tree_table.column("column3", width=110, anchor='center')
        self.Fixed_Income_Tree_table.column("column4", width=110, anchor='center')

        self.Fixed_Income_Tree_table.heading('#0', text='')
        self.Fixed_Income_Tree_table.heading("column1", text="Ticker", command=lambda: self.treeview_string_sort_column(
            self.Fixed_Income_Tree_table, "column1", False))
        self.Fixed_Income_Tree_table.heading("column2", text="Name", command=lambda: self.treeview_string_sort_column(
            self.Fixed_Income_Tree_table, "column2", False))
        self.Fixed_Income_Tree_table.heading("column3", text="Volatility(%)",
                                             command=lambda: self.treeview_sort_column(self.Fixed_Income_Tree_table,
                                                                                       "column3", False))
        self.Fixed_Income_Tree_table.heading("column4", text="Weight(%)",
                                             command=lambda: self.treeview_sort_column(self.Fixed_Income_Tree_table,
                                                                                       "column4", False))

        self.Fixed_Income_Tree_table.place(x=650, y=320)

        self.Fixed_Income_Ticker_Entry = AutocompleteEntry(self.Combined_list, self.Portfolio7_frame, width=55,
                                                           bg='#FF9C29')
        self.Fixed_Income_Ticker_Entry.place(x=650, y=520)

        self.Fixed_Income_Add_Button = ttk.Button(self.Portfolio7_frame, text='Add', width=10,
                                                  command=lambda: self.Insert_holdings(
                                                      Tree=self.Fixed_Income_Tree_table,
                                                      Entry=self.Fixed_Income_Ticker_Entry))
        self.Fixed_Income_Add_Button.place(x=1050, y=517)

        self.Fixed_Income_Delete_Button = ttk.Button(self.Portfolio7_frame, text='Delete', width=10,
                                                     command=self.Delete_holdings_Fixed_income)
        self.Fixed_Income_Delete_Button.place(x=1140, y=517)

        self.Fetch_Fixed_income_holdings()

        self.Fixed_income_Return_StrVar = tk.StringVar()
        self.Fixed_income_Volatility_StrVar = tk.StringVar()
        self.Fixed_income_Sharpe_StrVar = tk.StringVar()

        self.Fixed_income_Return_label = tk.Label(self.Portfolio7_frame, textvariable=self.Fixed_income_Return_StrVar)
        self.Fixed_income_Return_label.place(x=680, y=60)
        self.Fixed_income_Volatility_label = tk.Label(self.Portfolio7_frame,
                                                      textvariable=self.Fixed_income_Volatility_StrVar)
        self.Fixed_income_Volatility_label.place(x=680, y=80)
        self.Fixed_income_Sharpe_label = tk.Label(self.Portfolio7_frame, textvariable=self.Fixed_income_Sharpe_StrVar)
        self.Fixed_income_Sharpe_label.place(x=680, y=100)

        # self.Today_str
        self.Last_StrVar = tk.StringVar()
        self.Last_label = tk.Label(self.Portfolio7_frame, textvariable=self.Last_StrVar, font=('Arial', 12, 'bold'))
        self.Last_label.place(x=10, y=10)

        self.Conservative_Last_Label = tk.Label(self.Portfolio3_frame, textvariable=self.Last_StrVar,
                                                font=('Arial', 12, 'bold'))
        self.Conservative_Last_Label.place(x=1190, y=310)

        self.Moderate_Last_Label = tk.Label(self.Portfolio2_frame, textvariable=self.Last_StrVar,
                                            font=('Arial', 12, 'bold'))
        self.Moderate_Last_Label.place(x=1190, y=310)

        self.Aggressive_Last_Label = tk.Label(self.Portfolio1_frame, textvariable=self.Last_StrVar,
                                              font=('Arial', 12, 'bold'))
        self.Aggressive_Last_Label.place(x=1190, y=310)




        # Six Balance Portfolio===============================================================================================================
        self.BP_fig = Figure(figsize=(6, 4), dpi=100)
        self.BP_fig.set_tight_layout(True)
        self.BP_fig.patch.set_facecolor('#F0F0F0')

        self.BP_Chart = self.BP_fig.add_subplot(111, axisbg='#F0F0F0')
        self.BP_Chart.tick_params(axis='both', which='major', labelsize=8)
        self.BP_Chart.set_xlabel('Time', fontsize=8)
        self.BP_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.BP_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.BP_Chart.spines['bottom'].set_color('black')
        self.BP_Chart.spines['top'].set_color('black')
        self.BP_Chart.spines['left'].set_color('black')
        self.BP_Chart.spines['right'].set_color('black')
        self.BP_Chart.xaxis.label.set_color('black')
        self.BP_Chart.yaxis.label.set_color('black')

        self.BP_canvas = FigureCanvasTkAgg(self.BP_fig, master=self.Portfolio8_frame)
        self.BP_canvas.get_tk_widget().place(x=0, y=0)
        self.BP_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                           highlightbackground='#FFFFFF')

        # Balance Portfolio frame setting
        self.BP_Title = tk.Label(self.Portfolio8_frame, text='Six Balance Portfolio',
                                           font=('Arial', 22, 'bold'))
        self.BP_Title.place(x=620, y=10)

        ttk.Separator(self.Portfolio8_frame, orient=tk.HORIZONTAL).place(x=620, y=50, width=500)

        self.BP_Equity_Benchmark = tk.Label(self.Portfolio8_frame, text='Equity: MSCI World Index',
                                            font=('Arial',12,'bold'))
        self.BP_Equity_Benchmark.place(x=620,y=60)

        self.BP_FI_Benchmark = tk.Label(self.Portfolio8_frame, text='Fixed Income: Barclays U.S. Aggregate Index',
                                            font=('Arial',12,'bold'))
        self.BP_FI_Benchmark.place(x=620,y=90)

        #Horizontal line of table
        ttk.Separator(self.Portfolio8_frame, orient=tk.HORIZONTAL).place(x=620, y=180, width=550)
        ttk.Separator(self.Portfolio8_frame, orient=tk.HORIZONTAL).place(x=620, y=370, width=550)

        self.BP_table_Name1 = tk.Label(self.Portfolio8_frame, text='Fixed Income (20,80)',font=('Arial',12,'bold'))
        self.BP_table_Name2 = tk.Label(self.Portfolio8_frame, text='Conservative (30,70)',font=('Arial',12,'bold'))
        self.BP_table_Name3 = tk.Label(self.Portfolio8_frame, text='Moderate (40,60)',font=('Arial',12,'bold'))
        self.BP_table_Name4 = tk.Label(self.Portfolio8_frame, text='Growth (50,50)',font=('Arial',12,'bold'))
        self.BP_table_Name5 = tk.Label(self.Portfolio8_frame, text='Aggressive (60,40)',font=('Arial',12,'bold'))
        self.BP_table_Name6 = tk.Label(self.Portfolio8_frame, text='Very Aggressive (70,30)',font=('Arial',12,'bold'))
        self.BP_table_Name7 = tk.Label(self.Portfolio8_frame, text='MSCI World Index',font=('Arial',12,'bold'))
        self.BP_table_Name8 = tk.Label(self.Portfolio8_frame, text='Barclays U.S. Aggregate Index',font=('Arial',12,'bold'))

        self.BP_table_Name1.place(x=620,y=190)
        self.BP_table_Name2.place(x=620,y=220)
        self.BP_table_Name3.place(x=620,y=250)
        self.BP_table_Name4.place(x=620,y=280)
        self.BP_table_Name5.place(x=620,y=310)
        self.BP_table_Name6.place(x=620,y=340)
        self.BP_table_Name7.place(x=620,y=380)
        self.BP_table_Name8.place(x=620,y=410)

        self.BP_AR = tk.Label(self.Portfolio8_frame, text='Annual Return',font=('Arial',12,'bold'))
        self.BP_Vol = tk.Label(self.Portfolio8_frame, text='Volatility',font=('Arial',12,'bold'))
        self.BP_Sharpe = tk.Label(self.Portfolio8_frame, text='Sharpe',font=('Arial',12,'bold'))

        self.BP_AR.place(x=880,y=150)
        self.BP_Vol.place(x=1010,y=150)
        self.BP_Sharpe.place(x=1095,y=150)

        #Performance figure: Annual Return
        self.AR1_StrVar = tk.StringVar()
        self.AR1_label = tk.Label(self.Portfolio8_frame, textvariable=self.AR1_StrVar, font=('Arial', 12, 'bold'))
        self.AR1_label.place(x=990, y=190,anchor='ne')

        self.AR2_StrVar = tk.StringVar()
        self.AR2_label = tk.Label(self.Portfolio8_frame, textvariable=self.AR2_StrVar, font=('Arial', 12, 'bold'))
        self.AR2_label.place(x=990, y=220,anchor='ne')

        self.AR3_StrVar = tk.StringVar()
        self.AR3_label = tk.Label(self.Portfolio8_frame, textvariable=self.AR3_StrVar, font=('Arial', 12, 'bold'))
        self.AR3_label.place(x=990, y=250,anchor='ne')

        self.AR4_StrVar = tk.StringVar()
        self.AR4_label = tk.Label(self.Portfolio8_frame, textvariable=self.AR4_StrVar, font=('Arial', 12, 'bold'))
        self.AR4_label.place(x=990, y=280,anchor='ne')

        self.AR5_StrVar = tk.StringVar()
        self.AR5_label = tk.Label(self.Portfolio8_frame, textvariable=self.AR5_StrVar, font=('Arial', 12, 'bold'))
        self.AR5_label.place(x=990, y=310,anchor='ne')

        self.AR6_StrVar = tk.StringVar()
        self.AR6_label = tk.Label(self.Portfolio8_frame, textvariable=self.AR6_StrVar, font=('Arial', 12, 'bold'))
        self.AR6_label.place(x=990, y=340,anchor='ne')

        self.AR7_StrVar = tk.StringVar()
        self.AR7_label = tk.Label(self.Portfolio8_frame, textvariable=self.AR7_StrVar, font=('Arial', 12, 'bold'))
        self.AR7_label.place(x=990, y=380,anchor='ne')

        self.AR8_StrVar = tk.StringVar()
        self.AR8_label = tk.Label(self.Portfolio8_frame, textvariable=self.AR8_StrVar, font=('Arial', 12, 'bold'))
        self.AR8_label.place(x=990, y=410,anchor='ne')


        #Performance figure: Volatility
        self.Vol1_StrVar = tk.StringVar()
        self.Vol1_label = tk.Label(self.Portfolio8_frame, textvariable=self.Vol1_StrVar, font=('Arial', 12, 'bold'))
        self.Vol1_label.place(x=1075, y=190,anchor='ne')

        self.Vol2_StrVar = tk.StringVar()
        self.Vol2_label = tk.Label(self.Portfolio8_frame, textvariable=self.Vol2_StrVar, font=('Arial', 12, 'bold'))
        self.Vol2_label.place(x=1075, y=220,anchor='ne')

        self.Vol3_StrVar = tk.StringVar()
        self.Vol3_label = tk.Label(self.Portfolio8_frame, textvariable=self.Vol3_StrVar, font=('Arial', 12, 'bold'))
        self.Vol3_label.place(x=1075, y=250,anchor='ne')

        self.Vol4_StrVar = tk.StringVar()
        self.Vol4_label = tk.Label(self.Portfolio8_frame, textvariable=self.Vol4_StrVar, font=('Arial', 12, 'bold'))
        self.Vol4_label.place(x=1075, y=280,anchor='ne')

        self.Vol5_StrVar = tk.StringVar()
        self.Vol5_label = tk.Label(self.Portfolio8_frame, textvariable=self.Vol5_StrVar, font=('Arial', 12, 'bold'))
        self.Vol5_label.place(x=1075, y=310,anchor='ne')

        self.Vol6_StrVar = tk.StringVar()
        self.Vol6_label = tk.Label(self.Portfolio8_frame, textvariable=self.Vol6_StrVar, font=('Arial', 12, 'bold'))
        self.Vol6_label.place(x=1075, y=340,anchor='ne')

        self.Vol7_StrVar = tk.StringVar()
        self.Vol7_label = tk.Label(self.Portfolio8_frame, textvariable=self.Vol7_StrVar, font=('Arial', 12, 'bold'))
        self.Vol7_label.place(x=1075, y=380,anchor='ne')

        self.Vol8_StrVar = tk.StringVar()
        self.Vol8_label = tk.Label(self.Portfolio8_frame, textvariable=self.Vol8_StrVar, font=('Arial', 12, 'bold'))
        self.Vol8_label.place(x=1075, y=410,anchor='ne')

        #Performance figure: Sharpe
        self.Sharpe1_StrVar = tk.StringVar()
        self.Sharpe1_label = tk.Label(self.Portfolio8_frame, textvariable=self.Sharpe1_StrVar, font=('Arial', 12, 'bold'))
        self.Sharpe1_label.place(x=1145, y=190,anchor='ne')

        self.Sharpe2_StrVar = tk.StringVar()
        self.Sharpe2_label = tk.Label(self.Portfolio8_frame, textvariable=self.Sharpe2_StrVar, font=('Arial', 12, 'bold'))
        self.Sharpe2_label.place(x=1145, y=220,anchor='ne')

        self.Sharpe3_StrVar = tk.StringVar()
        self.Sharpe3_label = tk.Label(self.Portfolio8_frame, textvariable=self.Sharpe3_StrVar, font=('Arial', 12, 'bold'))
        self.Sharpe3_label.place(x=1145, y=250,anchor='ne')

        self.Sharpe4_StrVar = tk.StringVar()
        self.Sharpe4_label = tk.Label(self.Portfolio8_frame, textvariable=self.Sharpe4_StrVar, font=('Arial', 12, 'bold'))
        self.Sharpe4_label.place(x=1145, y=280,anchor='ne')

        self.Sharpe5_StrVar = tk.StringVar()
        self.Sharpe5_label = tk.Label(self.Portfolio8_frame, textvariable=self.Sharpe5_StrVar, font=('Arial', 12, 'bold'))
        self.Sharpe5_label.place(x=1145, y=310,anchor='ne')

        self.Sharpe6_StrVar = tk.StringVar()
        self.Sharpe6_label = tk.Label(self.Portfolio8_frame, textvariable=self.Sharpe6_StrVar, font=('Arial', 12, 'bold'))
        self.Sharpe6_label.place(x=1145, y=340,anchor='ne')

        self.Sharpe7_StrVar = tk.StringVar()
        self.Sharpe7_label = tk.Label(self.Portfolio8_frame, textvariable=self.Sharpe7_StrVar, font=('Arial', 12, 'bold'))
        self.Sharpe7_label.place(x=1145, y=380,anchor='ne')

        self.Sharpe8_StrVar = tk.StringVar()
        self.Sharpe8_label = tk.Label(self.Portfolio8_frame, textvariable=self.Sharpe8_StrVar, font=('Arial', 12, 'bold'))
        self.Sharpe8_label.place(x=1145, y=410,anchor='ne')

        #self.Six_Portfolio_Output_Button = ttk.Button(self.Portfolio8_frame, text='Performance Output', width=20,command = self.Six_Portfolio_to_excel)
        #self.Six_Portfolio_Output_Button.place(x=600, y=530)

        #===============================================================================================================
        # All Weather Portfolio=========================================================================================
        self.AWP_fig = Figure(figsize=(6, 4), dpi=100)
        self.AWP_fig.set_tight_layout(True)
        self.AWP_fig.patch.set_facecolor('#F0F0F0')

        self.AWP_Chart = self.AWP_fig.add_subplot(111, axisbg='#F0F0F0')
        self.AWP_Chart.tick_params(axis='both', which='major', labelsize=8)
        self.AWP_Chart.set_xlabel('Time', fontsize=8)
        self.AWP_Chart.set_ylabel('Cumulative Return', fontsize=8)

        self.AWP_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.AWP_Chart.spines['bottom'].set_color('black')
        self.AWP_Chart.spines['top'].set_color('black')
        self.AWP_Chart.spines['left'].set_color('black')
        self.AWP_Chart.spines['right'].set_color('black')
        self.AWP_Chart.xaxis.label.set_color('black')
        self.AWP_Chart.yaxis.label.set_color('black')

        self.AWP_canvas = FigureCanvasTkAgg(self.AWP_fig, master=self.Portfolio9_frame)
        self.AWP_canvas.get_tk_widget().place(x=0, y=0)
        self.AWP_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                           highlightbackground='#FFFFFF')

        # Balance Portfolio frame setting
        self.AWP_Title = tk.Label(self.Portfolio9_frame, text='All Weather Portfolio',
                                           font=('Arial', 22, 'bold'))
        self.AWP_Title.place(x=620, y=10)
        ttk.Separator(self.Portfolio9_frame, orient=tk.HORIZONTAL).place(x=620, y=50, width=500)

        # Pie Chart
        self.AWP_pie_fig = Figure(figsize=(4, 4), dpi=60)
        self.AWP_pie_fig.set_tight_layout(True)
        self.AWP_pie_fig.patch.set_facecolor('#F0F0F0')

        self.AWP_pie_Chart = self.AWP_pie_fig.add_subplot(111, axisbg='#F0F0F0')
        self.AWP_pie_Chart.tick_params(axis='both', which='major', labelsize=8)
        # self.Aggressive_pie_Chart.set_xlabel('Time', fontsize=8)
        # self.Aggressive_pie_Chart.set_ylabel('Cumulative Return', fontsize=8)
        self.AWP_pie_Chart.tick_params(axis='both', which='major', colors='black', labelsize=6)
        self.AWP_pie_Chart.spines['bottom'].set_color('#F0F0F0')
        self.AWP_pie_Chart.spines['top'].set_color('#F0F0F0')
        self.AWP_pie_Chart.spines['left'].set_color('#F0F0F0')
        self.AWP_pie_Chart.spines['right'].set_color('#F0F0F0')
        self.AWP_pie_Chart.xaxis.label.set_color('#F0F0F0')
        self.AWP_pie_Chart.yaxis.label.set_color('#F0F0F0')
        self.AWP_pie_Chart.tick_params(axis='x', colors='#F0F0F0')
        self.AWP_pie_Chart.tick_params(axis='y', colors='black')

        self.AWP_pie_canvas = FigureCanvasTkAgg(self.AWP_pie_fig, master=self.Portfolio9_frame)
        self.AWP_pie_canvas.get_tk_widget().place(x=620, y=80)
        self.AWP_pie_canvas.get_tk_widget().configure(background='#000000', highlightcolor='#FFFFFF',
                                                       highlightbackground='#FFFFFF')

        self.AWP_Tree_table = ttk.Treeview(self.Portfolio9_frame, height="10")

        self.AWP_Tree_table["columns"] = ("column1", "column2", 'column3', 'column4')
        self.AWP_Tree_table.column("#0", width=10, anchor='e')
        self.AWP_Tree_table.column("column1", width=60, anchor='center')
        self.AWP_Tree_table.column("column2", width=80, anchor='w')
        self.AWP_Tree_table.column("column3", width=250, anchor='w')
        self.AWP_Tree_table.column("column4", width=130, anchor='center')

        self.AWP_Tree_table.heading('#0', text='')
        self.AWP_Tree_table.heading("column1", text="Category")
        self.AWP_Tree_table.heading("column2", text="Ticker")
        self.AWP_Tree_table.heading("column3", text="Name")
        self.AWP_Tree_table.heading("column4", text="Weight(%)",
                                    command=lambda: self.treeview_sort_column(self.AWP_Tree_table, "column4", False))

        self.AWP_Tree_table.place(x=620, y=340)

        self.Quandrant_StrVar = tk.StringVar()
        self.Quandrant_label = tk.Label(self.Portfolio9_frame, textvariable=self.Quandrant_StrVar, font=('Arial', 24, 'bold'))
        #self.Quandrant_label = tk.Label(self.Portfolio9_frame, text='III', font=('Arial', 24, 'bold'))
        self.Quandrant_label.place(x=850, y=60)


        self.Quandrant_Return_StrVar = tk.StringVar()
        self.Quandrant_Volatility_StrVar = tk.StringVar()
        self.Quandrant_Sharpe_StrVar = tk.StringVar()

        self.Quandrant_Return = tk.Label(self.Portfolio9_frame, textvariable=self.Quandrant_Return_StrVar)
        self.Quandrant_Return.place(x=50, y=470)

        self.Quandrant_Volatility = tk.Label(self.Portfolio9_frame, textvariable=self.Quandrant_Volatility_StrVar)
        self.Quandrant_Volatility.place(x=50, y=500)

        self.Quandrant_Sharpe = tk.Label(self.Portfolio9_frame, textvariable=self.Quandrant_Sharpe_StrVar)
        self.Quandrant_Sharpe.place(x=50, y=530)


        #self.AWP_Output_Button = ttk.Button(self.Portfolio9_frame, text='Output as Excel', width=20,command = self.AWP_output)
        #self.AWP_Output_Button.place(x=600, y=530)







        #===============================================================================================================



        self.Save_Reload()

        self.Reload_Button = ttk.Button(self.Portfolio7_frame, text='Save & Reload', width=20, command=self.Save_Reload)
        self.Reload_Button.place(x=550, y=550)

        #self.Holdings_output_Button = ttk.Button(self.Portfolio7_frame, text='Output as Excel', width=20,command=self.Holdings_to_excel)
        #self.Holdings_output_Button.place(x=620, y=620)

        # ======================================================================
        self.root.mainloop()

        self.Equity_Benchmark = Database_Functions.Fetch(Ticker='MXWO Index').pct_change().dropna()
        self.FI_Benchmark = Database_Functions.Fetch(Ticker='LBUSTRUU Index').pct_change().dropna()
        #print self.Equity_Benchmark

    def Holdings_to_excel(self):
        #print Equity_Portfolio_CR
        #print FI_Portfolio_CR
        #print Equity_Portfolio_Weight
        #print Fixed_income_Portfolio_Weight

        Equity_CR = pd.DataFrame()
        Equity_CR['Equity - Cumulative Return'] = Equity_Portfolio_CR

        FI_CR = pd.DataFrame()
        FI_CR['Fixed Income - Cumulative Return'] = FI_Portfolio_CR
        #print FI_CR
        FileName = 'Holdings.xlsx'
        writer = pd.ExcelWriter("D:/Taishin_Platform/Portfolio_Holdings/" + str(FileName))

        Equity_CR.to_excel(writer, 'Equity - Cumulative Return')
        Equity_Portfolio_Weight.to_excel(writer, 'Equity - Weights')
        FI_CR.to_excel(writer, 'FI - Cumulative Return')
        Fixed_income_Portfolio_Weight.to_excel(writer, 'Fixed Income - Weights')

        writer.save()
        info = FileName + ' saved successfully.'
        self.Done_Messenger(info=info)

    def AWP_output(self):
        FileName = 'All_Weather_Portfolio.xlsx'
        writer = pd.ExcelWriter("D:/Taishin_Platform/Portfolio_Performance/" + str(FileName))
        Score_df = pd.DataFrame()
        Score_df['Quandrant'] = AWP_Score

        AWP_CR_df = pd.DataFrame()
        AWP_CR_df['Cumulative Return'] = AWP_CR

        AWP_CR_df.to_excel(writer, 'Cumulative Return')
        Score_df.to_excel(writer, 'Historical Quandrant')

        writer.save()
        info = FileName + ' saved successfully.'
        self.Done_Messenger(info=info)


        #print 'Done'
    def Six_Portfolio_to_excel(self):
        CR_df = All_Portfolo_Performace
        AR_df = (CR_df.resample('A',how='last').pct_change()*100.).dropna()
        MR_df = (CR_df.resample('M',how='last').pct_change()*100.).dropna()


        FileName = 'Six_Balanced_Portfolio.xlsx'
        writer = pd.ExcelWriter("D:/Taishin_Platform/Portfolio_Performance/" + str(FileName))

        CR_df.to_excel(writer, 'Summary')
        AR_df.to_excel(writer, 'Cumulative Return (Daily)')
        MR_df.to_excel(writer, 'Monthly Return')

        writer.save()
        info = FileName + ' saved successfully.'
        self.Done_Messenger(info=info)

    def Combined_list(self, list1, list2):
        if len(list1) != len(list2):
            print 'Lists number are not equal!!!!'
        else:
            Combined_list = []
            for i in range(len(list1)):
                Combined_list.append(list1[i] + ' , ' + list2[i])
        return Combined_list

    def Raise_frame(self, frame):
        frame.tkraise()



    def All_Weather_Portfolio(self,Lookback):
        def Equal_Weight_Portfolio(Tickers):
            Portfolio = pd.DataFrame()

            for Ticker in Tickers:
                Portfolio[Ticker] = Database_Functions.Fetch(Ticker)[Ticker]

            Portfolio_Return = Portfolio.pct_change().dropna()
            return Portfolio_Return.mean(axis=1)

        Growth_df = pd.read_excel("D:/Taishin_Platform/All_Weather_Portfolio/AW_data.xlsx", sheetname='Growth',
                                  index_col=0)
        Growth_df = Growth_df.dropna()
        Growth_df = Growth_df.sort_index()
        Growth_df.columns = ['US', 'Europe']
        Growth_df['Europe'] = Growth_df['Europe'].pct_change(12) * 100.
        Growth_df = Growth_df.round(1).dropna()
        Growth_df['Growth'] = Growth_df['US'] * 0.6 + Growth_df['Europe'] * 0.4

        Inflation_df = pd.read_excel("D:/Taishin_Platform/All_Weather_Portfolio/AW_data.xlsx", sheetname='Inflation',
                                     index_col=0)
        Inflation_df = Inflation_df.dropna()
        Inflation_df = Inflation_df.sort_index()
        Inflation_df.columns = ['US', 'Europe']
        Inflation_df['Europe'] = Inflation_df['Europe'].pct_change(12) * 100.
        Inflation_df = Inflation_df.round(1).dropna()
        Inflation_df['Inflation'] = Inflation_df['US'] * 0.6 + Inflation_df['Europe'] * 0.4

        Economic_Cycle = pd.concat([Growth_df['Growth'], Inflation_df['Inflation']], axis=1).dropna().diff(Lookback)
        Economic_Cycle['Quandrant'] = np.where((Economic_Cycle['Growth'] >= 0.) & (Economic_Cycle['Inflation'] >= 0.),
                                               1, np.nan)
        Economic_Cycle['Quandrant'] = np.where((Economic_Cycle['Growth'] <= 0.) & (Economic_Cycle['Inflation'] >= 0.),
                                               2, Economic_Cycle['Quandrant'])
        Economic_Cycle['Quandrant'] = np.where((Economic_Cycle['Growth'] <= 0.) & (Economic_Cycle['Inflation'] <= 0.),
                                               3, Economic_Cycle['Quandrant'])
        Economic_Cycle['Quandrant'] = np.where((Economic_Cycle['Growth'] >= 0.) & (Economic_Cycle['Inflation'] <= 0.),
                                               4, Economic_Cycle['Quandrant'])

        Economic_Cycle = Economic_Cycle.dropna()

        First_tickers = ['MXWO Index', 'LHVLTRUU Index', 'JPEICORE Index', 'LBUSTRUU Index', 'SPGSCI Index']
        First_Names = ['MSCI World Index','Barclays Capital High Yield Total Return Index Value Unhedged USD','JPMorgan EMBI Global Core Index',
                       'Barclays U.S. Aggregate Index','S&P GSCI Index']
        First_df = pd.DataFrame()
        First_df['Category'] = ['Equity','Fixed Income','Fixed Income','Fixed Income','Commodities']
        First_df['Ticker'] = First_tickers
        First_df['Name'] = First_Names
        First_df['Weight'] = [20,20,20,20,20]


        Second_tickers = ['JPEICORE Index', 'LBUSTRUU Index']
        Second_Names = ['JPMorgan EMBI Global Core Index','Barclays U.S. Aggregate Index']
        Second_df = pd.DataFrame()
        Second_df['Category'] = ['Fixed Income','Fixed Income']
        Second_df['Ticker'] = Second_tickers
        Second_df['Name'] = Second_Names
        Second_df['Weight'] = [50,50]


        Third_tickers = ['LBUSTRUU Index']
        Third_Names = ['Barclays U.S. Aggregate Index']
        Third_df = pd.DataFrame()
        Third_df['Category'] = ['Fixed Income']
        Third_df['Ticker'] = Third_tickers
        Third_df['Name'] = Third_Names
        Third_df['Weight'] = [100]


        Fourth_tickers = ['MXWO Index', 'LHVLTRUU Index', 'LBUSTRUU Index', 'SPGSCI Index']
        Fouth_Names = ['MSCI World Index','Barclays Capital High Yield Total Return Index Value Unhedged USD','Barclays U.S. Aggregate Index','S&P GSCI Index']
        Fourth_df = pd.DataFrame()
        Fourth_df['Category'] = ['Equity','Fixed Income','Fixed Income','Commodities']
        Fourth_df['Ticker'] = Fourth_tickers
        Fourth_df['Name'] = Fouth_Names
        Fourth_df['Weight'] = [25,25,25,25]


        All_Weather_df = pd.DataFrame()
        All_Weather_df['Quandrant'] = Economic_Cycle['Quandrant'].resample('D', how='last')
        All_Weather_df['Quandrant'] = All_Weather_df['Quandrant'].fillna(method='ffill')

        All_Weather_df = All_Weather_df[All_Weather_df.index > '2001-12-01']

        All_Weather_df['First'] = Equal_Weight_Portfolio(First_tickers)
        All_Weather_df['Second'] = Equal_Weight_Portfolio(Second_tickers)
        All_Weather_df['Third'] = Equal_Weight_Portfolio(Third_tickers)
        All_Weather_df['Fourth'] = Equal_Weight_Portfolio(Fourth_tickers)

        #print All_Weather_df[['First','Second','Third','Fourth']]

        All_Weather_df['Quandrant'] = All_Weather_df['Quandrant'].shift(20)
        All_Weather_df = All_Weather_df.dropna()

        All_Weather_df['AW_Return'] = np.where(All_Weather_df['Quandrant'] == 1, All_Weather_df['First'], np.nan)
        All_Weather_df['AW_Return'] = np.where(All_Weather_df['Quandrant'] == 2, All_Weather_df['Second'],
                                               All_Weather_df['AW_Return'])
        All_Weather_df['AW_Return'] = np.where(All_Weather_df['Quandrant'] == 3, All_Weather_df['Third'],
                                               All_Weather_df['AW_Return'])
        All_Weather_df['AW_Return'] = np.where(All_Weather_df['Quandrant'] == 4, All_Weather_df['Fourth'],
                                               All_Weather_df['AW_Return'])
        All_Weather_df['CR'] = All_Weather_df['AW_Return'].cumsum().apply(np.exp)
        #print All_Weather_df['CR'],All_Weather_df['Quandrant'][-1]
        if All_Weather_df['Quandrant'][-1] == 1:
            Return_Portfolio = First_df
        elif All_Weather_df['Quandrant'][-1] == 2:
            Return_Portfolio = Second_df
        elif All_Weather_df['Quandrant'][-1] == 3:
            Return_Portfolio = Third_df
        else:
            Return_Portfolio = Fourth_df
        #print All_Weather_df['CR'],All_Weather_df['Quandrant'][-1],Return_Portfolio
        #print Economic_Cycle['Quandrant']
        #print All_Weather_df['Quandrant'][-1]
        return All_Weather_df['CR'],Economic_Cycle['Quandrant'][-1],Return_Portfolio,Economic_Cycle['Quandrant']

    def Global_Volatility_Portfolio_Algorithm(self, Tickers, lookback):
        Process = ['Price', 'Daily_Return', 'Momentum', 'Volatility']

        Price_df = pd.DataFrame()

        for Ticker in Tickers:
            Price_df[Ticker] = Database_Functions.Fetch(Ticker=Ticker)[Ticker]
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
        #print Panel['Volatility'].ix[-1]*(250**0.5)*100.
        Panel['Volatility_Inverse'] = 1. / Panel['Volatility']

        Panel['Momentum_Rank'] = Panel['Volatility'].rank(axis=1, ascending=True)

        Panel['Selected_asset'] = np.where(Panel['Momentum_Rank'] <= 12, 1, 0)
        # Panel['Selected_asset'] = np.where(Panel['Momentum_Rank']<=len(Tickers)/3,1,0)
        Panel['Select_Volatility'] = Panel['Volatility_Inverse'] * Panel['Selected_asset']

        Volatility_weight = Panel['Select_Volatility'].replace(0, np.nan)
        Volatility_sum = Volatility_weight.sum(axis=1).replace(0, np.nan)

        Weight_df = pd.DataFrame()
        for Ticker in Volatility_weight.columns:
            Weight_df[Ticker] = Volatility_weight[Ticker] / Volatility_sum

        Panel['Volatility_weight'] = Weight_df

        Portfolio_Return_temp = Panel['Volatility_weight'] * Panel['Daily_Return']
        Portfolio_Return = Portfolio_Return_temp.sum(axis=1)
        Portfolio_CR = Portfolio_Return.cumsum().apply(np.exp)

        # Output:Last weight, Cumulative Return
        Turnover_temp = Panel['Volatility_weight'].diff().abs()
        Total_Turnover = Turnover_temp.abs().sum().sum()

        # print Panel['Volatility_weight'].ix[-1].fillna(0)

        return Panel['Volatility_weight'].ix[-1].fillna(0), Portfolio_CR,Panel['Volatility_weight'], Total_Turnover,Panel['Volatility'].ix[-1]*(250**0.5)

    def Save_Reload(self):
        Equity_Benchmark = Database_Functions.Fetch(Ticker='MXWO Index').pct_change().dropna()
        FI_Benchmark = Database_Functions.Fetch(Ticker='LBUSTRUU Index').pct_change().dropna()
        #print Equity_Benchmark

        Equity_tree_code = self.Equity_Tree_table.get_children()
        Fixed_income_tree_code = self.Fixed_Income_Tree_table.get_children()

        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
        cursor = connection.cursor()

        cursor.execute("DELETE FROM Equity_holdings")
        cursor.execute("DELETE FROM Fixed_income_holdings")

        Equity_list = []
        Fixed_income_list = []

        Equity_Name_list = []
        Fixed_income_Name_list = []

        Aggregate_Tickers = []
        Aggregate_Name = []
        Aggregate_Category = []

        for each in Equity_tree_code:
            Ticker = self.Equity_Tree_table.item(each)['values'][0]
            Name = self.Equity_Tree_table.item(each)['values'][1]
            Equity_list.append(Ticker)
            Equity_Name_list.append(Name)
            Aggregate_Tickers.append(Ticker)
            Aggregate_Name.append(Name)
            Aggregate_Category.append('Equity')
            cursor.execute("INSERT OR REPLACE INTO Equity_holdings VALUES (?,?,?)", ('Equity', str(Ticker), str(Name)))

        for each in Fixed_income_tree_code:
            Ticker = self.Fixed_Income_Tree_table.item(each)['values'][0]
            Name = self.Fixed_Income_Tree_table.item(each)['values'][1]
            Fixed_income_list.append(Ticker)
            Fixed_income_Name_list.append(Name)
            Aggregate_Tickers.append(Ticker)
            Aggregate_Name.append(Name)
            Aggregate_Category.append('Fixed Income')
            cursor.execute("INSERT OR REPLACE INTO Fixed_income_holdings VALUES (?,?,?)",
                           ('Fixed Income', str(Ticker), str(Name)))

        connection.commit()
        connection.close()

        # print Aggregate_Tickers
        # print Aggregate_Name


        Portfolio_Equity = self.Global_Volatility_Portfolio_Algorithm(Equity_list,lookback=60)
        global Equity_Portfolio_CR
        Equity_Portfolio_CR = Portfolio_Equity[1]

        Equity_compare = pd.DataFrame()
        Equity_compare['Equity'] = Equity_Portfolio_CR.pct_change()
        Equity_compare['Benchmark'] = Equity_Benchmark
        Equity_compare = Equity_compare.dropna()
        Equity_compare = Equity_compare.cumsum().apply(np.exp)


        global Equity_Portfolio_Weight
        Equity_Portfolio_Weight = Portfolio_Equity[2]
        #print Equity_Portfolio_Weight
        # Portfolio_Equity = self.Global_Volatility_Portfolio_Algorithm(Tickers=Equity_list,lookback=60)
        # Global_Volatility_Portfolio_Algorithm(Tickers,lookback)
        self.Equity_Chart.clear()
        self.Equity_Chart.plot(Equity_compare['Equity'].index, Equity_compare['Equity'], lw=1, color='#00A3DC', alpha=0.5,label='Equity(Low Volatility)')
        self.Equity_Chart.plot(Equity_compare['Benchmark'].index, Equity_compare['Benchmark'], lw=1, color='#01485E', alpha=0.5,label='MSCI World Index')
        self.Equity_Chart.legend(fontsize=8, loc=4, ncol=2, shadow=True).get_frame().set_facecolor('#F0F0F0')
        self.Equity_canvas.show()

        Horizon = (Portfolio_Equity[1].index[-1] - Portfolio_Equity[1].index[0]).days / 365.

        Equity_Return = round((Portfolio_Equity[1][-1] ** (1 / Horizon) - 1.) * 100., 2)
        Equity_Volatility = round(np.std(Portfolio_Equity[1].pct_change()) * (250 ** 0.5) * 100., 2)
        Equity_Sharpe = round(Equity_Return / Equity_Volatility, 2)

        self.Equity_Return_StrVar.set('Annual Return: ' + str(Equity_Return) + '%')
        self.Equity_Volatility_StrVar.set('Volatility: ' + str(Equity_Volatility) + '%')
        self.Equity_Sharpe_StrVar.set('Sharpe: ' + str(Equity_Sharpe))

        Portfolio_Fixed_income = self.Portfolio(Fixed_income_list)
        global FI_Portfolio_CR
        FI_Portfolio_CR = Portfolio_Fixed_income[1]

        FI_compare = pd.DataFrame()
        FI_compare['FI'] = FI_Portfolio_CR.pct_change()
        FI_compare['Benchmark'] = FI_Benchmark
        FI_compare = FI_compare.dropna()
        FI_compare = FI_compare.cumsum().apply(np.exp)



        global Fixed_income_Portfolio_Weight
        Fixed_income_Portfolio_Weight = Portfolio_Fixed_income[2]

        #Weights output
        global Aggressive_Weight_df
        global Moderate_Weight_df
        global Consevative_Weight_df
        Aggressive_Weight_df = pd.concat([Equity_Portfolio_Weight*0.8,Fixed_income_Portfolio_Weight*0.2],axis=1).dropna()
        Moderate_Weight_df = pd.concat([Equity_Portfolio_Weight*0.6,Fixed_income_Portfolio_Weight*0.4],axis=1).dropna()
        Consevative_Weight_df = pd.concat([Equity_Portfolio_Weight*0.2,Fixed_income_Portfolio_Weight*0.8],axis=1).dropna()



















        self.Fixed_Income_Chart.clear()
        self.Fixed_Income_Chart.plot(FI_compare['FI'].index, FI_compare['FI'], lw=1, color='#00A3DC',
                                     alpha=0.5,label='Fixed Income(Low Volatility)')
        self.Fixed_Income_Chart.plot(FI_compare['Benchmark'].index, FI_compare['Benchmark'], lw=1, color='#01485E', alpha=0.5,label='Barclays U.S. Aggregate Index')
        self.Fixed_Income_Chart.legend(fontsize=8, loc=4, ncol=2, shadow=True).get_frame().set_facecolor('#F0F0F0')

        self.Fixed_Income_canvas.show()

        Fixed_income_Horizon = (Portfolio_Fixed_income[1].index[-1] - Portfolio_Fixed_income[1].index[0]).days / 365.

        Fixed_income_Return = round((Portfolio_Fixed_income[1][-1] ** (1 / Fixed_income_Horizon) - 1.) * 100., 2)
        Fixed_income_Volatility = round(np.std(Portfolio_Fixed_income[1].pct_change()) * (250 ** 0.5) * 100., 2)
        Fixed_income_Sharpe = round(Fixed_income_Return / Fixed_income_Volatility, 2)

        self.Fixed_income_Return_StrVar.set('Annual Return: ' + str(Fixed_income_Return) + '%')
        self.Fixed_income_Volatility_StrVar.set('Volatility: ' + str(Fixed_income_Volatility) + '%')
        self.Fixed_income_Sharpe_StrVar.set('Sharpe: ' + str(Fixed_income_Sharpe))

        # ======================================================================

        Portfolio_Aggregate = self.Portfolio(Aggregate_Tickers)
        self.Aggregate_Chart.clear()
        self.Aggregate_Chart.plot(Portfolio_Aggregate[1].index, Portfolio_Aggregate[1], lw=1, color='#01485E',
                                  alpha=0.5)
        self.Aggregate_canvas.show()

        Aggregate_Horizon = (Portfolio_Aggregate[1].index[-1] - Portfolio_Aggregate[1].index[0]).days / 365.

        Aggregate_Return = round((Portfolio_Aggregate[1][-1] ** (1 / Aggregate_Horizon) - 1.) * 100., 2)
        Aggregate_Volatility = round(np.std(Portfolio_Aggregate[1].pct_change()) * (250 ** 0.5) * 100., 2)
        Aggregate_Sharpe = round(Aggregate_Return / Aggregate_Volatility, 2)

        self.Aggregate_Return_StrVar.set('Annual Return: ' + str(Aggregate_Return) + '%')
        self.Aggregate_Volatility_StrVar.set('Volatility: ' + str(Aggregate_Volatility) + '%')
        self.Aggregate_Sharpe_StrVar.set('Sharpe: ' + str(Aggregate_Sharpe))

        # Benchmark=============================================================

        # Barclays Agg
        Bond_Index = Database_Functions.Fetch(Ticker='LBUSTRUU Index').dropna()
        Bond_Index['Return'] = Bond_Index['LBUSTRUU Index'][1000:].pct_change()
        Bond_Index['CR'] = Bond_Index['Return'].cumsum().apply(np.exp)

        Bond_Index_horizon = (Bond_Index[1000:].index[-1] - Bond_Index[1000:].index[0]).days / 365.
        Bond_Benchmark_Return = (Bond_Index[1000:]['CR'][-1] ** (1. / Bond_Index_horizon) - 1.) * 100.
        Bond_Benchmark_Volatility = np.std(Bond_Index['Return'][1000:]) * (250. ** 0.5) * 100.

        # MSCI ACWI Index
        ACWI_Index = Database_Functions.Fetch(Ticker='MXWO Index').dropna()
        ACWI_Index['Return'] = ACWI_Index['MXWO Index'][1000:].pct_change()
        ACWI_Index['CR'] = ACWI_Index['Return'].cumsum().apply(np.exp)

        ACWI_Index_horizon = (ACWI_Index.index[1000:][-1] - ACWI_Index[1000:].index[0]).days / 365.
        ACWI_Benchmark_Return = (ACWI_Index[1000:]['CR'][-1] ** (1. / ACWI_Index_horizon) - 1.) * 100.
        ACWI_Benchmark_Volatility = np.std(ACWI_Index['Return'][1000:]) * (250. ** 0.5) * 100.

        # ======================================================================
        Equity_Portfolio_return = Portfolio_Equity[1].pct_change().dropna()
        Fixed_Income_Portfolio_return = Portfolio_Fixed_income[1].pct_change().dropna()
        Aggregate_Portfolio_return = Portfolio_Aggregate[1].pct_change().dropna()

        # Aggressive======================================================================
        Aggressive_Return = Equity_Portfolio_return * 0.8 + Fixed_Income_Portfolio_return * 0.2
        Aggressive_CR = Aggressive_Return.cumsum().apply(np.exp).dropna()

        Aggressive_Benchmark_return = ACWI_Index['Return'] * 0.8 + Bond_Index['Return'] * 0.2
        global Aggressive_compare
        Aggressive_compare = pd.concat([Aggressive_Return, Aggressive_Benchmark_return], axis=1).dropna()
        Aggressive_compare = Aggressive_compare.cumsum().apply(np.exp)
        Aggressive_compare.columns = ['Portfolio', 'Benchmark']

        # Aggressive_Up_time =  np.mean(np.where(Aggressive_compare['Portfolio'].resample('MS',how='first').pct_change()>0,1,0))
        # print Aggressive_Up_time



        self.Aggressive_Chart.clear()
        self.Aggressive_Chart.plot(Aggressive_compare.index, Aggressive_compare['Portfolio'], lw=1, color='#00A3DC',
                                   alpha=0.5, label='Portfolio')
        self.Aggressive_Chart.plot(Aggressive_compare.index, Aggressive_compare['Benchmark'], lw=1, color='#01485E',
                                   alpha=0.5, label='Benchmark(80,20)')
        self.Aggressive_Chart.legend(fontsize=8, loc=2, ncol=2, shadow=True).get_frame().set_facecolor('#F0F0F0')

        self.Aggressive_canvas.show()

        # Moderate========================================================================
        Moderate_Return = Equity_Portfolio_return * 0.6 + Fixed_Income_Portfolio_return * 0.4
        Moderate_CR = Moderate_Return.cumsum().apply(np.exp).dropna()

        Moderate_Benchmark_return = ACWI_Index['Return'] * 0.6 + Bond_Index['Return'] * 0.4
        global Moderate_compare
        Moderate_compare = pd.concat([Moderate_Return, Moderate_Benchmark_return], axis=1).dropna()
        Moderate_compare = Moderate_compare.cumsum().apply(np.exp)
        Moderate_compare.columns = ['Portfolio', 'Benchmark']

        # Moderate_Up_time =  np.mean(np.where(Moderate_compare['Portfolio'].resample('MS',how='first').pct_change()>0,1,0))
        # print Moderate_Up_time


        self.Moderate_Chart.clear()
        # self.Moderate_Chart.plot(Moderate_CR.index,Moderate_CR,lw=1,color='#01485E',alpha=0.5)
        self.Moderate_Chart.plot(Moderate_compare.index, Moderate_compare['Portfolio'], lw=1, color='#00A3DC',
                                 alpha=0.5, label='Portfolio')
        self.Moderate_Chart.plot(Moderate_compare.index, Moderate_compare['Benchmark'], lw=1, color='#01485E',
                                 alpha=0.5, label='Benchmark(60,40)')
        self.Moderate_Chart.legend(fontsize=8, loc=2, ncol=2, shadow=True).get_frame().set_facecolor('#F0F0F0')
        self.Moderate_canvas.show()

        # Conservative====================================================================
        Conservative_Return = Equity_Portfolio_return * 0.2 + Fixed_Income_Portfolio_return * 0.8
        Conservative_CR = Conservative_Return.cumsum().apply(np.exp).dropna()

        Conservative_Benchmark_return = ACWI_Index['Return'] * 0.2 + Bond_Index['Return'] * 0.8
        global Conservative_compare
        Conservative_compare = pd.concat([Conservative_Return, Conservative_Benchmark_return], axis=1).dropna()
        Conservative_compare = Conservative_compare.cumsum().apply(np.exp)
        Conservative_compare.columns = ['Portfolio', 'Benchmark']

        # Conservative_Up_time =  np.mean(np.where(Conservative_compare['Portfolio'].resample('MS',how='first').pct_change()>0,1,0))
        # print Conservative_Up_time


        self.Conservative_Chart.clear()
        # self.Conservative_Chart.plot(Conservative_CR.index,Conservative_CR,lw=1,color='#01485E',alpha=0.5)
        self.Conservative_Chart.plot(Conservative_compare.index, Conservative_compare['Portfolio'], lw=1,
                                     color='#00A3DC', alpha=0.5, label='Portfolio')
        self.Conservative_Chart.plot(Conservative_compare.index, Conservative_compare['Benchmark'], lw=1,
                                     color='#01485E', alpha=0.5, label='Benchmark(20,80)')
        self.Conservative_Chart.legend(fontsize=8, loc=2, ncol=2, shadow=True).get_frame().set_facecolor('#F0F0F0')
        self.Conservative_canvas.show()

        # Performance
        global Aggressive_AR
        global Aggressive_Std
        global Aggressive_Sharpe
        Horizon = (Aggressive_CR.index[-1] - Aggressive_CR.index[0]).days / 365.
        Aggressive_AR = round((Aggressive_CR[-1] ** (1. / Horizon) - 1.) * 100., 2)
        Aggressive_Std = round((np.std(Aggressive_Return) * (250. ** 0.5)) * 100., 2)
        Aggressive_Sharpe = round(Aggressive_AR / Aggressive_Std, 2)

        self.Aggressive_Return_StrVar.set('Annual Return: ' + str(Aggressive_AR) + '%')
        self.Aggressive_Volatility_StrVar.set('Volatility: ' + str(Aggressive_Std) + '%')
        self.Aggressive_Sharpe_StrVar.set('Sharpe: ' + str(Aggressive_Sharpe))

        self.Portfolio1_frame.update_idletasks()

        # ----------------------------------------------------------------------
        global Moderate_AR
        global Moderate_Std
        global Moderate_Sharpe
        Moderate_AR = round((Moderate_CR[-1] ** (1. / Horizon) - 1.) * 100., 2)
        Moderate_Std = round((np.std(Moderate_Return) * (250. ** 0.5)) * 100., 2)
        Moderate_Sharpe = round(Moderate_AR / Moderate_Std, 2)

        self.Moderate_Return_StrVar.set('Annual Return: ' + str(Moderate_AR) + '%')
        self.Moderate_Volatility_StrVar.set('Volatility: ' + str(Moderate_Std) + '%')
        self.Moderate_Sharpe_StrVar.set('Sharpe: ' + str(Moderate_Sharpe))

        self.Portfolio2_frame.update_idletasks()

        # ----------------------------------------------------------------------
        global Conservative_AR
        global Conservative_Std
        global Conservative_Sharpe
        Conservative_AR = round((Conservative_CR[-1] ** (1. / Horizon) - 1.) * 100., 2)
        Conservative_Std = round((np.std(Conservative_Return) * (250. ** 0.5)) * 100., 2)
        Conservative_Sharpe = round(Conservative_AR / Conservative_Std, 2)

        self.Conservative_Return_StrVar.set('Annual Return: ' + str(Conservative_AR) + '%')
        self.Conservative_Volatility_StrVar.set('Volatility: ' + str(Conservative_Std) + '%')
        self.Conservative_Sharpe_StrVar.set('Sharpe: ' + str(Conservative_Sharpe))

        self.Portfolio3_frame.update_idletasks()

        '''
        =======================================================================
        Bond Portfolio=========================================================
        =======================================================================
        '''

        Risky_tickers = ['JPEICORE Index', 'IBOXHY Index']
        Safety_tickers = ['G0Q0 Index']

        Risky_Portfolio = self.Portfolio(Tickers=Risky_tickers)
        Safety_Porfolio = self.Portfolio(Tickers=Safety_tickers)

        global Strategic_Bond_Weight
        Strategic_Bond_Weight = pd.concat([Risky_Portfolio[2],Safety_Porfolio[2]],axis=1)
        # print Risky_Portfolio,Safety_Porfolio

        # Signal_Generate_df = pd.DataFrame(Risky_Portfolio[1],columns = ['Risky_Portfolio'])
        Signal_Generate_df = pd.concat([Risky_Portfolio[1], Safety_Porfolio[1]], axis=1)
        Signal_Generate_df.columns = ['Risky_Portfolio', 'Safety_Portfolio']

        Signal_Generate_df['Risy_Return'] = Signal_Generate_df['Risky_Portfolio'].pct_change()
        Signal_Generate_df['Safety_Return'] = Signal_Generate_df['Safety_Portfolio'].pct_change()

        Signal_Generate_df['EMA'] = pd.ewma(Signal_Generate_df['Risky_Portfolio'], 60)
        Signal_Generate_df['Signal'] = np.where(Signal_Generate_df['Risky_Portfolio'] > Signal_Generate_df['EMA'], 1, 0)
        Signal_Generate_df['Bond_Portfolio_Return'] = np.where(Signal_Generate_df['Signal'].shift(1) == 1,
                                                               Signal_Generate_df['Risy_Return'],
                                                               Signal_Generate_df['Safety_Return'])
        Signal_Generate_df['Benchmark_Return'] = Bond_Index['Return']
        # print Signal_Generate_df
        Strategic_Bond_Weight['Signal'] = Signal_Generate_df['Signal']
        Strategic_Bond_Weight = Strategic_Bond_Weight.dropna()
        Strategic_Bond_Weight['JPEICORE Index'] = np.where(Strategic_Bond_Weight['Signal'] == 1.,Strategic_Bond_Weight['JPEICORE Index'],0)
        Strategic_Bond_Weight['IBOXHY Index'] = np.where(Strategic_Bond_Weight['Signal'] == 1.,Strategic_Bond_Weight['IBOXHY Index'],0)
        Strategic_Bond_Weight['G0Q0 Index'] = np.where(Strategic_Bond_Weight['Signal'] == 0.,Strategic_Bond_Weight['G0Q0 Index'],0)



        global Bond_Portfolio_compare
        global Bond_AR
        global Bond_Portfolio_Std
        global Bond_Sharpe

        Bond_Portfolio_compare = pd.DataFrame()
        Bond_Portfolio_compare['Portfolio'] = Signal_Generate_df['Bond_Portfolio_Return']
        Bond_Portfolio_Std = np.std(Bond_Portfolio_compare['Portfolio']) * (250 ** 0.5) * 100.
        Bond_Portfolio_compare['Benchmark'] = Signal_Generate_df['Benchmark_Return']
        Bond_Portfolio_compare = Bond_Portfolio_compare.dropna()

        Bond_Portfolio_compare = Bond_Portfolio_compare.cumsum().apply(np.exp)

        Bond_horizon = (Bond_Portfolio_compare.index[-1] - Bond_Portfolio_compare.index[0]).days / 365.
        Bond_AR = (Bond_Portfolio_compare['Portfolio'][-1] ** (1. / Bond_horizon) - 1) * 100.

        self.Bond_Chart.clear()
        self.Bond_Chart.plot(Bond_Portfolio_compare['Portfolio'].index, Bond_Portfolio_compare['Portfolio'], lw=1,
                             color='#00A3DC', alpha=0.5, label='Global Strategc Bond')
        self.Bond_Chart.plot(Bond_Portfolio_compare['Benchmark'].index, Bond_Portfolio_compare['Benchmark'], lw=1,
                             color='#01485E', alpha=0.5, label='Baclays U.S. Aggregate Index')
        self.Bond_Chart.legend(fontsize=8, loc=2, ncol=2, shadow=True).get_frame().set_facecolor('#F0F0F0')

        self.Bond_canvas.show()

        # Bond_Up_time =  np.mean(np.where(Bond_Portfolio_compare['Portfolio'].resample('MS',how='first').pct_change()>0,1,0))
        # print Bond_Up_time

        Bond_Sharpe = Bond_AR / Bond_Portfolio_Std

        self.Bond_Return_StrVar.set('Annual Return: ' + str(round(Bond_AR, 2)) + '%')
        self.Bond_Volatility_StrVar.set('Volatility: ' + str(round(Bond_Portfolio_Std, 2)) + '%')
        self.Bond_Sharpe_StrVar.set('Sharpe: ' + str(round(Bond_AR / Bond_Portfolio_Std, 2)))

        Last_bond_signal = Signal_Generate_df['Signal'][-1]

        if Last_bond_signal == 1:

            Bond_Portfolio_weight = pd.DataFrame(columns=['Weight'])
            Bond_Portfolio_weight['Weight'] = Risky_Portfolio[0] * 100.
            Bond_Portfolio_weight['Name'] = ['JPMorgan EMBI Global Core Index', 'iBoxx USD Liquid High Yield Index']
            # print Bond_Portfolio_weight

            self.Bond_Tree_table.delete(*self.Bond_Tree_table.get_children())

            for i in range(len(Bond_Portfolio_weight)):
                self.Bond_Tree_table.insert("", i, '', values=(
                'Fixed Income', Bond_Portfolio_weight.index[i], Bond_Portfolio_weight['Name'][i],
                round(Bond_Portfolio_weight['Weight'][i], 2)))

        else:

            Bond_Portfolio_weight = pd.DataFrame(columns=['Weight'])
            Bond_Portfolio_weight['Weight'] = Safety_Porfolio[0] * 100.
            Bond_Portfolio_weight['Name'] = ['BofA Merrill Lynch U.S. Treasury Index']
            # print Bond_Portfolio_weight

            self.Bond_Tree_table.delete(*self.Bond_Tree_table.get_children())

            for i in range(len(Bond_Portfolio_weight)):
                self.Bond_Tree_table.insert("", i, '', values=(
                'Fixed Income', Bond_Portfolio_weight.index[i], Bond_Portfolio_weight['Name'][i],
                round(Bond_Portfolio_weight['Weight'][i], 2)))

        # Conservative
        self.Bond_pie_Chart.clear()
        Bonc_Code = self.Bond_Tree_table.get_children()

        Bond_Tickers = []
        Bond_Weights = []

        for i in range(len(Bonc_Code)):
            Bond_Ticker = self.Bond_Tree_table.item(Bonc_Code[i])['values'][1]
            Bond_Weight = float(self.Bond_Tree_table.item(Bonc_Code[i])['values'][3])
            Bond_Tickers.append(Bond_Ticker)
            Bond_Weights.append(Bond_Weight)

        Bond_labels = tuple(Bond_Tickers)
        Bond_sizes = Bond_Weights

        self.Bond_pie_Chart.pie(Bond_sizes, colors=cm.Blues(np.arange(len(Bond_sizes)) / float(len(Bond_sizes))),
                                labels=Bond_labels, autopct='%1.1f%%', shadow=True, startangle=150,
                                labeldistance=0.8)  # ,autopct='%1.1f%%'
        self.Bond_pie_canvas.show()

        '''
        =======================================================================
        Global Momentum Portfolio==============================================
        =======================================================================
        '''

        GMP_info = self.Global_Momentum_Portfolio_Algorithm(Tickers=Aggregate_Tickers, lookback=200)
        global GMP_Weight_df
        GMP_Weight_df = GMP_info[2]

        GMP_table = pd.DataFrame(columns=['Category', 'Ticker', 'Name', 'Weight'])
        GMP_table['Category'] = Aggregate_Category
        GMP_table['Ticker'] = Aggregate_Tickers
        GMP_table['Name'] = Aggregate_Name
        GMP_table['Weight'] = np.array(GMP_info[0]) * 100.
        GMP_table = GMP_table.dropna().reset_index()

        global GMP_compare
        global GMP_AR
        global GMP_Std
        global GMP_Sharpe
        GMP_compare = pd.DataFrame()
        GMP_compare['Portfolio'] = GMP_info[1].pct_change()
        GMP_Return = GMP_compare['Portfolio']
        GMP_compare['Benchmark'] = ACWI_Index['Return']
        GMP_Std = np.std(GMP_compare['Portfolio']) * (250 ** 0.5) * 100.
        GMP_compare = GMP_compare.dropna()
        GMP_compare = GMP_compare.cumsum().apply(np.exp)

        # self.GMP_Chart
        self.GMP_Chart.clear()
        self.GMP_Chart.plot(GMP_compare['Portfolio'].index, GMP_compare['Portfolio'], lw=1, color='#00A3DC', alpha=0.5,
                            label='Global Momentum Portfolio')
        self.GMP_Chart.plot(GMP_compare['Benchmark'].index, GMP_compare['Benchmark'], lw=1, color='#01485E', alpha=0.5,
                            label='MSCI World Index')
        self.GMP_Chart.legend(fontsize=8, loc=2, ncol=2, shadow=True).get_frame().set_facecolor('#F0F0F0')

        self.GMP_canvas.show()

        # GMP_Up_time =  np.mean(np.where(GMP_compare['Portfolio'].resample('MS',how='first').pct_change()>0,1,0))
        # print GMP_Up_time



        GMP_horizon = (GMP_compare.index[-1] - GMP_compare.index[0]).days / 365.
        GMP_AR = (GMP_compare['Portfolio'][-1] ** (1. / GMP_horizon) - 1) * 100.
        GMP_Sharpe = GMP_AR / GMP_Std
        self.GMP_Return_StrVar.set('Annual Return: ' + str(round(GMP_AR, 2)) + '%')
        self.GMP_Volatility_StrVar.set('Volatility: ' + str(round(GMP_Std, 2)) + '%')
        self.GMP_Sharpe_StrVar.set('Sharpe: ' + str(round(GMP_AR / GMP_Std, 2)))

        self.GMP_Tree_table.delete(*self.GMP_Tree_table.get_children())

        for i in range(len(GMP_table)):
            self.GMP_Tree_table.insert("", i, '', values=(
            GMP_table['Category'][i], GMP_table['Ticker'][i], GMP_table['Name'][i], round(GMP_table['Weight'][i], 2)))

        self.GMP_pie_Chart.clear()
        self.GMP_pie2_Chart.clear()
        GMP_Code = self.GMP_Tree_table.get_children()

        GMP_Tickers = []
        GMP_Weights = []

        for i in range(len(GMP_table)):
            GMP_Ticker = self.GMP_Tree_table.item(GMP_Code[i])['values'][1]
            GMP_Weight = float(self.GMP_Tree_table.item(GMP_Code[i])['values'][3])
            GMP_Tickers.append(GMP_Ticker)
            GMP_Weights.append(GMP_Weight)

        GMP_labels = tuple(GMP_Tickers)

        GMP_sizes = GMP_Weights


        self.GMP_pie2_Chart.pie(GMP_sizes, colors=cm.Blues(np.arange(len(GMP_sizes)) / float(len(GMP_sizes))),
                                labels=GMP_labels, autopct='%1.1f%%', shadow=True, startangle=150,
                                labeldistance=0.8)  # ,autopct='%1.1f%%'
        self.GMP_pie2_canvas.show()

        Fixed_income_weight_GMP = GMP_table['Weight'][GMP_table.Category == 'Fixed Income'].sum()
        Equity_weight_GMP = GMP_table['Weight'][GMP_table.Category == 'Equity'].sum()

        GMP_category_size = [round(Equity_weight_GMP, 2), round(Fixed_income_weight_GMP, 2)]

        self.GMP_pie_Chart.pie(GMP_category_size,
                               colors=cm.Blues(np.arange(len(GMP_category_size)) / float(len(GMP_category_size))),
                               labels=self.labels, autopct='%1.1f%%', shadow=True, startangle=150,
                               labeldistance=0.8)  # ,autopct='%1.1f%%'
        self.GMP_pie_canvas.show()



        '''
        ================================================================================================================
        All Weather Portfolio===========================================================================================
        ================================================================================================================
        '''
        All_Weather_info = self.All_Weather_Portfolio(Lookback=2)
        global AWP_CR
        AWP_CR = All_Weather_info[0]
        AMP_Return = AWP_CR.pct_change().dropna()

        AMP_compare = pd.DataFrame()
        AMP_compare['AWP'] = AMP_Return
        AMP_compare['Benchmark'] = ACWI_Index['Return']
        AMP_compare =  AMP_compare.dropna()
        AMP_compare = AMP_compare.cumsum().apply(np.exp)

        #print AMP_Return
        #print AWP_CR
        AWP_Quandrant = All_Weather_info[1]
        AWP_table = All_Weather_info[2]

        for i in range(len(AWP_table)):
            self.AWP_Tree_table.insert("", i, '', values=(
                AWP_table['Category'][i], AWP_table['Ticker'][i], AWP_table['Name'][i], round(AWP_table['Weight'][i], 2)))

        AWP_labels = tuple(AWP_table['Ticker'].tolist())
        AWP_sizes = AWP_table['Weight'].tolist()

        self.AWP_pie_Chart.pie(AWP_sizes, colors=cm.Blues(np.arange(len(AWP_sizes)) / float(len(AWP_sizes))),
                                         labels=AWP_labels, autopct='%1.1f%%', shadow=True, startangle=150,
                                         labeldistance=0.8)  # ,autopct='%1.1f%%'
        self.AWP_pie_canvas.show()

        self.AWP_Chart.plot(AMP_compare.index,AMP_compare['AWP'], lw=1, alpha=0.5,color='#00A3DC', label='All Weather Portfolio')
        self.AWP_Chart.plot(AMP_compare.index,AMP_compare['Benchmark'], lw=1, alpha=0.5,color='#01485E', label='MSCI World Index')
        self.AWP_Chart.legend(fontsize=8, loc=2, ncol=2, shadow=True).get_frame().set_facecolor('#F0F0F0')

        self.AWP_canvas.show()
        if AWP_Quandrant == 1:
            self.Quandrant_StrVar.set('Curremt: I')
        elif AWP_Quandrant == 2:
            self.Quandrant_StrVar.set('Curremt: II')
        elif AWP_Quandrant == 3:
            self.Quandrant_StrVar.set('Curremt: III')
        elif AWP_Quandrant == 4:
            self.Quandrant_StrVar.set('Curremt: IV')
        else:
            self.Quandrant_StrVar.set('Error!')

        AWP_Performace = self.Performance_Calculate(AWP_CR)
        self.Quandrant_Return_StrVar.set('Annual Return: '+AWP_Performace[0])
        self.Quandrant_Volatility_StrVar.set('Volatility: '+AWP_Performace[1])
        self.Quandrant_Sharpe_StrVar.set('Sharpe: '+AWP_Performace[2])

        AMP_AR = float(AWP_Performace[0].split('%')[0])
        AMP_Vol = float(AWP_Performace[1].split('%')[0])

        global AWP_Score
        AWP_Score = All_Weather_info[3]



        '''
        Summary of Portfolio===================================================
        =======================================================================
        '''
        # Benchmark remark------------------------------------------------------
        # MSCI World Index
        World_Index = Database_Functions.Fetch(Ticker='MXWO Index').dropna()
        World_Index['World_Index'] = World_Index['MXWO Index'][2000:].pct_change()
        World_Index['Return'] = World_Index['MXWO Index'][5000:].pct_change()
        World_Index['CR'] = World_Index['Return'].cumsum().apply(np.exp)

        World_Index_horizon = (World_Index.index[5000:][-1] - World_Index[5000:].index[0]).days / 365.
        Equity_Benchmark_Return = (World_Index[5000:]['CR'][-1] ** (1. / World_Index_horizon) - 1.) * 100.
        Equity_Benchmark_Volatility = np.std(World_Index['Return'][5000:]) * (250. ** 0.5) * 100.

        # MSCI Emerging Market Index
        Emerging_Index = Database_Functions.Fetch(Ticker='MXEF Index').dropna()
        Emerging_Index['Return'] = Emerging_Index['MXEF Index'][5000:].pct_change()
        Emerging_Index['CR'] = Emerging_Index['Return'].cumsum().apply(np.exp)

        Emerging_Index_horizon = (Emerging_Index.index[5000:][-1] - Emerging_Index[5000:].index[0]).days / 365.
        Emerging_Benchmark_Return = (Emerging_Index[5000:]['CR'][-1] ** (1. / Emerging_Index_horizon) - 1.) * 100.
        Emerging_Benchmark_Volatility = np.std(Emerging_Index['Return'][5000:]) * (250. ** 0.5) * 100.

        # Barclays Agg
        Bond_Index = Database_Functions.Fetch(Ticker='LBUSTRUU Index').dropna()
        Bond_Index['Return1'] = Bond_Index['LBUSTRUU Index'][2000:].pct_change()
        Bond_Index['Return'] = Bond_Index['LBUSTRUU Index'][5000:].pct_change()
        Bond_Index['CR'] = Bond_Index['Return'].cumsum().apply(np.exp)

        Bond_Index_horizon = (Bond_Index[5000:].index[-1] - Bond_Index[5000:].index[0]).days / 365.
        Bond_Benchmark_Return = (Bond_Index[5000:]['CR'][-1] ** (1. / Bond_Index_horizon) - 1.) * 100.
        Bond_Benchmark_Volatility = np.std(Bond_Index['Return'][5000:]) * (250. ** 0.5) * 100.

        # MSCI ACWI Index
        ACWI_Index = Database_Functions.Fetch(Ticker='MXWD Index').dropna()
        ACWI_Index['Return'] = ACWI_Index['MXWD Index'][5000:].pct_change()
        ACWI_Index['CR'] = ACWI_Index['Return'].cumsum().apply(np.exp)

        ACWI_Index_horizon = (ACWI_Index.index[5000:][-1] - ACWI_Index[5000:].index[0]).days / 365.
        ACWI_Benchmark_Return = (ACWI_Index[5000:]['CR'][-1] ** (1. / ACWI_Index_horizon) - 1.) * 100.
        ACWI_Benchmark_Volatility = np.std(ACWI_Index['Return'][5000:]) * (250. ** 0.5) * 100.

        self.Overview_Chart.clear()
        self.Overview_Chart.set_xlabel('Standard Deviation(%)', fontsize=8)
        self.Overview_Chart.set_ylabel('Annual Return(%)', fontsize=8)

        self.Overview_Chart.scatter(x = round(float(Equity_Volatility),2),y = round(float(Equity_Return),2), edgecolors='black',s=200,color='#00A3DC')
        #self.Overview_Chart.scatter(x=round(float(Aggressive_Std), 2), y=round(float(Aggressive_AR), 2),
                                    #edgecolors='black', s=200, color='#00A3DC')
        #self.Overview_Chart.scatter(x=round(float(Moderate_Std), 2), y=round(float(Moderate_AR), 2), edgecolors='black',
                                    #s=200, color='#00A3DC')
        #self.Overview_Chart.scatter(x=round(float(Conservative_Std), 2), y=round(float(Conservative_AR), 2),
                                    #edgecolors='black', s=200, color='#00A3DC')
        self.Overview_Chart.scatter(x = round(float(Fixed_income_Volatility),2),y = round(float(Fixed_income_Return),2), edgecolors='black',s=200,color='#00A3DC')
        self.Overview_Chart.scatter(x=round(float(Equity_Benchmark_Volatility), 2),
                                    y=round(float(Equity_Benchmark_Return), 2), edgecolors='black', s=200,
                                    color='yellow')
        self.Overview_Chart.scatter(x=round(float(Bond_Benchmark_Volatility), 2),
                                    y=round(float(Bond_Benchmark_Return), 2), edgecolors='black', s=200, color='yellow')
        self.Overview_Chart.scatter(x=round(float(Emerging_Benchmark_Volatility), 2),
                                    y=round(float(Emerging_Benchmark_Return), 2), edgecolors='black', s=200,
                                    color='yellow')
        self.Overview_Chart.scatter(x=round(float(ACWI_Benchmark_Volatility), 2),
                                    y=round(float(ACWI_Benchmark_Return), 2), edgecolors='black', s=200, color='yellow')
        # self.Overview_Chart.scatter(x = round(float(Aggregate_Volatility),2),y = round(float(Aggregate_Return),2), edgecolors='black',s=200,color='#00A3DC')
        self.Overview_Chart.scatter(x=round(float(Bond_Portfolio_Std), 2), y=round(float(Bond_AR), 2),
                                    edgecolors='black', s=200, color='#00A3DC')
        self.Overview_Chart.scatter(x=round(float(GMP_Std), 2), y=round(float(GMP_AR), 2), edgecolors='black', s=200,
                                    color='#00A3DC')
        self.Overview_Chart.scatter(x=round(float(AMP_Vol), 2), y=round(float(AMP_AR), 2), edgecolors='black', s=200,
                                    color='#00A3DC')

        self.Overview_Chart.annotate('Low Vol Equity',(Equity_Volatility+0.5,Equity_Return+0.1),fontsize=8)
        #self.Overview_Chart.annotate('Aggressive', (Aggressive_Std + 0.5, Aggressive_AR + 0.1), fontsize=8)
        #self.Overview_Chart.annotate('Moderate', (Moderate_Std + 0.5, Moderate_AR + 0.1), fontsize=8)
        #self.Overview_Chart.annotate('Conservative', (Conservative_Std + 0.5, Conservative_AR + 0.1), fontsize=8)
        self.Overview_Chart.annotate('Low Vol Fixed Income',(Fixed_income_Volatility+0.5,Fixed_income_Return+0.1),fontsize=8)
        self.Overview_Chart.annotate('MSCI World Index',
                                     (Equity_Benchmark_Volatility + 0.5, Equity_Benchmark_Return + 0.1), fontsize=8)
        self.Overview_Chart.annotate('Barclays U.S. Aggregate Bond Index',
                                     (Bond_Benchmark_Volatility + 0.5, Bond_Benchmark_Return + 0.1), fontsize=8)
        self.Overview_Chart.annotate('MSCI Emerging Market Index',
                                     (Emerging_Benchmark_Volatility + 0.5, Emerging_Benchmark_Return + 0.1), fontsize=8)
        self.Overview_Chart.annotate('MSCI ACWI Index', (ACWI_Benchmark_Volatility + 0.5, ACWI_Benchmark_Return + 0.1),
                                     fontsize=8)
        # self.Overview_Chart.annotate('Aggregate',(Aggregate_Volatility+0.5,Aggregate_Return+0.1),fontsize=8)
        self.Overview_Chart.annotate('Global Strategic Bond', (Bond_Portfolio_Std + 0.5, Bond_AR + 0.1), fontsize=8)
        self.Overview_Chart.annotate('Global Momentum Portfolio', (GMP_Std + 0.5, GMP_AR + 0.1), fontsize=8)
        self.Overview_Chart.annotate('All Weather Portfolio', (AMP_Vol + 0.5, AMP_AR + 0.1), fontsize=8)
        Compare_df = pd.concat(
            [Equity_Portfolio_return, Fixed_Income_Portfolio_return, Aggressive_Return, Moderate_Return,
             Conservative_Return, Aggregate_Portfolio_return, Signal_Generate_df['Bond_Portfolio_Return'], GMP_Return,AMP_Return,World_Index['World_Index'],Bond_Index['Return1']],
            axis=1).dropna()
        print pd.concat(
            [Equity_Portfolio_return, Fixed_Income_Portfolio_return, Aggressive_Return, Moderate_Return,
             Conservative_Return, Aggregate_Portfolio_return, Signal_Generate_df['Bond_Portfolio_Return'], GMP_Return,AMP_Return,World_Index['World_Index'],Bond_Index['Return1']],
            axis=1)
        Portfolios_df = Compare_df.cumsum().apply(np.exp).dropna()

        #Portfolios_df1 = Compare_df.cumsum().apply(np.exp).dropna()
        self.Overview1_Chart.clear()
        self.Overview1_Chart.plot(Portfolios_df.index,Portfolios_df[0],lw=1,alpha=0.5,label='Low Vol Equity')
        #self.Overview1_Chart.plot(Portfolios_df.index, Portfolios_df[2], lw=1, alpha=0.5, label='Aggressive')
        #self.Overview1_Chart.plot(Portfolios_df.index, Portfolios_df[3], lw=1, alpha=0.5, label='Moderate')
        #self.Overview1_Chart.plot(Portfolios_df.index, Portfolios_df[4], lw=1, alpha=0.5, label='Conservative')
        self.Overview1_Chart.plot(Portfolios_df.index,Portfolios_df[1],lw=1,alpha=0.5,label='Low Vol Fixed Income')
        # self.Overview1_Chart.plot(Portfolios_df.index,Portfolios_df[5],lw=1,alpha=0.5,label='Aggregate')
        self.Overview1_Chart.plot(Portfolios_df.index, Portfolios_df['Bond_Portfolio_Return'], lw=1, alpha=0.5,
                                  label='Global Strategic Bond')
        self.Overview1_Chart.plot(Portfolios_df.index, Portfolios_df['Portfolio'], lw=1, alpha=0.5,
                                  label='Global Momentum Portfolio')
        self.Overview1_Chart.plot(Portfolios_df.index, Portfolios_df['CR'], lw=1, alpha=0.5,
                                  label='All Weather Portfolio')
        self.Overview1_Chart.plot(Portfolios_df.index, Portfolios_df['World_Index'], lw=1, alpha=0.5,
                                  label='MSCI World Index')
        self.Overview1_Chart.plot(Portfolios_df.index, Portfolios_df['Return1'], lw=1, alpha=0.5,
                                  label='Barclays U.S. Aggregate Index')

        global Portfolios_df_output
        Portfolios_df_output = Portfolios_df[[0,1,'Bond_Portfolio_Return','Portfolio','CR','World_Index','Return1']]
        Portfolios_df_output.columns = ['Low Vol Equity','Low Vol Fixed Income','Global Strategic Bond','Global Momentum','All Weather Portfolio','MSCI World Index','Barclays U.S. Aggregate Index']
        #print Portfolios_df_output
        '''
        Horizon = (Portfolios_df1.index[-1] - Portfolios_df1.index[0]).days/365.
        print (Portfolios_df1.ix[-1]**(1./Horizon)-1)/(Portfolios_df1.pct_change().std()*(250**0.5))
        print (Portfolios_df1.ix[-1]**(1./Horizon)-1)
        print Portfolios_df1.pct_change().std()*(250**0.5)
        '''
        # print Database_Functions.Fetch('MXWO Index')['MXWO Index'].dropna()
        # print Database_Functions.Fetch('MXEF Index')['MXEF Index'].dropna()
        # print Database_Functions.Fetch('MXAPJ Index')['MXAPJ Index'].dropna()





        Benchmark = pd.DataFrame()
        Benchmark['MXWO Index'] = Database_Functions.Fetch('MXWO Index')['MXWO Index'].dropna().pct_change()
        Benchmark['MXEF Index'] = Database_Functions.Fetch('MXEF Index')['MXEF Index'].dropna().pct_change()
        # Benchmark['MXAPJ Index'] = Database_Functions.Fetch('MXAPJ Index')['MXAPJ Index'].dropna()
        # Benchmark = pd.concat([MXWO,MXEF,MXAPJ],axis=1)

        # print Benchmark
        # ,GMP_Return,MXWO['MXWO Index'],MXEF['MXEF Index'],MXAPJ['MXAPJ Index']
        # Benchmark_df = pd.concat([Aggressive_Return,Moderate_Return,Conservative_Return,Signal_Generate_df['Bond_Portfolio_Return'],GMP_Return,Benchmark],axis=1).dropna()
        # Benchmark_df = Benchmark_df[Benchmark_df.index > '2009-01-06' ]
        # Benchmark_df = Benchmark_df[Benchmark_df.index < '2009-03-09' ]

        # Benchmark_CR_df = Benchmark_df.cumsum().dropna()
        # Benchmark_CR_df.columns = ['Aggressive','Moderate','Conservative','Strategic Bond','Global Momentum','MXWO Index','MXEF Index']
        # print Benchmark_df
        # Benchmark_CR_df[(Benchmark_CR_df.index > '2007-10-09') & (Benchmark_CR_df < '2008-11-20')].plot()
        # Interval1 = Benchmark_CR_df[Benchmark_CR_df.index > '2007-10-09' ]
        # Interval2 = Interval1[Interval1.index < '2008-11-20' ]
        # Benchmark_CR_df.plot()
        # print Interval2
        # Interval2.plot()


        self.Overview1_Chart.legend(fontsize=8, loc=2, ncol=2, shadow=True).get_frame().set_facecolor('#F0F0F0')
        # bbox_to_anchor=(0., 1.02, 1., .102),



        '''
        #Standardized : Next Mission
        self.Overview1_Chart.plot(Portfolio_Equity[1].index,Portfolio_Equity[1],lw=1,color='#01485E',alpha=0.5)
        self.Overview1_Chart.plot(Portfolio_Fixed_income[1].index,Portfolio_Fixed_income[1],lw=1,color='#01485E',alpha=0.5)
        self.Overview1_Chart.plot(Aggressive_CR.index,Aggressive_CR,lw=1,color='#01485E',alpha=0.5)
        self.Overview1_Chart.plot(Moderate_CR.index,Moderate_CR,lw=1,color='#01485E',alpha=0.5)
        self.Overview1_Chart.plot(Conservative_CR.index,Conservative_CR,lw=1,color='#01485E',alpha=0.5)
        '''
        # [ 0.   ,  0.125,  0.25 ,  0.375,  0.5  ,  0.625,  0.75 ,  0.875]

        # cm.Blues(np.arange(len(self.sizes))/float(len(self.sizes)))

        self.Overview_canvas.show()
        self.Overview1_canvas.show()

        '''
        #self.Aggressive_pie2_Chart.clear()
        #labels = 'S&P500', 'Emerging Market', 'Japan', 'EAFE', 'US Treasury Bond'

        self.labels = tuple(Aggressive_Tickers)
        
        self.sizes = Aggressive_Weights
        self.cs=cm.Set1(np.arange(len(Aggressive_Tickers))/float(len(Aggressive_Tickers)))
        #self.colors = ['yellowgreen', 'gold', 'lightskyblue', 'lightcoral']
        self.explode = (0, 0)  # only "explode" the 2nd slice (i.e. 'Hogs')
        
        self.Aggressive_pie2_Chart.pie(self.sizes,colors=cm.Blues(np.arange(len(self.sizes))/float(len(self.sizes))), labels=self.labels, autopct='%1.1f%%', shadow=True, startangle=150, labeldistance=0.8)#,autopct='%1.1f%%'   
        self.Aggressive_pie2_canvas.show()
        #colors=self.colors,
        '''

        # Global_Volatility_Portfolio_Algorithm(Tickers,lookback)














        Equity_df = Portfolio_Equity[0] * 100.
        Fixed_income_df = Portfolio_Fixed_income[0] * 100.
        Aggregate_df = pd.DataFrame(Portfolio_Aggregate[0] * 100.)

        '''
        Aggregate_df['Ticker'] = Aggregate_df.index
        Aggregate_df.index = range(len(Aggregate_df))
        Aggregate_df['Name'] = Aggregate_Name
        Aggregate_df['Category'] = Aggregate_Category

        for i in Aggregate_df['Ticker']:
            print i
        '''

        # Equity_df[Equity_df.columns[0]]['MXWO Index']


        self.Equity_Tree_table.delete(*self.Equity_Tree_table.get_children())
        self.Fixed_Income_Tree_table.delete(*self.Fixed_Income_Tree_table.get_children())

        for i in range(len(Equity_list)):
            STOCK = Database_Functions.Fetch(Ticker=Equity_list[i]).dropna().sort_index(ascending=True)
            STOCK['Rolling_std'] = pd.rolling_std(STOCK[str(Equity_list[i])].pct_change().dropna(), 200)
            # self.Equity_Tree_table.insert("",0,'',values=(str(Equity_list[i]),str(Equity_Name_list[i]),round(STOCK['Rolling_std'][-1]*(250**0.5),4)*100.,round(Equity_df[Equity_df.columns[0]][Equity_list[i]],2)))
            self.Equity_Tree_table.insert("", 0, '', values=(
            str(Equity_list[i]), str(Equity_Name_list[i]), round(Portfolio_Equity[4].ix[i], 4) * 100.,
            round(Equity_df.ix[i], 2)))

        for i in range(len(Fixed_income_list)):
            STOCK = Database_Functions.Fetch(Ticker=Fixed_income_list[i]).dropna().sort_index(ascending=True)
            STOCK['Rolling_std'] = pd.rolling_std(STOCK[str(Fixed_income_list[i])].pct_change().dropna(), 200)
            # self.Fixed_Income_Tree_table.insert("",0,'',values=(str(Fixed_income_list[i]),str(Fixed_income_Name_list[i]),round(STOCK['Rolling_std'][-1]*(250**0.5),4)*100.,round(Fixed_income_df[Fixed_income_df.columns[0]][Fixed_income_list[i]],2)))
            self.Fixed_Income_Tree_table.insert("", 0, '', values=(
            str(Fixed_income_list[i]), str(Fixed_income_Name_list[i]),
            round(STOCK['Rolling_std'][-1] * (250 ** 0.5), 4) * 100., round(Fixed_income_df.ix[i], 2)))
            '''
        for i in range(len(Aggregate_Tickers)):
            STOCK = Database_Functions.Fetch(Ticker=Aggregate_Tickers[i]).dropna().sort_index(ascending=True)
            STOCK['Rolling_std'] = pd.rolling_std(STOCK[str(Aggregate_Tickers[i])].pct_change().dropna(),200)
            #self.Fixed_Income_Tree_table.insert("",0,'',values=(str(Fixed_income_list[i]),str(Fixed_income_Name_list[i]),round(STOCK['Rolling_std'][-1]*(250**0.5),4)*100.,round(Fixed_income_df[Fixed_income_df.columns[0]][Fixed_income_list[i]],2)))
            self.Aggregate_Tree_table.insert("",0,'',values=(str(Aggregate_Tickers[i]),str(Aggregate_Name[i]),round(STOCK['Rolling_std'][-1]*(250**0.5),4)*100.,round(Aggregate_df.ix[i],2)))
            '''

        for i in range(len(Aggregate_Name)):
            STOCK = Database_Functions.Fetch(Ticker=Aggregate_df.index[i]).dropna().sort_index(ascending=True)
            STOCK['Rolling_std'] = pd.rolling_std(STOCK[str(Aggregate_df.index[i])].pct_change().dropna(), 200)
            self.Aggregate_Tree_table.insert("", 0, '', values=(
            str(Aggregate_Category[i]), str(Aggregate_df.index[i]), str(Aggregate_Name[i]),
            round(Aggregate_df.ix[i], 2)))

        self.Portfolio7_frame.update_idletasks()

        Equity_holdings_code = self.Equity_Tree_table.get_children()
        Fixed_income_holdings_code = self.Fixed_Income_Tree_table.get_children()

        Equity_Tickers = []
        Equity_Names = []
        Equity_Weights = []

        for each in Equity_holdings_code:
            Ticker = self.Equity_Tree_table.item(each)['values'][0]
            Name = self.Equity_Tree_table.item(each)['values'][1]
            Weight = self.Equity_Tree_table.item(each)['values'][3]

            Equity_Tickers.append(Ticker)
            Equity_Names.append(Name)
            Equity_Weights.append(float(Weight))

        Equity_holding_df = pd.DataFrame(columns=['Ticker', 'Name', 'Weight'])
        Equity_holding_df['Ticker'] = Equity_Tickers
        Equity_holding_df['Name'] = Equity_Names
        Equity_holding_df['Weight'] = Equity_Weights
        Equity_holding_df['Aggressive_Weight'] = Equity_holding_df['Weight'] * 0.8
        Equity_holding_df['Moderate_Weight'] = Equity_holding_df['Weight'] * 0.6
        Equity_holding_df['Conservative_Weight'] = Equity_holding_df['Weight'] * 0.2
        Equity_holding_df['Category'] = 'Equity'
        # print Equity_holding_df
        # print Equity_holding_df[['Category','Ticker','Name','Aggressive_Weight']]

        Fixed_income_Tickers = []
        Fixed_income_Names = []
        Fixed_income_Weights = []

        for each in Fixed_income_holdings_code:
            Ticker = self.Fixed_Income_Tree_table.item(each)['values'][0]
            Name = self.Fixed_Income_Tree_table.item(each)['values'][1]
            Weight = self.Fixed_Income_Tree_table.item(each)['values'][3]

            Fixed_income_Tickers.append(Ticker)
            Fixed_income_Names.append(Name)
            Fixed_income_Weights.append(float(Weight))

        Fixed_income_holding_df = pd.DataFrame(columns=['Ticker', 'Name', 'Weight'])
        Fixed_income_holding_df['Ticker'] = Fixed_income_Tickers
        Fixed_income_holding_df['Name'] = Fixed_income_Names
        Fixed_income_holding_df['Weight'] = Fixed_income_Weights
        Fixed_income_holding_df['Aggressive_Weight'] = Fixed_income_holding_df['Weight'] * 0.2
        Fixed_income_holding_df['Moderate_Weight'] = Fixed_income_holding_df['Weight'] * 0.4
        Fixed_income_holding_df['Conservative_Weight'] = Fixed_income_holding_df['Weight'] * 0.8
        Fixed_income_holding_df['Category'] = 'Fixed Income'
        # print Fixed_income_holding_df[['Category','Ticker','Name','Aggressive_Weight']]
        # print Fixed_income_holding_df

        Holdings_df = pd.concat([Equity_holding_df, Fixed_income_holding_df], axis=0)

        # print Holdings_df.sort('Moderate_Weight',ascending=False)
        self.Insert_specific_holdings_sort(All_holdings=Holdings_df)

        # Pie Chart=============================================================
        # Aggressive
        self.Aggressive_pie2_Chart.clear()
        Ag_Code = self.Aggressive_Tree_table.get_children()

        Aggressive_Tickers = []
        Aggressive_Weights = []

        for i in range(len(Ag_Code)):
            Ag_Ticker = self.Aggressive_Tree_table.item(Ag_Code[i])['values'][1]
            Ag_Weight = float(self.Aggressive_Tree_table.item(Ag_Code[i])['values'][3])
            Aggressive_Tickers.append(Ag_Ticker)
            Aggressive_Weights.append(Ag_Weight)

        Ag_labels = tuple(Aggressive_Tickers)
        Ag_sizes = Aggressive_Weights

        self.Aggressive_pie2_Chart.pie(Ag_sizes, colors=cm.Blues(np.arange(len(Ag_sizes)) / float(len(Ag_sizes))),
                                       labels=Ag_labels, autopct='%1.1f%%', shadow=True, startangle=150,
                                       labeldistance=0.8)  # ,autopct='%1.1f%%'
        self.Aggressive_pie2_canvas.show()

        # Moderate
        self.Moderate_pie2_Chart.clear()
        Mo_Code = self.Moderate_Tree_table.get_children()

        Moderate_Tickers = []
        Moderate_Weights = []

        for i in range(len(Mo_Code)):
            Mo_Ticker = self.Moderate_Tree_table.item(Mo_Code[i])['values'][1]
            Mo_Weight = float(self.Moderate_Tree_table.item(Mo_Code[i])['values'][3])
            Moderate_Tickers.append(Mo_Ticker)
            Moderate_Weights.append(Mo_Weight)

        Mo_labels = tuple(Moderate_Tickers)
        Mo_sizes = Moderate_Weights

        self.Moderate_pie2_Chart.pie(Mo_sizes, colors=cm.Blues(np.arange(len(Mo_sizes)) / float(len(Mo_sizes))),
                                     labels=Mo_labels, autopct='%1.1f%%', shadow=True, startangle=150,
                                     labeldistance=0.8)  # ,autopct='%1.1f%%'
        self.Moderate_pie2_canvas.show()

        # Conservative
        self.Conservative_pie2_Chart.clear()
        Co_Code = self.Conservative_Tree_table.get_children()

        Conservative_Tickers = []
        Conservative_Weights = []

        for i in range(len(Co_Code)):
            Co_Ticker = self.Conservative_Tree_table.item(Co_Code[i])['values'][1]
            Co_Weight = float(self.Conservative_Tree_table.item(Co_Code[i])['values'][3])
            Conservative_Tickers.append(Co_Ticker)
            Conservative_Weights.append(Co_Weight)

        Co_labels = tuple(Conservative_Tickers)
        Co_sizes = Conservative_Weights

        self.Conservative_pie2_Chart.pie(Co_sizes, colors=cm.Blues(np.arange(len(Co_sizes)) / float(len(Co_sizes))),
                                         labels=Co_labels, autopct='%1.1f%%', shadow=True, startangle=150,
                                         labeldistance=0.8)  # ,autopct='%1.1f%%'
        self.Conservative_pie2_canvas.show()

        # Aggregate
        self.Aggregate_pie2_Chart.clear()
        Agg_Code = self.Aggregate_Tree_table.get_children()

        Aggregate_Tickers = []
        Aggregate_Weights = []

        for i in range(len(Agg_Code)):
            Agg_Ticker = self.Aggregate_Tree_table.item(Agg_Code[i])['values'][1]
            Agg_Weight = float(self.Aggregate_Tree_table.item(Agg_Code[i])['values'][3])
            Aggregate_Tickers.append(Agg_Ticker)
            Aggregate_Weights.append(Agg_Weight)

        Agg_labels = tuple(Aggregate_Tickers)
        Agg_sizes = Aggregate_Weights

        self.Aggregate_pie2_Chart.pie(Agg_sizes, colors=cm.Blues(np.arange(len(Agg_sizes)) / float(len(Agg_sizes))),
                                      labels=Agg_labels, autopct='%1.1f%%', shadow=True, startangle=150,
                                      labeldistance=0.8)  # ,autopct='%1.1f%%'
        self.Aggregate_pie2_canvas.show()

        # Aggregate Category Pie Chart
        Aggregate_df.columns = ['Weight']
        Aggregate_df['Category'] = Aggregate_Category
        Fixed_income_weight_agg = Aggregate_df['Weight'][Aggregate_df.Category == 'Fixed Income'].sum()
        Equity_weight_agg = Aggregate_df['Weight'][Aggregate_df.Category == 'Equity'].sum()

        Agg_category_size = [round(Equity_weight_agg, 2), round(Fixed_income_weight_agg, 2)]

        self.Aggregate_pie_Chart.pie(Agg_category_size,
                                     colors=cm.Blues(np.arange(len(Agg_category_size)) / float(len(Agg_category_size))),
                                     labels=self.labels, autopct='%1.1f%%', shadow=True, startangle=150,
                                     labeldistance=0.8)  # ,autopct='%1.1f%%'
        self.Aggregate_pie_canvas.show()

        global Last_date
        Last_date = str(Aggressive_Return.index[-1])[:10]
        self.Last_StrVar.set(Last_date)

        '''
        ================================================================================================================
        Six Balance Portfolio===========================================================================================
        ================================================================================================================
        '''
        #Graph the NAV in charts
        Fixed_Inccome = self.Balance_Portfolio(Weight = (0.2,0.8))
        Conservative = self.Balance_Portfolio(Weight = (0.3,0.7))
        Moderate = self.Balance_Portfolio(Weight = (0.4,0.6))
        Growth = self.Balance_Portfolio(Weight = (0.5,0.5))
        Aggressive = self.Balance_Portfolio(Weight = (0.6,0.4))
        Very_Aggressive = self.Balance_Portfolio(Weight = (0.7,0.3))

        global All_Portfolo_Performace
        All_Portfolo_Performace = pd.DataFrame()
        All_Portfolo_Performace['Fixed_Income'] = Fixed_Inccome
        All_Portfolo_Performace['Conservative'] = Conservative
        All_Portfolo_Performace['Moderate'] = Moderate
        All_Portfolo_Performace['Growth'] = Growth
        All_Portfolo_Performace['Aggressive'] = Aggressive
        All_Portfolo_Performace['Very_Aggressive'] = Very_Aggressive
        All_Portfolo_Performace = All_Portfolo_Performace.dropna()



        self.BP_Chart.clear()
        self.BP_Chart.plot(Fixed_Inccome.index,Fixed_Inccome, lw=1, alpha=0.5, label='Fixed Income (20,80)')
        self.BP_Chart.plot(Conservative.index,Conservative, lw=1, alpha=0.5, label='Conservative (30,70)')
        self.BP_Chart.plot(Moderate.index,Moderate, lw=1, alpha=0.5, label='Moderate (40,60)')
        self.BP_Chart.plot(Growth.index,Growth, lw=1, alpha=0.5, label='Growth (50,50)')
        self.BP_Chart.plot(Aggressive.index,Aggressive, lw=1, alpha=0.5, label='Aggressive (60,40)')
        self.BP_Chart.plot(Very_Aggressive.index,Very_Aggressive, lw=1, alpha=0.5, label='Very Aggressive (70,30)')
        self.BP_Chart.legend(fontsize=8, loc=2, ncol=2, shadow=True).get_frame().set_facecolor('#F0F0F0')

        self.BP_canvas.show()


        #Show the performance figure
        Performance1 =  self.Performance_Calculate(Fixed_Inccome)
        Performance2 = self.Performance_Calculate(Conservative)
        Performance3 = self.Performance_Calculate(Moderate)
        Performance4 = self.Performance_Calculate(Growth)
        Performance5 = self.Performance_Calculate(Aggressive)
        Performance6 = self.Performance_Calculate(Very_Aggressive)

        Benchmark_df = pd.DataFrame()
        Benchmark_df['Equity'] = World_Index['MXWO Index']
        Benchmark_df['Bond'] = Bond_Index['LBUSTRUU Index']
        Benchmark_df = Benchmark_df[Benchmark_df.index > '1989-02-01'].pct_change()
        Benchmark_df = Benchmark_df.cumsum().apply(np.exp)

        Equity_Performance = self.Performance_Calculate(Benchmark_df['Equity'])
        Bond_Performance = self.Performance_Calculate(Benchmark_df['Bond'])

        self.AR1_StrVar.set(Performance1[0])
        self.AR2_StrVar.set(Performance2[0])
        self.AR3_StrVar.set(Performance3[0])
        self.AR4_StrVar.set(Performance4[0])
        self.AR5_StrVar.set(Performance5[0])
        self.AR6_StrVar.set(Performance6[0])
        self.AR7_StrVar.set(Equity_Performance[0])
        self.AR8_StrVar.set(Bond_Performance[0])

        self.Vol1_StrVar.set(Performance1[1])
        self.Vol2_StrVar.set(Performance2[1])
        self.Vol3_StrVar.set(Performance3[1])
        self.Vol4_StrVar.set(Performance4[1])
        self.Vol5_StrVar.set(Performance5[1])
        self.Vol6_StrVar.set(Performance6[1])
        self.Vol7_StrVar.set(Equity_Performance[1])
        self.Vol8_StrVar.set(Bond_Performance[1])

        self.Sharpe1_StrVar.set(Performance1[2])
        self.Sharpe2_StrVar.set(Performance2[2])
        self.Sharpe3_StrVar.set(Performance3[2])
        self.Sharpe4_StrVar.set(Performance4[2])
        self.Sharpe5_StrVar.set(Performance5[2])
        self.Sharpe6_StrVar.set(Performance6[2])
        self.Sharpe7_StrVar.set(Equity_Performance[2])
        self.Sharpe8_StrVar.set(Bond_Performance[2])

        #===============================================================================================================
        #All Weather Portfolio==========================================================================================

        #===============================================================================================================

        self.Portfolio9_frame.update_idletasks()
        self.Portfolio8_frame.update_idletasks()
        self.Portfolio7_frame.update_idletasks()
        self.Portfolio1_frame.update_idletasks()
        self.Portfolio2_frame.update_idletasks()
        self.Portfolio3_frame.update_idletasks()
        self.Portfolio5_frame.update_idletasks()

        #self.All_Portfolio_Output()







    def All_Portfolio_Output(self):
        Today = datetime.datetime.today()
        Time_For_FileName = Today.strftime('%Y%m%d%H%M')
        #print Portfolios_df_output
        Daily_Return = Portfolios_df_output.pct_change()*100.
        Monthly_Return = Portfolios_df_output.resample('M',how='last').pct_change()*100.


        Low_Vol_Holdings = pd.concat([Equity_Portfolio_Weight,Fixed_income_Portfolio_Weight],axis=1)
        Low_Vol_Holdings = Low_Vol_Holdings#.dropna()

        AWP_xlsx = pd.DataFrame()
        AWP_xlsx['Quandrant'] = AWP_Score

        FileName = Time_For_FileName + ' Portfolio_info_Volatility_Select'
        writer = pd.ExcelWriter("D:/Taishin_Platform/Portfolio_Performance/" + str(FileName)+".xlsx")

        Portfolios_df_output.to_excel(writer,'Cumulative Return')
        Daily_Return.to_excel(writer,'Daily Return(%)')
        Monthly_Return.to_excel(writer,'Monthly Return(%)')
        Low_Vol_Holdings.to_excel(writer,'Low Vol Holdings (Daily)')
        Low_Vol_Holdings.resample('MS',how='first').to_excel(writer, 'Low Vol Holdings (Monthly)')

        GMP_Weight_df.to_excel(writer,' Momentum Daily Weights')
        GMP_Weight_df.resample('MS',how='first').to_excel(writer,' Momentum Monthly Weights')
        Strategic_Bond_Weight.to_excel(writer,'Strategic Bond Weights')
        Strategic_Bond_Weight.resample('MS',how='first').to_excel(writer,'Strategic Bond Monthly Weights')
        AWP_xlsx.to_excel(writer,'AW Quandrant')

        writer.save()
        self.Done_Messenger(info=Time_For_FileName+'.xlsx created successfully!!')

    def Balance_Portfolio(self,Weight):
        df = pd.DataFrame()
        df['Equity'] = Database_Functions.Fetch('MXWO Index')['MXWO Index']
        df['Bond'] = Database_Functions.Fetch('LBUSTRUU Index')['LBUSTRUU Index']
        df = df.dropna()
        df = df[df.index > '1989-02-01']
        #df = df
        Return_df = df.pct_change()

        Initial_amount = 1000000
        Equity_amount = Initial_amount * Weight[0]
        Bond_amount = Initial_amount * Weight[1]
        #print Weight[0],Weight[1]

        Shares_df = (Initial_amount / df).resample('M', how='last')
        Shares_df['Equity'] = Shares_df['Equity'] * Weight[0]
        Shares_df['Bond'] = Shares_df['Bond'] * Weight[1]
        Shares_df = Shares_df.resample('D', how='last').fillna(method='ffill')
        Amount_df = (Shares_df * df).dropna()
        Amount_df['Sum'] = Amount_df.sum(axis=1)

        Weight_df = pd.DataFrame()
        Weight_df['Equity'] = Amount_df['Equity'] / Amount_df['Sum']
        Weight_df['Bond'] = Amount_df['Bond'] / Amount_df['Sum']
        Portfolio_EC = (Weight_df * Return_df).sum(axis=1).cumsum().apply(np.exp)

        return Portfolio_EC

    def Performance_Calculate(self,NAV):
        Horizon = (NAV.index[-1] - NAV.index[0]).days/365.

        AR = (NAV[-1]**(1./Horizon) - 1.)*100.
        Vol = (np.std(NAV.pct_change())*(250**0.5))*100.
        Sharpe = AR/Vol








        return str(round(AR,2))+'%',str(round(Vol,2))+'%',str(round(Sharpe,2))




    def Global_Momentum_Portfolio_Algorithm(self, Tickers, lookback):
        Process = ['Price', 'Daily_Return', 'Momentum', 'Volatility']

        Price_df = pd.DataFrame()

        for Ticker in Tickers:
            Price_df[Ticker] = Database_Functions.Fetch(Ticker=Ticker)[Ticker]
        Price_df = Price_df.dropna()

        Panel = pd.Panel(items=Process, major_axis=Price_df.index, minor_axis=Tickers)
        Panel['Price'] = Price_df
        Panel['Daily_Return'] = Panel['Price'].pct_change()

        Panel['Momentum'] = Panel['Price'].pct_change(lookback).resample('M', how='mean')
        Panel['Momentum'] = Panel['Momentum'].fillna(method='ffill')

        Panel['Volatility'] = pd.rolling_std(Panel['Daily_Return'], lookback).resample('M', how='mean')
        Panel['Volatility'] = Panel['Volatility'].fillna(method='ffill')

        Panel['Volatility_Inverse'] = 1. / Panel['Volatility']

        Panel['Momentum_Rank'] = Panel['Momentum'].rank(axis=1, ascending=False)
        Panel['Selected_asset'] = np.where(Panel['Momentum_Rank'] <= len(Tickers) / 3, 1, 0)

        Panel['Select_Volatility'] = Panel['Volatility_Inverse'] * Panel['Selected_asset']

        Volatility_weight = Panel['Select_Volatility'].replace(0, np.nan)
        Volatility_sum = Volatility_weight.sum(axis=1).replace(0, np.nan)

        Weight_df = pd.DataFrame()
        for Ticker in Volatility_weight.columns:
            Weight_df[Ticker] = Volatility_weight[Ticker] / Volatility_sum

        Panel['Volatility_weight'] = Weight_df

        Portfolio_Return_temp = Panel['Volatility_weight'] * Panel['Daily_Return']
        Portfolio_Return = Portfolio_Return_temp.sum(axis=1)
        Portfolio_CR = Portfolio_Return.cumsum().apply(np.exp)

        # Output:Last weight, Cumulative Return
        #print Panel['Volatility_weight']
        return Panel['Volatility_weight'].ix[-1].tolist(), Portfolio_CR,Panel['Volatility_weight']

    def Portfolio(self, Tickers):
        df = pd.DataFrame()

        for Ticker in Tickers:
            df[Ticker] = Database_Functions.Fetch(Ticker=Ticker)[Ticker]

        Return_df = df.dropna().pct_change()

        Volatility_df = 1. / pd.rolling_std(Return_df, 60) * (250 ** 0.5)
        Volatility_df['Sum'] = Volatility_df.sum(axis=1)

        Weight_df = pd.DataFrame()

        for Ticker in Tickers:
            Weight_df[Ticker] = Volatility_df[Ticker] / Volatility_df['Sum']

        Weight_df = Weight_df.resample('M', how='mean').resample('D', how='last')
        Weight_df = Weight_df.fillna(method='ffill').dropna()

        Holding_Return = (Weight_df * Return_df).dropna()

        Fund_Return = Holding_Return.sum(axis=1)
        CR = Fund_Return.cumsum().apply(np.exp)

        # print Weight_df.ix[-1].T
        # print Weight_df.ix[-1]

        return Weight_df.ix[-1].T, CR,Weight_df

    def Portfolio_dollar(self, Entry, Tree):

        def Save_to_Excel():
            File = str(Name_entry.get())
            if File == '':
                File = 'Client'

            xls_name = File + '_' + Kind + '_' + Last_date + '.xlsx'

            workbook = xlsxwriter.Workbook('D:/Taishin_Platform/Portfolio/' + str(xls_name))
            worksheet = workbook.add_worksheet()

            row = 3
            worksheet.write(row, 0, 'Category')
            worksheet.write(row, 1, 'Ticker')
            worksheet.write(row, 2, 'Name')
            worksheet.write(row, 3, 'Weight')
            worksheet.write(row, 4, 'Dollar')

            worksheet.write(0, 0, str(File))
            worksheet.write(0, 1, str(Kind))
            worksheet.write(0, 2, str(AUM))
            worksheet.write(0, 3, str(Last_date))

            xls_write_code = Holdings_Dollar_Tree_table.get_children()

            for i in range(len(Code)):
                Category = Holdings_Dollar_Tree_table.item(xls_write_code[i])['values'][0]
                Ticker = Holdings_Dollar_Tree_table.item(xls_write_code[i])['values'][1]
                Name = Holdings_Dollar_Tree_table.item(xls_write_code[i])['values'][2]
                Weight = float(Holdings_Dollar_Tree_table.item(xls_write_code[i])['values'][3])
                Dollar = Weight / 100. * AUM
                # print Category,Ticker,Name,Weight,Dollar

                row = row + 1

                worksheet.write(row, 0, Category)
                worksheet.write(row, 1, Ticker)
                worksheet.write(row, 2, Name)
                worksheet.write(row, 3, Weight)
                worksheet.write(row, 4, Dollar)

            workbook.close()
            self.Done_Messenger(info=xls_name + ' created successfully!')

        # global Kind

        if Tree == self.Moderate_Tree_table:
            Kind = 'Moderate'
        elif Tree == self.Aggressive_Tree_table:
            Kind = 'Aggressive'
        elif Tree == self.Conservative_Tree_table:
            Kind = 'Conservative'
        elif Tree == self.Aggregate_Tree_table:
            Kind = 'Aggregate'
        elif Tree == self.Bond_Tree_table:
            Kind = 'Global Strategic Bond'
        else:
            Kind = 'Global Momentum Portfolio'

        # try:
        AUM = float(Entry.get())
        Code = Tree.get_children()

        root = tk.Tk()
        root.wm_title('Portfolio Weight in Dollar')
        root.geometry('620x420')
        root.iconbitmap('D:/Taishin_Platform/pics/chart_diagram_analytics_business_flat_icon-512.ico')

        Holdings_Dollar_Tree_table = ttk.Treeview(root, height="12")

        Holdings_Dollar_Tree_table["columns"] = ("column1", "column2", 'column3', 'column4', 'column5')
        Holdings_Dollar_Tree_table.column("#0", width=10, anchor='e')
        Holdings_Dollar_Tree_table.column("column1", width=60, anchor='center')
        Holdings_Dollar_Tree_table.column("column2", width=80, anchor='w')
        Holdings_Dollar_Tree_table.column("column3", width=250, anchor='w')
        Holdings_Dollar_Tree_table.column("column4", width=80, anchor='e')
        Holdings_Dollar_Tree_table.column("column5", width=120, anchor='e')

        Holdings_Dollar_Tree_table.heading('#0', text='')
        Holdings_Dollar_Tree_table.heading("column1", text="Category")
        Holdings_Dollar_Tree_table.heading("column2", text="Ticker")
        Holdings_Dollar_Tree_table.heading("column3", text="Name")
        Holdings_Dollar_Tree_table.heading("column4", text="Weight(%)",
                                           command=lambda: self.treeview_sort_column(self.Holdings_Dollar_Tree_table,
                                                                                     "column4", False))
        Holdings_Dollar_Tree_table.heading("column5", text="Dollar",
                                           command=lambda: self.treeview_sort_column(self.Holdings_Dollar_Tree_table,
                                                                                     "column5", False))

        for i in range(len(Code)):
            Category = Tree.item(Code[i])['values'][0]
            Ticker = Tree.item(Code[i])['values'][1]
            Name = Tree.item(Code[i])['values'][2]
            Weight = float(Tree.item(Code[i])['values'][3])
            Dollar = Weight / 100. * AUM

            Holdings_Dollar_Tree_table.insert("", i, '', values=(
            str(Category), str(Ticker), str(Name), str(Weight), str(round(Dollar, 2))))

        Holdings_Dollar_Tree_table.place(x=10, y=10)

        Total_label = tk.Label(root, text='Total: $' + str(int(AUM)), font=('Arial', 12, 'bold'))
        Total_label.place(x=610, y=280, anchor='ne')

        ttk.Separator(root, orient=tk.HORIZONTAL).place(x=10, y=310, width=600)

        # --------------------------------------------------------------------------------------------

        style = ttk.Style()
        style.configure("Menu.TButton", foreground="black", background="#181818")

        Name_label = tk.Label(root, text='File Name', font=('Arial', 12, 'bold'))
        Name_label.place(x=360, y=328)

        Name_entry = tk.Entry(root)
        Name_entry.place(x=610, y=330, anchor='ne')

        To_Excel_Button = ttk.Button(root, text='Save as Excel  ', command=Save_to_Excel, style='Menu.TButton')
        To_Excel_Button.place(x=610, y=380, anchor='ne')

        root.mainloop()

        '''
        except:
            root = tk.Tk()
            root.wm_title('Error !')
            root.geometry('280x80')
            #root.iconbitmap('Icon.ico')
            
            Label = tk.Label(root,text = 'Empty portfolio value !!',font=('Arial',14,'bold'))
            Label.pack()
            
            Label2 = tk.Label(root,text = 'Please enter portfolio value in entrybox.',font=('Arial',12))
            Label2.pack()
            
            
            root.mainloop()
        '''

    def Done_Messenger(self, info):
        Message_root = tk.Tk()
        # Message_root.title('Done')
        Message_root.geometry('500x200')
        Message_root.iconbitmap('D:/Taishin_Platform/pics/chart_diagram_analytics_business_flat_icon-512.ico')
        Message_Label = tk.Label(Message_root, text=info, font=('Arial', 12, 'bold'), fg='black', background="#F0F0F0")
        Message_Label.place(x=40, y=65)

        OK_Button = ttk.Button(Message_root, width=15, text=u"  OK  ", command=Message_root.destroy)
        OK_Button.place(x=130, y=120)

        Message_root.config(background="#F0F0F0")
        tk.mainloop()

    def Insert_holdings(self, Tree, Entry):
        try:
            Ticker = str(Entry.get()).split(' , ')[0]
            Name = str(Entry.get()).split(' , ')[1]

            STOCK = Database_Functions.Fetch(Ticker=Ticker).dropna().sort_index(ascending=True)
            STOCK['Rolling_std'] = pd.rolling_std(STOCK[str(Ticker)].pct_change().dropna(), 200)

            # for i in range(5):
            Tree.insert("", 0, '', values=(
            str(Ticker), str(Name), str(round(STOCK['Rolling_std'][-1] * (250 ** 0.5), 4) * 100.) + '%'))
        except:
            pass

    def Insert_specific_holdings_sort(self, All_holdings):
        self.Aggressive_Tree_table.delete(*self.Aggressive_Tree_table.get_children())
        self.Moderate_Tree_table.delete(*self.Moderate_Tree_table.get_children())
        self.Conservative_Tree_table.delete(*self.Conservative_Tree_table.get_children())

        Aggressive_holdings = All_holdings.sort('Aggressive_Weight', ascending=False).reset_index()
        Moderate_holdings = All_holdings.sort('Moderate_Weight', ascending=False).reset_index()
        Conservative_holdings = All_holdings.sort('Conservative_Weight', ascending=False).reset_index()

        # print Aggressive_holdings,Moderate_holdings,  Conservative_holdings



        for i in range(len(All_holdings)):
            self.Aggressive_Tree_table.insert("", i, '', values=(
            Aggressive_holdings['Category'][i], Aggressive_holdings['Ticker'][i], Aggressive_holdings['Name'][i],
            round(Aggressive_holdings['Aggressive_Weight'][i], 2)))
            self.Moderate_Tree_table.insert("", i, '', values=(
            Moderate_holdings['Category'][i], Moderate_holdings['Ticker'][i], Moderate_holdings['Name'][i],
            round(Moderate_holdings['Moderate_Weight'][i], 2)))
            self.Conservative_Tree_table.insert("", i, '', values=(
            Conservative_holdings['Category'][i], Conservative_holdings['Ticker'][i], Conservative_holdings['Name'][i],
            round(Conservative_holdings['Conservative_Weight'][i], 2)))

    def Fetch_Equity_holdings(self):
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
        cursor = connection.cursor()
        cursor.execute("SELECT * FROM Equity_holdings")

        data = cursor.fetchall()

        connection.close()

        for i in range(len(data)):
            Ticker = data[i][1]
            Name = data[i][2]
            STOCK = Database_Functions.Fetch(Ticker=Ticker).dropna().sort_index(ascending=True)
            STOCK['Rolling_std'] = pd.rolling_std(STOCK[str(Ticker)].pct_change(), 200)

            self.Equity_Tree_table.insert("", i, '', values=(
            str(Ticker), str(Name), str(round(STOCK['Rolling_std'][-1] * (250 ** 0.5), 4) * 100.) + '%'))

    def Fetch_Fixed_income_holdings(self):
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
        cursor = connection.cursor()
        cursor.execute("SELECT * FROM Fixed_income_holdings")

        data = cursor.fetchall()

        connection.close()

        for i in range(len(data)):
            Ticker = data[i][1]
            Name = data[i][2]
            STOCK = Database_Functions.Fetch(Ticker=Ticker).dropna().sort_index(ascending=True)
            STOCK['Rolling_std'] = pd.rolling_std(STOCK[str(Ticker)].pct_change(), 200)

            self.Fixed_Income_Tree_table.insert("", i, '', values=(
            str(Ticker), str(Name), str(round(STOCK['Rolling_std'][-1] * (250 ** 0.5), 4) * 100.) + '%'))

    def Delete_holdings_Equity(self):
        selected_item = self.Equity_Tree_table.selection()[0]  ## get selected item
        self.Equity_Tree_table.delete(selected_item)

    def Delete_holdings_Fixed_income(self):
        selected_item = self.Fixed_Income_Tree_table.selection()[0]  ## get selected item
        self.Fixed_Income_Tree_table.delete(selected_item)

    def Performance_print(self, CR, FileName,Weight_df,AR,Std,Sharpe):
        Summary_df = pd.DataFrame(index = ['Annual Return','Volatility','Sharpe ratio'])
        Summary_df['Portfolio'] = [AR,Std,Sharpe]

        df = pd.DataFrame()
        df['Cumulative_Return'] = CR['Portfolio']
        df['Daily_Return'] = CR['Portfolio'].pct_change()
        df = df.dropna()

        Monthly_Return = pd.DataFrame()
        Monthly_Return['Monthly Return(%)'] = df['Cumulative_Return'].resample('M', how='last').pct_change() * 100.
        #print Monthly_Return.index
        Monthly_Return.index = Monthly_Return.index.format()
        Monthly_Return = Monthly_Return.dropna()

        Annual_Return = pd.DataFrame()
        Annual_Return['Annual Return(%)'] = df['Cumulative_Return'].resample('A', how='last').pct_change() * 100.
        Annual_Return.index = Annual_Return.index.format()
        Annual_Return = Annual_Return.dropna()

        writer = pd.ExcelWriter("D:/Taishin_Platform/Portfolio_Performance/" + str(FileName))

        Summary_df.to_excel(writer,'Summary')
        df.to_excel(writer, 'Cumulative Return (Daily)')
        Monthly_Return.to_excel(writer, 'Monthly Return')
        Annual_Return.to_excel(writer, 'Annual Return')
        Weight_df.to_excel(writer,'Historical Weight')

        writer.save()
        info = FileName + ' saved successfully.'
        self.Done_Messenger(info=info)





    def treeview_sort_column(self, tv, col, reverse):
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        l.sort(key=lambda t: float(t[0]), reverse=reverse)
        #      ^^^^^^^^^^^^^^^^^^^^^^^

        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)

        tv.heading(col,
                   command=lambda: self.treeview_sort_column(tv, col, not reverse))

    def treeview_string_sort_column(self, tv, col, reverse):
        l = [(tv.set(k, col), k) for k in tv.get_children('')]
        l.sort(key=lambda t: str(t[0]), reverse=reverse)
        #      ^^^^^^^^^^^^^^^^^^^^^^^

        for index, (val, k) in enumerate(l):
            tv.move(k, '', index)

        tv.heading(col,
                   command=lambda: self.treeview_string_sort_column(tv, col, not reverse))


if __name__ == "__main__":
    warnings.simplefilter(action="ignore", category=FutureWarning)
    app = GUI()
