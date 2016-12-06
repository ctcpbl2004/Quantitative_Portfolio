"""
Created on Wed Jun 22 14:20:52 2016

@author: Raymond
"""

import Tkinter as tk
import ttk
import sqlite3
#import Database_Functions
import FileDialog
import re
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import numpy as np
from tkFileDialog import askopenfilename
import xlsxwriter
import pandas as pd
import warnings
Menu_FONT = ('Arial',14)
Menu_Close_FONT = ('Arial',10)

NORM_FONT = ('Arial',12)
Entry_FONT = ('Arial',12,'bold')
Figure_FONT = ('Arial',20,'bold')
EXAPMLE_FONT = ('Arial',8)
#-----------------------Database_Functions--------------
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
        workbook = xlsxwriter.Workbook('D:/Taishin_Platform/Update/Update.xlsx')
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
        cursor.execute("INSERT OR REPLACE INTO Strategy_table_short VALUES (?, ?, ?, ?, ?, ?, ?, ?)",(Ticker, Name, '' ,'' ,'','','',''))
        connection.commit()
        connection.close()
    
    #Delete_data(Ticker = 'AUD',Name = 'AUDUSD')
    @staticmethod
    def Delete_Index_to_Strategy_table(Ticker,Name):
        connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
        cursor = connection.cursor()
        cursor.execute("DELETE FROM Strategy_table WHERE Ticker=? AND Name=?",(Ticker,Name))
        cursor.execute("DELETE FROM Strategy_table_short WHERE Ticker=? AND Name=?",(Ticker,Name))
        connection.commit()
        connection.close()
        Database_Functions.Fetch_All_Data()












#-----------------------Autocomplete--------------------
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
                    self.lb = tk.Listbox(font=NORM_FONT,width = 55, background="#181818", fg="white",selectforeground='black',
                                         selectbackground="#FF9C29",highlightcolor="#181818",activestyle=tk.NONE)
                    self.lb.bind("<Double-Button-1>", self.selection)
                    self.lb.bind("<Right>", self.selection)
                    self.lb.place(x=self.winfo_x()+20, y=self.winfo_y()*4+self.winfo_height()*4+2)
                    self.lb_up = True
                
                self.lb.delete(0, tk.END)
                for w in words:
                    self.lb.insert(tk.END,w)
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
                index = str(int(index)-1)                
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
                index = str(int(index)+1)        
                self.lb.selection_set(first=index)
                self.lb.activate(index) 

    def comparison(self):
        pattern = re.compile('.*' + self.var.get() + '.*')
        return [w for w in self.lista if re.match(pattern, w)]



#---------------------End of Autocomplete--------------
def Combined_list(list1,list2):
    if len(list1) != len(list2):
        print 'Lists number are not equal!!!!'
    else:
        Combined_list = []
        for i in range(len(list1)):
            Combined_list.append(list1[i]+' , '+list2[i])
    return Combined_list

def Get_data():
    Ticker = str(Ticker_Entry.get()).split(' , ')[0]
    STOCK = Database_Functions.Fetch(Ticker=Ticker).dropna().sort_index(ascending=False)
    
    Number = len(STOCK)
    for i in range(Number):
        Query_Table.insert("",i,values=(str(STOCK.index[i])[:10],STOCK[Ticker][i]))
    
    Stock_Chart.clear()
    Stock_Chart.plot( STOCK.index, STOCK[Ticker],lw=1,color='#01485E',alpha=0.5)
    Stock_Chart.fill_between(STOCK.index, STOCK[Ticker],facecolor='#107B8C')
    #Stock_Chart.grid(True,color='white')
    #Stock_Chart.set_ylabel('Cumulative Return(%)', fontsize=10)
    Stock_Chart.set_ylim([min(STOCK[Ticker])*0.9, max(STOCK[Ticker])*1.1])
    canvas1.show()
    
    STOCK['Return'] = STOCK[Ticker].pct_change()
    STOCK['Cummulative Return'] = STOCK['Return'].cumsum().apply(np.exp)
    
    Period = (STOCK.index[-1] - STOCK.index[0]).days/365.
    Annual_Return = STOCK['Cummulative Return'][-1]**(1/Period) -1 
    Standard_Deviation = np.std(STOCK['Return'])*(250**0.5)
    Sharpe_ratio = (Annual_Return - 0.02)/Standard_Deviation
    
    Annual_Return_Var.set(str(round(Annual_Return*100.,2))+'%')
    Standard_Deviation_Var.set(str(round(Standard_Deviation*100.,2))+'%')
    Sharpe_Var.set(str(round(Sharpe_ratio,2)))


def treeview_sort_column(tv, col, reverse):
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    l.sort(key=lambda t: float(t[0]), reverse=reverse)
    #      ^^^^^^^^^^^^^^^^^^^^^^^

    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)

    tv.heading(col,
              command=lambda :treeview_sort_column(tv, col, not reverse))

def treeview_string_sort_column(tv, col, reverse):
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    l.sort(key=lambda t: str(t[0]), reverse=reverse)
    #      ^^^^^^^^^^^^^^^^^^^^^^^

    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)

    tv.heading(col,
              command=lambda :treeview_string_sort_column(tv, col, not reverse))




def Import_Update_table():
    connection = sqlite3.connect('D:/Taishin_Platform/db/Taishin_Database.db')
    cursor = connection.cursor()
    Database_Functions.Fetch_All_Data()
    cursor.execute("SELECT * FROM Update_table")
    
    data = cursor.fetchall()
    
    Data_Table.delete(*Data_Table.get_children())    
    
    i = 0
    for row in data:
        i = i +1
        Data_Table.insert("",i,text=str(i),values=(row[0],row[1],int(row[2]),row[3],row[4]))





def Create_Update_File():
    Database_Functions.Create_Update_xlsx(Start='2016-06-01')
    
    Message_root = tk.Tk()
    Message_root.title('File completed')
    Message_root.geometry('300x200')
    Message_root.iconbitmap('D:/Taishin_Platform/pics/Data_Manager.ico')
    Message_Label = tk.Label(Message_root,text='Update.xlsx created successfully.',font=NORM_FONT,fg='black',background="#F0F0F0")
    Message_Label.place(x=40,y=65)
    
    
    OK_Button=tk.Button(Message_root,width=15, text =u"  OK  ",command = Message_root.destroy,font=NORM_FONT,relief='raised',fg='black',bg='#F0F0F0',activebackground='#F0F0F0')
    OK_Button.config(height=2)
    OK_Button.place(x=130, y=120)
    
    
    Message_root.config(background="#F0F0F0")


def Import_Data():

    def Add_New_Index_print():
        File = File_Entry.get()
        
        if File != '':
            pass
        else:
            Error_Messenger('Error, empty variables!')
        
        Database_Functions.Data_to_db(File)
        Import_Update_table()
        Done_Messenger(info='Done')
        tk.mainloop()
    
    def Select_file():
        tk.Tk().withdraw()
        filename = askopenfilename()
        File_Entry.delete(0,'end')
        File_Entry.insert(0,filename)
        
    #Main window===============================================================
    Add_root = tk.Tk()
    Add_root.title('Data Update')
    Add_root.geometry('500x200')
    Add_root.iconbitmap('D:/Taishin_Platform/pics/Data_Manager.ico')


    #Title Label =============================================================

    File_Label = tk.Label(Add_root,text='File (.xlsx)',font=NORM_FONT,fg='black',background="#F0F0F0")
    File_Label.place(x=40,y=40)
    
    
    #Example Label =============================================================

    File_Example_Label = tk.Label(Add_root,text='Ex: Data.xlsx',font=EXAPMLE_FONT,fg='black',background="#F0F0F0")
    File_Example_Label.place(x=50,y=60)

    #Variable Entry============================================================

    File_Entry = tk.Entry(Add_root, width = 40,justify="left",font=NORM_FONT,background = 'white',fg = 'black',bd=1)
    File_Entry.place(x=120,y=40)

    #Add Button==--============================================================
    Select_file_Button = ttk.Button(Add_root,text='  Select File  ',command=Select_file)
    Select_file_Button.place(x=150, y=100)





    Add_Index_Button = ttk.Button(Add_root,text='  Update Data  ',command=Add_New_Index_print)
    Add_Index_Button.place(x=150, y=150)
    

    Add_root.config(background="#F0F0F0")
    tk.mainloop()







def Add_New_Index():

    def Add_New_Index_print():
        Ticker = Ticker_Entry.get()
        Name = Name_Entry.get()
        File = File_Entry.get()
        
        if (Ticker != '')&(Name != '')&(File != ''):
            pass
        else:
            Error_Messenger('Error, empty variables!')
        
        Database_Functions.Add_Index(File_name=File,Ticker=Ticker,Name=Name)
        Import_Update_table()
        Database_Functions.Add_Index_to_Strategy_table(Ticker,Name)
        Done_Messenger(info='Done')
        
        
    def Select_file():
        tk.Tk().withdraw()
        filename = askopenfilename()
        File_Entry.delete(0,'end')
        File_Entry.insert(0,filename)
    
    
    
    #Main window===============================================================
    Add_root = tk.Tk()
    Add_root.title('Add New Index')
    Add_root.geometry('300x400')
    Add_root.iconbitmap('D:/Taishin_Platform/pics/Data_Manager.ico')


    #Title Label =============================================================
    Ticker_Label = tk.Label(Add_root,text='Ticker',font=NORM_FONT,fg='black',background="#F0F0F0")
    Ticker_Label.place(x=40,y=20)


    Name_Label = tk.Label(Add_root,text='Name',font=NORM_FONT,fg='black',background="#F0F0F0")
    Name_Label.place(x=40,y=80)

    File_Label = tk.Label(Add_root,text='File (.xlsx)',font=NORM_FONT,fg='black',background="#F0F0F0")
    File_Label.place(x=40,y=140)
    
    
    #Example Label =============================================================
    Ticker_Example_Label = tk.Label(Add_root,text='Ex: MXWO Index (Same as Bloomberg Excel API)',font=EXAPMLE_FONT,fg='black',background="#F0F0F0")
    Ticker_Example_Label.place(x=50,y=40)

    Name_Example_Label = tk.Label(Add_root,text='Ex: MSCI World Index',font=EXAPMLE_FONT,fg='black',background="#F0F0F0")
    Name_Example_Label.place(x=50,y=100)

    File_Example_Label = tk.Label(Add_root,text='Ex: Data.xlsx',font=EXAPMLE_FONT,fg='black',background="#F0F0F0")
    File_Example_Label.place(x=50,y=160)

    #Variable Entry============================================================
    Ticker_Entry = tk.Entry(Add_root, width = 15,justify="left",font=NORM_FONT,background = 'white',fg = 'black',bd=1)
    Ticker_Entry.place(x=120,y=20)

    Name_Entry = tk.Entry(Add_root, width = 15,justify="left",font=NORM_FONT,background = 'white',fg = 'black',bd=1)
    Name_Entry.place(x=120,y=80)

    File_Entry = tk.Entry(Add_root, width = 15,justify="left",font=NORM_FONT,background = 'white',fg = 'black',bd=1)
    File_Entry.place(x=120,y=140)

    #Add Button==--============================================================
    Select_file_Button = ttk.Button(Add_root,text='  Select File  ',command=Select_file)
    Select_file_Button.place(x=120, y=180)




    Add_Index_Button = ttk.Button(Add_root,text='  Add New Index  ',command=Add_New_Index_print)
    Add_Index_Button.place(x=150, y=300)
    

    Add_root.config(background="#F0F0F0")
    tk.mainloop()



def Delete_Index():
    
    
    def Delete_Index_print():
        Ticker = Ticker_Entry.get()
        Name = Name_Entry.get()
        
        if (Ticker != '')&(Name != ''):
            pass
        else:
            Error_Messenger('Error, empty variables!')
        
        Database_Functions.Delete_data(Ticker=Ticker,Name = Name)
        Import_Update_table()
        Database_Functions.Delete_Index_to_Strategy_table(Ticker,Name)
        Done_Messenger(info='Done')
        
        
    #Main window===============================================================
    Add_root = tk.Tk()
    Add_root.title('Delete Index')
    Add_root.geometry('300x400')
    Add_root.iconbitmap('D:/Taishin_Platform/pics/Data_Manager.ico')


    #Title Label =============================================================
    Ticker_Label = tk.Label(Add_root,text='Ticker',font=NORM_FONT,fg='black',background="#F0F0F0")
    Ticker_Label.place(x=40,y=20)


    Name_Label = tk.Label(Add_root,text='Name',font=NORM_FONT,fg='black',background="#F0F0F0")
    Name_Label.place(x=40,y=80)

    
    
    #Example Label =============================================================
    Ticker_Example_Label = tk.Label(Add_root,text='Ex: MXWO Index (Same as Bloomberg Excel API)',font=EXAPMLE_FONT,fg='black',background="#F0F0F0")
    Ticker_Example_Label.place(x=50,y=40)

    Name_Example_Label = tk.Label(Add_root,text='Ex: MSCI World Index',font=EXAPMLE_FONT,fg='black',background="#F0F0F0")
    Name_Example_Label.place(x=50,y=100)


    #Variable Entry============================================================
    Ticker_Entry = tk.Entry(Add_root, width = 15,justify="left",font=NORM_FONT,background = 'white',fg = 'black',bd=1)
    Ticker_Entry.place(x=120,y=20)

    Name_Entry = tk.Entry(Add_root, width = 15,justify="left",font=NORM_FONT,background = 'white',fg = 'black',bd=1)
    Name_Entry.place(x=120,y=80)


    #Add Button==--============================================================

    Add_Index_Button = ttk.Button(Add_root,text='  Delete  ',command=Delete_Index_print)
    Add_Index_Button.place(x=150, y=300)
    

    Add_root.config(background="#F0F0F0")
    tk.mainloop()







def Error_Messenger(Error_info):
    Message_root = tk.Tk()
    Message_root.title('Error')
    Message_root.geometry('300x200')
    Message_root.iconbitmap('D:/Taishin_Platform/pics/Data_Manager.ico')
    Message_Label = tk.Label(Message_root,text=Error_info,font=NORM_FONT,fg='black',background="#F0F0F0")
    Message_Label.place(x=40,y=65)
    
    
    OK_Button=ttk.Button(Message_root,width=15, text =u"  OK  ",command = Message_root.destroy)
    OK_Button.place(x=130, y=120)
  
    
    Message_root.config(background="#F0F0F0")
    tk.mainloop()
    
def Done_Messenger(info):
    Message_root = tk.Tk()
    #Message_root.title('Done')
    Message_root.geometry('300x200')
    Message_root.iconbitmap('D:/Taishin_Platform/pics/Data_Manager.ico')
    Message_Label = tk.Label(Message_root,text=info,font=NORM_FONT,fg='black',background="#F0F0F0")
    Message_Label.place(x=40,y=65)
    
    
    OK_Button=ttk.Button(Message_root,width=15, text =u"  OK  ",command = Message_root.destroy)
    OK_Button.place(x=130, y=120)
  
    
    Message_root.config(background="#F0F0F0")
    tk.mainloop()


#For tkinter window seting
def dragwin(event):
    x = root.winfo_pointerx() - root._offsetx
    y = root.winfo_pointery() - root._offsety
    root.geometry('+{x}+{y}'.format(x=x,y=y))

def clickwin(event):
    root._offsetx = event.x
    root._offsety = event.y






#==============================================================================
#Initialize
Database_Functions.Fetch_All_Data()








#==============================================================================
#GUI
#==============================================================================
warnings.simplefilter(action = "ignore", category = FutureWarning)
#Basic GUI Setting
root = tk.Tk()
root.wm_title('Data Manager 1.0')
root.geometry('1050x780')
root.iconbitmap('D:/Taishin_Platform/pics/Data_Manager.ico')
root.lift()
#root.attributes('-alpha', 0.0) #For icon
#root.overrideredirect(True)
root._offsetx = 0
root._offsety = 0
root.bind('<Button-1>',clickwin)
root.bind('<B1-Motion>',dragwin)

#root = tk.Toplevel(root)
#root.overrideredirect(1)
#==============================================================================
#Major Frame
#==============================================================================

Frame_Top=tk.Frame(root, width=2000, height=70, background="#282828")
Frame_Top.place(x=0, y=0)
#282828
Frame_Down=tk.Frame(root, width=1050, height=700, background="#F0F0F0")
Frame_Down.place(x=0, y=70)
#181818













#==============================================================================
#Button
#==============================================================================
style = ttk.Style()
style.configure("Menu.TButton", foreground="black", background="#181818")

#====================================================

Create_File_Button = ttk.Button(Frame_Top,text='Export Update File  ',command = Create_Update_File,style='Menu.TButton')
Create_File_Button.place(x=20, y=5)

Excel_icon = tk.PhotoImage(file = 'D:/Taishin_Platform/pics/Excel.gif')
Create_File_Button.config(image=Excel_icon,compound='left')
Excel_icon_Adj = Excel_icon.subsample(3,3)
Create_File_Button.config(image=Excel_icon_Adj)

#====================================================

Import_File_Button = ttk.Button(Frame_Top,text='Import Update File  ',command = Import_Data,style='Menu.TButton')
Import_File_Button.place(x=200, y=5)

Upload_icon = tk.PhotoImage(file = 'D:/Taishin_Platform/pics/Upload.gif')
Import_File_Button.config(image=Upload_icon,compound='left')
Upload_icon_Adj = Upload_icon.subsample(3,3)
Import_File_Button.config(image=Upload_icon_Adj)


#====================================================
Add_Index_Button = ttk.Button(Frame_Top,text='  Add New Index  ',command=Add_New_Index,style='Menu.TButton')
Add_Index_Button.place(x=380, y=5)

Add_icon = tk.PhotoImage(file = 'D:/Taishin_Platform/pics/Add1.gif')
Add_Index_Button.config(image=Add_icon,compound='left')
Add_icon_Adj = Add_icon.subsample(3,3)
Add_Index_Button.config(image=Add_icon_Adj)
#====================================================
Delete_Button = ttk.Button(Frame_Top,text='  Delete Index  ',command=Delete_Index,style='Menu.TButton')
Delete_Button.place(x=550, y=5)

Delete_icon = tk.PhotoImage(file = 'D:/Taishin_Platform/pics/delete.gif')
Delete_Button.config(image=Delete_icon,compound='left')
Delete_icon_Adj = Delete_icon.subsample(3,3)
Delete_Button.config(image=Delete_icon_Adj)


#==============================================================================
#Notebood
#==============================================================================
Notebook1 = ttk.Notebook(Frame_Down,height=620,width = 1000, style="TNotebook")
Notebook1.place(x=20,y=10)


Overview_frame = tk.Frame(Notebook1, background="white")
Data_Query_frame = tk.Frame(Notebook1, background="white")

Notebook1.add(Overview_frame, text="  Overview  ")
Notebook1.add(Data_Query_frame, text="  Data Query  ")



#==============================================================================
#Overview_frame in Notebook
#==============================================================================

Data_Table = ttk.Treeview(Overview_frame,height="28")

Data_Table["columns"]=("column1","column2",'column3','column4','column5')
Data_Table.column("#0",width=40, anchor='e')
Data_Table.column("column1", width=200, anchor='w' )
Data_Table.column("column2", width=380, anchor='w')
Data_Table.column("column3", width=120 , anchor='center')
Data_Table.column("column4", width=120 , anchor='center')
Data_Table.column("column5", width=120 , anchor='center')

Data_Table.heading('#0', text='')
Data_Table.heading("column1", text="Bloomberg Ticker",command= lambda : treeview_string_sort_column(Data_Table, "column1", False))
Data_Table.heading("column2", text="Name",command= lambda : treeview_string_sort_column(Data_Table, "column2", False))
Data_Table.heading("column3", text="Number",command= lambda : treeview_sort_column(Data_Table, "column3", False))
Data_Table.heading("column4", text="Start",command= lambda : treeview_string_sort_column(Data_Table, "column4", False))
Data_Table.heading("column5", text="Last",command= lambda : treeview_string_sort_column(Data_Table, "column5", False))

Data_Table.place(x=10, y=10)

#Import Taishin Database
Import_Update_table()

#==============================================================================
#Data_Query_frame in Notebook
#==============================================================================
Tickers_list = Database_Functions.Fetch_All_Tickers()
Combined_list = Combined_list(list1 = Tickers_list[0],list2 = Tickers_list[1])

Ticker_Entry = AutocompleteEntry(Combined_list, Data_Query_frame,width = 55,bg='#FF9C29',font=Entry_FONT)
Ticker_Entry.place(x=302, y=10)


Get_Button=tk.Button(Data_Query_frame,width=15, text =u"  Query  ",command=Get_data,font=NORM_FONT,relief='raised',fg='white',bg='black',activebackground='#FF9C29')
Get_Button.config(height=1)
Get_Button.place(x=820, y=5)


Query_Table = ttk.Treeview(Data_Query_frame,height="28")

Query_Table["columns"]=("column1","column2")
Query_Table.column("#0",width=40, anchor='e')
Query_Table.column("column1", width=100, anchor='center' )
Query_Table.column("column2", width=100, anchor='e')

Query_Table.heading('#0', text='')
Query_Table.heading("column1", text="Date",command= lambda : treeview_sort_column(Query_Table, "column1", False))
Query_Table.heading("column2", text="PX_LAST")

Query_Table.place(x=10, y=10)

#==============================================================================
#Chart in Notebook
#==============================================================================


fig = Figure(figsize=(6,4), dpi=120)
fig.set_tight_layout(True)
fig.patch.set_facecolor('#FFFFFF')

Stock_Chart = fig.add_subplot(111,axisbg='#FFFFFF')
Stock_Chart.tick_params(axis='both', which='major', labelsize=8)
#Stock_Chart.set_xlabel('Simulation Time', fontsize=10)
#Stock_Chart.set_ylabel('Stock price', fontsize=10)

Stock_Chart.tick_params(axis='both', which='major',colors='black', labelsize=6)
Stock_Chart.spines['bottom'].set_color('black')
Stock_Chart.spines['top'].set_color('black')
Stock_Chart.spines['left'].set_color('black')
Stock_Chart.spines['right'].set_color('black')
Stock_Chart.xaxis.label.set_color('black')
Stock_Chart.yaxis.label.set_color('black')


canvas1 = FigureCanvasTkAgg(fig, master=Data_Query_frame)
canvas1.get_tk_widget().place(x=260,y=40)
canvas1.get_tk_widget().configure(background='#FFFFFF',  highlightcolor='#FFFFFF', highlightbackground='#FFFFFF')

#==============================================================================
#Performance in Notebook
#==============================================================================

Annual_Return = tk.Label(Data_Query_frame,text='Annual Return',font=Entry_FONT,fg='black',background="white")
Annual_Return.place(x=480,y=520)

Annual_Return_Var = tk.StringVar()
Annual_Return_label = tk.Label(Data_Query_frame,textvariable=Annual_Return_Var,font=Figure_FONT,fg='black',background="white")
Annual_Return_label.place(x=598,y=550,anchor='ne')
#--------------------------------------

Standard_Deviation = tk.Label(Data_Query_frame,text='Standard Deviation',font=Entry_FONT,fg='black',background="white")
Standard_Deviation.place(x=650,y=520)

Standard_Deviation_Var = tk.StringVar()
Standard_Deviation_label = tk.Label(Data_Query_frame,textvariable=Standard_Deviation_Var,font=Figure_FONT,fg='black',background="white")
Standard_Deviation_label.place(x=800,y=550,anchor='ne')

#--------------------------------------


Sharpe = tk.Label(Data_Query_frame,text='Sharpe Ratio',font=Entry_FONT,fg='black',background="white")
Sharpe.place(x=850,y=520)

Sharpe_Var = tk.StringVar()
Sharpe_label = tk.Label(Data_Query_frame,textvariable=Sharpe_Var,font=Figure_FONT,fg='black',background="white")
Sharpe_label.place(x=953,y=550,anchor='ne')






#==============================================================================
root.config(background="#F0F0F0")
tk.mainloop()

#181818







