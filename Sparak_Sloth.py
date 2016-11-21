import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter.filedialog import askopenfilename
from PIL import Image, ImageTk
import pyautogui
from colour import Color
import openpyxl
from openpyxl.cell import get_column_letter
import os
import datetime

LARGE_FONT = ("Verdana", 12)
NORM_FONT = ("Verdana", 10)
SMALL_FONT = ("Verdana", 8)
BACKGROUND_COLOR = Color("#FCEBAE")
SPARAK_TIME_FORMAT = "%m%d%Y"


filename = ""
filepath = "C:"
defaultBatchNumber = "48"



#TODO: Create menu to see user statistics like how many times the application has been used

#TODO: Add to setting menu the option to always add to the description of entries that the entry was made by the Sparak Accounting Sloth


#TODO: Define method for initializing or clearing the array of payments 



#TODO: Check the array for errors
        # Distribution Code = 0 < X <= 2 Digits
        # Account Number = 0 < X <= 10 digits
        # Tran Code = 0 < X <= 2 Digits
        # Amount = < 0
        # Date = X <= Today
        # Description = 0 <= X <= 160 Characters


def popupmsg(msg):
    popup = tk.Tk()
    popup.wm_title("Sloth Message")
    label = ttk.Label(popup, text=msg, font=NORM_FONT)
    label.pack(side="top", fill="x", pady=10)
    B1 = ttk.Button(popup, text="Okay", command = popup.destroy)
    B1.pack()
    popup.mainloop()


class Sparak_Sloth(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        tk.Tk.iconbitmap(self, default="slothicon.ico")
        tk.Tk.wm_title(self, "Slothful Accounting")

        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}

        for F in (StartPage, EntryPage, DailyReports, SettingsPage):
            frame = F(container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, sticky="nsew")
        self.show_frame(StartPage)

    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()


class StartPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)
# --------------------- Background Image ---------------------
        width, height = 1024, 640
        image = Image.open('slothBackground.png')
        if image.size != (width, height):
            image = image.resize((width, height), Image.ANTIALIAS)

        image = ImageTk.PhotoImage(image)
        bg_label = tk.Label(self, image = image)
        bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        bg_label.image = image
# --------------------- Background Image ---------------------

        label = tk.Label(self, text=("Sparak Accounting Sloth"), bg=BACKGROUND_COLOR, font=LARGE_FONT)
        label.pack(pady=10,padx=10) 

        button1 = ttk.Button(self, text="Enter Transactions",
                            command=lambda: controller.show_frame(EntryPage))
        button1.pack(pady=1, ipadx=6)
        button2 = ttk.Button(self, text="Collect Daily Reports",
                            command=lambda: controller.show_frame(DailyReports))
        button2.pack(pady=1)
        button3 = ttk.Button(self, text="Settings",
                            command=lambda: controller.show_frame(SettingsPage))
        button3.pack(pady=1, ipadx=22)
        button4 = ttk.Button(self, text="Quit Application",
                            command=quit)
        button4.pack(pady=1, ipadx=12)
    
   
class EntryPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)
# --------------------- Background Image ---------------------
        width, height = 1024, 640
        image = Image.open('slothBackground.png')
        if image.size != (width, height):
            image = image.resize((width, height), Image.ANTIALIAS)

        image = ImageTk.PhotoImage(image)
        bg_label = tk.Label(self, image = image)
        bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        bg_label.image = image
# --------------------- Background Image ---------------------


        label = tk.Label(self, text=("Accounting Transactions"), bg=BACKGROUND_COLOR, font=LARGE_FONT)
        label.grid(row=0, column=0, sticky=NSEW, columnspan=10)        
        
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)      
        self.grid_columnconfigure(2, weight=1)
        self.grid_rowconfigure(2, weight=1)  
        self.grid_columnconfigure(3, weight=1)
        self.grid_rowconfigure(3, weight=1)
        self.grid_columnconfigure(4, weight=1)
        self.grid_rowconfigure(4, weight=1)

        button1 = ttk.Button(self, text="Load Transactions",
                             command=lambda: selectFile(self))
        button1.grid(row=1, column=1, columnspan=1,  pady=10)
        button2 = ttk.Button(self, text="Enter Transactions")
        button2.grid(row=1, column=2, columnspan=1, pady=10) 
        button3 = ttk.Button(self, text="Clear Transactions",
                             command=lambda: clear_payment_box(self))
        button3.grid(row=1, column=3, columnspan=1, pady=10)
        button4 = ttk.Button(self, text="Return to Main Menu",
                            command=lambda: controller.show_frame(StartPage))
        button4.grid(row=1, column=4, columnspan=1, pady=10)

        

        tree = ttk.Treeview(self)
        tree['columns'] = ('distribution','account_number','tran_code','amount','effective_date','description',)
        tree.heading("#0", text='Tran ID', anchor=W)
        tree.column("#0",stretch=NO, width=10, anchor=W)
        tree.column('distribution', width=30, anchor=W)
        tree.heading('distribution', text='Distribution')        
        tree.column('account_number', width=30, anchor=W)
        tree.heading('account_number', text='Account')        
        tree.column('tran_code', width=30, anchor=E)
        tree.heading('tran_code', text='Tran Code')       
        tree.column('amount', width=30, anchor=E)
        tree.heading('amount', text='Amount')       
        tree.column('effective_date', width=30, anchor=W)
        tree.heading('effective_date', text='Eff Dt')        
        tree.column('description', width=100, anchor=W)
        tree.heading('description', text='Description')
        ttk.Style().configure("Treeview", font=NORM_FONT, background='grey',
                              foreground='white',fieldbackground=BACKGROUND_COLOR)
        
        tree.grid(row=2, column=0, columnspan=6, ipadx=150, ipady=120, pady=20)

        def selectFile(self):
            # Add adjustment to show all forms of excel files - xlsm, xls, etc
            filename = askopenfilename(filetypes = (("Excel Files","*.xlsx"),("CSV Files","*.csv")),
                               title = "Choose a file:")
            filename = str(filename)

            if filename:
                mylist = filename.split("/")
                mylistSize = len(mylist)
                filename = mylist[mylistSize-1]

                i = 0
                filepath = ""
                while (i < mylistSize - 1):
                    filepath = filepath + mylist[i] + "/"
                    i = i + 1
                openFile(filepath, filename)
            else:
                popupmsg("Error opening the file you requested or you did not select a file.")

        def openFile(filepath, filename):
            paymentArray = []

            os.chdir(filepath)
            wb = openpyxl.load_workbook(filename)
    
            # Add functionality to select which sheet with a dialogue  box or something
            sheet = wb.get_sheet_by_name('Payment Sheet')
            end_of_sheet_range = get_column_letter(sheet.max_column) + str(sheet.max_row)
    
            
            for rowOfCellObjects in sheet['A1':end_of_sheet_range]:
                for cellObj in rowOfCellObjects:
                    if type(cellObj.value) is datetime.datetime:
                        paymentArray.append(str(cellObj.value.strftime("%m/%d/%Y")))
                    elif type(cellObj.value) is type(None):
                        paymentArray.append("No Description")
                    else:
                        paymentArray.append(cellObj.value)


            entryFill(self, paymentArray)

        def entryFill(self, paymentArray):

            if len(paymentArray) % 6 == 0:
                for i in range(1, (len(paymentArray)+1)):
                    if i%6==0:
                        tree.insert("","end",values = (str(paymentArray[i - 6]), str(paymentArray[i - 5]), str(paymentArray[i - 4]),\
                                                       str(paymentArray[i - 3]), str(paymentArray[i - 2]), str(paymentArray[i - 1])))
            else:
                popupmsg("""There was an issue with loading your 
                            transactions. Please make sure you're
                            input file has the correct amount of 
                            entries and each part of the entry is
                            input correctly.""")

        def clear_payment_box(self):
            payment_box.delete(0, END)


class DailyReports(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)
# --------------------- Background Image ---------------------
        width, height = 1024, 640
        image = Image.open('slothBackground.png')
        if image.size != (width, height):
            image = image.resize((width, height), Image.ANTIALIAS)

        image = ImageTk.PhotoImage(image)
        bg_label = tk.Label(self, image = image)
        bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        bg_label.image = image
# --------------------- Background Image ---------------------
        
        label = tk.Label(self, text=("Daily Report Collection"), bg=BACKGROUND_COLOR, font=LARGE_FONT)
        label.pack(pady=10,padx=10)
        button1 = ttk.Button(self, text="Collect Daily Reports From Sparak",
                            command=lambda: controller.show_frame(StartPage))
        button1.pack()
        button2 = ttk.Button(self, text="Return to Main Menu",
                            command=lambda: controller.show_frame(StartPage))
        button2.pack(pady=1, ipadx=32)                


class SettingsPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self,parent)
# --------------------- Background Image ---------------------
        width, height = 1024, 640
        image = Image.open('slothBackground.png')
        if image.size != (width, height):
            image = image.resize((width, height), Image.ANTIALIAS)

        image = ImageTk.PhotoImage(image)
        bg_label = tk.Label(self, image = image)
        bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        bg_label.image = image
# --------------------- Background Image ---------------------

        label = tk.Label(self, text=("Settings"), bg=BACKGROUND_COLOR, font=LARGE_FONT)
        label.pack(pady=10,padx=10)
        button1 = ttk.Button(self, text="Return to Main Menu",
                            command=lambda: controller.show_frame(StartPage))
        button1.pack()                

   

app = Sparak_Sloth()
app.geometry("800x640")
app.resizable(0,0)
#app.after(0,openFile)
app.mainloop()