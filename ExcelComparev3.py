import pandas as pd
import tkinter as tk
from tkinter.ttk import Combobox
from tkinter import filedialog

root = tk.Tk()

#Changes the window size
root.geometry("400x400+120+120")

# Reads the excel file so you can work with it in python
#Will replace this with buttons to export your chosen file

#Home
df1 = pd.read_excel(r'C:\Users\Rislynn\Documents\Python Pandas Excels\VizData.xlsx')
df2 = pd.read_excel(r'C:\Users\Rislynn\Documents\Python Pandas Excels\CustomerData.xlsx')

#work
#df1 = pd.read_excel(r'C:\Users\Tonya.Reeves\Documents\Python\ExcelComparev1\VizData.xlsx')
#df2 = pd.read_excel(r'C:\Users\Tonya.Reeves\Documents\Python\ExcelComparev1\CustomerData.xlsx')

#Below to for potential use
#df1 = pd.DataFrame()
#df2 = pd.DataFrame()

# Returns the list of columns
df1columns = df1.columns
df2columns = df2.columns

#Initializes a list of the columns to be used in the combo box
df1column_list = []
df2column_list = []

def column_list_creator(dfxcolumns,dfxcol_list):
#Creates a list out of the data frame columns
    #dfxcolumns are the columns of the data frame
    #dfxcol_list is the list you would like the columns inserted into
    for i in range(len(dfxcolumns)):
       dfxcol_list.append(dfxcolumns[i])
    return dfxcol_list

df1column_list = column_list_creator(df1columns,df1column_list)
df2column_list = column_list_creator(df2columns,df2column_list)

#Will be used for letting the user import the excel
#currently doesn't work as intended
def getexcel1():
    #global df
    import_file_path = filedialog.askopenfilename()
    df1 = pd.read_excel(import_file_path)


def getexcel2():
    #global df2
    import_file_path = filedialog.askopenfilename()
    df2 = pd.read_excel(import_file_path)
    return df2 #print(df2)


#End user will add values to this list using the drop down
df1columns_keep = []
df2columns_keep = []
# this list must be in the order you want in the resulting excel.
# First set of columns will be the predicate columns and after that must be ordered to match df2

def clear_listbox(lb_dfxcols_keep):
    #Deletes all entries of the list box
    lb_dfxcols_keep.delete(0, "end")

def update_listbox(dfxcolumns_keep,lb_dfxcols_keep):
    #populates the listbox
    for cols in dfxcolumns_keep:
        lb_dfxcols_keep.insert("end", cols)

def add_col(combodfx,dfxcolumns_keep,lb_dfxcols_keep):
    #Adds the selected value from the drop down to our list
    x = combodfx.get() #Gets the currently selected value from the combo (drop) box
    if x != 'select':
        dfxcolumns_keep.append(x) #adds the value of the drop box to the list of columns
        #Refreshes the list box to reflect the changes
        clear_listbox(lb_dfxcols_keep)
        update_listbox(dfxcolumns_keep, lb_dfxcols_keep)

def del_col(combodfx,dfxcolumns_keep,lb_dfxcols_keep):
    x = combodfx.get()
    if x != 'select':
        dfxcolumns_keep.remove(x)
        #Refreshes the list box to reflect the changes
        clear_listbox(lb_dfxcols_keep)
        update_listbox(dfxcolumns_keep, lb_dfxcols_keep)

def move_down(combodfx,dfxcolumns_keep,lb_dfxcols_keep):
    x = combodfx.get()
    if x != 'select':
        oldindex = dfxcolumns_keep.index(x)
        newindex = oldindex + 1
        dfxcolumns_keep.remove(x)
        dfxcolumns_keep.insert(newindex, x)
        #Refreshes the list box to reflect the changes
        clear_listbox(lb_dfxcols_keep)
        update_listbox(dfxcolumns_keep, lb_dfxcols_keep)

def move_up(combodfx,dfxcolumns_keep,lb_dfxcols_keep):
    x = combodfx.get()
    if x != 'select':
        oldindex = dfxcolumns_keep.index(x)
        newindex = oldindex - 1
        #Deletes the current entry
        dfxcolumns_keep.remove(x)
        #Adds the entry into the list at the desired location
        dfxcolumns_keep.insert(newindex, x)
        #Refreshes the list box to reflect the changes
        clear_listbox(lb_dfxcols_keep)
        update_listbox(dfxcolumns_keep, lb_dfxcols_keep)

#Need a function to take the columns you're keeping and create a dataframe
#Then merge the resulting data frames
#And create variance columns
#Lastly save the excel

#doesn't work yet
def update_combo_list():
    list = column_list_creator(df1columns,df1column_list)
    combodf1['values']= list

#Following 3 lines create the combo box aka drop down list
#Left
#combodf1 = Combobox(root, values=df1column_list, width=15)
combodf1 = Combobox(root, postcommand = update_combo_list, width=15)
combodf1.set("select")
combodf1.grid(row=2, column=0)



#Right
combodf2 = Combobox(root, values=df2column_list, width=15)
combodf2.set("select")
combodf2.grid(row=2, column=1)

#Creates the listboxes
#Left
lb_df1cols_keep = tk.Listbox(root)
lb_df1cols_keep.grid(row=8, column=0)

#Right
lb_df2cols_keep = tk.Listbox(root)
lb_df2cols_keep.grid(row=8, column=1)

# Functions for the buttons, they can't have arguments so need to use lambdas or wrappers
def add_col1():
    return add_col(combodf1, df1columns_keep, lb_df1cols_keep)
def del_col1():
    return del_col(combodf1, df1columns_keep, lb_df1cols_keep)
def move_down1():
    return move_down(combodf1, df1columns_keep, lb_df1cols_keep)
def move_up1():
    return move_up(combodf1, df1columns_keep, lb_df1cols_keep)

#Wrappers for df2
def add_col2():
    return add_col(combodf2, df2columns_keep, lb_df2cols_keep)
def del_col2():
    return del_col(combodf2, df2columns_keep, lb_df2cols_keep)
def move_down2():
    return move_down(combodf2, df2columns_keep, lb_df2cols_keep)
def move_up2():
    return move_up(combodf2, df2columns_keep, lb_df2cols_keep)

#Buttons
#Import Buttons
import_btn_1 = tk.Button(root, text="Import Excel 1", command=getexcel1).grid(row=0, column=0)
import_btn_2 = tk.Button(root, text="Import Excel 2", command=getexcel2).grid(row=0, column=1)

#Left Buttons
left_add = tk.Button(root, text="Add", command=add_col1).grid(row=4, column=0)
left_del = tk.Button(root, text="Delete", command=del_col1).grid(row=5, column=0)
left_up = tk.Button(root, text="Dwn", command=move_down1).grid(row=7, column=0)
left_dwn = tk.Button(root, text="Up", command=move_up1).grid(row=6, column=0)

#Right Buttons
right_add = tk.Button(root, text="Add", command=add_col2).grid(row=4, column=1)
right_del = tk.Button(root, text="Delete", command=del_col2).grid(row=5, column=1)
right_up = tk.Button(root, text="Dwn", command=move_down2).grid(row=7, column=1)
right_dwn = tk.Button(root, text="Up", command=move_up2).grid(row=6, column=1)

#Adds the text for Number Key Columns
lbl_title = tk.Label(root, text="Enter the number of key columns")
lbl_title.grid(row=9, column=0)

#Adds free entry field for number of key columns
txt_input = tk.Entry(root, width=15)
txt_input.grid(row=9, column=1)

#Button for Merging
merge_btn = tk.Button(root, text="Merge", command=move_up2).grid(row=10, column=1)

root.mainloop()

#Need to not let the user add the same column twice
#Columns to keep list size must be the same for both data frames

#Error handling for move up/down/delete if the items doesn't exist