import pandas as pd
import tkinter as tk
from tkinter.ttk import Combobox

root = tk.Tk()

# Reads the excel file so you can work with it in python
#Will replace this with buttons to export your chosen file
df1 = pd.read_excel(r'C:\Users\Rislynn\Documents\Python Pandas Excels\VizData.xlsx')
df2 = pd.read_excel(r'C:\Users\Rislynn\Documents\Python Pandas Excels\CustomerData.xlsx')

# Returns the list of columns
df1columns = df1.columns
df2columns = df2.columns

#Creates a list of the columns to be used in the combo box
df1column_list = []
df2column_list = []

#turn into a function and then run on df1column_list and df2column_list
for i in range(len(df1columns)):
    df1column_list.append(df1columns[i])

#in Progress doesn't work yet
#def column_list_creator(dfxcolumns,dfxcol_list):
#    for i in range(dfxcolumns):
#       dfxcol_list.append(dfxcolumns[i])
#   return dfxcol_list

#Use the function above to create your column list
#column_list_creator(df1columns,df1column_list)
#column_list_creator(df2columns,df2column_list)


#print(df1column_list.index('SiteID'))


#Adds the text at the top
lbl_title = tk.Label(root, text="Enter the number of key columns")
lbl_title.pack()

#Adds the free entry field
txt_input = tk.Entry(root, width=15)
txt_input.pack()


#Following 3 lines create the combo box aka drop down list
combo = Combobox(root, values=df1column_list, width=15)
combo.set("select")
combo.pack()

df1columns_keep = []
#Add values to this list using the drop down box
# this list must be in the order you want in the resulting excel.
# First set of columns will be the predicate columns and after that must be ordered to match df2

def clear_listbox():
    lb_df1cols_keep.delete(0, "end")

def update_listbox():
    #populates the listbox
    for cols in df1columns_keep:
        lb_df1cols_keep.insert("end", cols)

def add_col():
    #Adds the selected value from the drop down to our list
    x = combo.get() #Gets the currently selected value from the combo (drop) box
    df1columns_keep.append(x) #adds the value of the drop box to the list of columns
    #print(df1columns_keep)
    clear_listbox()
    update_listbox()

def del_col():
    x = combo.get()
    df1columns_keep.remove(x)
    #rint(df1columns_keep)
    clear_listbox()
    update_listbox()

def move_down():
    x = combo.get()
    oldindex = df1columns_keep.index(x)
    newindex = oldindex + 1
    df1columns_keep.remove(x)
    df1columns_keep.insert(newindex,x)
    #print(df1columns_keep)
    clear_listbox()
    update_listbox()

def move_up():
    x = combo.get()
    oldindex = df1columns_keep.index(x)
    newindex = oldindex - 1
    df1columns_keep.remove(x)
    df1columns_keep.insert(newindex,x)
    #print(df1columns_keep)
    clear_listbox()
    update_listbox()

#Need a function to take the columns you're keeping and create a dataframe
#Then merge the resulting data frames
#And create variance columns
#Lastly save the excel


#Buttons
button = tk.Button(root, text="Add", command=add_col).pack()
button = tk.Button(root, text="Delete", command=del_col).pack()
button = tk.Button(root, text="Move Down", command=move_down).pack()
button = tk.Button(root, text="Move Up", command=move_up).pack()

#Changes the window size
root.geometry("300x300+120+120")

#Creates the listbox
lb_df1cols_keep = tk.Listbox(root)
lb_df1cols_keep.pack()


#below will be needed to create the merged dataframes
#a = [1, 4]
#result3 = df1.iloc[:, a]
#print(result3)


root.mainloop()

#Need to not let the user add the same column twice
#Columns to keep list size must be the same for both data frames

#Error handling for move up/down/delete if the items doesn't exist