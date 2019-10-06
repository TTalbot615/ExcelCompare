import pandas as pd
import tkinter as tk
from tkinter.ttk import Combobox
from tkinter import filedialog

#Changes from v4
    #Add in an indicator that shows what file is currently exported
    #Need a way to differentiate in the merged excel file which are from customer and which are from viz
        #Maybe way to differentiate x and y ( x = viz for example) or add in additional rows at the top?
    #Add in an exit button
    #Make the up and down buttons work with list box instead of the combo box
    #Have the combo box entries decrease as they are added to the list box
    #Maybe sort the entries alphabetically for the combobox
    #add in some error handling if possible
    #Don't let the merge run if the lists are not the same size

root = tk.Tk()

# Changes the window size
root.geometry("405x440")    # height x width
root.configure(bg='#D4E6F1')

# End user adds/manipulates values to these lists using the drop down and the add/del/up/dwn buttons
# End user needs to order both of these lists so that matching columns have the same index
df1columns_keep = []
df2columns_keep = []

# Global Variables
# df1 = data frame 1 imported by user
# df2 = data frame 2 imported by the user
# df1columns_list list with the columns from data frame 1
# df2columns_list list with columns from data frame 2

file_imported = 0


def column_list_creator(dfxcolumns, dfxcol_list):
    # Creates a list out of the data frame columns
    # dfxcolumns are the columns of the data frame
    # dfxcol_list is the list you would like the columns inserted into
    for i in range(len(dfxcolumns)):
        dfxcol_list.append(dfxcolumns[i])
    return dfxcol_list


def getexcel1():
    # User will use import button to choose the first excel file and make it accessible to other functions
    global df1
    global df1column_list
    global file_imported
    global import_file_path1
    import_file_path1 = filedialog.askopenfilename()
    df1 = pd.read_excel(import_file_path1)
    df1columns = df1.columns
    df1column_list = []
    df1column_list = column_list_creator(df1columns, df1column_list)
    file_imported = 1


def getexcel2():
    # User will use import button to choose the second excel file and make it accessible to other functions
    global df2
    global df2column_list
    global file_imported
    global import_file_path2
    import_file_path2 = filedialog.askopenfilename()
    df2 = pd.read_excel(import_file_path2)
    df2columns = df2.columns
    df2column_list = []
    df2column_list = column_list_creator(df2columns, df2column_list)
    file_imported = 1


def clear_listbox(lb_dfxcols_keep):
    # Deletes all entries of the list box
    lb_dfxcols_keep.delete(0, "end")


def update_listbox(dfxcolumns_keep, lb_dfxcols_keep):
    # populates the listbox
    for cols in dfxcolumns_keep:
        lb_dfxcols_keep.insert("end", cols)


def add_col(combodfx, dfxcolumns_keep, lb_dfxcols_keep):
    # Adds the selected value from the drop down to our list
        # combodfx is the combo box for data frame x
        # dfcolumns_keep are the columns names added by the user they wish to use
        # lb_dfxcols_keep is the list box for data frame x
    x = combodfx.get()  # Gets the currently selected value from the combo (drop) box
    # print(df1column_list)
    if x != 'select' and x not in dfxcolumns_keep:  # user has selected a value AND value not already added
        dfxcolumns_keep.append(x)   # adds the value of the drop box to the list of columns
        # Refreshes the list box to reflect the changes
        clear_listbox(lb_dfxcols_keep)
        update_listbox(dfxcolumns_keep, lb_dfxcols_keep)


def del_col(combodfx, dfxcolumns_keep, lb_dfxcols_keep):
    x = combodfx.get()  # Gets the current value the user has selected
    if x != 'select':
        dfxcolumns_keep.remove(x)
        # Refreshes the list box to reflect the changes
        clear_listbox(lb_dfxcols_keep)
        update_listbox(dfxcolumns_keep, lb_dfxcols_keep)


def move_down(combodfx, dfxcolumns_keep, lb_dfxcols_keep):
    x = combodfx.get()
    if x != 'select':
        oldindex = dfxcolumns_keep.index(x)     # original index of the currently selected value
        newindex = oldindex + 1     # position we would like to move the element too
        dfxcolumns_keep.remove(x)   # deletes the original
        dfxcolumns_keep.insert(newindex, x)     # re-adds the element in the desired position
        # Refreshes the list box to reflect the changes
        clear_listbox(lb_dfxcols_keep)
        update_listbox(dfxcolumns_keep, lb_dfxcols_keep)


def move_up(combodfx, dfxcolumns_keep, lb_dfxcols_keep):
    x = combodfx.get()
    if x != 'select':
        oldindex = dfxcolumns_keep.index(x)
        newindex = oldindex - 1
        # Deletes the current entry
        dfxcolumns_keep.remove(x)
        # Adds the entry into the list at the desired location
        dfxcolumns_keep.insert(newindex, x)
        # Refreshes the list box to reflect the changes
        clear_listbox(lb_dfxcols_keep)
        update_listbox(dfxcolumns_keep, lb_dfxcols_keep)


# Functions for the buttons, they can't have arguments so need to use lambdas or wrappers
def add_col1():
    return add_col(combodf1, df1columns_keep, lb_df1cols_keep)


def del_col1():
    return del_col(combodf1, df1columns_keep, lb_df1cols_keep)


def move_down1():
    return move_down(combodf1, df1columns_keep, lb_df1cols_keep)


def move_up1():
    return move_up(combodf1, df1columns_keep, lb_df1cols_keep)


# lambdas for df2
def add_col2():
    return add_col(combodf2, df2columns_keep, rb_df2cols_keep)


def del_col2():
    return del_col(combodf2, df2columns_keep, rb_df2cols_keep)


def move_down2():
    return move_down(combodf2, df2columns_keep, rb_df2cols_keep)


def move_up2():
    return move_up(combodf2, df2columns_keep, rb_df2cols_keep)


def update_combo_list(combodfx, dfxcolumn_list):
    combodfx['values'] = dfxcolumn_list


def update_combo_list1():
    return update_combo_list(combodf1, df1column_list)


def update_combo_list2():
    return update_combo_list(combodf2, df2column_list)


################################### Following functions are for the merge button
def dfxcolumns_keep_orig_index(dfxcolumns_keep, dfxcolumns_list):
    # Creates a list that displays the original index of the columns specified by the user
    dfxcolumns_keep_orig_index = []
    for i in range(len(dfxcolumns_keep)):
        dfkeepi = dfxcolumns_keep[i]  # ith entry of dfxcolumns_keep
        origindex = dfxcolumns_list.index(dfkeepi)
        dfxcolumns_keep_orig_index.append(origindex)
    return dfxcolumns_keep_orig_index


def cleanup_df(d, a):
    # Deletes and reorders the columns of a data frame
    # d is a dataframe
    # a is an array with the index of the column in the order you would like now
    #[0,1,5,6,7,8,2,3,4] column 5 would move to the third position for example
    return d.iloc[:, a]


def merge_df(d, f,k):  #dataframe1, dataframe2, keys = # columns at the beginning to use for join
    key_col_array = []
    for i in range(k):
        key_col_array.append(d.columns[i])  # Adds the ith column name to the array (d1 and d2 need same first key columns)
    result = pd.merge(d, f, how='outer', on=key_col_array)  # Joins df and df2 on the first num_key_columns
    return result


def var_col(d, k):  #Takes merged_df and makes variances columns, k = number of key columns
    d0 = d.fillna(0)    # Replaces all the nulls with 0
    c = (len(d.columns) - k)//2     #First k columns of a merged dfs are the key columns
    for i in range(k, 2+c):
        first_col_name = d0.columns[i]
        second_col_name = d0.columns[i+c]
        var_col_name = "Var " + first_col_name[0:len(d0.columns[i])-2]
        if type(d0.iloc[1, i]) == str:  #Checks to see if the first entry of the column is a string, probalby needs to be more robust, like check more than just the first entry
            d0[var_col_name] = d0[first_col_name] == d0[second_col_name]
        else:
            d0[var_col_name] = d0[first_col_name] - d0[second_col_name]
    return d0


def merged_var_df():
    df1colkeep_index = dfxcolumns_keep_orig_index(df1columns_keep,df1column_list)
    df2colkeep_index = dfxcolumns_keep_orig_index(df2columns_keep,df2column_list)
    df1_clean = cleanup_df(df1,df1colkeep_index)
    df2_clean = cleanup_df(df2,df2colkeep_index)
    k = key_input.get()
    k = int(k)
    merged_df = merge_df(df1_clean, df2_clean, k)
    var_df = var_col(merged_df, k)
    return var_df


def save_dataframe():
    save_path = filedialog.asksaveasfile(mode='w', defaultextension=".csv")
    var_df_final = merged_var_df()
    print(var_df_final)
    var_df_final.to_csv(save_path, index=False, line_terminator='\n')   # line terminator avoid empty spaces after rows


root.grid_columnconfigure(0, minsize=35)
root.grid_columnconfigure(1, minsize=15)
root.grid_columnconfigure(2, minsize=15)
root.grid_columnconfigure(3, minsize=15)


# Following 3 lines create the combo box aka drop down list
# Left
combodf1 = Combobox(root, postcommand=update_combo_list1, width=20)
combodf1.set("select")
combodf1.grid(row=2, column=0, sticky='W'+'N'+'S'+'E', padx=(20,0), pady=5, columnspan=2)
# Right
combodf2 = Combobox(root, postcommand=update_combo_list2, width=20)
combodf2.set("select")
combodf2.grid(row=2, column=3, sticky='W'+'N'+'S'+'E', pady=5, columnspan=2)

# Creates the list boxes
# Left
lb_df1cols_keep = tk.Listbox(root)
lb_df1cols_keep.grid(row=8, column=0, sticky='W'+'N'+'S'+'E', padx=(20,0), pady=5, columnspan=2)
# Right
rb_df2cols_keep = tk.Listbox(root)
rb_df2cols_keep.grid(row=8, column=3, sticky='W'+'N'+'S'+'E', pady=5, columnspan=2)

# Buttons
# Import Buttons
import_btn_1 = tk.Button(root, text="Import Excel 1", command=getexcel1, width=20).grid(row=0, column=0, padx=5, pady=5,columnspan =2)
import_btn_2 = tk.Button(root, text="Import Excel 2", command=getexcel2, width=20).grid(row=0, column=3, padx=5, pady=5,columnspan =2)

# Imported File Paths
lbl_filepath1 = tk.Label(root, text="In progress1", bg='#D4E6F1').grid(row=1, column=0, columnspan=2, padx=5, pady=5)
lbl_filepath2= tk.Label(root, text="In progress2",bg='#D4E6F1').grid(row=1, column=3, columnspan=2, padx=5, pady=5)

# Left Buttons
left_add = tk.Button(root, text="Add", command=add_col1, width=10).grid(row=4, column=1) #, sticky='W'+'N'+'S'+'E',padx=5
left_del = tk.Button(root, text="Del", command=del_col1, width=10).grid(row=4, column=0, padx=(20, 0)) # , sticky='W'+'N'+'S'+'E',padx=20
left_up = tk.Button(root, text="Dwn", command=move_down1, width=10).grid(row=9, column=1)
left_dwn = tk.Button(root, text="Up", command=move_up1, width=10).grid(row=9, column=0, padx=(20, 0))

# Right Buttons
right_add = tk.Button(root, text="Add", command=add_col2, width=10).grid(row=4, column=4)
right_del = tk.Button(root, text="Del", command=del_col2, width=10).grid(row=4, column=3)
right_up = tk.Button(root, text="Dwn", command=move_down2, width=10).grid(row=9, column=4)
right_dwn = tk.Button(root, text="Up", command=move_up2, width=10).grid(row=9, column=3)

# Adds the text for Number Key Columns
lbl_title = tk.Label(root, text="# Key Cols" , bg='#D4E6F1')
lbl_title.grid(row=10, column=1, sticky='W', pady=20, padx=(20, 0))

# Adds free entry field for number of key columns
key_input = tk.Entry(root,width=10)
key_input.grid(row=10, column=2,columnspan=2,sticky='W')

# Button for Merging
merge_btn = tk.Button(root, text="Merge", command=save_dataframe, width=20).grid(row=11, column=1,
                    sticky='W'+'N'+'E'+'S', columnspan=3, padx=(5, 0))

root.mainloop()

# Need to not let the user add the same column twice
# Columns to keep list size must be the same for both data frames

# Error handling for move up/down/delete if the items doesn't exist