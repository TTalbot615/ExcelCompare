import pandas as pd
import tkinter as tk
from tkinter.ttk import Combobox
from tkinter import filedialog
import tkinter.messagebox

# Desired Changes from v6
    # Need a way to differentiate in the merged excel file which are from customer and which are from viz
        # Maybe way to differentiate x and y ( x = viz for example) or add in additional rows at the top?
    # add in some error handling if possible
    # Move up the column totals
    # Give a warning if the 'key columns' have nothing in common
    # For the output would like a customer data tab, viz data tab, comparison data tab, summary information

# Completed Changes
    # Fixed issue with variance column
    # The delete button (not delete all) does not work because of the, it's looking for an item in the combo box that no longer exists (since the list updates as you add items)
    # Change the delete button to work with the listbox instead and move the buttons so the add buttons are above the list and the delete buttons are above the listbox
    # Don't let the merge run if the lists are not the same size
    # add in another pop up box if they need to set the number of key columns
    # Add in an indicator that shows what file is currently exported
    # add in the number of columns currently selected
    # Write a function that finds last instance of / in the file path and then returns the characters after that
    # import string variable to show only characters after that index use str.rfind(sub,start,end)
    # Create a general string for the column numbers (in case you decide to update it later) and then update all the functions
root = tk.Tk()
root.title('Excel Compare')

# Changes the window size
root.geometry("380x560")    # width x height
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

# listbox list with the columns to display in the list box, updates once elements are add/del
#left
import_excel1_lbl_text = tk.StringVar()
import_excel1_lbl_text.set('Please import a file')
#right
import_excel2_lbl_text = tk.StringVar()
import_excel2_lbl_text.set('Please import a file')

# Variable for the number of columns selected by user
#Left
selected_column_count1 = tk.StringVar()
general_column_string = 'Total selected columns: '
selected_column_count1.set(general_column_string+str(len(df1columns_keep)))
#Right
selected_column_count2 = tk.StringVar()
selected_column_count2.set(general_column_string+str(len(df2columns_keep)))


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
    global import_file_path1
    global listbox1
    import_file_path1 = filedialog.askopenfilename()
    df1 = pd.read_excel(import_file_path1)
    df1columns = df1.columns
    df1column_list = []
    df1column_list = column_list_creator(df1columns, df1column_list)
    # Initialize listbox1 as a copy of df1column_list
    # This way when we add delete from list box it does not alter our original list
    listbox1 = []
    for x in df1column_list:
        listbox1.append(x)
        listbox1.sort()
    # Below will get the file name of the file path and update the variable for the label
    len_file_path = len(import_file_path1)
    last_slash_index = import_file_path1.rfind('/')     # Finds the index of the last slash in file path
    import_file_name = import_file_path1[last_slash_index+1: len_file_path]
    import_excel1_lbl_text.set(import_file_name)


def getexcel2():
    # User will use import button to choose the second excel file and make it accessible to other functions
    global df2
    global df2column_list
    global import_file_path2
    global listbox2
    import_file_path2 = filedialog.askopenfilename()
    df2 = pd.read_excel(import_file_path2)
    df2columns = df2.columns
    df2column_list = []
    df2column_list = column_list_creator(df2columns, df2column_list)
    listbox2 = []
    for x in df2column_list:
        listbox2.append(x)
        listbox2.sort()
    #Updates text for label
    # Below will get the file name of the file path and update the variable for the label
    len_file_path = len(import_file_path2)
    last_slash_index = import_file_path2.rfind('/')  # Finds the index of the last slash in file path
    import_file_name2 = import_file_path2[last_slash_index + 1: len_file_path]
    import_excel2_lbl_text.set(import_file_name2)


def clear_listbox(lb_dfxcols_keep):
    # Deletes all entries of the list box
    lb_dfxcols_keep.delete(0, "end")


def update_listbox(dfxcolumns_keep, lb_dfxcols_keep):
    # populates the listbox
    for cols in dfxcolumns_keep:
        lb_dfxcols_keep.insert("end", cols)


def add_col(combodfx, dfxcolumns_keep, lb_dfxcols_keep, listbox_x, selected_column_countx):
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
        listbox_x.remove(x)
        #updates the column count
        selected_column_countx.set(general_column_string + str(len(dfxcolumns_keep)))


def add_all(dfxcolumns_keep, lb_dfxcols_keep, listbox_x, selected_column_countx):
    for x in listbox_x:
        dfxcolumns_keep.append(x)
    del listbox_x[:]
    clear_listbox(lb_dfxcols_keep)
    update_listbox(dfxcolumns_keep, lb_dfxcols_keep)
    # updates the column count
    selected_column_countx.set(general_column_string + str(len(dfxcolumns_keep)))


def del_col(dfxcolumns_keep, lb_dfxcols_keep, listbox_x, selected_column_countx):
    x = lb_dfxcols_keep.get(lb_dfxcols_keep.curselection())  # Gets the current user selection
    cur_index = dfxcolumns_keep.index(x)
    dfxcolumns_keep.remove(x)
    # Refreshes the list box to reflect the changes
    clear_listbox(lb_dfxcols_keep)
    update_listbox(dfxcolumns_keep, lb_dfxcols_keep)
    listbox_x.append(x)
    lb_dfxcols_keep.selection_set(cur_index)
    # updates the column count
    selected_column_countx.set(general_column_string + str(len(dfxcolumns_keep)))


def del_all(dfxcolumns_keep, lb_dfxcols_keep, listbox_x, selected_column_countx):
    for x in dfxcolumns_keep:
        listbox_x.append(x)
    del dfxcolumns_keep[:]
    clear_listbox(lb_dfxcols_keep)
    update_listbox(dfxcolumns_keep, lb_dfxcols_keep)
    # updates the column count
    selected_column_countx.set(general_column_string + str(len(dfxcolumns_keep)))



def move_down(dfxcolumns_keep, lb_dfxcols_keep):
        x = lb_dfxcols_keep.get(lb_dfxcols_keep.curselection())     # Gets the current user selection
        oldindex = dfxcolumns_keep.index(x)     # original index of the currently selected value
        newindex = oldindex + 1     # position we would like to move the element too
        dfxcolumns_keep.remove(x)   # deletes the original
        dfxcolumns_keep.insert(newindex, x)     # re-adds the element in the desired position
        # Refreshes the list box to reflect the changes
        clear_listbox(lb_dfxcols_keep)
        update_listbox(dfxcolumns_keep, lb_dfxcols_keep)
        lb_dfxcols_keep.selection_set(newindex)     # Sets the selection to the new position in case user wants to move down again


def move_up(dfxcolumns_keep, lb_dfxcols_keep):
    x = lb_dfxcols_keep.get(lb_dfxcols_keep.curselection())
    oldindex = dfxcolumns_keep.index(x)
    newindex = oldindex - 1
    # Deletes the current entry
    dfxcolumns_keep.remove(x)
    # Adds the entry into the list at the desired location
    dfxcolumns_keep.insert(newindex, x)
    # Refreshes the list box to reflect the changes
    clear_listbox(lb_dfxcols_keep)
    update_listbox(dfxcolumns_keep, lb_dfxcols_keep)
    lb_dfxcols_keep.selection_set(newindex)


# Functions for the buttons, they can't have arguments so need to use lambdas or wrappers
def add_col1():
    return add_col(combodf1, df1columns_keep, lb_df1cols_keep, listbox1, selected_column_count1)


def add_all1():
    return add_all(df1columns_keep, lb_df1cols_keep, listbox1, selected_column_count1)


def del_col1():
    return del_col(df1columns_keep, lb_df1cols_keep, listbox1, selected_column_count1)


def del_all1():
    return del_all(df1columns_keep, lb_df1cols_keep, listbox1, selected_column_count1)


def move_down1():
    return move_down(df1columns_keep, lb_df1cols_keep)


def move_up1():
    return move_up(df1columns_keep, lb_df1cols_keep)


# lambdas for df2
def add_col2():
    return add_col(combodf2, df2columns_keep, rb_df2cols_keep, listbox2, selected_column_count2)


def add_all2():
    return add_all(df2columns_keep, rb_df2cols_keep, listbox2, selected_column_count2)


def del_all2():
    return del_all(df2columns_keep, rb_df2cols_keep, listbox2, selected_column_count2)


def del_col2():
    return del_col(df2columns_keep, rb_df2cols_keep, listbox2, selected_column_count2)


def move_down2():
    return move_down(df2columns_keep, rb_df2cols_keep)


def move_up2():
    return move_up(df2columns_keep, rb_df2cols_keep)


#################################################Combo Box Functions
def update_combo_list(combodfx, listbox_x):
    combodfx['values'] = listbox_x


def update_combo_list1():
    return update_combo_list(combodf1, listbox1)


def update_combo_list2():
    return update_combo_list(combodf2, listbox2)


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
    result = pd.merge(d, f, how='outer', on=key_col_array)  # Joins df and df2 on the first num_key_columns]
    return result


# Appears something going wrong in var, seems to be on the first column?
# has to do with the nubmer of keys (1+c) is needed if k =1 but I think (2+c) is needed if k = 2
# As it is right now the values are correct but if k = 2 it doesn't do all of the columns
# Could also maybe have something to do with the number of columns I use
def var_col(d, k):  #Takes merged_df and makes variances columns, k = number of key columns
    d0 = d.fillna(0)    # Replaces all the nulls with 0
    c = (len(d.columns) - k)//2     #First k columns of a merged dfs are the key columns
    for i in range(k, k+c):
        first_col_name = d0.columns[i]
        second_col_name = d0.columns[i+c]
        var_col_name = "Var " + first_col_name[0:len(d0.columns[i])-2]
        if type(d0.iloc[1, i]) == str:  #Checks to see if the first entry of the column is a string,
            # probably needs to be more robust, like check more than just the first entry
            d0[var_col_name] = d0[first_col_name] == d0[second_col_name]
        else:
            d0[var_col_name] = d0[first_col_name] - d0[second_col_name]
    return d0


def merged_var_df():
    df1colkeep_index = dfxcolumns_keep_orig_index(df1columns_keep, df1column_list)
    df2colkeep_index = dfxcolumns_keep_orig_index(df2columns_keep, df2column_list)
    df1_clean = cleanup_df(df1, df1colkeep_index)
    df2_clean = cleanup_df(df2, df2colkeep_index)
    k = key_input.get()
    k = int(k)
    merged_df = merge_df(df1_clean, df2_clean, k)
    var_df = var_col(merged_df, k)
    return var_df


def save_dataframe():
    if not key_input.get():
        tk.messagebox.showinfo('Warning', 'Please input the number of key columns')
    if len(df1columns_keep) != len(df2columns_keep):
        tk.messagebox.showinfo('Warning', 'The number of columns needs to be the same')
    else:
        save_path = filedialog.asksaveasfile(mode='w', defaultextension=".csv")
        var_df_final = merged_var_df()
        var_df_final.to_csv(save_path, index=False, line_terminator='\n')   # line terminator avoid empty spaces after rows



# Potentially going to make a button to color the key columns
def listcolor():
    # Try to color the list boxes for the keys
    if not key_input.get():     # if no input there is at least one key
        x = 0
    else:
        x = int(key_input.get()) - 1
    return x


# Below changes the listbox row color based on the number of keys
def color_key_cols():
    y = listcolor()
    for i in range(y):
        lb_df1cols_keep.itemconfig(i, {'bg': '#2980b9'})


def close_window():
    root.destroy()


# Sets the min column size
root.grid_columnconfigure(0, minsize=35)
root.grid_columnconfigure(1, minsize=15)
root.grid_columnconfigure(2, minsize=15)
root.grid_columnconfigure(3, minsize=15)


# Following 3 lines create the combo box aka drop down list
# Left
combodf1 = Combobox(root, postcommand=update_combo_list1, width=20)
combodf1.set("select")
combodf1.grid(row=2, column=0, sticky='W'+'N'+'S'+'E', padx=(20, 0), pady=5, columnspan=2)
# Right
combodf2 = Combobox(root, postcommand=update_combo_list2, width=20)
combodf2.set("select")
combodf2.grid(row=2, column=3, sticky='W'+'N'+'S'+'E', pady=5, columnspan=2)

# Creates the list boxes
# Left
lb_df1cols_keep = tk.Listbox(root)
lb_df1cols_keep.grid(row=4, column=0, sticky='W'+'N'+'S'+'E', padx=(20, 0), pady=5, columnspan=2)
# Right
rb_df2cols_keep = tk.Listbox(root)
rb_df2cols_keep.grid(row=4, column=3, sticky='W'+'N'+'S'+'E', pady=5, columnspan=2)

# Buttons
# Import Buttons
import_btn_1 = tk.Button(root, text="Import Excel 1", command=getexcel1, width=20).grid(row=0, column=0, padx=5, pady=5
                                                                                        , columnspan=2)
import_btn_2 = tk.Button(root, text="Import Excel 2", command=getexcel2, width=20).grid(row=0, column=3, padx=5, pady=5
                                                                                        , columnspan=2)


# Imported File Paths
lbl_filepath1 = tk.Label(root, textvariable=import_excel1_lbl_text, bg='#D4E6F1').grid(row=1, column=0, columnspan=2, padx=5, pady=5)
lbl_filepath2 = tk.Label(root, textvariable=import_excel2_lbl_text, bg='#D4E6F1').grid(row=1, column=3, columnspan=2, padx=5, pady=5)

# Left Buttons
left_add = tk.Button(root, text="Add", command=add_col1, width=10, bg='#A9DFBF').grid(row=3, column=1)
left_del = tk.Button(root, text="Del", command=del_col1, width=10, bg='#F5B7B1').grid(row=6, column=1)
left_dwn = tk.Button(root, text="Dwn", command=move_down1, width=10).grid(row=8, column=0, columnspan=2, padx=(20, 0))
left_up = tk.Button(root, text="Up", command=move_up1, width=10).grid(row=7, column=0, columnspan=2, padx=(20, 0))
left_add_all = tk.Button(root, text="Add All", command=add_all1, width=10, bg='#A9DFBF').grid(row=3, column=0, padx=(20, 0))
left_del_all = tk.Button(root, text="Del All", command=del_all1, width=10, bg='#F5B7B1').grid(row=6, column=0, padx=(20, 0))

# Right Buttons
right_add = tk.Button(root, text="Add", command=add_col2, width=10, bg='#A9DFBF').grid(row=3, column=4)
right_del = tk.Button(root, text="Del", command=del_col2, width=10, bg='#F5B7B1').grid(row=6, column=4)
right_dwn = tk.Button(root, text="Dwn", command=move_down2, width=10).grid(row=8, column=3, columnspan=2)
right_up = tk.Button(root, text="Up", command=move_up2, width=10).grid(row=7, column=3, columnspan=2)
right_add_all = tk.Button(root, text="Add All", command=add_all2, width=10, bg='#A9DFBF').grid(row=3, column=3)
right_del_all = tk.Button(root, text="Del All", command=del_all2, width=10, bg='#F5B7B1').grid(row=6, column=3)

# Labels for the number of columns selected
lbl_column_selected1 = tk.Label(root, textvariable=selected_column_count1, bg='#D4E6F1').grid(row=5, column=0, columnspan=2)
lbl_column_selected2 = tk.Label(root, textvariable=selected_column_count2, bg='#D4E6F1').grid(row=5, column=3, columnspan=2)


# Adds the text for Number Key Columns
lbl_key_col = tk.Label(root, text="# Key Cols", bg='#D4E6F1').grid(row=9, column=1, sticky='W', pady=20, padx=(20, 0))

# Adds free entry field for number of key columns
key_input = tk.Entry(root, width=10)
key_input.grid(row=9, column=2, columnspan=2, sticky='W')


# Button for Merging
merge_btn = tk.Button(root, text="Merge", command=save_dataframe, width=20).grid(row=10, column=1,
                    sticky='W'+'N'+'E'+'S', columnspan=3, padx=(5, 0))

# Button to quit
quit_btn = tk.Button(root, text="Quit", command=close_window, width=20).grid(row=11, column=1,
                    sticky='W'+'N'+'E'+'S', columnspan=3, padx=(5, 0), pady=20)

root.mainloop()

# Need to not let the user add the same column twice
# Columns to keep list size must be the same for both data frames

# Error handling for move up/down/delete if the items doesn't exist