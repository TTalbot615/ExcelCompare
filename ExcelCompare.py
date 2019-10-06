import pandas as pd

# Reads the excel file so you can work with it in python
df1 = pd.read_excel(r'C:\Users\Rislynn\Documents\Python Pandas Excels\VizData.xlsx')
df2 = pd.read_excel(r'C:\Users\Rislynn\Documents\Python Pandas Excels\CustomerData.xlsx')

# Takes a data frame as an input and returns a data frame with only the specified columns
# You specify the columns with an array of 0 and 1 whose length = # columns
# a value of 1 in the ith postion (i = 1 to # columns) indicates that column should be retained
def delete_cols(d,a):  # d is a dataframe and a is array of 0,1 that indicate if column in that position shoudl be kept
    # if len(a) != len(d.columns):
    #   print("Array is not long enough")
    col_to_keep = []  # initializes array that will contain the values of the columns to keep

    for i in range(len(a)):
        if a[i] == 1:  # If the ith entry is 1 then you will add this column to the list of ones to keep
            col_to_keep.append(i)

    return d.iloc[:, col_to_keep]

# print(delete_cols(df1,[1, 1, 1, 1, 1, 1, 1, 1, 0, 1, 0]))

def reorder_cols(d,a):  # d is a dataframe, a is an array with the index of the column in the order you would like now
#[0,1,5,6,7,8,2,3,4] column 5 would move to the third position for example
    return d.iloc[:, a]

#print(reorder_cols(df1,[0,1,5,6,7,8,2,3,4]))

def merge_df(d, f, k):  #dataframe1, dataframe2, keys = # columns at the beginning to use for join
    key_col_array = []
    for i in range(k):
        key_col_array.append(d.columns[i])  # Adds the ith column name to the array (d1 and d2 need same first key columns)
    result = pd.merge(d, f, how='outer', on=key_col_array)  # Joins df and df2 on the first num_key_columns
    return result
    #print(merge_df(df1,df2,2))

#function to create the variances out of a merged
#first columns = index + # key columns

def var_col(d,k):  #Takes merged_df and makes variances columns, k = number of key columns
    c = (len(d.columns) - k)//2
    for i in range(k, 2+c):
        first_col_name = d.columns[i]
        second_col_name = d.columns[i+c]
        var_col_name = "Var " + first_col_name[0:len(d.columns[i])-2]
        if type(d.iloc[1, i]) == str:  #Checks to see if the first entry of the column is a string, probalby needs to be more robust, like check more than just the first entry
            d[var_col_name] = d[first_col_name] == d[second_col_name]
        else:
            d[var_col_name] = d[first_col_name] - d[second_col_name]
    return d

######################################################################
#Steps:
#1 Import both excels
#2 Run delete_cols on both excels
#3 Run Reorder Cols on both excels
#4 Run merge_df on both data frames with the number of keys
#5 Run var_col on the resulting data frame from the last step

clean_df1 = delete_cols(df1, [1, 1, 1, 1, 1, 1, 1, 1, 0, 0, 0])
clean_df2 = delete_cols(df2, [1, 1, 1, 1, 1, 1, 1, 1])

reordered_df1 = reorder_cols(clean_df1, [0, 1, 5, 6, 7, 2, 3, 4]) #need to add in error handling if # cols <> length of array
#print(reordered_df1)
reordered_df2 = reorder_cols(clean_df2, [0, 4, 5, 6, 7, 1, 2, 3])

merged_dfs = merge_df(reordered_df1, reordered_df2, 2)
print(merged_dfs)


var_col(merged_dfs, 2)

print(merged_dfs)
merged_dfs.to_excel(r'C:\Users\Rislynn\Documents\Python Pandas Excels\ExcelCompare.xlsx')
##################

