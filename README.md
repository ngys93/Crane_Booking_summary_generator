# Crane_Booking_summary_generator
import pandas as pd
from collections import Counter
import numpy as np
working_hours = ('Crane Type','08:00','09:00','10:00','11:00','12:00','13:00','14:00','15:00','16:00','17:00')
OT_hours = ('Crane Type','17:00:00','18:00:00','19:00:00','20:00:00','21:00:00','22:00:00','23:00:00','00:00:00','01:00:00','02:00:00','03:00:00',
            '04:00:00','05:00:00','06:00:00','07:00:00')

#read_file
def read_excel(path):
    df_master = pd.read_excel(path,None)#read all sheets
    return df_master

def summary(df_input,df_summary):
    global lst_normal
    #pass in empty dataframe with crane type and subcon from 07:00 to 17:00
    #check if crane type in df_summary
    sheet_name = tuple(df_input.keys())
    for sheet in sheet_name:
        df = df_input[sheet]
        #df concat dataframe
        df_summary = pd.concat([df,df_summary])
    #for col in list(df_summary.columns):
    #    if col not in lst_normal:
    #        df_summary.drop(columns=[col],inplace=True)
    return df_summary
    
def normal_hours(df):
    #select out normal hours
    global working_hours
    df_col = list(df.columns)
    df_col1 = tuple(df_col[:11])
    df_1 = df.copy(deep=True)
    for col in df_col:
        if col not in df_col1:
            df_1.drop(columns=[col],inplace=True)
    #add in column for subcontractor and sum
    df_1.index = np.arange(0,len(df_1['Crane Type']))
    df_1['Sub Con']=[0 for i in range(len(df_1['Crane Type']))]
    df_1['Count']=[0 for i in range(len(df_1['Crane Type']))]
    for indx in df_1.index:
        row = list(df_1.iloc[indx][1:11])#select row
        c_dic = Counter(list(row))
        for i in c_dic.keys():
            if i !=0 and i!='nan':
                df_1.iloc[indx,11]=i
                print(df_1.iloc[indx][11])
                #df_1.iloc[indx][12]=c_dic.get(i)
            
    #convert to format
    return df_1
def add_dic(val,length):#add team dic
    dic = {val:[0 for i in range(length)]}
    print(dic)
    return dic

def align_format(df):
    df_result = pd.DataFrame(columns = ['Crane Type'])
    df_result['Crane Type'] = df['Crane Type']
    df_result.drop_duplicates(subset = ['Crane Type'],
                              keep = "first", inplace = True)
    df_result.index = df_result['Crane Type']
    for crane_type in df_result.index:
        if crane_type in df['Crane Type']:#check if crane_type is in df
            #if crane_type in df
            df_1 = df.loc[crane_type]#select all crane types
            #select each row
            #reindex selected dataframe
            df_1.index = np.arange(0,len(df_1['Crane Type']))
        
            
    return df_result

def write_BCA(path,df1,df2):
    with pd.ExcelWriter(path) as writer:
        df1.to_excel(writer, sheet_name = 'result', index=True)
        df2.to_excel(writer, sheet_name = 'summary', index=True)
       
if __name__ == '__main__':
    path1 = str(input("Please enter in file location of test file:\n"))
    path1 = path1.replace('\\','/')
    path2 = str(input("Please enter in file location of result file:\n"))
    path2 = path2.replace('\\','/')
    df_master = read_excel(path1)
    lst_1=list(df_master.keys())
    df_summary = pd.DataFrame(columns = ['Crane Type'])
    df_summary.fillna('NA')  
    df_summary = summary(df_master,df_summary)
    df_normal = normal_hours(df_summary.copy(deep=True))
    df_result = align_format(df_normal.copy(deep=True))
    write_BCA(path2,df_result,df_normal)
