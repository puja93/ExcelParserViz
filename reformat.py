import pandas as pd
import numpy as np

#Global variables
year_list = ['2018','2019','2020','2021', '18', '19', '20', '21']

month_list = ['jan', 'feb', 'mar', 'apr', 'mei', 'may','jun', 
              
              'jul', 'agu', 'aug', 'sep', 'okt', 'oct', 'nov',
              
              'des', 'dec']

month_num_list = ['1', '2', '3', '4', '5', '5', '6', 
                  
                  '7', '8', '8', '9', '10', '10', '11',
                  
                  '12','12']


def filter_column(df,list_str,mode='retain'):
    unwanted_columns = []
    
    for str_marker in list_str:
        for col in df:
            temp = df[col].astype(str).str.contains(str_marker).sum()
            if(temp):
                unwanted_columns.append(col)

    res = [unwanted_columns if (mode=='remove') else list(set(df.columns)- set(unwanted_columns))][0]

    df.drop(columns=res,inplace=True)

    return df



def filter_row(df,list_str,mode='retain'):
    unwanted_rows = []
    for str_marker in list_str:
        for row in df.index:
            temp = df.loc[row].astype(str).str.contains(str_marker).sum()
            if(temp):
                unwanted_rows.append(row)
                
    res = [unwanted_rows if (mode=='remove') else list(set(df.index)- set(unwanted_rows))][0]

    df.drop(res,inplace=True)
    
    return df


def list_mapper(compared_value, from_list, to_list=[], reverse=False):
    
    #Mode True, check from_list in compared_value
    if(reverse):
        try:
            if(to_list == []):
                res = [y for y in from_list if compared_value in y][0]
            else:
                res = [y for x,y in zip(from_list, to_list) if compared_value in x][0]
        except:
            res = '-'   
            
    #Mode False, check compared_value in from_list
    else:
        try:
            if(to_list == []):
                res = [y for y in from_list if y in compared_value][0]
            else:
                res = [y for x,y in zip(from_list, to_list) if x in compared_value][0]
        except:
            res = '-'
    
    return res
    
def get_year(compared_value):
    
    year = list_mapper(compared_value.lower(),year_list)
    
    temp = int(year)
    if(temp < 2000):
        year = str(2000 + temp)
        
    return year
    
def get_month(compared_value):
    
    month = list_mapper(compared_value.lower(),month_list,month_num_list)
    return month
    
    
def unpivot(df, id_vars=[], value_vars=[], value_name='value', var_name='variable'):
    
    #Check column name
    #Assign first row automatically as column
    if( df.columns.dtype == 'int'):
        df.columns = df.iloc[0,:].values
        df.drop(df.index[0], inplace=True)
    
    #Check which one to unpivot
    if(not id_vars):
        id_vars = set(df.columns) - set(value_vars)
    else:
        value_vars = set(df.columns) - set(id_vars)
        
    df = pd.melt(df, id_vars = id_vars, value_vars = value_vars, value_name=value_name, var_name=var_name) 
    
    return df