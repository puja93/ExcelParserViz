import os
from os import listdir
from os.path import isfile, join, dirname
import pandas as pd
import re
import numpy as np
import border_partition
import random_walk
import contour
import style_partition
import datetime
from openpyxl import load_workbook


def get_file_list(directory_name,debug=False):
    f_path_list = []
    f_name_list = []
    #f_dir = os.getcwd()+ '/' + directory_name + '/'
    f_dir = directory_name
    f_reports = [f for f in listdir(f_dir) if (isfile(join(f_dir, f)) and (('.XLSX' in f) or ('.xlsx' in f)))]
    f_path_list = [os.path.join(f_dir,f) for f in f_reports]

    if(debug):
        print('Directory name: {}'.format(f_dir))
        print('Available files: {}\n'.format(f_reports))

    return f_path_list


def get_sheet_list(sheets, mode='all', value='',debug=False):
    
    if(mode=='all'):
        sheet_list = [sheet.name for sheet in sheets if sheet.visibility==0]
    elif(mode=='hybrid'): #contain string&number
        sheet_list = [sheet.name for sheet in sheets if (bool(re.search(r'\d', sheet.name)) and sheet.visibility==0)]
    elif(mode=='numeric'):
        sheet_list = [sheet.name for sheet in sheets if (sheet.name.isnumeric() and sheet.visibility==0)]
    elif(mode=='alphabet'):
        sheet_list = [sheet.name for sheet in sheets if (not bool(re.search(r'\d', sheet.name)) and sheet.visibility==0)]
    elif(mode=='match'):
        sheet_list = [sheet.name for sheet in sheets if (bool(re.match(value, sheet.name)) and sheet.visibility==0)]
    elif(mode=='contain'):
        sheet_list = [sheet.name for sheet in sheets if (bool(value in sheet.name) and sheet.visibility==0)]
    else:
        sheet_list = None

    if(debug):
        print('Available sheets: {}\n'.format(sheet_list))
    return sheet_list


def wb_to_list(fname,sname):

    #Efficiency 1 : Clean unnecesary xml attributes
    workbook = load_workbook(filename=fname, read_only=True, data_only=True)
    ws = workbook[sname]

    #Efficiency 2 : Generator
    for row in ws.rows:
        yield [cell.value for cell in row]



def split_excel_table(fname,sname,thres=1,debug=False):
    #sht = mr_sheets[0]
    lst = wb_to_list(fname, sname)
    df = pd.DataFrame.from_records(lst)

    
    #df = excel_file.parse(sname,header=None)

    #SEGMENTATION
    borders = border_partition.auto_parse(fname,sname,debug=debug)
    borders = borders.fillna(0)
    borders = borders.astype('int')
    data_arr = borders.values

    origin_arr,mask_arr, final_arr = random_walk.auto_segment(data_arr, threshold=thres, visualize=debug)
    table_contours = contour.calculate(final_arr,visualize=debug)
    split_points = np.int64(contour.get_splitting_points(table_contours))
    
    
#     if(debug):
#         print('Processing file {} on sheet {}'.format(excel_file,sname))
#         print('Original array\n', origin_arr)
#         print('After random walk\n', final_arr)
#         print('After contour calculation\n',table_contours)
#         print('Split points\n',split_points)

    #SPLIT
    df_list = []
    for sp in split_points:
        df_list.append(df.loc[sp[0]:sp[1],sp[2]:sp[3]])

    return df_list



def experimental_split(excel_file,sname,style,debug=False):
    #sht = mr_sheets[0]
    if(debug):
        print('Processing file {} on sheet {}'.format(excel_file,sname))
    df = excel_file.parse(sname,header=None)

    #SEGMENTATION
    borders = style_partition.auto_parse(excel_file,sname,mode=style,debug=debug)
    borders = borders.fillna(0)
#     borders = borders.astype('int')
#     data_arr = borders.values

    for col in borders:
        borders[col] = borders[col].astype(int)
    data_arr = borders.values


    origin_arr,mask_arr, final_arr = random_walk.auto_segment(data_arr,visualize=debug)
    table_contours = contour.calculate(final_arr,visualize=debug)
    split_points = np.int64(contour.get_splitting_points(table_contours))
    
    if(debug):
        print('Processing file {} on sheet {}'.format(excel_file,sname))
        print('Original array\n', origin_arr)
        print('After random walk\n', final_arr)
        print('After contour calculation\n',table_contours)
        print('Split points\n',split_points)


    #SPLIT
    df_list = []
    for sp in split_points:
        df_list.append(df.iloc[sp[0]:sp[1],sp[2]:sp[3]])

    return df_list



def debug_output(df,prime_key_list):

    print('\n\033[5m' + 'DATAFRAME SAMPLE')
    print('\033[0m')
    print(df.head(10))
    
    dup_df = df[df.duplicated(subset=prime_key_list, keep=False)]
    print('\n\n', '\033[5m' + 'DUPLICATED DATA BASED ON PRIMARY KEYS', prime_key_list)
    print('\033[0m')
    print(dup_df)

    print('\n\n','\033[5m' + 'DATAFRAME INFO')
    print('\033[0m')
    df.info()

    print('\n\n','\033[5m' + 'UNIQUE VALUES FOR EACH COLUMNS')
    print('\033[0m')

    for col in df:
        print(df[col].nunique(),'unique values in column',col, 'are:')
        print(df[col].unique(),'\n')


        
def save_csv(df, output_name ,write=False):
    now = datetime.datetime.now()
    time_now = now.strftime("%d%m%Y-%H%M%S")
    filename = output_name +'-'+time_now+'.csv'
    print('Output filename:',filename)
    if(write):
        
        rel_path = 'result/' + filename
        full_path = os.path.join(dirname(os.getcwd()),rel_path)
        df.to_csv(full_path,index=False)
    