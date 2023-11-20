
from pathlib import Path
import numpy as np
import pandas as pd
from IPython.display import display
import tabula
import openpyxl as op
from openpyxl.drawing.image import Image
from datetime import datetime, timedelta

trav_df = None
processes = [] #list of processes
p_df = None
due_date = None
timezone = None

def pdf_to_df(file_name):
    global trav_df
    df = tabula.read_pdf(file_name, pages='all')

    concatenated_df = df[0]
    for d in df[1:]:
        concatenated_df = pd.concat([concatenated_df,d], axis=1)
    #print(concatenated_df)
    trav_df = concatenated_df

def create_excel(): #fill in the variable data
    workbook = op.load_workbook('TRAVELER TEMPLATE.xlsx')
    ws = workbook.active
    #res = str(df.loc[0,'Purchase Order'])

    #print( df.loc[0,'Purchase Order'])
    ws['G4'] = trav_df.loc[0,'Purchase Order']
    ws['G5'] = trav_df.loc[0,'Customer Part ID']
    ws['G6'] = format_pn()
    ws['B9'] = trav_df.loc[0,'Quantity']
    ws['C9'] = trav_df.loc[0,'Due Date']
    ws['B13'] = trav_df.loc[0,'Material']
    ws['E13'] = calculate_quantity()
    ws['D9'] = calculate_shop_due()

    #first add processes to queue

    workbook.save("Traveller-1.xlsx")

def calculate_shop_due():
    shop_day = due_date - timedelta(days=1)
    shop_str = shop_day.strftime('%m/%d/%Y')
    shop_str = shop_str + "," + timezone
    return shop_str

def calculate_due_date():
    global due_date
    global timezone

    unformatted = trav_df.loc[0,'Due Date']
    date = unformatted.split(',')
    date_num = date[0]
   
    date_format = '%m/%d/%Y'
    due_date = datetime.strptime(date_num,date_format)
    timezone = date[1]
    print("Date:",due_date)

def format_pn():
    part = trav_df.loc[0,'Part Name']
    part = part.split('.')
    return part[0]

def calculate_quantity():
    quantity = trav_df.loc[0,'Quantity']
    if(quantity <= 10 ):
        shop_quantity = quantity + 1
    else:
        shop_quantity = quantity + 2

    return shop_quantity

def find_deburr_description():
    desc = ""
    if ('Finish' in trav_df.columns and trav_df.loc[0,'Finish'] != 'Standard'):
        desc = 'Deburr and Clean'
    else:
        desc = 'Deburr'
    return desc

def create_queue(t_status): #create queue for all processes
    global processes
    if(t_status == True):
        processes.append("Turning")
    processes.append("Operations")
    processes.append("Deburr")

    if ('Part Marking' in trav_df.columns and trav_df.loc[0,'Finish'] != 'Bag and Tag'):
        processes.append("Part Marking")

    if('Finish' in trav_df.columns and trav_df.loc[0,'Finish'] != 'Standard'):
        processes.append("Finish")
    if ('Inserts' in trav_df.columns):
        processes.append("Inserts")
    processes.append("Final Inspection")
    processes.append("Bag and Tag")

def create_p_df(line_items):
    global processes
    #dataframe with items in the format Process, Description, Due Date
    global p_df
    column_names = ['Process','Description','Due Date']
    p_df = pd.DataFrame(columns = column_names)
    

    processes = reversed(processes) #going from last process to first process
    for p in processes:

        match p:
            case 'Bag and Tag':
                #calculate date
                if line_items > 5:
                    dd = due_date - timedelta(days = 2)
                else:
                    dd = due_date - timedelta(days = 1)
                #description is always the same
                desc = "Package to prevent shipping damage"
            case 'Final Inspection':
                if line_items > 5:
                    dd = due_date - timedelta(days = 2)
                else:
                    dd = due_date - timedelta(days = 1)
                desc = "QC"
            case 'Inserts':
                day = p_df.loc[p_df['Process'] == 'Final Inspection', 'Due Date'].values[0]
                day = datetime.utcfromtimestamp(day.astype('datetime64[s]').astype(int))
                dd = day - timedelta(days = 3)

                desc = trav_df.loc[0,'Inserts']

                #print("Day", day)
                
            case _:
                dd = due_date
                desc = "Not Done Yet"

        p_df['Due Date'] = pd.to_datetime(p_df['Due Date'])
        row = {"Process": p,"Description": desc,"Due Date":dd}
        #p_df = p_df.append(row, ignore_index=True)
        p_df.loc[len(p_df)] = row
        
    print("Process data frame:\n",p_df)



        
if __name__ == '__main__':

    file_name = '0571C6E-traveler.pdf'
    #file_name = '057531C-traveler.pdf'
    path =  Path('jfiles',file_name)
    pdf_to_df(path)
    calculate_due_date()
    create_queue(False)
    create_p_df(2)
    create_excel()