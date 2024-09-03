import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Color
from openpyxl.worksheet.table import Table, TableStyleInfo
import xlwings as xw
import time
import os

t1 = time.time()

base_file_location = r"C:\Users\aasabu\Desktop\Sales and Marketing\Python codes\Requirement Subclass"
reqDump_file_location = base_file_location + r"\Req_US_Dump.xlsx"
output_location = base_file_location + r"\ReqSubClassOutput.xlsx"

reqDump = pd.read_excel(reqDump_file_location)

def check_sub_class(row,array):

    if str(row['SubClassification']) == str(row['SubClassification2']):
        array.append(str(row['Requirement ID']))
        return 'True'
    return ''

def gap_check(row,array):

    if str(row['SubClassification']) == 'Gap - Process Deviation' or str(row['SubClassification']) == 'Gap - WRICEF-(E)-Enhancement' or str(row['SubClassification']) == 'Gap - Extended configuration not aligned to Best Practices':
        if str(row['SubClassification']) != str(row['SubClassification2']):
            for x in array:
                if str(row['Requirement ID']) == x:
                    return ''
            return 'True'
            #return 'No US Subclass matches req subclass'
        
    return ''

def fit_to_gap_check(row):
    if str(row['SubClassification']) != 'Gap - Process Deviation' and str(row['SubClassification']) != 'Gap - WRICEF-(E)-Enhancement' and str(row['SubClassification']) != 'Gap - Extended configuration not aligned to Best Practices':
        if str(row['SubClassification2']) == 'Gap - Process Deviation' or str(row['SubClassification2']) == 'Gap - WRICEF-(E)-Enhancement' or str(row['SubClassification2']) == 'Gap - Extended configuration not aligned to Best Practices':
            return 'True'
    
req_Array = []

reqDump['Check'] = reqDump.apply(check_sub_class,args=(req_Array,),axis=1)
reqDump['GAP Req Check'] = reqDump.apply(gap_check,args=(req_Array,),axis=1)
reqDump['Fit-GAP Issues'] = reqDump.apply(fit_to_gap_check,axis=1)

reqDump['FIT - 1:1 Match'] = ''
reqDump['FIT - atleast 1 Match'] = ''
reqDump['FIT - no match'] = ''

for req_id in reqDump['Requirement ID'].unique():

    req_group= reqDump[reqDump['Requirement ID'] == req_id]

    if req_group['GAP Req Check'].any():
        continue
    if req_group['Fit-GAP Issues'].any():
        continue
    if req_group['Check'].all():
        reqDump.loc[reqDump['Requirement ID'] == req_id, 'FIT - 1:1 Match'] = 'True'
    elif req_group['Check'].any():
        reqDump.loc[reqDump['Requirement ID'] == req_id, 'FIT - atleast 1 Match'] = 'True'
    else:
        reqDump.loc[reqDump['Requirement ID'] == req_id, 'FIT - no match'] = 'True'

print(req_Array)

with pd.ExcelWriter(output_location, engine='openpyxl') as writer:
    reqDump.to_excel(writer, sheet_name='BaseFile', index=False)

print("Completed..")
t2 = time.time()
print(t2-t1)