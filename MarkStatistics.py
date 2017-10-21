# statistics the score
# MarkStatistics.py
# BonizLee
# -*- coding: utf-8 -*-

import json
from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
import os.path

def init():   
    global PATH
    PATH = os.path.dirname(os.path.realpath(__file__))+os.path.sep
    fp = open(PATH+'config.json')    
    global COMMOM_DATA
    COMMOM_DATA = json.load(fp)
    print(COMMOM_DATA)
    global MAXNUMBER 
    MAXNUMBER = COMMOM_DATA['maxnumber']
    global PROJECT 
    PROJECT = COMMOM_DATA['project']
    global STUDENT_MARK
    STUDENT_MARK = [0 for i in range(MAXNUMBER)]
    global FILETYPE
    FILETYPE = '.'+COMMOM_DATA['filetype']

def summary():      
    for subject in COMMOM_DATA['subject']:        
        judges = subject['judges']
        subjectmark =  [ [ 0 for i in range(MAXNUMBER) ] for j in range(judges) ]
        for n in range(0,judges):            
            filename = PATH+PROJECT+subject['filename']+str(n+1)+FILETYPE
            workbook = load_workbook(filename,data_only=True)
            sheetnames =workbook.get_sheet_names()
            sheet = workbook.get_sheet_by_name(sheetnames[0])
            start_cell = sheet[subject['markcell']]
            start_col = column_index_from_string(start_cell.column)
            start_row = start_cell.row
            for i in range(0,MAXNUMBER):
                markcell = sheet.cell(row=start_row,column=start_col+i)                
                subjectmark[n][i] = markcell.value
        print('评分表'+subject['filename']+'完成')
        print(subjectmark)
        calc_type = subject['calculate']
        if calc_type == 1:
            average(subjectmark,judges)
        elif calc_type == 2:
            without_max_min_average(subjectmark,judges)         

def average( subjectmark,judges ):    
    for i in range(0,MAXNUMBER):
        sum = 0
        for n in range(0,judges):
            sum += subjectmark[n][i]
        STUDENT_MARK[i] += sum / judges

def without_max_min_average(  subjectmark,judges ):
    for i in range(0,MAXNUMBER):
        max = 0
        sum = 0
        min = 0
        for n in range(0,judges):
            value = subjectmark[n][i]
            sum += value
            if value > max:
                max = value
            if value < min:
                min = value
        STUDENT_MARK[i] += sum - min -max / (judges -2)

def write_excel():
    filename = PATH+PROJECT + 'summary'+FILETYPE   
    workbook = Workbook()
    sheet = workbook.get_active_sheet()
    sheet.cell(row=1,column=1).value='工位号'
    sheet.cell(row=1,column=2).value='成绩'
    for i in range(len(STUDENT_MARK)):
        sheet.cell(row=2+i,column=1).value= i+1
        sheet.cell(row=2+i,column=2).value= STUDENT_MARK[i]
    workbook.save(filename)

if __name__ == "__main__":
    init()
    print('加载配置完成')
    summary()
    print('汇总完成完成')
    print(STUDENT_MARK)
    write_excel()
    print('保存完成')
    while True:
        q = input("输入Q退出程序：")
        if q == 'Q' or q == 'q':
            break 