import xlrd
import re
import time
import sys
from fuzzywuzzy import fuzz

def getListFromXLSX(filename, withId=False):
    dataList =[]
    
    book = xlrd.open_workbook('Контрагенты.xlsx')
    sheet = book.sheet_by_index(0)
    for row in range(sheet.nrows-1):
        if withId:
            dataList.append((sheet.row_values(row+1)[0], sheet.row_values(row+1)[1]))
        else:
            dataList.append(sheet.row_values(row+1)[1])
    return dataList

def unification(inputList):
    pattern = r'''(?x)
            [,"'@_#!~/%:;=<>]
            |(\-й\b)|(\-ая\b)|(\-ый\b)|(\-е\b)|(\-я\b)|(\-ое\b)|(\-го\b) 
            |([\?\+\^\&\[\]\(\)\\\|\*])
           '''
    unifList = []
    for str in inputList:
        unifStr = str.lower()
        #it uses for correct proccessing string like that '1 - ая'  further
        unifStr = re.sub(r'(\s*\-\s*)',
            lambda match: '-', 
            unifStr)
        unifStr = re.sub(pattern,  
            lambda match: ' ', 
            unifStr)
        unifStr = re.sub(r'\-',  
            lambda match: ' ', 
            unifStr)
        unifStr = re.sub(r'\s{2,}',
            lambda match: ' ', 
            unifStr)
        unifStr = unifStr.strip()
        unifList.append(unifStr)
    return  sorted(unifList)

def deleteDuplicate(inputList, threshold=95, threshold2=80, trace=False):
    listWithoutDupl = inputList
    listDupl = []
    
    if trace:
        count = 0
    
    for str in listWithoutDupl:
        if trace:
            count += 1
            sys.stdout.write('{0}\r'.format(count))
        
        index = listWithoutDupl.index(str)
        i = 1
        while index + i < len(listWithoutDupl):
            compareStr = listWithoutDupl[index + i]
            ratio = fuzz.ratio(str, compareStr)
            if ratio < threshold2:
                break
            elif ratio > threshold:
                key = 'y'
                if re.findall(r'\d+', str) != []:
                    if re.findall(r'\d+', str) != re.findall(r'\d+', compareStr):                    
                        key = 'n'
                if key == 'y':
                    listWithoutDupl.remove(compareStr)
                    listDupl.append((str, compareStr))
            i+=1
    if trace:
        sys.stdout.write('\n')
    return  (listWithoutDupl, listDupl)
