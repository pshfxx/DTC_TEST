import os
import win32com.client
import re
import pandas as pd


def listit(t):
    return list(map(listit, t)) if isinstance(t, (list, tuple)) else t


def Excel_Load(i_file):
    excel = win32com.client.Dispatch("Excel.Application")
    excel_file = excel.Workbooks.Open(i_file)

    sheet_dtc = excel_file.Sheets('DTC List')
    data_dtc = sheet_dtc.UsedRange
    data_dtc = data_dtc.Value
    data_dtc = listit(data_dtc)

    sheet_ioc = excel_file.Sheets('(자동반영)I-O Control')
    data_ioc = sheet_ioc.UsedRange
    data_ioc = data_ioc.Value
    data_ioc = listit(data_ioc)

    excel_file.Close(True)
    excel.Quit()

    return data_dtc, data_ioc


def search_file():
    file_path = os.getcwd() + '\\' + 'Input'
    file_list = os.listdir(file_path)
    file_list_xlsx = [file for file in file_list if file.endswith(".xlsx")]

    file_1 = file_path + '\\' + file_list_xlsx[0]
    file_2 = file_path + '\\' + file_list_xlsx[1]

    return file_1, file_2, file_list_xlsx


def excute_Total(i_data):
    for temprow in i_data:
        for idx, val in enumerate(temprow):
            if type(val) is float:
                temprow[idx] = unicode(int(val))

    row1 = i_data[1];
    fieldBuffer = []
    for index, temp in enumerate(row1):
        if row1[index]:
            row1V = row1[index]
            row1V = re.sub(r'\s+', "_", row1V)
            fieldBuffer.append(row1V)
    #print(fieldBuffer)
    excelData = pd.DataFrame(i_data[2:], columns=fieldBuffer)
    try:
        excelData['Description'] = excelData['Description'].replace('\s+', '', regex=True)
    except:
        excelData['Description'] = excelData['DTC']
    return excelData

def excute_Total_Ioc(i_data):
    for temprow in i_data:
        for idx, val in enumerate(temprow):
            if type(val) is float:
                temprow[idx] = unicode(int(val))

    row1 = i_data[7];
    fieldBuffer = []
    for index, temp in enumerate(row1):
        if row1[index]:
            row1V = row1[index]
            row1V = re.sub(r'\s+', "_", row1V)
            fieldBuffer.append(row1V)
    #print(fieldBuffer)
    excelData = pd.DataFrame(i_data[7:], columns=fieldBuffer)
    try:
        excelData['Description'] = excelData['Description'].replace('\s+', '', regex=True)
    except:
        excelData['Description'] = excelData['I/O LID']
    return excelData


def make_Total_list(data1, data2):
    def dtc_list(i_data):
        list_data = []
        for idx, DTC in enumerate(i_data['DTC']):
            if i_data[u'\uc801\uc6a9_\uc720\ubb34'][idx] == 'O':
                list_data.append(DTC)

        return list_data

    dtc_list_A = dtc_list(data1)
    dtc_list_B = dtc_list(data2)

    return dtc_list_A, dtc_list_B


def comp_list(i_total, i_data):
    list_data = []
    for temp in i_total:
        write_jdg = False
        for temp1 in i_data:
            if temp == temp1:
                write_jdg = True
                break
        if write_jdg:
            list_data.append('O')
        else:
            list_data.append('X')

    return list_data
