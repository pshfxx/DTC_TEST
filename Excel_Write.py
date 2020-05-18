import pandas as pd
import os
import win32com.client


def createFolder(dir):
    try:
        if not os.path.exists(dir):
            os.makedirs(dir)
    except OSError:
        print('Error : Creating directory. ' + dir)

def createList(i_data,i_total):
    list_data = []
    cnt = 0
    max_cnt = len(i_total)
    for idx, Desc in enumerate(i_data['Description']):
        #if ((i_des_list['DTC'][idx] == i_total[cnt])):
        if (cnt >= max_cnt):
            return list_data
        elif (i_data['DTC'][idx] == i_total[cnt]):
            list_data.append(Desc)
            cnt += 1

def Make_Excel(i_total, i_data1, i_data2, i_file_name, i_des_list):
    file = os.getcwd()
    file_path = file + '\\' + 'Output'
    createFolder(file_path)

    file_ = file_path + '\\' + 'Result.xlsx'

    writer = pd.ExcelWriter(file_)
    '''
    df_total = pd.DataFrame(i_total)
    df_data1 = pd.DataFrame(i_data1)
    df_data2 = pd.DataFrame(i_data2)

    df_total.to_excel(writer, sheet_name='Result', columns=None, index=False, startrow=1, startcol=1)
    df_data1.to_excel(writer, sheet_name='Result', index=False, startrow=1, startcol=2)
    df_data2.to_excel(writer, sheet_name='Result', index=False, startrow=1, startcol=3)
    '''
    ###### 포멧 만들기###############
    list_data = createList(i_des_list, i_total)
    ################################
    #df = pd.DataFrame({"DTC_CODE": [], i_file_name[0]: [], i_file_name[1]: []})
    df = pd.DataFrame({"DTC_CODE": [], "Description": [], i_file_name[0]: [], i_file_name[1]: []})
    df["DTC_CODE"] = i_total
    df["Description"] = list_data
    df[i_file_name[0]] = i_data1
    df[i_file_name[1]] = i_data2

    df.to_excel(writer, sheet_name='Result', index=False, startrow=1, startcol=1)

    writer.save()

    return
