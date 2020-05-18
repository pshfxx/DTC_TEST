import Excel_Write as EW
import Excel_Load as EL
import os
import pandas as pd
import numpy as np

if __name__ == '__main__':

    file_1, file_2, file_list_xlsx = EL.search_file()

    data_DTC_1, data_IOC_1 = EL.Excel_Load(file_1)
    data_DTC_2, data_IOC_2 = EL.Excel_Load(file_2)

    data_DTC_T1 = EL.excute_Total(data_DTC_1)
    data_DTC_T2 = EL.excute_Total(data_DTC_2)

    data_IOC_T1 = EL.excute_Total_Ioc(data_IOC_1)
    data_IOC_T2 = EL.excute_Total_Ioc(data_IOC_2)

    print("Read Complete")
    a, b = EL.make_Total_list(data_DTC_T1, data_DTC_T2)
    #b = EL.make_Total_list(data_T2)

    total_list = a + b


    total_list = list(set(total_list))

    total_list = sorted(total_list)


    list_A = EL.comp_list(total_list, a)
    list_B = EL.comp_list(total_list, b)
    #hex(list_A[0])

    EW.Make_Excel(total_list, list_A, list_B, file_list_xlsx)
    print("Write Complete")
    '''
    '''