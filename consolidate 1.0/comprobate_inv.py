import os

import pandas as pd
from tkinter import messagebox
import locale
import datetime

fileList = []
INV_INITIAL = 1
INV_END = 0


# def equals_products(id_1, id_2):
#     return

for file in os.listdir('./'):
    if file.__contains__('.xlsx') and file[0] != '~' and file.__contains__('Reporte'):
        fileList.append(file)

if len(fileList) != 2:
    messagebox.showwarning('Comprobacion de Inventario',
                           'Debe de haber dos reportes para realizar la comprobación de correspondencia entre inventarios.')
else:
    dict_products = {}

    start_name = 1
    step_name = 14

    start_inv_init = 5
    step_inv_init = 14

    start_inv_end = 9
    step_inv_end = 14

    dict_inv_init = {}
    dict_inv_end = {}
    dict_inv = {0: dict_inv_end, 1: dict_inv_init}

    count = 0
    for file in fileList:
        print('FILE IS: ', file)
        start_inv = 0
        step_inv = 0

        if count == 0:
            start_inv = start_inv_end
            step_inv = step_inv_end
        else:
            start_inv = start_inv_init
            step_inv = step_inv_init

        current_possition_inv = start_inv
        current_possition_name = start_name

        try:
            # print(file)
            data_sheet_1 = pd.read_excel(file, sheet_name=0)  # Total Prod-Centro
            dict_products[count] = data_sheet_1[data_sheet_1.columns[0]][1:200]
            # print(data_sheet_1[data_sheet_1.columns[0]][1:200])
            # print('*********************************************************************')
            # print(len(dict_products[count]))
            # print(len(data_sheet_1.columns))
            # print(data_sheet_1.columns[1])
            # print('len(data_sheet_1.columns): ', len(data_sheet_1.columns))
            while current_possition_inv < len(data_sheet_1.columns):
                name_temp = data_sheet_1.columns[current_possition_name]
                dict_inv[count][name_temp] = data_sheet_1[data_sheet_1.columns[current_possition_inv]][1:200].fillna(0)

                # current_possition_inv += start_inv
                current_possition_inv += step_inv
                current_possition_name += step_name
        except Exception as e:
            print(repr(e))

        count += 1

    dict_inv_correct = {}
    if not dict_products[0].equals(dict_products[1]):
        print('Corregir correlacion de ID productos')
        further = 0
        less = 0
        if dict_products[0].size > dict_products[1].size:
            further = len(dict_products[0])
            less = len(dict_products[1])
        else:
            further = len(dict_products[1])
            less = len(dict_products[0])

        # for i in range(further.size):
        #     if i>less.size:
        #         dict_products
    else:
        print('Son iguales!')

    list_columns = []
    # print('Size END: ', len(dict_inv_end))
    # print('Size INIT: ', len(dict_inv_init))
    for key in dict_inv_end.keys():
        print(key)
        # print(dict_inv_end[key])

        # print(dict_inv_init[key])

        column = pd.DataFrame(
            {
                'Desfasaje': [key] + ['OK' if value == 0 else value for value in list(
                    dict_inv_end[key] - dict_inv_init[key]
                )]}
        )



