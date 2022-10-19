import os

import pandas as pd
from tkinter import *
from tkinter import messagebox
import locale
import datetime

current_Date = datetime.datetime.now()
date = str(current_Date.day) + '-' + str(current_Date.month) + '-' + str(current_Date.year) + '  ' + str(
    current_Date.hour) + 'h y ' + str(current_Date.minute) + 'm'
print(date)


def positive_and_negative_numbers_print(values, star_row, end_row, col_num, worksheet, workbook):
    for i in range(star_row, end_row):

        if values[i - 1][col_num] is not None and values[i - 1][col_num] < 0:
            worksheet.write(i, col_num, values[i - 1][col_num], workbook.add_format(
                {
                    'bold': True,
                    'font_color': '#FF0000',
                    'num_format': '#,#0.0'
                })
                            )
        elif values[i - 1][col_num] is not None and values[i - 1][col_num] > 0:
            worksheet.write(i, col_num, values[i - 1][col_num], workbook.add_format(
                {
                    'bold': True,
                    'font_color': '#4D975D',
                    'num_format': '#,#0.0'
                })
                            )


def differ_than(values, star_row, end_row, col_num_value, than_column, worksheet, workbook):
    for i in range(star_row, end_row):
        temp_val = values[i - 1][than_column]
        if temp_val is None:
            temp_val = 0

        if values[i - 1][col_num_value] is not None:
            if values[i - 1][col_num_value] < temp_val:
                worksheet.write(i, col_num_value, values[i - 1][col_num_value], workbook.add_format(
                    {
                        'bold': True,
                        'font_color': '#FF8F8F',
                        'num_format': '#,#0.0'
                    })
                                )
            elif values[i - 1][col_num_value] > temp_val:
                worksheet.write(i, col_num_value, values[i - 1][col_num_value], workbook.add_format(
                    {
                        'bold': True,
                        'font_color': '#85C192',
                        'num_format': '#,#0.0'
                    })
                                )
            elif values[i - 1][col_num_value] == temp_val:
                worksheet.write(i, col_num_value, values[i - 1][col_num_value], workbook.add_format(
                    {
                        'font_color': '#757575',
                        'num_format': '#,#0.0'
                    })
                                )


def set_color_column(worksheet, init, column, array, color):
    for i in range(init, len(array)):
        value = array[i][column] if str(array[i][column]) != 'nan' else ''
        worksheet.write(i + 1, column, value, workbook.add_format({
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': color,
            'bold': True,
            'font_color': '#000000',
            'num_format': '#,#0.0'
        }))


def negative_numbers_print(values,star_row, end_row, col_num, worksheet, workbook):
    for i in range(star_row, end_row):
        if values[i - 1][col_num] is not None and values[i - 1][col_num] < 0:
            worksheet.write(i, col_num, values[i - 1][col_num], workbook.add_format(
                {
                    'bold': True,
                    'font_color': '#FF0000',
                    'num_format': '#,#0.0'
                })
            )

def only_one_decimal(number, symbol=False):
    value = locale.currency(number, grouping=True, symbol=True)
    return value[0: len(value)-3].replace(',', '.')

fileList = []
sellManagersNames = []
seller_account = {}
types_account = {}
id_unique_account = {}
totalManager = {}
point_of_sale = {'Vendedor': [], 'Asistente': [], 'Dueno': []}
point_admin_logistic = {'AE': [], 'SM': [], 'CD': [], 'LCA': [], 'LTD': [], 'CEO': []}
# Var of Sheet Three
other_income = []
other_egress = []
buy_egress = []
net_flow = []
real_cash = []
desf_cash = []
COLOR_ACCOUNT_ID = ['#FCD5B4', '#B1A0C7', '#76933C', '#DA9694', '#92CDDC', '#DEC400', '#CC5D26', '#C4D79B', '#F95555',
                    '#4F81BD', '#808080', '#DE58C8', '#00CC99', '#0099FF']
COUNT_COLOR = 0
total_ir_group = 0 # Sheet number 8
total_ft_group = 0 # Sheet number 8
total_fr_group = 0 # Sheet number 8


# Agregar al nombre del reporte la hora del sistema y fecha.
name_new_report = 'Reporte [' + date + '].xlsx'
writer = pd.ExcelWriter(name_new_report, engine='xlsxwriter')
locale.setlocale(locale.LC_ALL, 'es_CU')

for file in os.listdir('./'):
    if file.__contains__('.xlsx') and file[0] != '~' and file.__contains__('Reporte') == False:
        fileList.append(file)
        init = file.find('-') + 1
        end = file.find('.') - 2
        sellManagersNames.append(file[init: end].split()[0])

    products_id = []
    managerDict = {}
    count = 0
    correctFiles = True

for file in fileList:
    try:
        print(file)
        data_sheet_1 = pd.read_excel(file, sheet_name=1)  # Monday
        data_sheet_8 = pd.read_excel(file, sheet_name=8)  # Seller Summary
        data_sheet_9 = pd.read_excel(file, sheet_name=9)  # Sell Manager Summary

        products_id = list(data_sheet_8['RESUMEN'][1:201])

        try:
            # Extracting the necessary data
            dayDict = [data_sheet_8[data_sheet_8.columns[29]][1:201]]  # Cant Venta (AD)
            dayDict += [data_sheet_8[data_sheet_8.columns[38]][1:201]]  # Inv.IR (AM)
            dayDict += [data_sheet_8[data_sheet_8.columns[39]][1:201]]  # In (AN)
            dayDict += [data_sheet_8[data_sheet_8.columns[40]][1:201]]  # Out (AO)
            dayDict += [data_sheet_8[data_sheet_8.columns[41]][1:201]]  # Inv.FT (AP)
            dayDict += [data_sheet_8[data_sheet_8.columns[42]][1:201]]  # Inv.FR (AQ)
            dayDict += [data_sheet_8[data_sheet_8.columns[42]][1:201] - data_sheet_8[data_sheet_8.columns[41]][
                                                                        1:201]]  # Desfasaje
            dayDict += [data_sheet_8[data_sheet_8.columns[50]][1:201]]  # Cost.V (AY)
            dayDict += [data_sheet_8[data_sheet_8.columns[51]][1:201]]  # Ing.V (AZ)
            dayDict += [data_sheet_8[data_sheet_8.columns[52]][1:201]]  # Result (BA)
            dayDict += [data_sheet_8[data_sheet_8.columns[54]][1:201]]  # Merc.V (BC)
            dayDict += [data_sheet_8[data_sheet_8.columns[55]][1:201] * -1]  # Desf.Inv (BD)
            dayDict += [data_sheet_8[data_sheet_8.columns[31]][1:201]]  # Cant.C(AF)
            dayDict += [data_sheet_8[data_sheet_8.columns[32]][1:201]]  # Egr.C (AG)

            # Account Values
            id_temp = list(data_sheet_1[data_sheet_1.columns[23]][35:236])  # id
            name_temp = list(data_sheet_1[data_sheet_1.columns[22]][35:236])  # nombres
            ir_temp = list(data_sheet_1[data_sheet_1.columns[24]][35:236])  # ir
            ft_temp = list(data_sheet_9[data_sheet_9.columns[12]][27:228])  # ft
            fr_temp = list(data_sheet_9[data_sheet_9.columns[11]][27:228])  # fr

            # Total of Account Group
            if type(total_ir_group) is int:
                total_ir_group = data_sheet_1[data_sheet_1.columns[23]][23:33]
                total_ft_group = data_sheet_9[data_sheet_9.columns[11]][15:25]
                total_fr_group = data_sheet_9[data_sheet_9.columns[10]][15:25]
            else:
                total_ir_group += data_sheet_1[data_sheet_1.columns[23]][23:33]
                total_ft_group += data_sheet_9[data_sheet_9.columns[11]][15:25]
                total_fr_group += data_sheet_9[data_sheet_9.columns[10]][15:25]
            # print(total_ir_group, "\n\n\n")

            def search_id(list_temp):
                set_temp = set()
                for i in range(len(list_temp)):
                    if pd.isna(list_temp[i]) == False and list_temp[i] != 0:
                        set_temp.add(i)

                return set_temp


            set_position_temp = set()
            set_position_temp |= search_id(ir_temp)

            set_position_temp |= search_id(ft_temp)
            set_position_temp |= search_id(fr_temp)
            id_temp_2 = []
            name_temp_2 = []
            ir_temp_2 = []
            ft_temp_2 = []
            fr_temp_2 = []

            for pos in set_position_temp:
                id_temp_2.append(id_temp[pos])
                name_temp_2.append(name_temp[pos])
                ir_temp_2.append(float(ir_temp[pos]))
                ft_temp_2.append(float(ft_temp[pos]))
                fr_temp_2.append(float(fr_temp[pos]))

            id_temp = id_temp_2
            ir_temp = ir_temp_2
            ft_temp = ft_temp_2
            fr_temp = fr_temp_2
            name_temp = name_temp_2
            cc = [sellManagersNames[count]] + [None for i in range(len(id_temp) - 1)]

            if len(id_temp) > 0:
                seller_account[sellManagersNames[count]] = {'name': name_temp, 'id': id_temp,
                                                            'cc': cc, 'IR': ir_temp,
                                                            'FT': ft_temp, 'FR': fr_temp}
                for i in range(len(id_temp)):
                    if id_temp[i] in id_unique_account:
                        color = id_unique_account[id_temp[i]]['color']
                        if color == 'none':
                            color = COLOR_ACCOUNT_ID[COUNT_COLOR]
                            COUNT_COLOR += 1
                        id_unique_account[id_temp[i]] = {'name': name_temp[i],
                                                         'IR': id_unique_account[id_temp[i]]['IR'] + ir_temp[i],
                                                         'FT': id_unique_account[id_temp[i]]['FT'] + ft_temp[i],
                                                         'FR': id_unique_account[id_temp[i]]['FR'] + fr_temp[i],
                                                         'repeat': True,
                                                         'color': color}
                    else:
                        id_unique_account[id_temp[i]] = {'name': name_temp[i], 'IR': ir_temp[i], 'FT': ft_temp[i],
                                                         'FR': fr_temp[i],
                                                         'repeat': False, 'color': 'none'}

            # Account Values ~2
            types_account[sellManagersNames[count]] = {'type': list(data_sheet_1[data_sheet_1.columns[22]][23:33]),
                                                       'ir': list(data_sheet_1[data_sheet_1.columns[23]][23:33]),
                                                       'ft': list(data_sheet_9[data_sheet_9.columns[11]][15:25]),
                                                       'fr': list(data_sheet_9[data_sheet_9.columns[10]][15:25])}

            # Var of Sheet Three
            other_income.append(float(data_sheet_8[data_sheet_8.columns[34]][10]))
            other_egress.append(float(data_sheet_8[data_sheet_8.columns[34]][20]))
            buy_egress.append(float(data_sheet_8[data_sheet_8.columns[33]][7]))
            net_flow.append(float(data_sheet_8[data_sheet_8.columns[33]][13]))
            real_cash.append(float(data_sheet_9[data_sheet_9.columns[6]][18]))
            desf_cash.append(float(data_sheet_8[data_sheet_8.columns[33]][20]))

            point_of_sale['Vendedor'].append(float(data_sheet_8[data_sheet_8.columns[35]][33]))
            point_of_sale['Asistente'].append(float(data_sheet_8[data_sheet_8.columns[35]][34]))
            point_of_sale['Dueno'].append(float(data_sheet_8[data_sheet_8.columns[35]][35]))

            point_admin_logistic['AE'].append(float(data_sheet_9[data_sheet_9.columns[9]][1]))
            point_admin_logistic['SM'].append(float(data_sheet_9[data_sheet_9.columns[9]][2]))
            point_admin_logistic['CD'].append(float(data_sheet_9[data_sheet_9.columns[9]][3]))
            point_admin_logistic['LCA'].append(float(data_sheet_9[data_sheet_9.columns[9]][4]))
            point_admin_logistic['LTD'].append(float(data_sheet_9[data_sheet_9.columns[9]][5]))
            point_admin_logistic['CEO'].append(float(data_sheet_9[data_sheet_9.columns[9]][6]))

            # Get information of move in account


        except Exception as e:
            print(repr(e))
            # text = " " + file
            # messagebox.showwarning('Problema en: ', text)
            # correctFiles = False

        if correctFiles:
            total = {}
            arr = data_sheet_8.to_numpy()

            total['Cost.V'] = round(sum(dayDict[7]), 2)
            total['Ing.V'] = round(sum(dayDict[8]), 2)
            total['Result'] = round(sum(dayDict[9]), 2)
            total['Merc.V'] = round(sum(dayDict[10]), 2)
            total['Desc.DI'] = round(sum(dayDict[11]), 2)
            total['Egr.C'] = round(sum(dayDict[13]), 2)

            managerDict[sellManagersNames[count]] = dayDict
            totalManager[sellManagersNames[count]] = total
            count = count + 1
    except Exception as e:
        print(repr(e))
        messagebox.showwarning('Consolidate', 'Verifique todos los archivos excel\n Puede que tenga alguno no estandar')
        correctFiles = False

if correctFiles and len(sellManagersNames) > 0:
    # Var of Sheet Two
    total_amount = 0
    total_inv_sem_ant = 0
    total_inv_teorico = 0
    total_inv_real = 0
    total_desfasaje = 0
    total_cost_v = 0
    total_venta = 0
    total_result = 0
    total_merc_v = 0
    total_desc_inv = 0

    column_RESUMEN = pd.DataFrame(
        {
            'RESUMEN': ['ID'] + products_id
        }
    )
    values_sm = {'Cost.V': [], 'Ing.V': [], 'Egr.C': [], 'Result': [], 'Merc.V': [], 'Desc.DI': []}
    consolidate = [column_RESUMEN]
    for sm in sellManagersNames:
        column_amount_sell = pd.DataFrame(
            {
                sm: ['Cant.V.'] + [None if value == 0 else value for value in list(managerDict[sm][0])]
            }
        )

        column_inv_init_real = pd.DataFrame(
            {
                '': ['Inv.IR'] + [None if value == 0 else value for value in list(managerDict[sm][1])]
            }
        )
        column_in = pd.DataFrame(
            {
                '': ['In'] + [None if value == 0 else value for value in list(managerDict[sm][2])]
            }
        )
        column_out = pd.DataFrame(
            {
                '': ['Out'] + [None if value == 0 else value for value in list(managerDict[sm][3])]
            }
        )
        column_inv_end_teory = pd.DataFrame(
            {
                '': ['Inv.FT'] + [None if value == 0 else value for value in list(managerDict[sm][4])]
            }
        )
        column_inv_end_real = pd.DataFrame(
            {
                '': ['Inv.FR'] + [None if value == 0 else value for value in list(managerDict[sm][5])]
            }
        )
        column_desf = pd.DataFrame(
            {
                '': ['Desf.Inv'] + [None if value == 0 else value for value in list(managerDict[sm][6])]
            }
        )
        column_cost = pd.DataFrame(
            {
                only_one_decimal(totalManager[sm]['Cost.V'], symbol=True): ['Cost.V.'] + [None if value == 0 else value for value in list(managerDict[sm][7])]
            }
        )

        column_sell = pd.DataFrame(
            {
                only_one_decimal(totalManager[sm]['Ing.V'], symbol=True): ['Ing.V'] + [None if value == 0 else value for value in list(managerDict[sm][8])]
            }
        )
        column_result = pd.DataFrame(
            {
                only_one_decimal(totalManager[sm]['Result'], symbol=True): ['Result'] + [None if value == 0 else value for value in list(managerDict[sm][9])]
            }
        )
        column_merc = pd.DataFrame(
            {
                only_one_decimal(totalManager[sm]['Merc.V'], symbol=True): ['Merc.V'] + [None if value == 0 else value for value in list(managerDict[sm][10])]
            }
        )

        column_desc_inv = pd.DataFrame(
            {
                only_one_decimal(totalManager[sm]['Desc.DI'], symbol=True): ['Desc.DI'] + [None if value == 0 else value for value in list(managerDict[sm][11])]
            }
        )
        column_amount_buy = pd.DataFrame(
            {
                'Cant.C': ['Cant.C'] + [None if value == 0 else value for value in list(managerDict[sm][12])]
            }
        )
        column_buy = pd.DataFrame(
            {
                only_one_decimal(totalManager[sm]['Egr.C'], symbol=True): ['Egr.C'] + [None if value == 0 else value for value in list(managerDict[sm][13])]
            }
        )

        consolidate = consolidate + [column_amount_sell, column_sell, column_amount_buy, column_buy,
                                     column_inv_init_real, column_in, column_out, column_inv_end_teory,
                                     column_inv_end_real, column_desf, column_cost, column_result, column_merc,
                                     column_desc_inv]
        values_sm['Cost.V'].append(totalManager[sm]['Cost.V'])
        values_sm['Ing.V'].append(totalManager[sm]['Ing.V'])
        values_sm['Egr.C'].append(totalManager[sm]['Egr.C'])
        values_sm['Result'].append(totalManager[sm]['Result'])
        values_sm['Merc.V'].append(totalManager[sm]['Merc.V'])
        values_sm['Desc.DI'].append(totalManager[sm]['Desc.DI'])

        #     columns of sheet two
        if type(total_amount) == int:
            total_amount = managerDict[sm][0]
            total_inv_sem_ant = managerDict[sm][1]
            total_inv_teorico = managerDict[sm][4]
            total_inv_real = managerDict[sm][5]
            total_desfasaje = managerDict[sm][6]
            total_cost_v = managerDict[sm][7]
            total_venta = managerDict[sm][8]
            total_result = managerDict[sm][9]
            total_merc_v = managerDict[sm][10]
            total_desc_inv = managerDict[sm][11]
            total_amount_buy = managerDict[sm][12]
            total_egress_buy = managerDict[sm][13]
        else:
            total_amount += managerDict[sm][0]
            total_inv_sem_ant += managerDict[sm][1]
            total_inv_teorico += managerDict[sm][4]
            total_inv_real += managerDict[sm][5]
            total_desfasaje += managerDict[sm][6]
            total_cost_v += managerDict[sm][7]
            total_venta += managerDict[sm][8]
            total_result += managerDict[sm][9]
            total_merc_v += managerDict[sm][10]
            total_desc_inv += managerDict[sm][11]
            total_amount_buy += managerDict[sm][12]
            total_egress_buy += managerDict[sm][13]

    sheet_frame_1 = pd.concat(consolidate, axis=1)
    # result.to_excel("Consolidate2.xlsx", index=False)
    sheet_frame_1.to_excel(writer, index=False, sheet_name='Total Prod-Centro')
    column_RESUMEN_sheet_2 = pd.DataFrame({'RESUMEN': ['ID'] + products_id})
    column_CantV_sheet_2 = pd.DataFrame({'': ['Cant.V'] + [None if value == 0 else value for value in list(total_amount)]})
    column_CantC_sheet_2 = pd.DataFrame({'': ['Cant.C'] + [None if value == 0 else value for value in list(total_amount_buy)]})
    column_Inv_Sem_Ant_sheet_2 = pd.DataFrame({'': ['Inv.IR'] + [None if value == 0 else value for value in list(total_inv_sem_ant)]})
    column_Inv_Teorico_sheet_2 = pd.DataFrame({'': ['Inv.FT'] + [None if value == 0 else value for value in list(total_inv_teorico)]})
    column_Inv_Real_sheet_2 = pd.DataFrame({'': ['Inv.FR'] + [None if value == 0 else value for value in list(total_inv_real)]})
    column_Desfasaje_sheet_2 = pd.DataFrame(
        {
            only_one_decimal(sum(total_desfasaje)): ['Desf.Inv'] + [None if value == 0 else value for value in list(total_desfasaje)]})
    column_CostV_sheet_2 = pd.DataFrame(
        {
            only_one_decimal(sum(total_cost_v)): ['Cost.V.'] + [None if value == 0 else value for value in list(total_cost_v)]})
    column_Venta_sheet_2 = pd.DataFrame(
        {only_one_decimal(sum(total_venta)): ['Ing.V'] + [None if value == 0 else value for value in list(total_venta)]})
    column_Compra_sheet_2 = pd.DataFrame(
        {only_one_decimal(sum(total_egress_buy)): ['Egr.C'] + [None if value == 0 else value for value in list(total_egress_buy)]})
    column_Result_sheet_2 = pd.DataFrame(
        {only_one_decimal(sum(total_result)): ['Result'] + [None if value == 0 else value for value in list(total_result)]})
    column_Merc_V_sheet_2 = pd.DataFrame(
        {only_one_decimal(sum(total_merc_v)): ['Merc.V.'] + [None if value == 0 else value for value in list(total_merc_v)]})
    column_Desf_Inv_sheet_2 = pd.DataFrame(
        {only_one_decimal(sum(total_desc_inv)): ['Desc.DI'] + [None if value == 0 else value for value in list(total_desc_inv)]})

    sheet_frame_2 = pd.concat(
        [column_RESUMEN_sheet_2, column_CantV_sheet_2, column_CantC_sheet_2, column_Inv_Sem_Ant_sheet_2,
         column_Inv_Teorico_sheet_2,
         column_Inv_Real_sheet_2, column_Desfasaje_sheet_2, column_CostV_sheet_2, column_Venta_sheet_2,
         column_Compra_sheet_2,
         column_Result_sheet_2, column_Merc_V_sheet_2, column_Desf_Inv_sheet_2], axis=1)
    sheet_frame_2.to_excel(writer, index=False, sheet_name='Total Prod.')
    # three sheet
    temp_result = only_one_decimal(sum(values_sm['Result']))
    temp_desfInv = only_one_decimal(sum(values_sm['Desc.DI']))
    temp_desfEfectivo = only_one_decimal(sum(desf_cash))

    sheet_frame_4 = pd.DataFrame({
        'RESUMEN': ['Cent. Contable'] + sellManagersNames,
        'Punto de Venta': ['Vendedor'] + list([x if x > 0 else 0 for x in point_of_sale['Vendedor']]),
        'Asistente': ['Asistente'] + list([x if x > 0 else 0 for x in point_of_sale['Asistente']]),
        'Dueno': ['Dueño'] + list([x if x > 0 else 0 for x in point_of_sale['Dueno']]),
        'Punto Administración y Logística': ['AE'] + list([x if x > 0 else 0 for x in point_admin_logistic['AE']]),
        ' ': ['SM'] + list([x if x > 0 else 0 for x in point_admin_logistic['SM']]),
        '  ': ['CD'] + list([x if x > 0 else 0 for x in point_admin_logistic['CD']]),
        '   ': ['LCA'] + list([x if x > 0 else 0 for x in point_admin_logistic['LCA']]),
        '    ': ['LTD'] + list([x if x > 0 else 0 for x in point_admin_logistic['LTD']]),
        '     ': ['CEO'] + list([x if x > 0 else 0 for x in point_admin_logistic['CEO']])
    })
    temp_sheet_4 = pd.DataFrame({
        'Total point_of_sale': sheet_frame_4['Punto de Venta'][1:] + sheet_frame_4['Asistente'][1:] + sheet_frame_4[
                                                                                                          'Dueno'][1:],
        'Total point_admin_logistic': sheet_frame_4['Punto Administración y Logística'][1:] + sheet_frame_4[' '][1:] +
                                      sheet_frame_4['  '][1:] + sheet_frame_4['   '][1:] + sheet_frame_4['    '][1:] +
                                      sheet_frame_4['     '][1:],
        'Result Clean': values_sm['Result'],
        'Desf.Inv Clean': [x if x < 0 else 0 for x in values_sm['Desc.DI']],
        'Desf.Cash Clean': [x if x < 0 else 0 for x in desf_cash]
    })

    column_RESUMEN_sheet_3 = pd.DataFrame({'RESUMEN': ['Cent.Contable'] + sellManagersNames})
    column_CostV_sheet_3 = pd.DataFrame(
        {only_one_decimal(sum(values_sm['Cost.V'])): ['Cost.V'] + [None if value == 0 else value for value in values_sm['Cost.V']]})
    column_Venta_sheet_3 = pd.DataFrame(
        {only_one_decimal(sum(values_sm['Ing.V'])): ['Ing.V'] + [None if value == 0 else value for value in values_sm['Ing.V']]})
    column_Compra_sheet_3 = pd.DataFrame(
        {only_one_decimal(sum(buy_egress)): ['Egr.C'] + [None if value == 0 else value for value in buy_egress]})
    column_Result_sheet_3 = pd.DataFrame({temp_result: ['Result'] + [None if value == 0 else value for value in values_sm['Result']]})
    column_MercV_sheet_3 = pd.DataFrame(
        {only_one_decimal(sum(values_sm['Merc.V'])): ['Merc.V'] + [None if value == 0 else value for value in values_sm['Merc.V']]})
    column_Desf_Inv_sheet_3 = pd.DataFrame({temp_desfInv + '      ': ['Desc.DI'] + [None if value == 0 else value for value in values_sm['Desc.DI']]})
    column_Other_Ing_sheet_3 = pd.DataFrame(
        {only_one_decimal(sum(other_income)): ['Otros Ingr.'] + [None if value == 0 else value for value in other_income]})
    column_Other_Egr_sheet_3 = pd.DataFrame(
        {only_one_decimal(sum(other_egress)): ['Otros Egr'] + [None if value == 0 else value for value in other_egress]})
    column_Flujo_sheet_3 = pd.DataFrame(
        {only_one_decimal(sum(net_flow)): ['Flujo Neto'] + [None if value == 0 else value for value in net_flow]})
    column_Real_Cash_sheet_3 = pd.DataFrame(
        {only_one_decimal(sum(real_cash)): ['Efect. Real'] + [None if value == 0 else value for value in real_cash]})

    final_score = temp_sheet_4['Result Clean'] + temp_sheet_4['Desf.Inv Clean'] + temp_sheet_4['Desf.Cash Clean']
    column_Final_sheet_3 = pd.DataFrame({'': ['Final Score'] + list([x if x > 0 else None for x in final_score])})

    column_PointV_sheet_3 = pd.DataFrame({only_one_decimal(sum(temp_sheet_4['Total point_of_sale'])): ['Punto Venta'] + [None if value == 0 else value for value in temp_sheet_4['Total point_of_sale']]})
    column_Admin_sheet_3 = pd.DataFrame({only_one_decimal(sum(temp_sheet_4['Total point_admin_logistic'])): ['Admin y Logistica'] + [None if value == 0 else value for value in temp_sheet_4['Total point_admin_logistic']]})

    dividends = temp_sheet_4['Total point_admin_logistic'] + temp_sheet_4['Total point_of_sale']
    column_Dividen_sheet_3 = pd.DataFrame(
        {only_one_decimal(sum(dividends)): ['Dividendos'] + [None if value == 0 else value for value in dividends]})

    temp_k = [x if x > 0 else 0 for x in final_score] - dividends
    column_K_Retenido_sheet_3 = pd.DataFrame({only_one_decimal(sum([x if x > 0 else 0 for x in final_score] - dividends)): ['K-Retenido'] + list(
        [None if value == 0 else value for value in temp_k])})

    sheet_frame_3 = pd.concat(
        [column_RESUMEN_sheet_3, column_CostV_sheet_3, column_Venta_sheet_3, column_Compra_sheet_3,
         column_Result_sheet_3,
         column_MercV_sheet_3, column_Desf_Inv_sheet_3, column_Other_Ing_sheet_3, column_Other_Egr_sheet_3,
         column_Flujo_sheet_3,
         column_Real_Cash_sheet_3, column_Final_sheet_3, column_PointV_sheet_3, column_Admin_sheet_3,
         column_Dividen_sheet_3,
         column_K_Retenido_sheet_3], axis=1)


    sheet_frame_3.to_excel(writer, index=False, sheet_name='Total por Centros')
    sheet_frame_4.to_excel(writer, index=False, sheet_name='Total Salario-Centro')

    names_temp = []
    id_temp = []
    cc_temp = []
    ir_temp = []
    ft_temp = []
    fr_temp = []

    for seller in seller_account.keys():
        names_temp += seller_account[seller]['name']
        id_temp += seller_account[seller]['id']
        cc_temp += seller_account[seller]['cc']
        ir_temp += seller_account[seller]['IR']
        ft_temp += seller_account[seller]['FT']
        fr_temp += seller_account[seller]['FR']

    sheet_5_column_name = pd.DataFrame(
        {
            'RESUMEN': ['Nombre'] + names_temp
        }
    )
    sheet_5_column_id = pd.DataFrame(
        {
            '': ['ID'] + [element for element in id_temp]
        }
    )
    sheet_5_column_cc = pd.DataFrame(
        {
            '': ['Centro Contable'] + [element for element in cc_temp]
        }
    )
    total_ir = 0
    total_ft = 0
    total_fr = 0
    for i in range(len(ir_temp)):
        if not pd.isna(ir_temp[i]):
            total_ir += ir_temp[i]
        if not pd.isna(ft_temp[i]):
            total_ft += ft_temp[i]
        if not pd.isna(fr_temp[i]):
            total_fr += fr_temp[i]

    sheet_5_column_ir = pd.DataFrame(
        {
            only_one_decimal(total_ir): ['Est.IR'] + ir_temp
        }
    )
    sheet_5_column_ft = pd.DataFrame(
        {
            only_one_decimal(total_ft): ['Est.FT'] + ft_temp
        }
    )
    sheet_5_column_fr = pd.DataFrame(
        {
            only_one_decimal(total_fr): ['Est.FR'] + fr_temp
        }
    )

    sheet_frame_5 = pd.concat(
        [sheet_5_column_name, sheet_5_column_id, sheet_5_column_cc, sheet_5_column_ir, sheet_5_column_ft,
         sheet_5_column_fr], axis=1)

    sheet_frame_5.to_excel(writer, index=False, sheet_name='Total Cuenta-Centro')

    list_name_total = []
    list_ir_total = []
    list_ft_total = []
    list_fr_total = []

    # id_unique_account[id_temp[i]] = {'name': name_temp[i], 'IR': ir_temp[i], 'FT': ft_temp[i], 'FR': fr_temp[i],
    #                                  'repeat': False, 'color': 'none'}

    for id in id_unique_account.keys():
        list_name_total.append(id_unique_account[id]['name'])
        list_ir_total.append(id_unique_account[id]['IR'])
        list_ft_total.append(id_unique_account[id]['FT'])
        list_fr_total.append(id_unique_account[id]['FR'])

    total_ir = 0
    total_ft = 0
    total_fr = 0
    for i in range(len(list_ir_total)):
        if not pd.isna(list_ir_total[i]):
            total_ir += list_ir_total[i]
        if not pd.isna(list_ft_total[i]):
            total_ft += list_ft_total[i]
        if not pd.isna(list_fr_total[i]):
            total_fr += list_fr_total[i]

    sheet_6_column_name = pd.DataFrame(
        {
            'RESUMEN': ['Nombre'] + list_name_total
        }
    )

    sheet_6_column_id = pd.DataFrame(
        {
            '': ['ID'] + list(id_unique_account.keys())
        }
    )

    sheet_6_column_ir = pd.DataFrame(
        {
            only_one_decimal(total_ir): ['Est.IR'] + list_ir_total
        }
    )
    sheet_6_column_ft = pd.DataFrame(
        {
            only_one_decimal(total_ft): ['Est.FT'] + list_ft_total
        }
    )
    sheet_6_column_fr = pd.DataFrame(
        {
            only_one_decimal(total_fr): ['Est.FR'] + list_fr_total
        }
    )

    sheet_frame_6 = pd.concat(
        [sheet_6_column_name, sheet_6_column_id, sheet_6_column_ir, sheet_6_column_ft,
         sheet_6_column_fr], axis=1)

    sheet_frame_6.to_excel(writer, index=False, sheet_name='Total Cuenta')


    # Developing sheet 7: Types Accounts for files
    sheet_7_column_type = pd.DataFrame(
        {
            'RESUMEN': ['Cuentas'] + list(types_account[sellManagersNames[0]]['type'])
        }
    )
    consolidate = [sheet_7_column_type]
    for sm in sellManagersNames:


        sheet_7_column_ir = pd.DataFrame(
            {
                sm: ['Est.IR'] + list(types_account[sm]['ir'])
            }
        )
        sheet_7_column_ft = pd.DataFrame(
            {
                '': ['Est.FT'] + list(types_account[sm]['ft'])
            }
        )
        sheet_7_column_fr = pd.DataFrame(
            {
                '': ['Est.FR'] + list(types_account[sm]['fr'])
            }
        )

        consolidate = consolidate + [sheet_7_column_ir, sheet_7_column_ft, sheet_7_column_fr]
    sheet_frame_7 = pd.concat(consolidate, axis=1)

    sheet_frame_7.to_excel(writer, index=False, sheet_name='Total Grupo Cuenta-Centro')

    # Developing sheet 8: Total Accounts Group
    sheet_8_column_type = pd.DataFrame(
        {
            'RESUMEN': ['Cuentas'] + list(types_account[sellManagersNames[0]]['type'])
        }
    )

    sheet_8_column_ir = pd.DataFrame(
        {
            '': ['Est.IR'] + list(total_ir_group)
        }
    )
    sheet_8_column_ft = pd.DataFrame(
        {
            '': ['Est.FT'] + list(total_ft_group)
        }
    )
    sheet_8_column_fr = pd.DataFrame(
        {
            '': ['Est.FR'] + list(total_fr_group)
        }
    )

    consolidate = [sheet_8_column_type, sheet_8_column_ir, sheet_8_column_ft, sheet_8_column_fr]
    sheet_frame_8 = pd.concat(consolidate, axis=1)

    sheet_frame_8.to_excel(writer, index=False, sheet_name='Total Grupo de Cuentas')

    # Edditing Sheets:

    workbook = writer.book
    worksheet_1 = writer.sheets['Total Prod-Centro']
    worksheet_2 = writer.sheets['Total Prod.']
    worksheet_3 = writer.sheets['Total por Centros']
    worksheet_4 = writer.sheets['Total Salario-Centro']
    worksheet_5 = writer.sheets['Total Cuenta-Centro']
    worksheet_6 = writer.sheets['Total Cuenta']
    worksheet_7 = writer.sheets['Total Grupo Cuenta-Centro']
    worksheet_8 = writer.sheets['Total Grupo de Cuentas']

    worksheet_1.freeze_panes(2, 1)
    worksheet_2.freeze_panes(2, 0)
    worksheet_3.freeze_panes(2, 1)
    worksheet_4.freeze_panes(2, 1)
    worksheet_5.freeze_panes(2, 0)
    worksheet_6.freeze_panes(2, 0)
    worksheet_7.freeze_panes(0, 1)

    header_format = workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#00C85A',
            'bold': True,
            'font_color': '#FFFFFF',
            'num_format': '#,#0.0'
        }
    )
    money_fmt = workbook.add_format({'num_format': '#,#0.0'})

    star = 1
    end = 8
    color = '#00C85A'

    # Game of Queen
    for col_num, value in enumerate(sheet_frame_1.columns.values):

        if col_num == star:
            if color == '#00642D' or col_num == 1:
                color = '#00C85A'
            else:
                color = '#00642D'
            worksheet_1.merge_range(0, star, 0, end, value, workbook.add_format(
                {
                    'valign': 'vcenter',
                    'align': 'center',
                    'bg_color': color,
                    'bold': True,
                    'font_color': '#FFFFFF',
                    'font_size': 13.2,
                    # 'num_format': '#,#0.0'
                }))
            star = end + 7
            end = star + 7

        elif col_num >= end - 13:
            worksheet_1.write(0, col_num, value, workbook.add_format(
                {
                    'valign': 'vcenter',
                    'align': 'center',
                    'bg_color': color,
                    'bold': True,
                    'font_color': '#FFFFFF',
                    # 'num_format': '#,#0.0'
                }))

    color = '#D9D9D9'
    star = 2
    array_frame_5 = sheet_frame_5.to_numpy()
    for seller in seller_account.keys():
        count_account = len(seller_account[seller]['id'])
        end = count_account + star

        for i in range(star, end):
            for j in range(2, 6):
                value = array_frame_5[i - 1, j]
                if pd.isna(value):
                    value = None
                worksheet_5.write(i, j, value, workbook.add_format(
                    {
                        'valign': 'vcenter',
                        'bg_color': color,
                        'font_color': '#000000',
                        'font_size': 11,
                        'num_format': '#,#0.0'
                    }))
        if star + 1 < end:
            worksheet_5.merge_range(star, 2, end - 1, 2, seller, workbook.add_format(
                {
                    'valign': 'vcenter',
                    'align': 'center',
                    'bg_color': color,
                    'font_color': '#000000',
                    'font_size': 11,
                    'num_format': '#,#0.0'
                }))

        star = end

        if color == '#D9D9D9':
            color = '#B7DEE8'
        else:
            color = '#D9D9D9'

    star = 1
    color = '#00C85A'
    # Set color to first cell (Resumen)
    array_frame_7 = sheet_frame_7.to_numpy()
    worksheet_7.write(0, 0, 'RESUMEN', workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': color,
            'bold': True,
            'font_color': '#FFFFFF',
            'num_format': '#,#0.0'
        }))

    for col_num, value in enumerate(sheet_frame_7.columns.values):

        if col_num == star:
            if color == '#00642D' or col_num == 1:
                color = '#00C85A'
            else:
                color = '#00642D'
            worksheet_7.merge_range(0, star, 0, star+2, value, workbook.add_format(
                {
                    'valign': 'vcenter',
                    'align': 'center',
                    'bg_color': color,
                    'bold': True,
                    'font_color': '#FFFFFF',
                    'font_size': 13.2,
                    'num_format': '#,#0.0'
                }))
            star += 3


    # Editing Main Menu
    for col_num, value in enumerate(sheet_frame_2.columns.values):
        worksheet_2.write(0, col_num, value, header_format)
    for col_num, value in enumerate(sheet_frame_3.columns.values):
        worksheet_3.write(0, col_num, value, header_format)
    for col_num, value in enumerate(sheet_frame_4.columns.values):
        if col_num == 1:
            worksheet_4.merge_range(0, 1, 0, 3, value, workbook.add_format(
                {
                    'valign': 'vcenter',
                    'align': 'center',
                    'bg_color': '#00C85A',
                    'bold': True,
                    'font_color': '#FFFFFF',
                    'font_size': 11,
                    'left': 1,
                    'right': 1,
                    'num_format': '#,#0.0'
                }))
        elif col_num == 4:
            worksheet_4.merge_range(0, 4, 0, 9, value, workbook.add_format(
                {
                    'valign': 'vcenter',
                    'align': 'center',
                    'bg_color': '#00C85A',
                    'bold': True,
                    'font_color': '#FFFFFF',
                    'font_size': 11,
                    'right': 1,
                    'num_format': '#,#0.0'
                }))
        elif col_num == 0:
            worksheet_4.write(0, col_num, value, header_format)

        if col_num != 0:
            worksheet_4.write(1, col_num, sheet_frame_4[sheet_frame_4.columns[col_num]][0], workbook.add_format(
                {
                    'valign': 'vcenter',
                    'align': 'center',
                    'bg_color': '#000000',
                    'bold': True,
                    'font_color': '#FFFF00',
                    'num_format': '#,#0.0'

                }))
        else:
            worksheet_4.write(1, col_num, sheet_frame_4[sheet_frame_4.columns[col_num]][0], workbook.add_format(
                {
                    'valign': 'vcenter',
                    'align': 'center',
                    'bg_color': '#5B9BD5',
                    'bold': True,
                    'font_color': '#000000',
                    'num_format': '#,#0.0'

                }))
    for col_num, value in enumerate(sheet_frame_5.columns.values):
        worksheet_5.write(0, col_num, value, header_format)
    for col_num, value in enumerate(sheet_frame_6.columns.values):
        worksheet_6.write(0, col_num, value, header_format)
    # Editing Sub-Menu of sheet 5 and 6 that realy is sheet number three and four
    array_frame_6 = sheet_frame_6.to_numpy()
    for i in range(6):
        worksheet_5.write(1, i, array_frame_5[0, i], workbook.add_format({
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#548235',
            'bold': True,
            'font_color': '#FFFFFF',
            'num_format': '#,#0.0'
        }))
    for i in range(5):
        worksheet_6.write(1, i, array_frame_6[0, i], workbook.add_format({
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#548235',
            'bold': True,
            'font_color': '#FFFFFF',
            'num_format': '#,#0.0'
        }))


    format_7 = workbook.add_format({
                'bg_color': '#E2E2E2',
                'font_color': '#000000',
                'left': 1,
                'right': 1,
                'top': 1,
                'bottom': 1,
                'border_color': '#C4C4C4',
                'num_format': '#,#0.0'
            })

    for i in range(1, len(sellManagersNames), 2):
        for j in range(1, 11):      # Amount Account
            worksheet_7.write(j + 1, i * 3 + 1, array_frame_7[j, i*3+1], format_7)
            worksheet_7.write(j + 1, i * 3 + 2, array_frame_7[j, i*3+2], format_7)
            worksheet_7.write(j + 1, i * 3 + 3, array_frame_7[j, i*3+3], format_7)

    # Sheet number: 8
    for col_num, value in enumerate(sheet_frame_8.columns.values):
        worksheet_8.write(0, col_num, value, header_format)

    #Sub Menu
    for i in range(len(sellManagersNames)*3 + 1):
        worksheet_7.write(1, i, array_frame_7[0, i], workbook.add_format({
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#548235',
            'font_color': '#FFFFFF',
            'bold': True,


        }))
    array_frame_8 = sheet_frame_8.to_numpy()
    for i in range(4):
        worksheet_8.write(1, i, array_frame_8[0, i], workbook.add_format({
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#548235',
            'font_color': '#FFFFFF',
            'bold': True,
        }))

    # Styling sub-menu
    worksheet_1.write(1, 0, sheet_frame_1[sheet_frame_1.columns[0]][0], workbook.add_format({
        'valign': 'vcenter',
        'align': 'center',
        'bg_color': '#FFFFE7',
        'bold': True,
        'font_color': '#000000',
        'num_format': '#,#0.0'
    }))
    array_frame_1 = sheet_frame_1.to_numpy()

    count = 1
    for i in range(len(sellManagersNames)):
        worksheet_1.write(1, count, array_frame_1[0][count], workbook.add_format(
            {
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#800000',
                'bold': True,
                'font_color': '#FFFFFF',
                'num_format': '#,#0.0'
            })
                          )
        count += 1
        worksheet_1.write(1, count, array_frame_1[0][count], workbook.add_format(
            {
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#005828',
                'bold': True,
                'font_color': '#FFFFFF',
                'num_format': '#,#0.0'
            })
                          )
        count += 1
        worksheet_1.write(1, count, array_frame_1[0][count], workbook.add_format(
            {
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#595959',
                'bold': True,
                'font_color': '#FFFFFF',
                'num_format': '#,#0.0'
            })
                          )
        count += 1
        worksheet_1.write(1, count, array_frame_1[0][count], workbook.add_format(
            {
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#3A3838',
                'bold': True,
                'font_color': '#FFFFFF',
                'num_format': '#,#0.0'
            })
                          )
        count += 1
        worksheet_1.write(1, count, array_frame_1[0][count], workbook.add_format(
            {
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#002060',
                'bold': True,
                'font_color': '#FFFFFF',
                'num_format': '#,#0.0'
            })
                          )
        count += 1
        worksheet_1.write(1, count, array_frame_1[0][count], workbook.add_format(
            {
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#305496',
                'bold': True,
                'font_color': '#FFFFFF',
                'num_format': '#,#0.0'
            })
                          )
        count += 1
        worksheet_1.write(1, count, array_frame_1[0][count], workbook.add_format(
            {
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#9E4F00',
                'bold': True,
                'font_color': '#FFFFFF',
                'num_format': '#,#0.0'
            })
                          )
        count += 1
        worksheet_1.write(1, count, array_frame_1[0][count], workbook.add_format(
            {
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#002060',
                'bold': True,
                'font_color': '#FFFFFF',
                'num_format': '#,#0.0'
            })
                          )


        count += 1
        differ_than(array_frame_1, 2, 201, count, count-1, worksheet_1, workbook)  # Real Inventary
        worksheet_1.write(1, count, array_frame_1[0][count], workbook.add_format(
            {
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#002060',
                'bold': True,
                'font_color': '#FFFFFF',
                'num_format': '#,#0.0'
            })
                          )

        count += 1
        positive_and_negative_numbers_print(array_frame_1, 2, 201, count, worksheet_1, workbook)  # Inventory Mismatch
        worksheet_1.write(1, count, array_frame_1[0][count], workbook.add_format(
            {
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#000000',
                'bold': True,
                'font_color': '#FFFFFF',
                'num_format': '#,#0.0'
            })
                          )
        count += 1
        worksheet_1.write(1, count, array_frame_1[0][count], workbook.add_format(
            {
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#000000',
                'bold': True,
                'font_color': '#FFFFFF',
                'num_format': '#,#0.0'
            })
                          )
        count += 1
        worksheet_1.write(1, count, array_frame_1[0][count], workbook.add_format(
            {
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#000000',
                'bold': True,
                'font_color': '#FFFFFF',
                'num_format': '#,#0.0'
            })
                          )
        count += 1
        worksheet_1.write(1, count, array_frame_1[0][count], workbook.add_format(
            {
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#000000',
                'bold': True,
                'font_color': '#FFFFFF',
                'num_format': '#,#0.0'
            })
                          )
        count += 1
        positive_and_negative_numbers_print(array_frame_1, 2, 201, count, worksheet_1, workbook)  # Inventory Mismatch
        worksheet_1.write(1, count, array_frame_1[0][count], workbook.add_format(
            {
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#000000',
                'bold': True,
                'font_color': '#FFFFFF',
                'num_format': '#,#0.0'
            })
                          )
        # negative_numbers_print(array_frame_1, 2, 201, count, worksheet_1, workbook)
        count += 1

    set_color_column(worksheet_1, 1, 0, sheet_frame_1.to_numpy(), '#FFFFF7')
    set_color_column(worksheet_2, 1, 0, sheet_frame_2.to_numpy(), '#FFFFF7')
    set_color_column(worksheet_3, 1, 0, sheet_frame_3.to_numpy(), '#88B6E0')
    set_color_column(worksheet_4, 1, 0, sheet_frame_4.to_numpy(), '#88B6E0')
    set_color_column(worksheet_7, 1, 0, sheet_frame_7.to_numpy(), '#D0FBCD')
    set_color_column(worksheet_8, 1, 0, sheet_frame_8.to_numpy(), '#D0FBCD')

    worksheet_1.set_column('A:AZZ', 12, money_fmt)
    worksheet_2.set_column('A:AZZ', 12, money_fmt)
    worksheet_3.set_column('A:AZZ', 14, money_fmt)
    worksheet_4.set_column('A:AZZ', 14, money_fmt)
    worksheet_5.set_column('A:F', 16, money_fmt)
    worksheet_6.set_column('A:E', 16, money_fmt)
    worksheet_7.set_column('A:A', 18, money_fmt)
    worksheet_7.set_column('B:AZ', 14, money_fmt)
    worksheet_8.set_column('A:A', 18, money_fmt)
    worksheet_8.set_column('B:E', 14, money_fmt)

    worksheet_2.write(1, 0, sheet_frame_2[sheet_frame_2.columns[0]][0], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#FFFFE7',
            'bold': True,
            'font_color': '#000000',
            'num_format': '#,#0.0'
        }))

    # Estilo de la segunda hoja.. sub menu
    array_frame_2 = sheet_frame_2.to_numpy()
    worksheet_2.write(1, 1, array_frame_2[0][1], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#800000',
            'bold': True,
            'font_color': '#FFFFFF',
            'num_format': '#,#0.0'
        })
                      )
    worksheet_2.write(1, 2, array_frame_2[0][2], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#595959',
            'bold': True,
            'font_color': '#FFFFFF',
            'num_format': '#,#0.0'
        })
                      )
    worksheet_2.write(1, 3, array_frame_2[0][3], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#002060',
            'bold': True,
            'font_color': '#FFFFFF',
            'num_format': '#,#0.0'
        })
                      )
    worksheet_2.write(1, 4, array_frame_2[0][4], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#002060',
            'bold': True,
            'font_color': '#FFFFFF',
            'num_format': '#,#0.0'
        })
                      )
    differ_than(array_frame_2, 2, 201, 5, 4, worksheet_2, workbook)
    positive_and_negative_numbers_print(array_frame_2, 2, 201, 6, worksheet_2, workbook)
    worksheet_2.write(1, 5, array_frame_2[0][5], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#002060',
            'bold': True,
            'font_color': '#FFFFFF',
            'num_format': '#,#0.0'
        })
                      )
    # positive_and_negative_numbers_print(array_frame_2, 2, 201, 5, worksheet_2, workbook)

    worksheet_2.write(1, 6, array_frame_2[0][6], workbook.add_format({
        'valign': 'vcenter',
        'align': 'center',
        'bg_color': '#000000',
        'bold': True,
        'font_color': '#FFFFFF',
        'num_format': '#,#0.0'
    }))
    worksheet_2.write(1, 7, array_frame_2[0][7], workbook.add_format({
        'valign': 'vcenter',
        'align': 'center',
        'bg_color': '#000000',
        'bold': True,
        'font_color': '#FFFFFF',
        'num_format': '#,#0.0'
    }))
    worksheet_2.write(1, 8, array_frame_2[0][8], workbook.add_format({
        'valign': 'vcenter',
        'align': 'center',
        'bg_color': '#000000',
        'bold': True,
        'font_color': '#FFFFFF',
        'num_format': '#,#0.0'
    }))
    worksheet_2.write(1, 9, array_frame_2[0][9], workbook.add_format({
        'valign': 'vcenter',
        'align': 'center',
        'bg_color': '#000000',
        'bold': True,
        'font_color': '#FFFFFF',
        'num_format': '#,#0.0'
    }))
    worksheet_2.write(1, 10, array_frame_2[0][10], workbook.add_format({
        'valign': 'vcenter',
        'align': 'center',
        'bg_color': '#000000',
        'bold': True,
        'font_color': '#FFFFFF',
        'num_format': '#,#0.0'
    }))
    worksheet_2.write(1, 11, array_frame_2[0][11], workbook.add_format({
        'valign': 'vcenter',
        'align': 'center',
        'bg_color': '#000000',
        'bold': True,
        'font_color': '#FFFFFF',
        'num_format': '#,#0.0'
    }))
    negative_numbers_print(array_frame_2, 2, 201, 12, worksheet_2, workbook)
    worksheet_2.write(1, 12, array_frame_2[0][12], workbook.add_format({
        'valign': 'vcenter',
        'align': 'center',
        'bg_color': '#000000',
        'bold': True,
        'font_color': '#FFFFFF',
        'num_format': '#,#0.0'
    }))
    # negative_numbers_print(array_frame_2, 2, 201, 10, worksheet_2, workbook)

    # Estilo de la tercera hoja
    array_frame_3 = sheet_frame_3.to_numpy()
    worksheet_3.write(1, 0, array_frame_3[0][0], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#5B9BD5',
            'bold': True,
            'font_color': '#000000',
            'num_format': '#,#0.0'
        })
                      )
    worksheet_3.write(1, 1, array_frame_3[0][1], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#000000',
            'bold': True,
            'font_color': '#FFFFFF',
            'num_format': '#,#0.0'
        })
                      )
    worksheet_3.write(1, 2, array_frame_3[0][2], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#000000',
            'bold': True,
            'font_color': '#FFFFFF',
            'num_format': '#,#0.0'
        })
                      )
    worksheet_3.write(1, 3, array_frame_3[0][3], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#000000',
            'bold': True,
            'font_color': '#FFFFFF',
            'num_format': '#,#0.0'
        })
                      )
    worksheet_3.write(1, 4, array_frame_3[0][4], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#000000',
            'bold': True,
            'font_color': '#FFFFFF',
            'num_format': '#,#0.0'
        })
                      )
    worksheet_3.write(1, 5, array_frame_3[0][5], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#000000',
            'bold': True,
            'font_color': '#FFFFFF',
            'num_format': '#,#0.0'
        })
                      )
    # negative_numbers_print(array_frame_3, 2, len(sellManagersNames) + 2, 5, worksheet_3, workbook)

    worksheet_3.write(1, 6, array_frame_3[0][6], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#C6E0B4',
            'bold': True,
            'font_color': '#000000',
            'num_format': '#,#0.0'
        })
                      )
    worksheet_3.write(1, 7, array_frame_3[0][7], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#AEAAAA',
            'bold': True,
            'font_color': '#000000',
            'num_format': '#,##0.00'
        })
                      )
    worksheet_3.write(1, 8, array_frame_3[0][8], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#FF0000',
            'bold': True,
            'font_color': '#000000',
            'num_format': '#,#0.0'
        })
                      )
    worksheet_3.write(1, 9, array_frame_3[0][9], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#FFC000',
            'bold': True,
            'font_color': '#000000',
            'num_format': '#,#0.0'
        })
                      )
    worksheet_3.write(1, 10, array_frame_3[0][10], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#FFFF00',
            'bold': True,
            'font_color': '#000000',
            'num_format': '#,#0.0'
        })
                      )
    # positive_and_negative_numbers_print(array_frame_3, 2, len(sellManagersNames) + 2, 10, worksheet_3, workbook)

    worksheet_3.write(1, 11, array_frame_3[0][11], workbook.add_format(
        {
            'valign': 'vcenter',
            'align': 'center',
            'bg_color': '#0EC1F2',
            'bold': True,
            'font_color': '#000000',
            'num_format': '#,#0.0'
        })
                      )
    for i in range(12, 16):
        worksheet_3.write(1, i, array_frame_3[0][i], workbook.add_format(
            {
                'valign': 'vcenter',
                'align': 'center',
                'bg_color': '#000000',
                'bold': True,
                'font_color': '#FFFF00',
                'num_format': '#,#0.0'
            })
                          )
    # Repeat colors for unique for repeated identifiers
    for id in id_unique_account.keys():
        if id_unique_account[id]['repeat']:
            for i in range(1, len(array_frame_5)):
                if array_frame_5[i, 1] == id:
                    worksheet_5.write(i + 1, 1, array_frame_5[i][1], workbook.add_format(
                        {
                            'valign': 'vcenter',
                            'align': 'center',
                            'bg_color': id_unique_account[id]['color'],
                            'bold': True,
                            'font_color': '#000000',
                            'num_format': '#,#0.0'
                        })
                                      )

    writer.save()
    messagebox.showinfo('Consolidate', 'Proceso Finalizado')
