from src.tables.res1 import Res1
from src.tools.worker_file import WorkerFile

ATTRIBUTES_GET = {'id': 1, 'Exist_Ini_Lu': 3, 'Cant.C': 4, 'IN_1': 5,
              'IN_2': 6, 'IN_3': 7, 'IN_4': 8, 'IN_5': 9, 'IN_6': 10
, 'IN_7': 11, 'IN_8': 12, 'IN_9': 13, 'IN_10': 14, 'IN_11': 15
, 'IN_12': 16, 'IN_13': 17, 'IN_14': 18, 'IN_15': 19, 'IN_16': 21
, 'IN_17': 22,
              'Cant.V': 24, 'OU_1': 25, 'OU_2': 26, 'OU_3': 27, 'OU_4': 28,
'OU_5': 29, 'OU_6': 30, 'OU_7': 31, 'OU_8': 32, 'OU_9': 33, 'OU_10': 34, 'OU_11': 35,
'OU_12': 36, 'OU_13': 37, 'OU_14': 38, 'OU_15': 39, 'OU_16': 41, 'OU_17': 42,
              'Exist.Real': 45, 'CB': 47, 'CN': 48, 'Venta': 50, 'Compra': 51, }

ATTRIBUTES_SET = {'id': 1, 'Exist_Ini_Lu': 3, 'Cant.C': 4, 'IN_1': 5,
              'IN_2': 6, 'IN_3': 7, 'IN_4': 8, 'IN_5': 9, 'IN_6': 10
, 'IN_7': 11, 'IN_8': 12, 'IN_9': 13, 'IN_10': 14, 'IN_11': 15
, 'IN_12': 16, 'IN_13': 17, 'IN_14': 18, 'IN_15': 19, 'IN_16': 21
, 'IN_17': 22,
              'Cant.V': 24, 'OU_1': 25, 'OU_2': 26, 'OU_3': 27, 'OU_4': 28,
'OU_5': 29, 'OU_6': 30, 'OU_7': 31, 'OU_8': 32, 'OU_9': 33, 'OU_10': 34, 'OU_11': 35,
'OU_12': 36, 'OU_13': 37, 'OU_14': 38, 'OU_15': 39, 'OU_16': 41, 'OU_17': 42,
              'Exist.Real': 45, 'Venta': 47, 'Compra': 48, 'CB': 50, 'CN': 51}


# sheet(Resumen)
class ResumeOneManage:

    @staticmethod
    def get_resume_one(worksheet, is_pivot=False):
        products_dictionary = {}
        max_row = worksheet.max_row + 1
        for row in range(3, max_row):
            if worksheet.cell(row=row, column=ATTRIBUTES_GET['id']).value is not None:
                # Conditional of Pivot ('aaa')
                # id = None
                # gross_cost = None
                # net_cost = None
                id = worksheet.cell(row=row, column=ATTRIBUTES_GET['id']).value
                init_existence = worksheet.cell(row=row, column=ATTRIBUTES_GET['Exist_Ini_Lu']).value
                buy_amount = worksheet.cell(row=row, column=ATTRIBUTES_GET['Cant.C']).value
                sale_amount = worksheet.cell(row=row, column=ATTRIBUTES_GET['Cant.V']).value
                in_list = [worksheet.cell(row=row, column=ATTRIBUTES_GET['IN_'+str(i)]).value for i in range(1, 18)]
                ou_list = [worksheet.cell(row=row, column=ATTRIBUTES_GET['OU_' + str(i)]).value for i in range(1, 18)]
                cb = worksheet.cell(row=row, column=ATTRIBUTES_GET['CB']).value
                cn = worksheet.cell(row=row, column=ATTRIBUTES_GET['CN']).value

                real_existence = worksheet.cell(row=row, column=ATTRIBUTES_GET['Exist.Real']).value
                sale = worksheet.cell(row=row, column=ATTRIBUTES_GET['Venta']).value
                purchase = worksheet.cell(row=row, column=ATTRIBUTES_GET['Compra']).value


                    # print('cb is ', str(cb), ' tipo:',type(cb))
                products_dictionary[id] = Res1(id, init_existence, buy_amount, sale_amount, in_list, ou_list, real_existence, sale, purchase, cb, cn)
            # else:
                # products_dictionary['id'] = Product(None, None, None, None, [], [], None, None, None)
        return products_dictionary

    @staticmethod
    def set_resume_one(worksheet, worker_report: WorkerFile, is_pivot=False):
        products = sorted(worker_report.get_res_1().keys())
        counter_row = 3
        for key in products:
            product = worker_report.get_res_1()[key]
            # product = product.value
            # if is_pivot:
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['id'], value=product.get_id())
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['Exist_Ini_Lu'], value=product.get_init_existence())
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['Cant.C'], value=product.get_buy_amount())
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['Cant.V'], value=product.get_sale_amount())
            for i in range(len(product.get_in_list())):
                worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['IN_' + str(i+1)], value=product.get_in_list()[i])
            for i in range(len(product.get_ou_list())):
                worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['OU_' + str(i+1)], value=product.get_ou_list()[i])

            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['Exist.Real'], value=product.get_real_existence())
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['Venta'], value=product.get_sale())
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['Compra'], value=product.get_purchase())

            # if is_pivot:
            # print(product.get_id(), ' - cb:', product.get_cb())
            # print(product.get_id(), ' - cn:', product.get_cn())
            # worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['CN'], value=product.get_cn())
            # worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['CB'], value=product.get_cb())

            counter_row += 1

