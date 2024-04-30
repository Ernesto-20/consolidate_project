from src.tables.res2 import Res2
from src.tools.worker_file import WorkerFile

ATTRIBUTES_SET = {'ID C': 1, 'Nombre Cuen': 3, 'Gr.': 4, 'SUMA_Deb': 5, 'SUMA_Hab': 6,
              'DEBE_Cierre': 13, 'HABER_Cierre': 14}
ATTRIBUTES_GET = {'ID C': 1, 'Nombre Cuen': 3, 'Gr.': 4, 'SUMA_Deb': 19, 'SUMA_Hab': 20,
              'DEBE_Cierre': 27, 'HABER_Cierre': 28}

class ResumeTwoManage:

    @staticmethod
    def get_resume_two(worksheet, is_pivot=False):
        balance_dictionary = []
        max_row = worksheet.max_row + 1
        for row in range(3, max_row+1):
            if worksheet.cell(row=row, column=ATTRIBUTES_GET['ID C']).value is not None:

                account_id = None
                account_name = None
                apgi = None

                if is_pivot:
                    account_id = worksheet.cell(row=row, column=ATTRIBUTES_GET['ID C']).value
                    account_name = worksheet.cell(row=row, column=ATTRIBUTES_GET['Nombre Cuen']).value
                    apgi = worksheet.cell(row=row, column=ATTRIBUTES_GET['Gr.']).value

                sum_deb = worksheet.cell(row=row, column=ATTRIBUTES_GET['SUMA_Deb']).value
                sum_hab = worksheet.cell(row=row, column=ATTRIBUTES_GET['SUMA_Hab']).value
                deb_cierre = worksheet.cell(row=row, column=ATTRIBUTES_GET['DEBE_Cierre']).value
                hab_cierre = worksheet.cell(row=row, column=ATTRIBUTES_GET['HABER_Cierre']).value

                balance_dictionary.append(Res2(account_id, account_name, apgi, sum_deb, sum_hab, deb_cierre, hab_cierre))
            else:
                balance_dictionary.append(Res2(None, None, None, None, None, None, None))

        return balance_dictionary

    @staticmethod
    def set_resume_one(worksheet, worker_report: WorkerFile, is_pivot=False):
        balance = worker_report.get_res_2()
        counter_row = 3
        for item in balance:
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['SUMA_Hab'], value=item.get_sum_hab())
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['SUMA_Deb'], value=item.get_sum_deb())
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['HABER_Cierre'], value=item.get_hab_cierre())
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['DEBE_Cierre'], value=item.get_deb_cierre())

            if is_pivot:
                worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['ID C'], value=item.get_account_id())
                worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['Nombre Cuen'], value=item.get_account_name())
                worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['Gr.'], value=item.get_apgi())
            counter_row += 1
