from src.tables.res2 import Res2
from src.tables.res3 import Res3
from src.tools.worker_file import WorkerFile

ATTRIBUTES_SET = {'Suma_Ing': 19, 'Suma_Egr': 20, 'Saldo_CUP': 19, 'Saldo_DIV': 20}
ATTRIBUTES_GET = {'Suma_Ing': 17, 'Suma_Egr': 18, 'Saldo_CUP': 17, 'Saldo_DIV': 18}

class ResumeThreeManage:

    @staticmethod
    def get_resume_three(worksheet,):
        resume_list = []
        max_row = worksheet.max_row + 1
        for row in range(3, max_row+1):
            if worksheet.cell(row=row, column=ATTRIBUTES_GET['Suma_Ing']).value is not None:
                sum_ing = worksheet.cell(row=row, column=ATTRIBUTES_GET['Suma_Ing']).value
                sum_egr = worksheet.cell(row=row, column=ATTRIBUTES_GET['Suma_Egr']).value
                saldo_cup = worksheet.cell(row=row, column=ATTRIBUTES_GET['Saldo_CUP']).value
                saldo_div = worksheet.cell(row=row, column=ATTRIBUTES_GET['Saldo_DIV']).value

                resume_list.append(Res3(sum_ing, sum_egr, saldo_cup, saldo_div))
            else:
                resume_list.append(Res3(None, None, None, None,))

        return resume_list

    @staticmethod
    def set_resume_three(worksheet, worker_report: WorkerFile):
        res_list = worker_report.get_res_3()
        counter_row = 3
        for item in res_list:
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['Suma_Ing'], value=item.get_sum_ing())
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['Suma_Egr'], value=item.get_sum_egr())

            # worksheet.cell(row=counter_row+19, column=ATTRIBUTES_SET['Saldo_CUP'], value=item.get_saldo_cup())
            # worksheet.cell(row=counter_row+19, column=ATTRIBUTES_SET['Saldo_DIV'], value=item.get_saldo_div())
            counter_row += 1
