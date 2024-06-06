from src.tables.res2 import Res2
from src.tables.res3 import Res3
from src.tools.worker_file import WorkerFile

ATTRIBUTES_SET = {'Concept': 19, 'Suma_Ing': 20, 'Suma_Egr': 21, 'Id_Cue': 19, 'Saldo_CUP': 20, 'Saldo_DIV': 22, 'REAL_DIV': 21}
ATTRIBUTES_GET = {'Concept': 1, 'Suma_Ing': 17, 'Suma_Egr': 18, 'Id_Cue': 1, 'Saldo_CUP': 18, 'Saldo_DIV': 20, 'REAL_DIV': 19}

class ResumeThreeManage:

    @staticmethod
    def get_resume_three(worksheet,):
        resume_list = []
        for row in range(3, 17):
            if worksheet.cell(row=row, column=ATTRIBUTES_GET['Suma_Ing']).value is not None:
                concept   = worksheet.cell(row=row, column=ATTRIBUTES_GET['Concept']).value
                sum_ing = worksheet.cell(row=row, column=ATTRIBUTES_GET['Suma_Ing']).value
                sum_egr = worksheet.cell(row=row, column=ATTRIBUTES_GET['Suma_Egr']).value

                id_cue    = worksheet.cell(row=row+16, column=ATTRIBUTES_GET['Id_Cue']).value
                saldo_cup = worksheet.cell(row=row+16, column=ATTRIBUTES_GET['Saldo_CUP']).value
                real_div  = worksheet.cell(row=row+16, column=ATTRIBUTES_GET['REAL_DIV']).value
                saldo_div = worksheet.cell(row=row+16, column=ATTRIBUTES_GET['Saldo_DIV']).value

                resume_list.append(Res3(concept, sum_ing, sum_egr, id_cue, saldo_cup, saldo_div, real_div))
            else:
                resume_list.append(Res3(None, None, None, None, None, None, None,))

        return resume_list

    @staticmethod
    def set_resume_three(worksheet, worker_report: WorkerFile):
        res_list = worker_report.get_res_3()
        counter_row = 3
        for item in res_list:
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['Concept'], value=item.get_concept())
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['Suma_Ing'], value=item.get_sum_ing())
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['Suma_Egr'], value=item.get_sum_egr())

            worksheet.cell(row=counter_row+16, column=ATTRIBUTES_SET['Id_Cue'], value=item.get_id_cuen())
            worksheet.cell(row=counter_row+16, column=ATTRIBUTES_SET['Saldo_CUP'], value=item.get_saldo_cup())
            worksheet.cell(row=counter_row+16, column=ATTRIBUTES_SET['REAL_DIV'], value=item.get_real_div())
            worksheet.cell(row=counter_row+16, column=ATTRIBUTES_SET['Saldo_DIV'], value=item.get_saldo_div())
            counter_row += 1
