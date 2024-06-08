from src.tables.div import Div
from src.tools.worker_file import WorkerFile

ATTRIBUTES_SET = {'Cargo': 23, 'Dividendo': 24,}
ATTRIBUTES_GET = {'Cargo': 9, 'Dividendo': 10}

class ResumeDivManage:

    @staticmethod
    def get_resume_div(worksheet):
        div_dictionary = []
        max_row = worksheet.max_row + 1
        for row in range(5, max_row+1):
            if worksheet.cell(row=row, column=ATTRIBUTES_GET['Dividendo']).value is not None:
                cargo = worksheet.cell(row=row, column=ATTRIBUTES_GET['Cargo']).value
                div = worksheet.cell(row=row, column=ATTRIBUTES_GET['Dividendo']).value

                div_dictionary.append(Div(cargo, div))
            else:
                div_dictionary.append(Div(None, None))

        return div_dictionary

    @staticmethod
    def set_resume_div(worksheet, worker_report: WorkerFile, is_pivot: False):
        div = worker_report.get_div()
        counter_row = 3
        for item in div:
            if is_pivot:
                worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['Cargo'], value=item.get_cargo())
                
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['Dividendo'], value=item.get_div())
            counter_row += 1
