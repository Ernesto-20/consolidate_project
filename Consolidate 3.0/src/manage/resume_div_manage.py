from src.tables.div import Div
from src.tools.worker_file import WorkerFile

ATTRIBUTES_SET = {'Dividendo': 23,}
ATTRIBUTES_GET = {'Dividendo': 10}

class ResumeDivManage:

    @staticmethod
    def get_resume_div(worksheet):
        div_dictionary = []
        max_row = worksheet.max_row + 1
        for row in range(5, max_row+1):
            if worksheet.cell(row=row, column=ATTRIBUTES_GET['Dividendo']).value is not None:
                div = worksheet.cell(row=row, column=ATTRIBUTES_GET['Dividendo']).value

                div_dictionary.append(Div(div))
            else:
                div_dictionary.append(Div(None,))

        return div_dictionary

    @staticmethod
    def set_resume_div(worksheet, worker_report: WorkerFile):
        div = worker_report.get_div()
        counter_row = 3
        for item in div:
            worksheet.cell(row=counter_row, column=ATTRIBUTES_SET['Dividendo'], value=item.get_div())
            counter_row += 1
