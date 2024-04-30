from src.tables.merchandise_control import MerchandiseControl
from src.tools.worker_file import WorkerFile

ATTRIBUTES = {'IN': 29, 'OUT': 30, 'Id Conc': 31}

# sheet(Resumen)
class MerchandiseControlManage:

    @staticmethod
    def get_merchandise_control_source(worksheet, is_pivot=False):
        merchandise_dictionary = []
        for row in range(5, 18):
            if worksheet.cell(row=row, column=ATTRIBUTES['IN']).value is not None:
                input = worksheet.cell(row=row, column=ATTRIBUTES['IN']).value
                output = worksheet.cell(row=row, column=ATTRIBUTES['OUT']).value

                id_concept = None

                if is_pivot:
                    id_concept = worksheet.cell(row=row, column=ATTRIBUTES['Id Conc']).value

                merchandise_dictionary.append(MerchandiseControl(input, output, id_concept))
            else:
                merchandise_dictionary.append(MerchandiseControl(None, None, None))

        return merchandise_dictionary

    @staticmethod
    def set_merchandise_control_source(worksheet, worker_report: WorkerFile, is_pivot=False):
        merchandise = worker_report.get_merchandise_control()
        counter_row = 5
        for item in merchandise:
            worksheet.cell(row=counter_row, column=ATTRIBUTES['IN'], value=item.get_input())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['OUT'], value=item.get_output())
            if is_pivot:
                worksheet.cell(row=counter_row, column=ATTRIBUTES['Id Conc'], value=item.get_id_concept())
            counter_row += 1
