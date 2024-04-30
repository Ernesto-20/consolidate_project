from src.tables.coin_control import CoinControl
from src.tools.worker_file import WorkerFile

ATTRIBUTES = {'DEBE': 29, 'HABER': 30, 'Id Conc': 31}


class CoinControlManage:

    @staticmethod
    def get_coin_control_source(worksheet):
        merchandise_dictionary = {}
        for row in range(21, 34):
            if worksheet.cell(row=row, column=ATTRIBUTES['DEBE']).value is not None:
                debit = worksheet.cell(row=row, column=ATTRIBUTES['DEBE']).value
                credit = worksheet.cell(row=row, column=ATTRIBUTES['HABER']).value
                id_concept = worksheet.cell(row=row, column=ATTRIBUTES['Id Conc']).value

                merchandise_dictionary[id_concept] = CoinControl(debit, credit, id_concept)

        return merchandise_dictionary

    @staticmethod
    def set_coin_control_source(worksheet, worker_report: WorkerFile):
        coin_inventory = worker_report.get_coin_control()
        counter_row = 21
        for k in coin_inventory.keys():
            worksheet.cell(row=counter_row, column=ATTRIBUTES['DEBE'], value=coin_inventory[k].get_debit())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['HABER'], value=coin_inventory[k].get_credit())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['Id Conc'], value=coin_inventory[k].get_id_concept())
            counter_row += 1
