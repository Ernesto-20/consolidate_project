from src.tables.balance import Balance
from src.tools.worker_file import WorkerFile

ATTRIBUTES = {'ID Cuen': 2, 'DEBE': 3, 'HABER': 4, 'Nombre': 11, 'APGI': 12}

# sheet(Balance)
class BalanceManage:

    @staticmethod
    def get_balance_source(worksheet, is_pivot=False):
        balance_dictionary = []
        max_row = worksheet.max_row + 1
        for row in range(5, max_row+1):
            if worksheet.cell(row=row, column=ATTRIBUTES['ID Cuen']).value is not None:

                account_id = None
                account_name = None
                apgi = None
                debit = worksheet.cell(row=row, column=ATTRIBUTES['DEBE']).value
                credit = worksheet.cell(row=row, column=ATTRIBUTES['HABER']).value
                if is_pivot:
                    account_id = worksheet.cell(row=row, column=ATTRIBUTES['ID Cuen']).value
                    account_name = worksheet.cell(row=row, column=ATTRIBUTES['Nombre']).value
                    apgi = worksheet.cell(row=row, column=ATTRIBUTES['APGI']).value

                balance_dictionary.append(Balance(account_id, debit, credit, account_name, apgi))
            else:
                balance_dictionary.append(Balance(None, None, None, None, None))

        return balance_dictionary

    @staticmethod
    def set_balance_source(worksheet, worker_report: WorkerFile, is_pivot=False):
        balance = worker_report.get_balance()
        counter_row = 5
        for item in balance:
            worksheet.cell(row=counter_row, column=ATTRIBUTES['DEBE'], value=item.get_debit())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['HABER'], value=item.get_credit())

            if is_pivot:
                worksheet.cell(row=counter_row, column=ATTRIBUTES['ID Cuen'], value=item.get_account_id())
                worksheet.cell(row=counter_row, column=ATTRIBUTES['Nombre'], value=item.get_account_name())
                worksheet.cell(row=counter_row, column=ATTRIBUTES['APGI'], value=item.get_apgi())
            counter_row += 1
