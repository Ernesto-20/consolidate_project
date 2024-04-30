from src.tables.coin_inventory import CoinInventory
from src.tools.worker_file import WorkerFile

ATTRIBUTES = {'IN': 22, 'OUT': 23, 'Id Cuen': 24, 'Exist': 26}

# sheet(Resumen)
class CoinInventoryManage:

    @staticmethod
    def get_coin_inventory_source(worksheet, is_pivot=False):
        balance_dictionary = []
        for row in range(5, 27):
            if worksheet.cell(row=row, column=ATTRIBUTES['Id Cuen']).value is not None:
                account_id = None

                input = worksheet.cell(row=row, column=ATTRIBUTES['IN']).value
                output = worksheet.cell(row=row, column=ATTRIBUTES['OUT']).value

                if is_pivot:
                    account_id = worksheet.cell(row=row, column=ATTRIBUTES['Id Cuen']).value

                existence = worksheet.cell(row=row, column=ATTRIBUTES['Exist']).value

                balance_dictionary.append(CoinInventory(input, output, account_id, existence))
        else:
            balance_dictionary.append(CoinInventory(None, None, None, None))

        return balance_dictionary

    @staticmethod
    def set_coin_inventory_source(worksheet, worker_report: WorkerFile, is_pivot=False):
        balance = worker_report.get_coin_inventory()
        counter_row = 5
        for account in balance:
            worksheet.cell(row=counter_row, column=ATTRIBUTES['IN'], value=account.get_input())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['OUT'], value=account.get_output())

            if is_pivot:
                worksheet.cell(row=counter_row, column=ATTRIBUTES['Id Cuen'], value=account.get_account_id())

            worksheet.cell(row=counter_row, column=ATTRIBUTES['Exist'], value=account.get_existence())
            counter_row += 1
