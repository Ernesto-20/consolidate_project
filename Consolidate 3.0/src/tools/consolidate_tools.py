from src.tools.worker_file import WorkerFile
from openpyxl import load_workbook
from tkinter import messagebox
import sys

def to_consolidate_data(worker_files):
    products = None
    for worker_data in worker_files:
        if products is None:
            products = worker_data.products
        else:
            for key_p in worker_data.products.key():
                if key_p not in products:
                    products[key_p] = worker_data.products[key_p]
                else:
                    # products[key_p].set_gross_cost(products[key_p].get_gross_cost() + worker_data.products[key_p].get_gross_cost())
                    # products[key_p].net_cost(products[key_p].net_cost() + worker_data.products[key_p].net_cost())
                    products[key_p].buy_amount(products[key_p].buy_amount() + worker_data.products[key_p].buy_amount())
                    products[key_p].internal_input(products[key_p].internal_input() + worker_data.products[key_p].internal_input())
                    products[key_p].external_input(products[key_p].external_input() + worker_data.products[key_p].external_input())
                    products[key_p].amount_sell(products[key_p].amount_sell() + worker_data.products[key_p].amount_sell())
                    products[key_p].internal_output(products[key_p].internal_output() + worker_data.products[key_p].internal_output())
                    products[key_p].income(products[key_p].income() + worker_data.products[key_p].income())
                    products[key_p].egress(products[key_p].egress() + worker_data.products[key_p].egress())
                    products[key_p].theoretical_stock(products[key_p].theoretical_stock() + worker_data.products[key_p].theoretical_stock())
                    products[key_p].real_stock(products[key_p].real_stock() + worker_data.products[key_p].real_stock())

    return WorkerFile(name='Reporte', res_1=products)

def load_sheet(workbook, sheet_name: str, workbook_name: str):
    ws = None
    try:
        ws = workbook[sheet_name]
    except KeyError:
        # Exception: No existe esa hoja
        messagebox.showerror(title='Error', message=
        'El archivo "' + workbook_name + '" no contiene la hoja "' + sheet_name + '".\n\n Puede que en el archivo, el nombre "' + sheet_name + '" contenga algun espacio en blanco o caracter extraño.')
        sys.exit()

    return ws
