from src.tables.product import Product
from src.tools.worker_file import WorkerFile

ATTRIBUTES = {'id': 1, 'CB': 3, 'CN': 4, 'CntC': 5, 'IN Int': 6, 'IN Ext': 7,
              'CantV': 8, 'OUT Int': 9, 'OUT Ext': 10, 'Ing': 11, 'Egr': 12,
              'Exist Teor': 13, 'E.F.R': 14}

# sheet(Resumen)
class ProductsManage:

    @staticmethod
    def get_products(worksheet, is_pivot=False):
        products_dictionary = []
        max_row = worksheet.max_row + 1
        for row in range(5, max_row):
            if worksheet.cell(row=row, column=ATTRIBUTES['id']).value is not None:
                # Conditional of Pivot ('aaa')
                id = None
                gross_cost = None
                net_cost = None

                if is_pivot:
                    id = worksheet.cell(row=row, column=ATTRIBUTES['id']).value
                    gross_cost = worksheet.cell(row=row, column=ATTRIBUTES['CB']).value
                    net_cost = worksheet.cell(row=row, column=ATTRIBUTES['CN']).value

                buy_amount = worksheet.cell(row=row, column=ATTRIBUTES['CntC']).value
                internal_input = worksheet.cell(row=row, column=ATTRIBUTES['IN Int']).value
                external_input = worksheet.cell(row=row, column=ATTRIBUTES['IN Ext']).value
                amount_sell = worksheet.cell(row=row, column=ATTRIBUTES['CantV']).value
                internal_output = worksheet.cell(row=row, column=ATTRIBUTES['OUT Int']).value
                external_output = worksheet.cell(row=row, column=ATTRIBUTES['OUT Ext']).value
                income = worksheet.cell(row=row, column=ATTRIBUTES['Ing']).value
                egress = worksheet.cell(row=row, column=ATTRIBUTES['Egr']).value
                theoretical_stock = worksheet.cell(row=row, column=ATTRIBUTES['Exist Teor']).value
                real_stock = worksheet.cell(row=row, column=ATTRIBUTES['E.F.R']).value
                products_dictionary.append(Product(id, gross_cost, net_cost, buy_amount, internal_input, external_input,
                                                  amount_sell,
                                                  internal_output, external_output, income, egress, theoretical_stock,
                                                  real_stock))
            else:
                products_dictionary.append(Product(None, None, None, None, None, None, None, None, None, None, None, None, None))
        return products_dictionary

    @staticmethod
    def set_products(worksheet, worker_report: WorkerFile, is_pivot=False):
        products = worker_report.get_products()
        counter_row = 5
        for product in products:
            if is_pivot:
                worksheet.cell(row=counter_row, column=ATTRIBUTES['id'], value=product.get_id())
                worksheet.cell(row=counter_row, column=ATTRIBUTES['CB'], value=product.get_gross_cost())
                worksheet.cell(row=counter_row, column=ATTRIBUTES['CN'], value=product.get_net_cost())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['CntC'], value=product.get_buy_amount())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['IN Int'], value=product.get_internal_input())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['IN Ext'], value=product.get_external_input())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['CantV'], value=product.get_amount_sell())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['OUT Int'], value=product.get_internal_output())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['OUT Ext'], value=product.get_external_output())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['Ing'], value=product.get_income())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['Egr'], value=product.get_egress())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['Exist Teor'], value=product.get_theoretical_stock())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['E.F.R'], value=product.get_real_stock())
            counter_row += 1

