from src.tables.product_price import ProductPrice
from src.tools.file_model import FileModel

ATTRIBUTES = {'Prec.i 1': 4, 'Prec 1': 10, 'Prec (B) 1': 34,
              'Prec (BO) 1': 40, 'Prec (C) 1': 46, 'Prec (D) 1': 52, }

# sheet(Prec)
class ProductPriceManage:

    @staticmethod
    def get_products_price(worksheet):
        products = []
        max_row = worksheet.max_row + 1
        for row in range(5, max_row):
            prec_i1 = worksheet.cell(row=row, column=ATTRIBUTES['Prec.i 1']).value
            prec_1 = worksheet.cell(row=row, column=ATTRIBUTES['Prec 1']).value
            prec_b1 = worksheet.cell(row=row, column=ATTRIBUTES['Prec (B) 1']).value
            prec_bo1 = worksheet.cell(row=row, column=ATTRIBUTES['Prec (BO) 1']).value
            prec_c1 = worksheet.cell(row=row, column=ATTRIBUTES['Prec (C) 1']).value
            prec_d1 = worksheet.cell(row=row, column=ATTRIBUTES['Prec (D) 1']).value

            products.append(ProductPrice(prec_i1, prec_1, prec_b1, prec_bo1, prec_c1, prec_d1))
        return products

    @staticmethod
    def set_products_price(worksheet, file_model: FileModel):
        products = file_model.get_product_price()
        counter_row = 5
        for product in products:
            worksheet.cell(row=counter_row, column=ATTRIBUTES['Prec.i 1'], value=product.get_pric_i1())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['Prec 1'], value=product.get_pric_1())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['Prec (B) 1'], value=product.get_pric_b1())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['Prec (BO) 1'], value=product.get_pric_bo1())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['Prec (C) 1'], value=product.get_pric_c1())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['Prec (D) 1'], value=product.get_pric_d1())
            counter_row += 1
