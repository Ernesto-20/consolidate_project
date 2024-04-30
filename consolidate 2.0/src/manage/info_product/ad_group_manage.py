from src.tables.info_product.ad_group import ADGroup
from src.tools.file_model import FileModel

ATTRIBUTES = {'name (A)': 2, 'Cod(A)': 3, 'name (D)': 24, 'Cod(D)': 25}


# sheet(Prod)
class ADGroupManage:

    @staticmethod
    def get_ad_group(worksheet):
        products = []
        max_row = worksheet.max_row + 1
        for row in range(5, max_row):
            name_a = None
            code_a = None
            name_d = None
            code_d = None

            if worksheet.cell(row=row, column=ATTRIBUTES['Cod(A)']).value is not None:
                name_a = worksheet.cell(row=row, column=ATTRIBUTES['name (A)']).value
                code_a = worksheet.cell(row=row, column=ATTRIBUTES['Cod(A)']).value

            if worksheet.cell(row=row, column=ATTRIBUTES['Cod(A)']).value is not None:
                name_d = worksheet.cell(row=row, column=ATTRIBUTES['name (D)']).value
                code_d = worksheet.cell(row=row, column=ATTRIBUTES['Cod(D)']).value

            products.append(ADGroup(name_a, code_a, name_d, code_d))
        return products

    @staticmethod
    def set_ad_group(worksheet, file_model: FileModel):
        products = file_model.get_ad_group()
        counter_row = 5
        for product in products:
            worksheet.cell(row=counter_row, column=ATTRIBUTES['name (A)'], value=product.get_name_a())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['Cod(A)'], value=product.get_code_a())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['name (D)'], value=product.get_name_d())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['Cod(D)'], value=product.get_code_d())
            counter_row += 1
