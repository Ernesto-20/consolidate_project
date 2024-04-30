from src.tables.info_product.bc_group import BCGroup
from src.tools.file_model import FileModel

ATTRIBUTES = {'name (B)': 9, 'Nom (B)': 10, 'Esp. (B)': 11, 'A cm (B)': 12, 'L cm (B)': 13,
              'name (C)': 17, 'Nom (C)': 18, 'Esp. (C)': 19, 'A cm (C)': 20, 'L cm (C)': 21}


# sheet(Prod)
class BCGroupManage:

    @staticmethod
    def get_bc_group(worksheet):
        products = []
        max_row = worksheet.max_row + 1
        for row in range(5, max_row+1):
            name_b = None
            nom_b = None
            esp_b = None
            width_b = None
            height_b = None
            name_c = None
            nom_c = None
            esp_c = None
            width_c = None
            height_c = None

            if worksheet.cell(row=row, column=ATTRIBUTES['Nom (B)']).value is not None:
                name_b = worksheet.cell(row=row, column=ATTRIBUTES['name (B)']).value
                nom_b = worksheet.cell(row=row, column=ATTRIBUTES['Nom (B)']).value
                esp_b = worksheet.cell(row=row, column=ATTRIBUTES['Esp. (B)']).value
                width_b = worksheet.cell(row=row, column=ATTRIBUTES['A cm (B)']).value
                height_b = worksheet.cell(row=row, column=ATTRIBUTES['L cm (B)']).value
            if worksheet.cell(row=row, column=ATTRIBUTES['Nom (C)']).value is not None:
                name_c = worksheet.cell(row=row, column=ATTRIBUTES['name (C)']).value
                nom_c = worksheet.cell(row=row, column=ATTRIBUTES['Nom (C)']).value
                esp_c = worksheet.cell(row=row, column=ATTRIBUTES['Esp. (C)']).value
                width_c = worksheet.cell(row=row, column=ATTRIBUTES['A cm (C)']).value
                height_c = worksheet.cell(row=row, column=ATTRIBUTES['L cm (C)']).value

            products.append(BCGroup(name_b, nom_b, esp_b, width_b, height_b, name_c, nom_c, esp_c, width_c, height_c))
        return products

    @staticmethod
    def set_bc_group(worksheet, file_model: FileModel):
        products = file_model.get_bc_group()
        counter_row = 5
        for product in products:
            worksheet.cell(row=counter_row, column=ATTRIBUTES['name (B)'], value=product.get_name_b())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['Nom (B)'], value=product.get_nom_b())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['Esp. (B)'], value=product.get_esp_b())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['A cm (B)'], value=product.get_a_cm_b())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['L cm (B)'], value=product.get_l_cm_b())

            worksheet.cell(row=counter_row, column=ATTRIBUTES['name (C)'], value=product.get_name_c())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['Nom (C)'], value=product.get_nom_c())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['Esp. (C)'], value=product.get_esp_c())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['A cm (C)'], value=product.get_a_cm_c())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['L cm (C)'], value=product.get_l_cm_c())

            counter_row += 1
