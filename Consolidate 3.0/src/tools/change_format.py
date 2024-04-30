# import jpype
# import asposecells
# jpype.startJVM()
# from asposecells.api import Workbook


# class ChangeFormat:
#     @staticmethod
#     def xlsb_to_xlsx(file: str):
#         output = 'temp_format.xlsx'
#         workbook = Workbook(file)
#         workbook.save(output)

#         return output

#     @staticmethod
#     def xlsx_to_xlsb(file: str):
#         output = file[:len(file)-1]+'b'
#         print(output)
#         workbook = Workbook(file)
#         workbook.save(output)

#         return output
