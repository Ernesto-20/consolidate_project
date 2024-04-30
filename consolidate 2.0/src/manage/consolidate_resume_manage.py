from src.tools.file_model import FileModel

ATTRIBUTES = {'Cbrt 1': 20,
              'CB': 28}
# "Cbrt 1" possition 20 in Cons-sem...
# "CB" possition 28 in Cont-sem...

# sheet(Cue-Con)
class ConsolidateResumeManage:

    @staticmethod
    def get_gross_cost(worksheet):
        gross_cost = []
        max_row = worksheet.max_row + 1
        for row in range(5, max_row):
            cost = worksheet.cell(row=row, column=ATTRIBUTES['Cbrt 1']).value

            gross_cost.append(cost)
        return gross_cost

    @staticmethod
    def set_gross_cost(worksheet, file_model: FileModel):
        gross_cost = file_model.get_gross_cost()
        counter_row = 5
        for cost in gross_cost:
            worksheet.cell(row=counter_row, column=ATTRIBUTES['CB'], value=cost)
            counter_row += 1