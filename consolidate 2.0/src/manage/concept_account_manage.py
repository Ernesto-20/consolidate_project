from src.tables.concepct_account import ConceptAccount
from src.tools.file_model import FileModel

ATTRIBUTES = {'ID Cuen': 1, 'account_name': 2, 'APGI': 3,
              'concept_name': 7, 'id_concept': 8, 'id_account': 9}


# sheet(Cue-Con)
class ConceptAccountManage:

    @staticmethod
    def get_concept_account(worksheet):
        concepts = []
        max_row = worksheet.max_row + 1
        for row in range(5, max_row):
            account_id = worksheet.cell(row=row, column=ATTRIBUTES['ID Cuen']).value
            account_name = worksheet.cell(row=row, column=ATTRIBUTES['account_name']).value
            apgi = worksheet.cell(row=row, column=ATTRIBUTES['APGI']).value
            concept_name = worksheet.cell(row=row, column=ATTRIBUTES['concept_name']).value
            concept_id = worksheet.cell(row=row, column=ATTRIBUTES['id_concept']).value
            account_id_2 = worksheet.cell(row=row, column=ATTRIBUTES['id_account']).value

            concepts.append(ConceptAccount(account_id, account_name, apgi, concept_name, concept_id, account_id_2))
        return concepts

    @staticmethod
    def set_concept_account(worksheet, file_model: FileModel):
        concepts = file_model.get_concept_account()
        counter_row = 5
        for concept in concepts:
            worksheet.cell(row=counter_row, column=ATTRIBUTES['ID Cuen'], value=concept.get_account_id())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['account_name'], value=concept.get_account_name())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['APGI'], value=concept.get_apgi())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['concept_name'], value=concept.get_concept_name())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['id_concept'], value=concept.get_concept_id())
            worksheet.cell(row=counter_row, column=ATTRIBUTES['id_account'], value=concept.get_account_id_2())
            counter_row += 1
