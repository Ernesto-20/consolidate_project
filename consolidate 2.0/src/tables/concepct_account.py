
class ConceptAccount:

    def __init__(self, account_id, account_name, apgi, concept_name, conpcect_id, account_id_2):
        self.__account_id = account_id
        self.__account_name = account_name
        self.__apgi = apgi
        self.__concept_name = concept_name
        self.__conpcect_id = conpcect_id
        self.__account_id_2 = account_id_2

    def get_account_id(self): return self.__account_id

    def get_account_name(self): return self.__account_name

    def get_apgi(self): return self.__apgi

    def get_concept_name(self): return self.__concept_name

    def get_concept_id(self): return self.__conpcect_id

    def get_account_id_2(self): return self.__account_id_2