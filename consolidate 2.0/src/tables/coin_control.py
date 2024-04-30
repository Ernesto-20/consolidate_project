

class CoinControl:

    def __init__(self, debit, credit, id_concept):
        self.__debit = debit
        self.__credit = credit
        self.__id_concept = id_concept

    def get_debit(self): return self.__debit

    def get_credit(self): return self.__credit

    def get_id_concept(self): return self.__id_concept

    def set_debit(self, value): self.__debit = value

    def set_credit(self, value): self.__credit = value

    def set_id_concept(self, value): self.__id_concept = value
