class Res3:
    # This class stores all the attributes necessary for the consolidation of the Res3 sheet corresponding to a record.
    def __init__(self, concept, sum_ing, sum_egr, id_cuen, saldo_cup, saldo_div, real_div):
        self.__concept = concept
        self.__sum_ing = sum_ing
        self.__sum_egr = sum_egr
        self.__id_cuen = id_cuen
        self.__saldo_cup = saldo_cup
        self.__saldo_div = saldo_div
        self.__real_div = real_div

    def get_concept(self): return self.__concept

    def get_sum_ing(self): return self.__sum_ing

    def get_sum_egr(self): return self.__sum_egr

    def get_id_cuen(self): return self.__id_cuen

    def get_saldo_cup(self): return self.__saldo_cup

    def get_saldo_div(self): return self.__saldo_div
    
    def get_real_div(self): return self.__real_div
