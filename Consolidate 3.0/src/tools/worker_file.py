
class WorkerFile:

    def __init__(self, name: str, res_1: dict, res_2: dict, res_3_ingreso_egreso: dict, res_3_saldo: dict, div: dict):
        self.__name = name
        self.__res_1 = res_1
        self.__res_2 = res_2
        self.__res_3_ingreso_egreso = res_3_ingreso_egreso
        self.__res_3_saldo = res_3_saldo
        self.__div = div

    def get_name(self): return self.__name

    def get_res_1(self): return self.__res_1

    def get_res_2(self): return self.__res_2

    def get_res_3_ingreso_egreso(self): return self.__res_3_ingreso_egreso
    def get_res_3_saldo(self): return self.__res_3_saldo

    def get_div(self): return self.__div



