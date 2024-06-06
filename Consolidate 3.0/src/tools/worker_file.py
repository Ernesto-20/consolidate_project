
class WorkerFile:

    def __init__(self, name: str, res_1: dict, res_2: dict, res_3: dict, div: dict):
        self.__name = name
        self.__res_1 = res_1
        self.__res_2 = res_2
        self.__res_3 = res_3
        self.__div = div

    def get_name(self): return self.__name

    def get_res_1(self): return self.__res_1

    def get_res_2(self): return self.__res_2

    def get_res_3(self): return self.__res_3

    def get_div(self): return self.__div



