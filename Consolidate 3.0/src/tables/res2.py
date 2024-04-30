class Res2:

    def __init__(self, account_id, account_name, apgi, sum_deb, sum_hab, deb_cierre, hab_cierre):
        self.__account_id = account_id
        self.__account_name = account_name
        self.__apgi = apgi
        self.__sum_deb = sum_deb
        self.__sum_hab = sum_hab
        self.__deb_cierre = deb_cierre
        self.__hab_cierre = hab_cierre

    def get_account_id(self): return self.__account_id

    def get_account_name(self): return self.__account_name

    def get_apgi(self): return self.__apgi

    def get_sum_deb(self): return self.__sum_deb

    def get_sum_hab(self): return self.__sum_hab

    def get_deb_cierre(self): return self.__deb_cierre

    def get_hab_cierre(self): return self.__hab_cierre