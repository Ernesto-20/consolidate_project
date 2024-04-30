class BCGroup:
    def __init__(self, name_b, nom_b, esp_b, a_cm_b, l_cm_b,
                 name_c, nom_c, esp_c, a_cm_c, l_cm_c):
        self.__name_b = name_b
        self.__nom_b = nom_b
        self.__esp_b = esp_b
        self.__a_cm_b = a_cm_b
        self.__l_cm_b = l_cm_b

        self.__name_c = name_c
        self.__nom_c = nom_c
        self.__esp_c = esp_c
        self.__a_cm_c = a_cm_c
        self.__l_cm_c = l_cm_c

    def get_name_b(self): return self.__name_b

    def get_nom_b(self): return self.__nom_b

    def get_esp_b(self): return self.__esp_b

    def get_a_cm_b(self): return self.__a_cm_b

    def get_l_cm_b(self): return self.__l_cm_b

    def get_name_c(self): return self.__name_c

    def get_nom_c(self): return self.__nom_c

    def get_esp_c(self): return self.__esp_c

    def get_a_cm_c(self): return self.__a_cm_c

    def get_l_cm_c(self): return self.__l_cm_c
