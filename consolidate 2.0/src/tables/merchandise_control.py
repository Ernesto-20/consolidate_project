

class MerchandiseControl:

    def __init__(self, input, output, id_concept):
        self.__input = input
        self.__output = output
        self.__id_concept = id_concept

    def get_input(self): return self.__input

    def get_output(self): return self.__output

    def get_id_concept(self): return self.__id_concept

    def set_input(self, value): self.__input = value

    def set_output(self, value): self.__output = value

    def set_id_concept(self, value): self.__id_concept = value
