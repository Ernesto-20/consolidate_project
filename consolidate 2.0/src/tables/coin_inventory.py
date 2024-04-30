

class CoinInventory:

    def __init__(self, input, output, account_id, existence):
        self.__input = input
        self.__output = output
        self.__account_id = account_id
        self.__existence = existence

    def get_input(self): return self.__input

    def get_output(self): return self.__output

    def get_account_id(self): return self.__account_id

    def get_existence(self): return self.__existence

    def set_input(self, value): self.__input = value

    def set_output(self, value): self.__output = value

    def set_account_id(self, value): self.__account_id = value

    def set_existence(self, value): self.__existence = value




