class Balance:

    def __init__(self, account_id, debit, credit, account_name, apgi):
        self.__account_id = account_id
        self.__debit = debit
        self.__credit = credit
        self.__account_name = account_name
        self.__apgi = apgi

    def get_account_id(self): return self.__account_id

    def get_debit(self): return self.__debit

    def get_credit(self): return self.__credit

    def get_account_name(self): return self.__account_name

    def get_apgi(self): return self.__apgi

    def set_account_id(self, value): self.__account_id = value

    def set_debit(self, value): self.__debit = value

    def set_credit(self, value): self.__credit = value

    def set_account_name(self, value): self.__account_name = value

    def set_apgi(self, value): self.__apgi = value
