
class WorkerFile:

    def __init__(self, name: str, products: dict, coin_inventory: dict, merchandise_control: dict, balance: dict):
        self.__name = name
        self.__products = products
        self.__coin_inventory = coin_inventory
        self.__merchandise_control = merchandise_control
        self.__balance = balance

    def get_name(self): return self.__name

    def get_products(self): return self.__products

    def get_coin_inventory(self): return self.__coin_inventory

    def get_merchandise_control(self): return self.__merchandise_control

    def get_balance(self): return self.__balance

