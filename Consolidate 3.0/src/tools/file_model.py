from src.tables.info_product.ad_group import ADGroup
from src.tables.info_product.bc_group import BCGroup
from src.tables.product_price import ProductPrice
from src.tables.concepct_account import ConceptAccount


class FileModel:

    def __init__(self, ad_group: ADGroup, bc_group: BCGroup, product_price: ProductPrice,
                 concept_account: ConceptAccount, gross_cost: list):
        self.__ad_group = ad_group
        self.__bc_group = bc_group
        self.__product_price = product_price
        self.__concept_account = concept_account
        self.__gross_cost = gross_cost

    def get_ad_group(self): return self.__ad_group

    def get_bc_group(self): return self.__bc_group

    def get_product_price(self): return self.__product_price

    def get_concept_account(self): return self.__concept_account

    def get_gross_cost(self): return self.__gross_cost
