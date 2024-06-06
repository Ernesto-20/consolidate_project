class Res1:
    # This class stores all the attributes necessary for the consolidation of the Res1 sheet corresponding to a record.
    
    def __init__(self, id, init_existence, buy_amount, sale_amount,
                in_list, ou_list, real_existence, sale, purchase, cb, cn):
        self.id = id
        self.init_existence = init_existence
        self.buy_amount = buy_amount
        self.sale_amount = sale_amount
        self.in_list = in_list
        self.ou_list = ou_list
        self.real_existence = real_existence
        self.sale = sale
        self.purchase = purchase
        self.cb = cb
        self.cn = cn

    def __eq__(self, other):
        return self.id == other.get_id()

    def get_id(self): return self.id

    def get_init_existence(self): return self.init_existence

    def get_buy_amount(self): return self.buy_amount

    def get_sale_amount(self): return self.sale_amount

    def get_in_list(self): return self.in_list

    def get_ou_list(self): return self.ou_list

    def get_real_existence(self): return self.real_existence

    def get_sale(self): return self.sale

    def get_purchase(self): return self.purchase

    def get_cb(self): return self.cb

    def get_cn(self): return self.cn