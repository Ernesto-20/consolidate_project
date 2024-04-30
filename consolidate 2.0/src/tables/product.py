class Product:
    def __init__(self, id, gross_cost, net_cost, buy_amount,
                 internal_input, external_input, amount_sell,
                 internal_output, external_output, income, egress,
                 theoretical_stock, real_stock):
        self.id = id
        self.gross_cost = gross_cost
        self.net_cost = net_cost
        self.buy_amount = buy_amount
        self.internal_input = internal_input
        self.external_input = external_input
        self.amount_sell = amount_sell
        self.internal_output = internal_output
        self.external_output = external_output
        self.income = income
        self.egress = egress
        self.theoretical_stock = theoretical_stock
        self.real_stock = real_stock


    def get_id(self): return self.id

    def get_gross_cost(self): return self.gross_cost

    def get_net_cost(self): return self.net_cost

    def get_buy_amount(self): return self.buy_amount

    def get_internal_input(self): return self.internal_input

    def get_external_input(self): return self.external_input

    def get_amount_sell(self): return self.amount_sell

    def get_internal_output(self): return self.internal_output

    def get_external_output(self): return self.external_output

    def get_income(self): return self.income

    def get_egress(self): return self.egress

    def get_theoretical_stock(self): return self.theoretical_stock

    def get_real_stock(self): return self.real_stock

    def set_gross_cost(self, value): self.gross_cost = value

    def set_net_cost(self, value): self.net_cost = value

    def set_buy_amount(self, value): self.buy_amount = value

    def set_internal_input(self, value): self.internal_input = value

    def set_external_input(self, value): self.external_input = value

    def set_amount_sell(self, value): self.amount_sell = value

    def set_internal_output(self, value): self.internal_output = value

    def set_external_output(self, value): self.external_output = value

    def set_income(self, value): self.income = value

    def set_egress(self, value): self.egress = value

    def set_theoretical_stock(self, value): self.theoretical_stock = value

    def set_real_stock(self, value): self.real_stock = value
