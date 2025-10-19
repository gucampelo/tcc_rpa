class Operation:
    def __init__(self, nmr_po, client, value, term_days, rate_type, spread_requested,
                  cost_requested, rate_requested, parcel_flow):
        self.__nmr_po = nmr_po
        self.__client = client
        self.__value = value
        self.__term_days = term_days
        self.__rate_type = rate_type
        self.__parcel_flow = parcel_flow

        # Requested values
        self.__spread_requested = spread_requested
        self.__cost_requested = cost_requested
        self.__rate_requested = rate_requested

        # Approved values
        self.__spread_approved = None
        self.__cost_approved = None
        self.__rate_approved = None

    # -----------------------------
    # Getters e Setters (encapsulamento)
    # -----------------------------
    @property
    def nmr_po(self):
        return self.__nmr_po

    @property
    def client(self):
        return self.__client

    @property
    def value(self):
        return self.__value

    @value.setter
    def value(self, new_value):
        if new_value < 0:
            raise ValueError("Value cannot be negative.")
        self.__value = new_value

    @property
    def term_days(self):
        return self.__term_days

    @term_days.setter
    def term_days(self, days):
        if days <= 0:
            raise ValueError("Term days must be greater than zero.")
        self.__term_days = days

    @property
    def rate_type(self):
        return self.__rate_type

    @rate_type.setter
    def rate_type(self, rate_type):
        if rate_type not in ["Pré-Fixada", "Pós-Fixada"]:
            raise ValueError("Invalid rate type. Must be 'Pré-Fixada' or 'Pós-Fixada'.")
        self.__rate_type = rate_type

    @property
    def parcel_flow(self):
        return self.__parcel_flow

    @parcel_flow.setter
    def parcel_flow(self, flow):
        self.__parcel_flow = flow

    # Requested
    @property
    def spread_requested(self):
        return self.__spread_requested

    @spread_requested.setter
    def spread_requested(self, spread):
        self.__spread_requested = spread

    @property
    def cost_requested(self):
        return self.__cost_requested

    @cost_requested.setter
    def cost_requested(self, cost):
        self.__cost_requested = cost

    @property
    def rate_requested(self):
        return self.__rate_requested

    @rate_requested.setter
    def rate_requested(self, rate):
        self.__rate_requested = rate

    # Approved
    @property
    def spread_approved(self):
        return self.__spread_approved

    @spread_approved.setter
    def spread_approved(self, spread):
        self.__spread_approved = spread

    @property
    def cost_approved(self):
        return self.__cost_approved

    @cost_approved.setter
    def cost_approved(self, cost):
        self.__cost_approved = cost

    @property
    def rate_approved(self):
        return self.__rate_approved

    @rate_approved.setter
    def rate_approved(self, rate):
        self.__rate_approved = rate

    # -----------------------------
    # Métodos de negócio
    # -----------------------------
    def calculate_rate(self, cost_approved):
        # Converter taxa e custo de anual para mensal
        rate_monthly = (1 + self.__rate_requested) ** (1/12) - 1
        cost_monthly = (1 + cost_approved) ** (1/12) - 1

        # Calcular o spread mensal
        spread_monthly = rate_monthly - cost_monthly

        self.cost_approved = cost_approved
        self.rate_approved = self.__rate_requested
        # Converter o spread mensal de volta para anual
        self.spread_approved = round((1 + spread_monthly) ** 12 - 1, 4)

    # -----------------------------
    # Representação textual
    # -----------------------------
    def __repr__(self):
        return f"<Operation {self.__nmr_po} - Value: {self.__value}>"
