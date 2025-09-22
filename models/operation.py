class Operation:
    def __init__(self, operation_id, client, value, term_days, rate_type, spread, cost, parcel_flow):
        self.operation_id = operation_id
        self.client = client  # Association with Client
        self.value = value
        self.term_days = term_days
        self.rate_type = rate_type
        self.spread = spread
        self.cost = cost
        self.parcel_flow = parcel_flow
        self.rate = None

    def calculate_rate(self):
        # Placeholder - business rule to be added
        self.rate = self.spread + self.cost
        return self.rate

    def __repr__(self):
        return f"<Operation {self.operation_id} - Value: {self.value}>"
