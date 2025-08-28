class Operation:
    def __init__(self, operation_id, client, value, term_days, rate_type, spread, cost):
        self.operation_id = operation_id
        self.client = client  # Association with Client
        self.value = value
        self.term_days = term_days
        self.rate_type = rate_type
        self.spread = spread
        self.cost = cost
        self.final_rate = None

    def calculate_rate(self):
        # Placeholder - business rule to be added
        self.final_rate = self.spread + self.cost
        return self.final_rate

    def __repr__(self):
        return f"<Operation {self.operation_id} - Value: {self.value}>"
