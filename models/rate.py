# models/rate.py

class Rate:
    """
    Representa a taxa calculada de uma operação financeira.
    """
    def __init__(self, value: float, kind: str, scenario: str):
        self.value = value        # valor da taxa
        self.kind = kind          # tipo da taxa ('PRE' ou 'POS')
        self.scenario = scenario  # cenário do cálculo ou metadata

    def __repr__(self):
        return f"<Rate value={self.value}, kind={self.kind}, scenario={self.scenario}>"
