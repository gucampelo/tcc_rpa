class Client:
    def __init__(self, client_id, name, cpf_cnpj, segment, rating):
        self.name = name
        self.cpf_cnpj = cpf_cnpj
        self.segment = segment
        self.rating = rating

    def __repr__(self):
        return f"<Client {self.name} - Rating: {self.rating}>"
