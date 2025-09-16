class Client:
    def __init__(self, client_id, name, cpf_cnpj, segmento, rating):
        self.client_id = client_id
        self.name = name
        self.cpf_cnpj = cpf_cnpj
        self.segmento = segmento
        self.rating = rating

    def __repr__(self):
        return f"<Client {self.name} - Rating: {self.rating}>"
