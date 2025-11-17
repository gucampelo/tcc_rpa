class Client:
    def __init__(self, nmr_po, name, cpf_cnpj, segment, rating):
        self.nmr_po = nmr_po
        self.name = name
        self.cpf_cnpj = cpf_cnpj
        self.segment = segment
        self.rating = rating

    def __repr__(self):
        return f"<Client {self.name} - Rating: {self.rating}>"
