class Client:
    def __init__(self, nmr_po, name, cpf_cnpj, segment, rating):
        self._name = name
        self._cpf_cnpj = cpf_cnpj
        self._segment = segment
        self._rating = rating

    def __repr__(self):
        return f"<Client {self.name} - Rating: {self.rating}>"


    @property
    def name(self):
        return self._name

    @name.setter
    def name(self, value):
        self._name = value

    @property
    def cpf_cnpj(self):
        return self._cpf_cnpj

    @cpf_cnpj.setter
    def cpf_cnpj(self, value):
        self._cpf_cnpj = value

    @property
    def segment(self):
        return self._segment

    @segment.setter
    def segment(self, value):
        self._segment = value

    @property
    def rating(self):
        return self._rating

    @rating.setter
    def rating(self, value):
        self._rating = value
