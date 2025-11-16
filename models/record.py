# models/record.py
from datetime import datetime


class Record:
    """
    Representa a resposta final a ser enviada ao SharePoint.
    """
    def __init__(self, nmr_po, requester, email_solc, status,
                 rate, justification):
        self.__nmr_po = nmr_po      # ID da operação
        self.__requester = requester
        self.__email_solc = email_solc
                          
        self.__status = status                  # 'APROVADO', 'RECUSADO', 'ERRO'                  # instância de Rate ou None
        self.__justification = justification    # motivo de recusa ou erro
        self.__processed_at = datetime.now()
        self.__processed_by = "RPA"

    @property
    def nmr_po(self):
        return self.__nmr_po

    @nmr_po.setter
    def nmr_po(self, value):
        self.__nmr_po = value

    @property
    def requester(self):
        return self.__requester

    @requester.setter
    def requester(self, value):
        self.__requester = value

    @property
    def email_solc(self):
        return self.__email_solc

    @email_solc.setter
    def email_solc(self, value):
        self.__email_solc = value

    @property
    def status(self):
        return self.__status

    @status.setter
    def status(self, value):
        self.__status = value

    @property
    def rate(self):
        return self.__rate

    @rate.setter
    def rate(self, value):
        self.__rate = value

    @property
    def justification(self):
        return self.__justification

    @justification.setter
    def justification(self, value):
        self.__justification = value

    @property
    def processed_at(self):
        return self.__processed_at

    @processed_at.setter
    def processed_at(self, value):
        self.__processed_at = value

    @property
    def processed_by(self):
        return self.__processed_by

    @processed_by.setter
    def processed_by(self, value):
        self.__processed_by = value

    # ...restante da classe...

    @classmethod
    def from_operation(cls, operation, rate=None):
        """
        Cria um Record a partir de uma operação e uma Rate calculada (opcional).
        """
        return cls(
            nmr_po=operation.nmr_po,
            status="PENDENTE",
            rate=rate,
            justification=None
        )

    def __repr__(self):
        rate_value = self.rate.value if self.rate else None
        return (f"<Record nmr_po={self.nmr_po}, "
                f"status={self.status}, rate={rate_value}>")
