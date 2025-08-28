# models/record.py
from datetime import datetime

class Record:
    """
    Representa a resposta final a ser enviada ao SharePoint.
    """
    def __init__(self, operation_id: int, status: str,
                 final_rate=None, justification: str = None):
        self.operation_id = operation_id      # ID da operação
        self.status = status                  # 'APROVADO', 'RECUSADO', 'ERRO'
        self.final_rate = final_rate          # instância de Rate ou None
        self.justification = justification    # motivo de recusa ou erro
        self.processed_at = datetime.now()
        self.processed_by = "RPA"

    @classmethod
    def from_operation(cls, operation, rate=None):
        """
        Cria um Record a partir de uma operação e uma Rate calculada (opcional).
        """
        return cls(
            operation_id=operation.operation_id,
            status="PENDENTE",
            final_rate=rate,
            justification=None
        )

    def __repr__(self):
        rate_value = self.final_rate.value if self.final_rate else None
        return (f"<Record operation_id={self.operation_id}, "
                f"status={self.status}, rate={rate_value}>")
