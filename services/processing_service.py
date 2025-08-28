# services/processing_service.py
import time
from models.client import Client
from models.operation import Operation
from models.rate import Rate
from models.record import Record

class ProcessingService:
    def __init__(self, data_queue, api_service, excel_service, sharepoint_service):
        self.data_queue = data_queue
        self.api_service = api_service
        self.excel_service = excel_service
        self.sharepoint_service = sharepoint_service

    def run(self):
        print("[ProcessingService] Iniciado.")
        while True:
            df = self.data_queue.get()  # bloqueante até ter algo
            print(f"[ProcessingService] Recebido dataframe ({len(df)} linhas).")
            
            for _, row in df.iterrows():
                try:
                    # 1. Criar cliente
                    client = Client(
                        client_id=row.get("ID_CLIENTE", ""),
                        name=row.get("CLIENTE", ""),
                        cpf_cnpj=row.get("CPF_CNPJ", ""),
                        segment=row.get("SEGMENTO", ""),
                        rating=row.get("RATING", 0)
                    )

                    # 2. Criar operação
                    operation = Operation(
                        operation_id=row.get("ID", 0),
                        client=client,
                        value=row.get("VALOR", 0.0),
                        term_days=row.get("PRAZO_DIAS", 0),
                        rate_type=row.get("TIPO_TAXA", ""),
                        spread_req=row.get("SPREAD_SOLC", 0.0),
                        cost_req=row.get("CUSTO_SOLC", 0.0),
                        rate_req=row.get("TAXA_SOLC", None),
                        spread_locked=row.get("SPREAD_TRV", None),
                        cost_locked=row.get("CUSTO_TRV", None),
                        rate_locked=row.get("TAXA_TRV", None),
                        installment_flow=row.get("FLUXO_PARCELAS", None),
                        requester=row.get("SOLICITANTE", ""),
                        requester_email=row.get("EMAIL_SOLC", "")
                    )

                    # 3. Validar cliente
                    if not self.api_service.validate_client(client):
                        print(f"[ProcessingService] Cliente inválido: {client.client_id}")
                        record = Record.from_operation(operation, None)
                        record.status = "RECUSADO"
                        record.justification = "Cliente inválido"
                        self.sharepoint_service._upload_item(record)
                        continue

                    # 4. Validar operação
                    if not self.api_service.validate_operation(operation):
                        print(f"[ProcessingService] Operação inválida: {operation.operation_id}")
                        record = Record.from_operation(operation, None)
                        record.status = "RECUSADO"
                        record.justification = "Operação inválida"
                        self.sharepoint_service._upload_item(record)
                        continue

                    # 5. Calcular taxa
                    rate = self.excel_service.calculate_rate(operation)

                    # 6. Criar record e enviar para SharePoint
                    record = Record.from_operation(operation, rate)
                    record.status = "APROVADO"
                    self.sharepoint_service._upload_item(record)
                    print(f"[ProcessingService] Operação processada: {operation.operation_id}")

                except Exception as e:
                    print(f"[ProcessingService] Erro ao processar linha: {e}")

            self.data_queue.task_done()
            print("[ProcessingService] Lote processado.\n")
