import time
import pandas as pd
from services.api_service import APIService
from services.excel_service import ExcelService
from models.client import Client
from models.operation import Operation
from models.record import Record

class ProcessingService:
    def __init__(self, dequeue_method, upload_result, excel_process=None):
        """
        Parameters
        ----------
        dequeue_method : callable
            Função que remove e retorna o próximo arquivo da fila (ex: queue.get).
        upload_result : callable
            Função que envia o resultado ao SharePoint.
        excel_process : subprocess.Popen, opcional
            Processo externo do LibreOffice já iniciado (reutilizável entre serviços).
        """
        self.dequeue_method = dequeue_method
        self.upload_result = upload_result
        self.excel_process = excel_process
        self.api_service = APIService()

        # Conecta à planilha
        self.excel_service = ExcelService("/home/gucampe/Documentos/TCC/Projeto/masterFunding.xlsx")
        print("[PROCESSING] Serviço inicializado.")

    def start(self):
        print("[PROCESSING] Iniciando...")
        # Garante que o LibreOffice esteja acessível (ExcelService faz isso internamente)
        self.excel_service.connect()

        while True:
            try:
                # Aguarda o próximo arquivo
                file_path = self.dequeue_method()
                print(f"[PROCESSING] Novo arquivo recebido: {file_path}")

                # Processa o CSV
                self.process_file(file_path)

            except Exception as e:
                print(f"[PROCESSING] Erro: {e}")
                time.sleep(5)

    def process_file(self, file_path: str):
        print(f"[PROCESSING] Lendo arquivo {file_path}...")
        df = pd.read_csv(file_path).fillna("")

        for _, row in df.iterrows():
            client = Client(row["ID"], row["CLIENTE"], row["CPF_CNPJ"], row["SEGMENTO"], row["RATING"])
            operation = Operation(
                row["ID"], client, row["VALOR"], row["PRAZO_DIAS"],
                row["TIPO_TAXA"], row["SPREAD_SOLC"], row["CUSTO_SOLC"], row['TAXA_SOLC'], row["FLUXO_PARCELAS"]
            )
            record = Record(
                email_solc=row["EMAIL_SOLC"], operation_id=row["ID"],
                status="PENDENTE", solicitante=row["SOLICITANTE"],
                rate=None, justification=None
            )

            print(f"[PROCESSING] Validando operação {operation.nrm_po} do cliente {client.name}")

            # Validação simulada (substituir futuramente pela chamada real à API)
            is_valid, motivo = True, "Aprovado para teste"

            if is_valid:
                print(f"[PROCESSING] Operação {record.operation_id} aprovada. Motivo: {motivo}")

                self.excel_service.preencher_dados(operation.value, operation.rate_type, operation.parcel_flow)

                cost_approved = round(self.excel_service.rodar_macro(), 4)
                operation.calculate_rate(cost_approved)
                
                if cost_approved != operation.cost_requested:
                    record.status = "APROVADO COM ALTERAÇÃO"
                else:
                    record.status = "APROVADO"
                
                print(f"[PROCESSING] ✅ Cliente {client.name} - Operação {operation.nrm_po} - Taxa {operation.rate_approved * 100}%")
            else:
                print(f"[PROCESSING] ❌ Operação {record.operation_id} recusada. Motivo: {motivo}")
                record.status = "RECUSADO"
                record.justification = motivo

            try:
                self.upload_result(client, operation, record)
                print(f"[PROCESSING] Resultado enviado para SharePoint: {record.operation_id}")
            except Exception as e:
                print(f"[PROCESSING] Erro ao enviar para SharePoint: {e}")

        print("[PROCESSING] Arquivo concluído.")
