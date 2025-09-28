# services/processing_service.py
import time, subprocess, pandas as pd
from services.api_service import APIService
from services.excel_service import ExcelService
from models.client import Client
from models.operation import Operation
from models.record import Record

class ProcessingService:
    def __init__(self, dequeue_method, upload_result):
        self.dequeue_method = dequeue_method
        self.upload_result = upload_result

        subprocess.Popen([
            "libreoffice",
            "--accept=socket,host=localhost,port=2002;urp;",
            "--norestore", "--nofirststartwizard", "--nologo"
        ])
        time.sleep(3)  # tempo para o LibreOffice iniciar
        self.excel_service = ExcelService("/home/gucampe/Documentos/TCC/Projeto/masterFunding.xlsx")
        self.api_service = APIService()

    def start(self):
        print("[PROCESSING] Service started...")
        # Conecta uma vez ao Excel/LibreOffice
        self.excel_service.connect()

        while True:
            try:
                # Espera até que um novo arquivo apareça na fila
                file_path = self.dequeue_method()
                print(f"[PROCESSING] Novo arquivo recebido: {file_path}")

                # Processa o CSV
                self.process_file(file_path)

                # Marca como concluído na fila
                self.task_queue.task_done()

            except Exception as e:
                print(f"[PROCESSING] Erro: {e}")
                time.sleep(5)  # espera um pouco antes de tentar de novo

    def process_file(self, file_path: str):
        print(f"[PROCESSING] Lendo arquivo {file_path}...")
        df = pd.read_csv(file_path)

        # Pré-processamento simples: substitui NaN
        df = df.fillna("")

        # Itera sobre cada linha do CSV
        for _, row in df.iterrows():
            client = Client(row["ID"], row["CLIENTE"], row["CPF_CNPJ"], row["SEGMENTO"], row["RATING"])
            operation = Operation(
                row["ID"],
                client,
                row["VALOR"],
                row["PRAZO_DIAS"],
                row["TIPO_TAXA"],
                row["SPREAD_SOLC"],
                row["CUSTO_SOLC"],
                row["FLUXO_PARCELAS"]
            )
            record = Record(email_solc=row["EMAIL_SOLC"], operation_id=row["ID"], status="PENDENTE", solicitante=row["SOLICITANTE"], rate=None, justification=None )

            print(f"[PROCESSING] Validando operação {operation.operation_id} do cliente {client.name}")

            # Validação via API (temporariamente sempre True)
            is_valid, motivo = True, "Aprovado para teste"
            # Para produção: is_valid, motivo = self.api_service.validate(record)

            if is_valid:
                print(f"[PROCESSING] Operação {record.operation_id} aprovada! Motivo: {motivo}")

                # Cálculo da taxa via Excel
                # Espera-se que row['FLUXO_PARCELAS'] seja lista de parcelas no formato:
                # [{"data": "10/03/26", "principal": 250000}, ...]
                
                self.excel_service.preencher_dados(operation.value, operation.rate_type, operation.parcel_flow)

                # Rodar a macro "GerarDecimal" e recuperar valor
                cost = self.excel_service.rodar_macro("GerarDecimal")
                record.rate = cost

                print(f"[PROCESSING] ✅ Cliente {client.name} - Operação {operation.operation_id} - Taxa {cost}%")
            else:
                print(f"[PROCESSING] Operação {record.operation_id} recusada. Motivo: {motivo}")
                record.status = "RECUSADO"
                record.justification = motivo
                # Aqui você pode chamar SharePointService para registrar o status

            try:
                self.upload_result(client, operation, record)
                print(f"[PROCESSING] Registro enviado para SharePoint: Operação {record.operation_id}")
            except Exception as e:
                print(f"[PROCESSING] Erro ao enviar para SharePoint: {e}")

        print("[PROCESSING] Arquivo concluído.")
