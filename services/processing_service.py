# services/processing_service.py
import pandas as pd
import queue
import time
from services.api_service import APIService
from services.excel_service import ExcelService
from models.client import Client
from models.operation import Operation
from models.record import Record

class ProcessingService:
    def __init__(self, task_queue: queue.Queue, api_service: APIService, excel_service: ExcelService):
        self.task_queue = task_queue
        self.api_service = api_service
        self.excel_service = excel_service

    def start(self):
        print("[PROCESSING] Service started...")
        while True:
            try:
                # Espera até que um novo arquivo apareça na fila
                file_path = self.task_queue.get()
                print(f"[PROCESSING] Novo arquivo recebido: {type(file_path)}")

                # Processa CSV
                self.process_file(file_path)

                # Marca como concluído na fila
                self.task_queue.task_done()

            except Exception as e:
                print(f"[PROCESSING] Erro: {e}")
                time.sleep(5)  # espera um pouco antes de tentar de novo

    def process_file(self, file_path: str):
        print(f"[PROCESSING] Lendo arquivo {file_path}...")
        df = pd.read_csv(file_path)

        # Pré-processamento simples
        df = df.fillna("")  # substitui NaN

        # Iterar sobre linhas do CSV
        for _, row in df.iterrows():
            client = Client(row["ID"], row["CLIENTE"], row["CPF_CNPJ"], row["SEGMENTO"], row["RATING"])
            operation = Operation(row["ID"], client, row["VALOR"], row["PRAZO_DIAS"], row["TIPO_TAXA"], row["SPREAD_SOLC"], row["CUSTO_SOLC"])
            record = Record(client, operation)

            print(f"[PROCESSING] Validando operação {operation.operation_id} do cliente {client.name}")

            # Validação via API Flask
            is_valid, motivo = self.api_service.validate(record)

            if is_valid:
                print(f"[PROCESSING] Operação {record.operation_id} aprovada! Motivo: {motivo}")
                # aqui chama o ExcelService para calcular taxa
            else:
                print(f"[PROCESSING] Operação {record.operation_id} recusada. Motivo: {motivo}")
                # aqui você já pode registrar no SharePoint como recusada
            # Cálculo da taxa via Excel
            rate = self.excel_service.calculate_rate(operation.amount, client.rating)
            record.set_rate(rate)

            print(f"[PROCESSING] ✅ Cliente {client.name} - Operação {operation.operation_id} - Taxa {rate}%")

        print("[PROCESSING] Arquivo concluído.")
