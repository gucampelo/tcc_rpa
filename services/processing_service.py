import time
import pandas as pd
from services.validate_service import ValidateService
from services.excel_service import ExcelService
from models.client import Client
from models.operation import Operation
from models.record import Record
import settings

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
        self.validate_service = ValidateService()

        # Conecta à planilha
        self.excel_service = ExcelService(settings.EXCEL_FILE)
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

    def get_pending_operations(self, df):
         # Agrupa por NMR_PO e mantém apenas grupos com 1 registro
        df_pending = df.groupby('NMR_PO').filter(lambda g: len(g) == 1)
        return df_pending

    def process_file(self, file_path: str):
        print(f"[PROCESSING] Lendo arquivo {file_path}...")
        df = pd.read_csv(file_path).fillna("")

        df_pending = self.get_pending_operations(df)
        print(f"[INFO] Encontradas {len(df_pending)} operações pendentes.")


        for _, row in df_pending.iterrows():
            client = Client(nmr_po = row["NMR_PO"], name= row["CLIENTE"], 
                cpf_cnpj=row["CPF_CNPJ"], segment=row["SEGMENTO"], rating=row["RATING"])
            operation = Operation(
                nmr_po = row["NMR_PO"],client = client, product=row['PRODUTO'],operation_type=row['TIPO_OPERACAO'], 
                value=row["VALOR"], guarantee=row['GARANTIA'], guarantee_percentage=row['PORCEN_GARANTIA'],
                term_days=row["PRAZO_DIAS"], rate_type=row["TIPO_TAXA"], spread_requested=row["SPREAD_SOLC"], cost_requested=row["CUSTO_SOLC"], rate_requested=row['TAXA_SOLC'], parcel_flow=row["FLUXO_PARCELAS"], trade_defense=row["DEFESA_COMERCIAL"]
            )
            record = Record(
                email_solc=row["EMAIL_SOLC"], nmr_po=row["NMR_PO"],
                status="PENDENTE", requester=row["SOLICITANTE"],justification=None
            )

            print(f"[PROCESSING] Validando operação {operation.nmr_po} do cliente {client.name}")

            # Validação simulada (substituir futuramente pela chamada real à API)
            response = self.validate_service.validate_operation(client, operation)
            is_valid, status = response["valido"], response["motivo"]
            
            if is_valid:
                print(f"[PROCESSING] Operação {operation.nmr_po}: Validada")

                self.excel_service.preencher_dados(operation.value, operation.rate_type, operation.parcel_flow)

                cost_approved = round(self.excel_service.rodar_macro(), 4)
                operation.calculate_rate(cost_approved)
                
                if cost_approved != operation.cost_requested:
                    record.status = "APROVADO COM ALTERAÇÃO"
                else:
                    record.status = "APROVADO"
                record.justification = status
                print(f"[PROCESSING] ✅ Cliente {client.name} - Operação {operation.nmr_po} - Taxa {operation.rate_approved * 100}%")
            else:
                print(f"[PROCESSING] ❌ Operação {record.nmr_po}: Rejeitada")
                record.status = "RECUSADO"
                record.justification = status

            try:
                self.upload_result(client, operation, record)
                print(f"[PROCESSING] Resultado enviado para SharePoint: {record.nmr_po} - {record.status}")
            except Exception as e:
                print(f"[PROCESSING] Erro ao enviar para SharePoint: {e}")

        print("[PROCESSING] Arquivo concluído.")
