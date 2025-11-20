import time
import pandas as pd
from services.validate_service import ValidateService
from services.excel_service import ExcelService
from models.client import Client
from models.operation import Operation
from models.record import Record
import settings, os, json

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

        # Arquivo de cache dos IDs já processados
        self.cache_file = "processed_ids.json"
        self.processed_ids = self.load_cache()

        # Conecta à planilha
        self.excel_service = ExcelService(settings.EXCEL_FILE)
        print("[PROCESSING] Serviço inicializado.")

    # -------------------------
    # Cache de IDs processados
    # -------------------------
    def load_cache(self):
        if not os.path.exists(self.cache_file):
            return set()
        try:
            with open(self.cache_file, "r") as f:
                return set(json.load(f))
        except:
            print("[CACHE] Erro ao ler cache. Criando novo…")
            return set()

    def save_cache(self):
        try:
            with open(self.cache_file, "w") as f:
                json.dump(list(self.processed_ids), f)
        except Exception as e:
            print(f"[CACHE] Erro ao salvar cache: {e}")

    # -------------------------

    def start(self):
        print("[PROCESSING] Iniciando…")
        self.excel_service.connect()

        while True:
            try:
                file_path = self.dequeue_method()
                #print(f"[PROCESSING] Novo arquivo recebido: {file_path}")
                self.process_file(file_path)

            except Exception as e:
                print(f"[PROCESSING] Erro: {e}")
                time.sleep(5)

    def get_pending_operations(self, df):
        """
        Retorna apenas grupos de NMR_PO com 1 registro.
        """
        return df.groupby("NMR_PO").filter(lambda g: len(g) == 1)

    def process_file(self, file_path: str):
        print(f"[PROCESSING] Lendo arquivo {file_path}…")
        df = pd.read_csv(file_path, dtype={"CPF_CNPJ": str}).fillna("")

        # Primeiro filtro: grupos com 1 registro apenas
        df_pending = self.get_pending_operations(df)
        #print(f"[INFO] Encontradas {len(df_pending)} operações pendentes.") if df_pending.size > 0 else None 

        # Segundo filtro: remover IDs já processados
        df_pending = df_pending[~df_pending["ID"].isin(self.processed_ids)]
        #print(f"[INFO] Encontradas: {len(df_pending)} operações restantes.")

        for _, row in df_pending.iterrows():

            # ID da solicitação
            record_id = str(row["ID"])

            # Segurança: caso ainda esteja no cache
            if record_id in self.processed_ids:
                #print(f"[CACHE] Ignorando operação ID {record_id} (já processada).")
                continue

            # Criar modelo Client
            client = Client(
                nmr_po=row["NMR_PO"],
                name=row["CLIENTE"],
                cpf_cnpj=row["CPF_CNPJ"],
                segment=row["SEGMENTO"],
                rating=row["RATING"]
            )

            # Criar modelo Operation
            operation = Operation(
                nmr_po=row["NMR_PO"], client=client,
                product=row['PRODUTO'], operation_type=row['TIPO_OPERACAO'],
                value=row["VALOR"], guarantee=row['GARANTIA'],
                guarantee_percentage=row['PORCEN_GARANTIA'],
                term_days=row["PRAZO_DIAS"], rate_type=row["TIPO_TAXA"],
                spread_requested=row["SPREAD_SOLC"], cost_requested=row["CUSTO_SOLC"],
                rate_requested=row['TAXA_SOLC'], parcel_flow=row["FLUXO_PARCELAS"],
                trade_defense=row["DEFESA_COMERCIAL"]
            )

            # Criar modelo Record
            record = Record(
                email_solc=row["EMAIL_SOLC"],
                nmr_po=row["NMR_PO"],
                status="PENDENTE",
                requester=row["SOLICITANTE"],
                justification=None
            )

            print(f"[PROCESSING] Validando operação {operation.nmr_po} - Cliente {client.name}")

            response = self.validate_service.validate_operation(client, operation)
            is_valid, status = response["valido"], response["motivo"]

            if is_valid:
                print(f"[PROCESSING] Operação {operation.nmr_po}: Validada")

                self.excel_service.preencher_dados(
                    operation.value, operation.rate_type, operation.parcel_flow
                )
                cost_approved = round(self.excel_service.rodar_macro(), 4)
                operation.calculate_rate(cost_approved)

                if cost_approved != operation.cost_requested:
                    record.status = "APROVADO COM ALTERAÇÃO"
                else:
                    record.status = "APROVADO"

                record.justification = status

                print(f"[PROCESSING] ✅ {client.name} - PO {operation.nmr_po} - Taxa {operation.rate_approved * 100}%")

            else:
                print(f"[PROCESSING] ❌ Operação {record.nmr_po}: Rejeitada")
                record.status = "RECUSADO"
                record.justification = status

            try:
                # Envia resultado ao SharePoint
                self.upload_result(client, operation, record)

                # Marca como processada
                self.processed_ids.add(record_id)
                self.save_cache()

                print(f"[PROCESSING] Resultado enviado e cache atualizado para ID {record_id}")

            except Exception as e:
                print(f"[PROCESSING] Erro ao enviar ao SharePoint: {e}")

        print("[PROCESSING] Arquivo concluído.")
