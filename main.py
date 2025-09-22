import threading
import subprocess
import time
import queue

from services.sharepoint_service import SharePointService
from services.watchdog_service import WatchdogService
from services.processing_service import ProcessingService
from services.excel_service import ExcelService
from services.api_service import APIService
import settings

# -----------------------------
# Inicializa LibreOffice headless
# -----------------------------
subprocess.Popen([
    "libreoffice",
    "--accept=socket,host=localhost,port=2002;urp;",
    "--norestore", "--nofirststartwizard", "--nologo"
])

time.sleep(3)  # tempo para o LibreOffice iniciar

# -----------------------------
# Filas compartilhadas
# -----------------------------
data_queue = queue.Queue()       # CSVs baixados para processamento
driver_lock = threading.Lock()   # lock para sincronizar Selenium

# -----------------------------
# Instância única do SharePoint
# -----------------------------
sp_service = SharePointService(
    download_path=settings.PASTA_DOWNLOAD,
    url_sharepoint=settings.URL_SHAREPOINT,
    usuario=settings.USUARIO,
    senha=settings.SENHA
)
sp_service._init_driver()
sp_service._login()

# -----------------------------
# Thread: download de CSV do SharePoint
# -----------------------------
def start_sharepoint():
    while True:
        try:
            df = sp_service._download_csv()
            if df is not None:
                data_queue.put(df)  # envia para processamento
                print("[SP] CSV baixado e colocado na fila.")
        except Exception as e:
            print(f"[SP] Erro ao baixar CSV: {e}")
        time.sleep(20)

# -----------------------------
# Thread: observa a pasta de downloads (Watchdog)
# -----------------------------
def start_watchdog():
    wd = WatchdogService(
        path=settings.PASTA_DOWNLOAD,
        data_queue=data_queue
    )
    wd.run()

# -----------------------------
# Thread: processa CSVs, calcula taxa e envia ao SharePoint
# -----------------------------
def start_processing():
    excel_service = ExcelService("/home/gucampe/Documentos/TCC/Projeto/masterFunding.xlsx")
    api_service = APIService()
    
    proc = ProcessingService(
        task_queue=data_queue,
        excel_service=excel_service,
        api_service=api_service,
        sharepoint_service=sp_service  # <-- passa a instância compartilhada
    )
    proc.start()  # valida dados, roda Excel e grava no SharePoint

# -----------------------------
# Main
# -----------------------------
if __name__ == "__main__":
    threads = [
        threading.Thread(target=start_sharepoint, daemon=True),
        threading.Thread(target=start_watchdog, daemon=True),
        threading.Thread(target=start_processing, daemon=True)
    ]

    for t in threads:
        t.start()

    # Mantém o main thread ativo
    for t in threads:
        t.join()
