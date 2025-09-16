import threading
from services.sharepoint_service import SharePointService
from services.watchdog_service import WatchdogService
from services.processing_service import ProcessingService
from services.excel_service import ExcelService
from services.api_service import APIService
import settings
import queue

# Fila compartilhada
data_queue = queue.Queue()

def start_sharepoint():
    sp = SharePointService(
        download_path=settings.PASTA_DOWNLOAD,
        url_sharepoint=settings.URL_SHAREPOINT,
        usuario=settings.USUARIO,
        senha=settings.SENHA
    )
    sp.run_download(interval=20)   # loop infinito de extração

def start_watchdog():
    wd = WatchdogService(
        path=settings.PASTA_DOWNLOAD,
        data_queue=data_queue
    )
    wd.run()   # observa a pasta e dispara eventos

def start_processing():
    excel_service = ExcelService()
    api_service = APIService()
    
    proc = ProcessingService(
        task_queue=data_queue,
        excel_service=excel_service,
        api_service=api_service
    )
    proc.start()   # valida dados, roda excel, grava no sharepoint

if __name__ == "__main__":
    t1 = threading.Thread(target=start_sharepoint, daemon=True)
    t2 = threading.Thread(target=start_watchdog, daemon=True)
    t3 = threading.Thread(target=start_processing, daemon=True)

    t1.start()
    t2.start()
    t3.start()

    t1.join()
    t2.join()
    t3.join()
