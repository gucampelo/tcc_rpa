import threading
from services.sharepoint_service import SharePointService
from services.watchdog_service import WatchdogService
#from services.processing_service import ProcessingService
import settings
import queue

data_queue = queue.Queue()

def start_sharepoint():
    sp = SharePointService(download_path=settings.PASTA_DOWNLOAD, url_sharepoint=settings.URL_SHAREPOINT, usuario=settings.USUARIO, senha=settings.SENHA)
    sp.run_download(interval=20)   # loop infinito de extração

def start_watchdog():
    wd = WatchdogService(path=settings.PASTA_DOWNLOAD, data_queue=data_queue)
    wd.run()   # observa a pasta e dispara eventos

def start_processing():
    #proc = ProcessingService()
    #proc.run()   # valida dados, roda excel, grava no sharepoint
    ...

if __name__ == "__main__":
    t1 = threading.Thread(target=start_sharepoint)
    t2 = threading.Thread(target=start_watchdog)
    t3 = threading.Thread(target=start_processing)

    t1.start()
    t2.start()
    t3.start()

    t1.join()
    t2.join()
    t3.join()
