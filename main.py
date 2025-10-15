import threading, time, queue, settings
from services.sharepoint_service import SharePointService
from services.watchdog_service import WatchdogService
from services.processing_service import ProcessingService


class RPA:
    def __init__(self):
        
        # -----------------------------
        # Fila
        # -----------------------------
        self.__data_queue = queue.Queue()
     

        # -----------------------------
        # Instância única do SharePoint
        # -----------------------------
        self.__sharepoint_service = SharePointService(
            download_path=settings.DOWNLOAD_DIR,
            url_sharepoint=settings.URL_SHAREPOINT,
            user=settings.USER,
            password=settings.PASSWORD,
        )
    
        self.__watchdog_service = WatchdogService(
            path=settings.DOWNLOAD_DIR,
            enqueue_method=self.enqueue,
        )

        self.__processing_service = ProcessingService(
            dequeue_method=self.dequeue,
            upload_result=self.__sharepoint_service.upload_item,
        )

    # -----------------------------
    # Thread: download de CSV do SharePoint
    # -----------------------------
    @property
    def data_queue(self):
        return self.__data_queue

    # Método público para adicionar um item na fila
    def enqueue(self, item):
        self.__data_queue.put(item)

    def dequeue(self):
        return self.__data_queue.get()
    
    def start_sharepoint(self):
        self.__sharepoint_service.connect()
        while True:
            try:
                self.__sharepoint_service.download_csv()
            except Exception as e:
                print(f"[SP] Erro ao baixar CSV: {e}")
            time.sleep(20)

    # -----------------------------
    # Thread: observa a pasta de downloads
    # -----------------------------
    def start_watchdog(self):
        self.__watchdog_service.run()

    # -----------------------------
    # Thread: processa CSVs, calcula taxa e envia ao SharePoint
    # -----------------------------
    def start_processing(self):
        self.__processing_service.start()

    # -----------------------------
    # Inicializa as threads
    # -----------------------------
    def run(self):
        threads = [
            threading.Thread(target=self.start_sharepoint, daemon=True),
            threading.Thread(target=self.start_watchdog, daemon=True),
            threading.Thread(target=self.start_processing, daemon=True)
        ]

        for t in threads:
            t.start()

        # Mantém o main thread ativo
        for t in threads:
            t.join()


# -----------------------------
# Entry point
# -----------------------------
if __name__ == "__main__":
    rpa = RPA()
    rpa.run()
