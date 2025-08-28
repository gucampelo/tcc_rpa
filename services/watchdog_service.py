# services/watchdog_service.py
import time
import queue
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pandas as pd
import os

class WatchdogHandler(FileSystemEventHandler):
    def __init__(self, data_queue):
        super().__init__()
        self.data_queue = data_queue

    def on_created(self, event):
        """Dispara quando um novo arquivo é criado"""
        if not event.is_directory and event.src_path.endswith(".csv"):
            print(f"[Watchdog] Novo CSV detectado: {event.src_path}")
            # dá um tempinho pro arquivo terminar de salvar
            time.sleep(2)
            try:
                df = pd.read_csv(event.src_path, sep=",")
                self.data_queue.put(df)
                print(f"[Watchdog] CSV enviado para fila de processamento ({len(df)} linhas).")
            except Exception as e:
                print(f"[Watchdog] Erro ao ler CSV: {e}")

class WatchdogService:
    def __init__(self, path: str, data_queue: queue.Queue):
        self.path = path
        self.data_queue = data_queue
        self.observer = Observer()

    def run(self):
        event_handler = WatchdogHandler(self.data_queue)
        self.observer.schedule(event_handler, self.path, recursive=False)
        self.observer.start()
        print(f"[Watchdog] Observando pasta: {self.path}")

        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            self.stop()

    def stop(self):
        self.observer.stop()
        self.observer.join()
