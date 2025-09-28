# services/watchdog_service.py
import time
import queue
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pandas as pd
import os

class WatchdogHandler(FileSystemEventHandler):
    def __init__(self, enqueue_method):
        super().__init__()
        self.enqueue_method = enqueue_method

    def on_created(self, event):
        """Dispara quando um novo arquivo é criado"""
        if not event.is_directory and event.src_path.endswith(".csv"):
            print(f"[Watchdog] Novo CSV detectado: {event.src_path}")
            # dá um tempinho pro arquivo terminar de salvar
            time.sleep(2)
            try:
                self.enqueue_method(event.src_path)
                print(f"[Watchdog] CSV enviado para fila de processamento ({event.src_path} linhas).")
            except Exception as e:
                print(f"[Watchdog] Erro ao ler CSV: {e}")

class WatchdogService:
    def __init__(self, path: str, enqueue_method):
        self.path = path
        self.enqueue_method = enqueue_method
        self.observer = Observer()

    def run(self):
        event_handler = WatchdogHandler(self.enqueue_method)
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
