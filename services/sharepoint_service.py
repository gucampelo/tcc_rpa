# services/sharepoint_service.py
import time
import threading
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class SharePointService:
    def __init__(self, download_path: str, url_sharepoint, usuario, senha):
        self.download_path = download_path
        self.url_sharepoint = url_sharepoint
        self.usuario = usuario
        self.senha = senha
        self.driver = None
        self.lock = threading.Lock()  # lock para sincronizar Selenium

    def _init_driver(self):
        options = webdriver.ChromeOptions()
        prefs = {"download.default_directory": self.download_path}
        options.add_experimental_option("prefs", prefs)
        options.add_argument("--start-maximized")
        self.driver = webdriver.Chrome(options=options)

    def _login(self):
        self.driver.get(self.url_sharepoint)
        # Login SharePoint
        WebDriverWait(self.driver, 20).until(
            EC.presence_of_element_located((By.ID, "usernameEntry"))
        ).send_keys(self.usuario + Keys.ENTER)

        WebDriverWait(self.driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Outras maneiras de entrar')]"))
        ).click()
        WebDriverWait(self.driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Use sua senha')]"))
        ).click()

        WebDriverWait(self.driver, 20).until(
            EC.presence_of_element_located((By.ID, "passwordEntry"))
        ).send_keys(self.senha + Keys.ENTER)

        WebDriverWait(self.driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@data-testid="primaryButton"]'))
        ).click()

    def _download_csv(self):
        with self.lock:
            time.sleep(5)
            iframe = self.driver.find_element(By.TAG_NAME, "iframe")
            self.driver.switch_to.frame(iframe)
            try:
                export_btn = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@data-automationid="exportListCommand"]'))
                )
                export_btn.click()
                export_csv_btn = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@data-automationid="exportListToCSVCommand"]'))
                )
                export_csv_btn.click()
                print("[SP] Exportação iniciada...")
            except Exception as e:
                print(f"[SP] Erro ao clicar no botão de exportação: {e}")
            finally:
                self.driver.switch_to.default_content()

    def _upload_item(self, client, operation, record):
        """
        Registra um item na lista SharePoint.
    : dic campos {"Cliente": ..., "Operacao": ..., "Taxa": ..., "Status": ...}
        """
        with self.lock:
            try:
                iframe = self.driver.find_element(By.TAG_NAME, "iframe")
                self.driver.switch_to.frame(iframe)
                # Exemplo: navegar até a lista de criação de item
                new_item_btn = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[1]/div[1]/div/div/span[1]/button[1]'))
                )
                new_item_btn.click()

               
            except Exception as e:
                print(f"[SP] Erro ao registrar item: {e}")
            finally:
                self.driver.switch_to.default_content()

    def run_download(self, interval: int = 20):
        self._init_driver()
        self._login()
        while True:
            try:
                self._download_csv()
            except Exception as e:
                print(f"[SP] Erro ao baixar CSV: {e}")
            time.sleep(interval)

    def run_upload(self, data_queue):
        while True:
            df = data_queue.get()  # Pega os dados mais recentes
            for _, row in df.iterrows():
                item_data = {
                    "Cliente": row["CLIENTE"],
                    "Operacao": row["ID"],
                    "Taxa": row["TAXA"],
                    "Status": row["STATUS"]
                }
                self._upload_item(item_data)
            print("[SP] Todos os itens enviados.")
            data_queue.task_done()
