# services/sharepoint_service.py
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd

class SharePointService:
    def __init__(self, download_path: str, url_sharepoint, usuario, senha):
        self.download_path = download_path
        self.driver = None
        self.url_sharepoint = url_sharepoint
        self.usuario = usuario
        self.senha = senha

    def _init_driver(self):
        options = webdriver.ChromeOptions()
        prefs = {"download.default_directory": self.download_path}
        options.add_experimental_option("prefs", prefs)
        options.add_argument("--start-maximized")
        self.driver = webdriver.Chrome(options=options)

    def _login(self):
        self.driver.get(self.url_sharepoint)
        #time.sleep(500)
        # Login
        WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.ID, "usernameEntry"))).send_keys(self.usuario + Keys.ENTER)
        WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Outras maneiras de entrar')]"))).click()
        WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Use sua senha')]"))).click()

        WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.ID, "passwordEntry"))).send_keys(self.senha + Keys.ENTER)
        self.driver.find_element(By.XPATH, '//*[@data-testid="primaryButton"]')
        WebDriverWait(self.driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@data-testid="primaryButton"]'))).click()


    def _download_csv(self):
        time.sleep(5)
        # Acessa a lista e exportacão
        iframe =self.driver.find_element(By.TAG_NAME, "iframe")
        time.sleep(3)
        self.driver.switch_to.frame(iframe)
        time.sleep(3)
        try:
            self.driver.find_element(By.XPATH, '//*[@data-automationid="exportListCommand"]').click()
            time.sleep(5)
            export_csv_btn = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@data-automationid="exportListToCSVCommand"]')))
            export_csv_btn.click()        
            self.driver.switch_to.default_content()
            print("Exportação iniciada...")
        except Exception as e:
            print(f"Erro ao clicar no botão de exportação: {e}")
            self.driver.switch_to.default_content()  

    def _upload_item(self, item_data):
        """Registra um item no SharePoint"""
        # lógica de inserir item via Selenium
        pass

    def run_download(self, interval):
        self._init_driver()
        self._login()
        while True:
            try:
                df = self._download_csv()
                print("Dados enviados para processamento.")
            except Exception as e:
                print(f"Erro ao baixar CSV: {e}")
            time.sleep(interval)

    def run_upload(self, data_queue):
        self._init_driver()
        self._login()
        while True:
            df = data_queue.get()  # pega os dados mais recentes
            for _, row in df.iterrows():
                self._upload_item(row)
            print("Itens registrados no SharePoint.")
