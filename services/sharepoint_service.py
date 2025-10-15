# services/sharepoint_service.py
import time, threading
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class SharePointService:
    """
    Serviço responsável por interagir com o SharePoint.
    Responsabilidades:
      - Conectar e autenticar no portal.
      - Fazer download de arquivos CSV da lista.
      - Realizar upload (inserção) de resultados.
    """

    def __init__(self, download_path: str, url_sharepoint: str, user: str, password: str):
        self.download_path = download_path
        self.url_sharepoint = url_sharepoint
        self.__user = user
        self.__password = password
        self.__driver = None
        self.__lock = threading.Lock()  # cada instância tem seu lock interno

    # -----------------------------
    # Inicialização e login
    # -----------------------------
    def __init_driver(self):
        options = webdriver.ChromeOptions()
        prefs = {"download.default_directory": self.download_path}
        options.add_experimental_option("prefs", prefs)
        options.add_argument("--start-maximized")
        self.__driver = webdriver.Chrome(options=options)

    def __login(self):
        """Executa login automatizado no SharePoint."""
        self.__driver.get(self.url_sharepoint)
        WebDriverWait(self.__driver, 20).until(
            EC.presence_of_element_located((By.ID, "usernameEntry"))
        ).send_keys(self.__user + Keys.ENTER)

        WebDriverWait(self.__driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Outras maneiras de entrar')]"))
        ).click()
        WebDriverWait(self.__driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'Use sua password')]"))
        ).click()

        WebDriverWait(self.__driver, 20).until(
            EC.presence_of_element_located((By.ID, "passwordEntry"))
        ).send_keys(self.__password + Keys.ENTER)

        WebDriverWait(self.__driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@data-testid="primaryButton"]'))
        ).click()

    def connect(self):
        """Inicializa o driver e faz login. Chamado pela RPA."""
        with self.__lock:
            self.__init_driver()
            self.__login()

    # -----------------------------
    # Download de CSV
    # -----------------------------
    def __download_csv(self):
        """Método interno que clica nos botões de exportação do SharePoint."""
        time.sleep(5)
        iframe = self.__driver.find_element(By.TAG_NAME, "iframe")
        self.__driver.switch_to.frame(iframe)

        try:
            export_btn = WebDriverWait(self.__driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@data-automationid="exportListCommand"]'))
            )
            export_btn.click()

            export_csv_btn = WebDriverWait(self.__driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@data-automationid="exportListToCSVCommand"]'))
            )
            export_csv_btn.click()

            print("[SP] Exportação iniciada...")
        except Exception as e:
            print(f"[SP] Erro ao clicar no botão de exportação: {e}")
        finally:
            self.__driver.switch_to.default_content()

    def download_csv(self):
        """Método público usado pela thread principal (RPA)."""
        with self.__lock:
            self.__download_csv()

    # -----------------------------
    # Upload de item (registro de resultado)
    # -----------------------------
    def __upload_item(self, client, operation, record):
        """
        Registra um item na lista SharePoint com base nos objetos
        client, operation e record.
        """
        iframe = self.__driver.find_element(By.TAG_NAME, "iframe")
        self.__driver.switch_to.frame(iframe)
        try:
            # Botão Novo
            WebDriverWait(self.__driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[1]/div[1]/div/div/span[1]/button[1]'))
            ).click()

            # Campos
            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="NRM_PO, vazio, editor de campo. "]'))
            ).send_keys(operation.nrm_po)
            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="CPF_CNPJ, vazio, editor de campo. "]'))
            ).send_keys(client.cpf_cnpj)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="CLIENTE, vazio, editor de campo. "]'))
            ).send_keys(client.name)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="segment, vazio, editor de campo. "]'))
            ).send_keys(client.segment)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="RATING, vazio, editor de campo. "]'))
            ).send_keys(client.rating)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="requester, vazio, editor de campo. "]'))
            ).send_keys(record.requester)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="EMAIL_SOLC, vazio, editor de campo. "]'))
            ).send_keys(record.email_solc)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="VALOR, vazio, editor de campo. "]'))
            ).send_keys(operation.value)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="PRAZO_DIAS, vazio, editor de campo. "]'))
            ).send_keys(operation.term_days)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="TIPO_TAXA, vazio, editor de campo. "]'))
            ).send_keys(operation.rate_type)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="TAXA_SOLC, vazio, editor de campo. "]'))
            ).send_keys(operation.rate_requested)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="CUSTO_SOLC, vazio, editor de campo. "]'))
            ).send_keys(operation.cost_requested)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="SPREAD_SOLC, vazio, editor de campo. "]'))
            ).send_keys(operation.spread_requested)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="TAXA_TRV, vazio, editor de campo. "]'))
            ).send_keys(operation.rate_approved)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="CUSTO_TRV, vazio, editor de campo. "]'))
            ).send_keys(operation.cost_approved)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="SPREAD_TRV, vazio, editor de campo. "]'))
            ).send_keys(operation.spread_approved)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="FLUXO_PARCELAS, vazio, editor de campo. "]'))
            ).send_keys(operation.parcel_flow)

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@aria-label="STATUS, vazio, editor de campo. "]'))
            ).send_keys("Aprovado")

            WebDriverWait(self.__driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@data-automationid="ReactClientFormSaveButton"]'))
            ).click()

            print(f"[SP] Item enviado: {client.name} / {operation.value}")
        except Exception as e:
            print(f"[SP] Erro ao registrar item: {e}")
        finally:
            self.__driver.switch_to.default_content()

    def upload_item(self, client, operation, record):
        """Método público chamado pelo ProcessingService."""
        with self.__lock:
            self.__upload_item(client, operation, record)
