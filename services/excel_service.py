import uno # type: ignore
import subprocess
import time
import socket
import json


class ExcelService:
    
    def __init__(self, filepath: str, host: str = "localhost", port: int = 2002):
        self.__filepath = filepath
        self.__host = host
        self.__port = port
        self.__desktop = None
        self.__doc = None

    # -----------------------------
    # Métodos internos
    # -----------------------------
    def __is_port_open(self) -> bool:
        """Verifica se a porta do LibreOffice está aberta"""
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(1)
            return sock.connect_ex((self.__host, self.__port)) == 0

    def __start_libreoffice(self):
        """Inicia o LibreOffice em modo headless"""
        print("[EXCEL] Iniciando LibreOffice em modo headless...")
        subprocess.Popen([
            "libreoffice",
            f'--accept=socket,host={self.__host},port={self.__port};urp;',
            "--norestore", "--nofirststartwizard", "--nologo"
        ])
        time.sleep(3)

    def __executar_macro_interna(self, macro_name: str):
        """Executa uma macro dentro do arquivo"""
        script_provider = self.__doc.getScriptProvider()
        script = script_provider.getScript(
            f"vnd.sun.star.script:Standard.Module1.{macro_name}?language=Basic&location=application"
        )
        script.invoke((), (), ())

    # -----------------------------
    # Interface pública
    # -----------------------------
    def connect(self):
        """Conecta ao LibreOffice e carrega a planilha"""
        if not self.__is_port_open():
            self.__start_libreoffice()

        try:
            local_ctx = uno.getComponentContext()
            resolver = local_ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", local_ctx
            )
            ctx = resolver.resolve(
                f"uno:socket,host={self.__host},port={self.__port};urp;StarOffice.ComponentContext"
            )
            smgr = ctx.ServiceManager
            self.__desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
            url = uno.systemPathToFileUrl(self.__filepath)
            self.__doc = self.__desktop.loadComponentFromURL(url, "_blank", 0, ())
            print("[EXCEL] Conectado e planilha carregada com sucesso.")
        except Exception as e:
            print(f"[EXCEL] Erro ao conectar ao LibreOffice: {e}")

    def preencher_dados(self, valor, tipo_taxa, parcelas_json: str):
        """Preenche os campos de entrada no Excel"""
        plan = self.__doc.Sheets.getByIndex(0)
        plan.getCellRangeByName("C4").Value = valor
        plan.getCellRangeByName("F3").String = tipo_taxa

        # Limpa linhas anteriores
        for i in range(12, 50):
            plan.getCellByPosition(1, i).String = ""
            plan.getCellByPosition(2, i).Value = ""

        parcelas = json.loads(parcelas_json.replace("\n", "").strip())
        row = 12
        for parcela in parcelas:
            plan.getCellByPosition(1, row).String = parcela["data"]
            plan.getCellByPosition(2, row).Value = parcela["valor"]
            row += 1

    def rodar_macro(self) -> float:
        """Executa macro e retorna taxa gerada"""
        self.__executar_macro_interna("GerarDecimal")
        plan = self.__doc.Sheets.getByIndex(0)
        taxa = plan.getCellByPosition(5, 3).Value
        print(f"[EXCEL] Taxa gerada: {taxa}%")
        return taxa

    def salvar(self):
        """Salva alterações"""
        self.__doc.store()

    def fechar(self):
        """Fecha o arquivo"""
        self.__doc.close(True)
