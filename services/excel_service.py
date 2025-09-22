import uno  # type: ignore
import subprocess
import time
import socket
import json


class ExcelService:
    def __init__(self, filepath, host="localhost", port=2002):
        self.filepath = filepath
        self.desktop = None
        self.doc = None
        self.host = host
        self.port = port

    def _is_port_open(self):
        """Verifica se a porta do LibreOffice está aberta"""
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(1)
            return sock.connect_ex((self.host, self.port)) == 0

    def _start_libreoffice(self):
        """Inicia o LibreOffice headless se não estiver rodando"""
        subprocess.Popen([
            "libreoffice",
            f'--accept=socket,host={self.host},port={self.port};urp;',
            "--norestore", "--nofirststartwizard", "--nologo"
        ])
        # Dá um tempo para iniciar
        time.sleep(3)

    def connect(self):
        """Conecta ao LibreOffice, iniciando se necessário"""
        if not self._is_port_open():
            print("🔄 LibreOffice não está rodando. Iniciando em modo headless...")
            self._start_libreoffice()

        local_ctx = uno.getComponentContext()
        resolver = local_ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_ctx
        )
        ctx = resolver.resolve(
            f"uno:socket,host={self.host},port={self.port};urp;StarOffice.ComponentContext"
        )
        smgr = ctx.ServiceManager
        self.desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        url = uno.systemPathToFileUrl(self.filepath)
        self.doc = self.desktop.loadComponentFromURL(url, "_blank", 0, ())
        print("✅ Conectado ao LibreOffice e planilha carregada.")

    def preencher_dados(self, valor, tipo_taxa, parcelas):
        """
        Preenche os dados básicos na planilha:
        valor: número
        tipo_taxa: string ("Pré-Fixada" ou "Pós-Fixada")
        parcelas: lista de dicts [{data: "2026-10-10", principal: 250000}, ...]
        """
        plan = self.doc.Sheets.getByIndex(0)

        # Preenche valor (C4) e tipo de taxa (F3)
        plan.getCellRangeByName("C4").Value = valor
        plan.getCellRangeByName("F3").String = tipo_taxa

        # Limpa parcelas antigas
        for i in range(12, 50):
            plan.getCellByPosition(1, i).String = ""  # coluna C (datas)
            plan.getCellByPosition(2, i).Value = ""   # coluna D (principal)

        # Preenche parcelas a partir da linha 12
        row = 12  # índice zero-bparcelaased (linha 12 = 11)
        parcelas = json.loads(parcelas.replace("\n", "").strip())
        
        for parcela in parcelas:
            plan.getCellByPosition(1, row).String = parcela["data"]
            plan.getCellByPosition(2, row).Value = parcela["valor"]
            row += 1

    def rodar_macro(self, macro_name):
        plan = self.doc.Sheets.getByIndex(0)

        """Executa uma macro existente dentro do arquivo"""
        script_provider = self.doc.getScriptProvider()
        script = script_provider.getScript(
            f"vnd.sun.star.script:Standard.Module1.{macro_name}?language=Basic&location=application"
        )
        script.invoke((), (), ())
        return plan.getCellByPosition(5, 3).Value # rate

    def salvar(self):
        self.doc.store()

    def fechar(self):
        self.doc.close(True)
