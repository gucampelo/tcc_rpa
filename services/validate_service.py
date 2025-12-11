from services.api_service import APIService

class ValidateService:
    def __init__(self):
        self.__api_service = APIService()

    # -------------------------------------------------------------
    # NOVO MÉTODO → responsável apenas por acessar a API
    # -------------------------------------------------------------
    def get_client_api(self, cpf_cnpj: str) -> dict:
        """
        Encapsula a chamada à API externa.
        Retorna exatamente o dicionário fornecido pela API.
        """
        return self.__api_service.get_client({"cpf_cnpj": cpf_cnpj})

    # -------------------------------------------------------------
    # Método principal de validação
    # -------------------------------------------------------------
    def validate_operation(self, client, operation) -> dict:
        """
        Executa a validação completa:
        1. Consulta a API externa para obter dados do cliente
        2. Aplica regras internas de negócio
        """

        # 1) Chama a API para buscar informações do cliente
        resposta = self.get_client_api(client.cpf_cnpj)

        if "erro" in resposta:
            return {
                "valido": False,
                "motivo": "Erro ao acessar a API de clientes"
            }

        if not resposta.get("valido", False):
            return {
                "valido": False,
                "motivo": "Cliente não encontrado."
            }

        # DADOS DO CLIENTE VINDOS DA API
        cliente_api = resposta.get("cliente")

        # -------- Regras de Negócio Internas -------- #

        # Regra 1: cliente ativo
        if cliente_api.get("status", "").lower() != "ativo":
            return {
                "valido": False,
                "motivo": "Cliente inativo."
            }

        # Regra 2: rating mínimo para operação NOVA
        if cliente_api.get("rating", 0) < 6 and operation.operation_type == "NOVA":
            return {
                "valido": False,
                "motivo": "Cliente inelegível devido ao rating insuficiente."
            }

        # Regra 3: limites por segmento
        segmento = cliente_api.get("segmento")

        limites = {
            "Especial": 1_000_000,
            "E1": 10_000_000,
            "E2": 30_000_000,
            "E3": 50_000_000,
            "Agro": 100_000_000
        }

        # Segmento deve existir
        if segmento not in limites:
            return {
                "valido": False,
                "motivo": f"Segmento '{segmento}' inválido ou não informado."
            }

        # Converte valor da operação para float, robusto
        try:
            value = float(str(operation.value).replace(".", "").replace(",", ""))
        except (ValueError, TypeError):
            return {
                "valido": False,
                "motivo": f"Valor da operação inválido: {operation.value}"
            }

        # Compara com o limite do segmento
        if value > limites[segmento]:
            return {
                "valido": False,
                "motivo": "Valor da operação incompatível com o segmento do cliente."
            }

        # Regra 4: spread mínimo
        if operation.spread_requested < 0.02:
            return {
                "valido": False,
                "motivo": "Spread solicitado abaixo do mínimo aceitável (2%)."
            }

        # Regra 5: garantia mínima
        if operation.guarantee_percentage < 0.20:
            return {
                "valido": False,
                "motivo": "Garantia abaixo do mínimo exigido (20%)."
            }

        return {
            "valido": True,
            "motivo": "Cliente aprovado pelas regras de negócio."
        }
