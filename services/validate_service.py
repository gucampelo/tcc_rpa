from services.api_service import APIService

class ValidateService:
    def __init__(self, api_service: APIService):
        self.__api_service = api_service

    def validate_operation(self, client, operation) -> dict:
        """
        Executa a validação completa:
        - chama a API externa
        - aplica regras adicionais (rating, status etc.)
        """

        resposta = self.__api_service.validar_cliente(cliente.cpf_cnpj)

        if "erro" in resposta:
            return resposta  # problema de comunicação

        if not resposta.get("valido", False):
            return {
                "valido": False,
                "motivo": "Cliente não encontrado ou inválido"
            }

        cliente = resposta.get("cliente")

        # ---- Regras de Negócio Internas ----
        if cliente["status"].lower() != "ativo":
            return {
                "valido": False,
                "motivo": "Cliente inativo"
            }

        if cliente["rating"] < 6 and operation.operation_type == "NOVO":
            return {
                "valido": False,
                "motivo": "Cliente inelegível devido ao rating e tipo de operação indequados"
            }
        
        if cliente.segment == "Especial" and operation.value > 1000000:
            return {
                "valido": False,
                "motivo": "Valor da operação incompatível com o segmento"
            }
        
        if cliente.segment == "E1" and operation.value > 10000000:
            return {
                "valido": False,
                "motivo": "Valor da operação incompatível com o segmento"
            }
        
        if cliente.segment == "E2" and operation.value > 30000000:
            return {
                "valido": False,
                "motivo": "Valor da operação incompatível com o segmento"
            }
        
        if cliente.segment == "E3" and operation.value > 50000000:
            return {
                "valido": False,
                "motivo": "Valor da operação incompatível com o segmento"
            }
        
        if operation.spread_requested < 0.02:
            return {
                "valido": False,
                "motivo": "Spread solicitado abaixo do mínimo aceitável (2%)"
            }
        
        if operation.guarantee_percentage < 0.2:
            return {
                "valido": False,
                "motivo": "Porcentagem de garantia abaixo do mínimo aceitável (20%)"
            }

        return {
            "valido": True,
            "motivo": "Cliente aprovado pelas regras de negócio."
        }
