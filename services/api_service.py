import requests

class APIService:
    def __init__(self, base_url="http://127.0.0.1:5000"):
        self.__base_url = base_url

    def get_client(self, data: dict) -> dict:
        """
        Envia o CPF à API de validação e retorna o json contendo
        informações do cliente ou erro.
        """
        try:
            response = requests.post(
                f"{self.__base_url}/get_client",
                json=data,
                timeout=3
            )
            return response.json()

        except Exception as e:
            return {"erro": f"Falha na comunicação com a API: {str(e)}"}
