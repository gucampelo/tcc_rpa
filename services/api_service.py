import requests

class APIService:
    def __init__(self, base_url="http://127.0.0.1:5000"):
        self.base_url = base_url

    def validate(self, record):
        print(f"[API] Enviando record {record.operation_id} para validação...")
        try:
            payload = {
                "cliente_id": record.client_id,
                "operation_id": record.operation_id,
                "valor": record.value
            }
            resp = requests.post(f"{self.base_url}/validar", json=payload, timeout=5)
            resp.raise_for_status()
            resultado = resp.json()
            print(f"[API] Resposta: {resultado}")
            return resultado.get("valido", False), resultado.get("motivo", "Sem motivo")
        except Exception as e:
            print(f"[API] Erro ao validar: {e}")
            return False, "Erro de comunicação com API"
