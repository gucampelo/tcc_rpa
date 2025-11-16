import requests

class APIService:
    def __init__(self, base_url="http://127.0.0.1:5000"):
        self.base_url = base_url

    def get_client(self, data):
        response = requests.post(f"{self.base_url}/get_client", json=data)
        return response.json()