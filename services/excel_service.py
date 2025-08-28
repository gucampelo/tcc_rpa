class ExcelService:
    def calculate(self, record):
        print(f"Calculating rate for record {record['ID']}")
        # Example: spread + cost
        return record.get("SPREAD_SOLC", 0) + record.get("CUSTO_SOLC", 0)
