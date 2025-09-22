class RPAHandler:
    def __init__(self, sharepoint_service, api_service, excel_service):
        self.sharepoint_service = sharepoint_service
        self.api_service = api_service
        self.excel_service = excel_service

    def run(self):
        print("Starting RPA process...")
        # Step 1: extract data
        data = self.sharepoint_service.extract_data()

        # Step 2: validate data via API
        for record in data:
            if self.api_service.validate(record):
                # Step 3: calculate rate in Excel
                rate = self.excel_service.calculate(record)
                record["rate"] = rate

                # Step 4: insert result back to SharePoint
                self.sharepoint_service.insert_result(record)
