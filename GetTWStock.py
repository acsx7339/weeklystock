import requests

class TWStock:

    def CODE(self):
        # Define the API URL
        url = 'https://openapi.twse.com.tw/v1/exchangeReport/STOCK_DAY_ALL'

        # Send a GET request
        res = requests.get(url)

        # Parse the JSON response
        data = res.json()

        # Extract "Code" from each stock entry
        codes = [entry['Code'] for entry in data]
        filtered_codes = [code for code in codes if len(code) <= 4 and not code.startswith('0')]
        # Print the list of codes
        return filtered_codes
