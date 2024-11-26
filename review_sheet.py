import openpyxl
from datetime import datetime, timedelta
import twstock

class ReviewExcel():

    def __init__(self):
        self.file_name = "stock_review.xlsx"
        self.wb = openpyxl.load_workbook(self.file_name)
        seven_days_ago = datetime.now() - timedelta(days=7)
        self.past_date = seven_days_ago.strftime("%Y-%m-%d")
        print(self.past_date)
        self.sheet_name = self.past_date
        if self.sheet_name not in self.wb.sheetnames:
            raise ValueError(f"Sheet '{self.sheet_name}' does not exist in the workbook!")
        
    def get_price(self):
        sheet = self.wb[self.sheet_name]
        # Find the column index for "Name"
        header_row = 1  # Assuming headers are in the first row
        name_column_index = None
        for col_index, cell in enumerate(sheet[header_row], start=1):
            if cell.value == "stock":
                name_column_index = col_index
                break
        if name_column_index is None:
            raise ValueError("Column with header 'Name' not found!")
        # Get values from the "stock" column (excluding the header)
        stock_item = [sheet.cell(row=row, column=name_column_index).value for row in range(2, sheet.max_row + 1)]  # Start from row 2
        stocklist_price = []
        for item in stock_item:
            stock = twstock.Stock(item)
            past_price = stock.price[-10]
            stocklist_price.append(past_price)
        return stocklist_price

    def update_row(self, price):
        sheet = self.wb[self.sheet_name]
        # Find the next empty row (the first empty row in column 1)
        next_row = sheet.max_row + 1

        # Find the first row with data (assuming header is in row 1)
        header = "Close"
        header_column = sheet.max_column + 1  # Find the first empty column after existing data

        # Add the "Close" header in the first row
        sheet.cell(row=1, column=header_column, value=header)

        # Add the result_list values starting in row 2
        for i, value in enumerate(price, start=2):  # Start from row 2, as row 1 is for the header
            sheet.cell(row=i, column=header_column, value=value)

        # Save the workbook with the new data
        self.wb.save(self.file_name)

        # Print confirmation
        print(f"Added new header '{header}' and data {price} starting from column {header_column}.")


        

if __name__ == "__main__":
    rw = ReviewExcel()
    price = rw.get_price()
    rw.update_row(price)

