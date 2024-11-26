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
        
    def update_cell(self):
        sheet = self.wb[self.sheet_name]
        if self.sheet_name not in self.wb.sheetnames:
            raise ValueError(f"Sheet '{self.sheet_name}' does not exist in the workbook!")
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
        print(stocklist_price)

if __name__ == "__main__":
    rw = ReviewExcel()
    rw.update_cell()

