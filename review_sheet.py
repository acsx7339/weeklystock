import openpyxl
from datetime import datetime
from fugle_marketdata import RestClient

class ReviewExcel():

    def __init__(self):
        self.file_name = "stock_review.xlsx"
        self.wb = openpyxl.load_workbook(self.file_name)
        self.past_date = datetime.now(-7).strftime("%Y-%m-%d")
        self.sheet = self.wb.create_sheet(title=self.current_date)


    def update_cell(self, col_data, col_index=None):
        sheet_name = self.past_date
        if sheet_name not in self.wb.sheetnames:
            raise ValueError(f"工作表 {sheet_name} 不存在！")
        sheet = self.wb[sheet_name]
        self.sheet.append(["EndPrice"])
        if col_index is None:
            col_index = sheet.max_column + 1
        for i, value in enumerate(col_data, start=1):
            sheet.cell(row=i, column=col_index, value=value)

        # 保存工作簿
        self.wb.save(self.file_name)
        print(f"已向工作表 {sheet_name} 添加新列: {col_data}，位置: 第 {col_index} 列")


if __name__ == "__main__":
    pass
