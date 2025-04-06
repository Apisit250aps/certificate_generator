import pandas as pd

class ExcelManager:
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.excel_file = pd.ExcelFile(excel_path)
    
    def get_sheet_names(self):
        """คืนค่ารายชื่อชีททั้งหมดในไฟล์ Excel"""
        return self.excel_file.sheet_names
    
    def get_headers(self, sheet_name):
        """คืนค่าหัวคอลัมน์ของชีทที่ระบุ"""
        df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
        return list(df.columns)
    
    def get_data(self, sheet_name):
        """คืนค่าข้อมูลทั้งหมดในชีทที่ระบุ เป็นรูปแบบลิสต์ของลิสต์"""
        df = pd.read_excel(self.excel_file, sheet_name=sheet_name)
        return df.values.tolist()