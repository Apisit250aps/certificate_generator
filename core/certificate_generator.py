from PyQt6.QtCore import QThread, pyqtSignal
from core.pdf_manager import PDFManager
from core.excel_manager import ExcelManager


class CertificateGeneratorWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(
        self, pdf_path, excel_path, sheet_name, output_dir, name_format, column_values
    ):
        super().__init__()
        self.pdf_path = pdf_path
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.output_dir = output_dir
        self.name_format = name_format
        self.column_values = column_values

    def run(self):
        try:
            pdf_manager = PDFManager(self.pdf_path)
            excel_manager = ExcelManager(self.excel_path)
            data = excel_manager.get_data(self.sheet_name)

            total = len(data)
            for i, row in enumerate(data):
                # สร้างชื่อไฟล์จากค่าที่กำหนด
                filename_parts = []
                for value in self.column_values:
                    if value == -1:  # กรณีใช้ลำดับ
                        filename_parts.append(str(i + 1))
                    elif isinstance(value, int):  # กรณีใช้คอลัมน์จาก Excel
                        filename_parts.append(str(row[value]))
                    else:  # กรณีใช้ข้อความกำหนดเอง
                        filename_parts.append(value)

                filename = self.name_format.format(*filename_parts)
                # สร้างไฟล์ PDF
                output_path = f"{self.output_dir}/{filename}.pdf"
                pdf_manager.extract_page(i, output_path)

                # อัปเดตความคืบหน้า
                progress_val = int((i + 1) / total * 100)
                self.progress.emit(progress_val)

            self.finished.emit(f"เสร็จสิ้น: สร้างเกียรติบัตรทั้งหมด {total} ไฟล์")
        except Exception as e:
            self.error.emit(f"เกิดข้อผิดพลาด: {str(e)}")
