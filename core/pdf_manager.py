from PyPDF2 import PdfReader, PdfWriter
import os


class PDFManager:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.pdf = PdfReader(pdf_path)

    def get_page_count(self):
        """คืนค่าจำนวนหน้าทั้งหมดในไฟล์ PDF"""
        return len(self.pdf.pages)

    def extract_page(self, page_index, output_path):
        """สกัดหน้าที่ระบุและบันทึกเป็นไฟล์ใหม่

        Args:
            page_index: ลำดับหน้าที่ต้องการ (เริ่มจาก 0)
            output_path: พาธของไฟล์ output
        """
        if page_index >= len(self.pdf.pages):
            raise IndexError(
                f"หน้า {page_index+1} ไม่มีอยู่ในไฟล์ PDF ที่มีทั้งหมด {len(self.pdf.pages)} หน้า"
            )

        writer = PdfWriter()
        writer.add_page(self.pdf.pages[page_index])

        # สร้างโฟลเดอร์ถ้ายังไม่มี
        os.makedirs(os.path.dirname(output_path), exist_ok=True)

        with open(output_path, "wb") as output_file:
            writer.write(output_file)
