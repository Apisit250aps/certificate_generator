from PyQt6.QtWidgets import (QWidget, QGridLayout, QLabel, QLineEdit, 
                            QPushButton, QFileDialog, QGroupBox, QHBoxLayout)
from PyQt6.QtCore import pyqtSignal

class FileSelector(QWidget):
    excel_file_selected = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        
        # สร้าง Layout
        layout = QGridLayout(self)
        
        # สร้าง GroupBox
        self.group_box = QGroupBox("เลือกไฟล์")
        group_layout = QGridLayout(self.group_box)
        
        # สร้างวิดเจ็ตสำหรับเลือกไฟล์ PDF
        self.pdf_label = QLabel("ไฟล์เกียรติบัตร (PDF):")
        self.pdf_path = QLineEdit()
        self.pdf_path.setReadOnly(True)
        self.pdf_browse_btn = QPushButton("เลือกไฟล์...")
        self.pdf_browse_btn.clicked.connect(self.browse_pdf)
        
        # สร้างวิดเจ็ตสำหรับเลือกไฟล์ Excel
        self.excel_label = QLabel("ไฟล์รายชื่อ (Excel):")
        self.excel_path = QLineEdit()
        self.excel_path.setReadOnly(True)
        self.excel_browse_btn = QPushButton("เลือกไฟล์...")
        self.excel_browse_btn.clicked.connect(self.browse_excel)
        
        # เพิ่มวิดเจ็ตลงใน GroupBox
        group_layout.addWidget(self.pdf_label, 0, 0)
        group_layout.addWidget(self.pdf_path, 0, 1)
        group_layout.addWidget(self.pdf_browse_btn, 0, 2)
        
        group_layout.addWidget(self.excel_label, 1, 0)
        group_layout.addWidget(self.excel_path, 1, 1)
        group_layout.addWidget(self.excel_browse_btn, 1, 2)
        
        # เพิ่ม GroupBox ลงใน Layout หลัก
        layout.addWidget(self.group_box)
    
    def browse_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "เลือกไฟล์เกียรติบัตร", "", "PDF Files (*.pdf)"
        )
        if file_path:
            self.pdf_path.setText(file_path)
    
    def browse_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "เลือกไฟล์รายชื่อ", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.excel_path.setText(file_path)
            self.excel_file_selected.emit(file_path)
    
    def get_pdf_path(self):
        return self.pdf_path.text()
    
    def get_excel_path(self):
        return self.excel_path.text()