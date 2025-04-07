from PyQt6.QtWidgets import (QWidget, QGridLayout, QLabel, QLineEdit, 
                            QPushButton, QFileDialog, QGroupBox, QHBoxLayout,
                            QFrame, QSizePolicy)
from PyQt6.QtCore import pyqtSignal, Qt
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import QStandardPaths
class FileSelector(QWidget):
    excel_file_selected = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        
        # สร้าง Layout
        layout = QGridLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # สร้าง GroupBox
        self.group_box = QGroupBox("เลือกไฟล์")
        self.group_box.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)
        group_layout = QGridLayout(self.group_box)
        group_layout.setVerticalSpacing(15)
        
        # สร้างวิดเจ็ตสำหรับเลือกไฟล์ PDF
        self.pdf_label = QLabel("ไฟล์เกียรติบัตร (PDF):")
        self.pdf_path = QLineEdit()
        self.pdf_path.setReadOnly(True)
        self.pdf_path.setPlaceholderText("เลือกไฟล์ PDF...")
        
        self.pdf_browse_btn = QPushButton("เลือกไฟล์...")
        self.pdf_browse_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.pdf_browse_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                padding: 5px 10px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #1c6ea4;
            }
        """)
        self.pdf_browse_btn.clicked.connect(self.browse_pdf)
        
        # สร้างวิดเจ็ตสำหรับเลือกไฟล์ Excel
        self.excel_label = QLabel("ไฟล์รายชื่อ (Excel):")
        self.excel_path = QLineEdit()
        self.excel_path.setReadOnly(True)
        self.excel_path.setPlaceholderText("เลือกไฟล์ Excel...")
        
        self.excel_browse_btn = QPushButton("เลือกไฟล์...")
        self.excel_browse_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.excel_browse_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                padding: 5px 10px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #219955;
            }
            QPushButton:pressed {
                background-color: #1a7943;
            }
        """)
        self.excel_browse_btn.clicked.connect(self.browse_excel)
        
        # สร้าง Container สำหรับแต่ละแถว
        pdf_container = QFrame()
        pdf_container.setFrameShape(QFrame.Shape.NoFrame)
        pdf_layout = QHBoxLayout(pdf_container)
        pdf_layout.setContentsMargins(0, 0, 0, 0)
        pdf_layout.addWidget(self.pdf_path, 5)
        pdf_layout.addWidget(self.pdf_browse_btn, 1)
        
        excel_container = QFrame()
        excel_container.setFrameShape(QFrame.Shape.NoFrame)
        excel_layout = QHBoxLayout(excel_container)
        excel_layout.setContentsMargins(0, 0, 0, 0)
        excel_layout.addWidget(self.excel_path, 5)
        excel_layout.addWidget(self.excel_browse_btn, 1)
        
        # เพิ่มวิดเจ็ตลงใน GroupBox
        group_layout.addWidget(self.pdf_label, 0, 0)
        group_layout.addWidget(pdf_container, 0, 1)
        
        group_layout.addWidget(self.excel_label, 1, 0)
        group_layout.addWidget(excel_container, 1, 1)
        
        # เพิ่ม GroupBox ลงใน Layout หลัก
        layout.addWidget(self.group_box)
    
    def browse_pdf(self):
    # รับพาธของโฟลเดอร์ Documents
        documents_path = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
        
        file_path, _ = QFileDialog.getOpenFileName(
            self, "เลือกไฟล์เกียรติบัตร", documents_path, "PDF Files (*.pdf)"
        )
        if file_path:
            self.pdf_path.setText(file_path)
            # อัปเดตสถานะว่าได้เลือกไฟล์แล้ว
            self.pdf_path.setStyleSheet("background-color: #e8f5e9; color: #1b5e20;")

    def browse_excel(self):
        # รับพาธของโฟลเดอร์ Documents
        documents_path = QStandardPaths.writableLocation(QStandardPaths.StandardLocation.DocumentsLocation)
        
        file_path, _ = QFileDialog.getOpenFileName(
            self, "เลือกไฟล์รายชื่อ", documents_path, "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.excel_path.setText(file_path)
            # อัปเดตสถานะว่าได้เลือกไฟล์แล้ว
            self.excel_path.setStyleSheet("background-color: #e8f5e9; color: #1b5e20;")
            self.excel_file_selected.emit(file_path)
        
    def get_pdf_path(self):
        return self.pdf_path.text()
    
    def get_excel_path(self):
        return self.excel_path.text()