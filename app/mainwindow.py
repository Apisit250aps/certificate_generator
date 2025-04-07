from PyQt6.QtWidgets import (QMainWindow, QWidget, QGridLayout, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QMessageBox,
                            QStatusBar, QProgressBar, QSplitter, QFrame)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt6.QtGui import QIcon, QFont, QColor, QPalette

from app.widgets.file_selector import FileSelector
from app.widgets.excel_viewer import ExcelViewer
from app.widgets.output_config import OutputConfig
from core.pdf_manager import PDFManager
from core.excel_manager import ExcelManager
from core.certificate_generator import CertificateGeneratorWorker

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Certificate Generator")
        self.setMinimumSize(1000, 700)
        
        # ตั้งค่า Style ทั่วไป
        self.setup_styles()
        
        # สร้าง Central Widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # สร้าง Layout หลัก
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)
        
        # สร้าง QSplitter เพื่อให้ปรับขนาดได้
        splitter = QSplitter(Qt.Orientation.Vertical)
        
        # สร้างวิดเจ็ตต่างๆ
        self.file_selector = FileSelector()
        
        # สร้าง Container สำหรับ Excel Viewer และ Output Config
        bottom_container = QWidget()
        bottom_layout = QHBoxLayout(bottom_container)
        bottom_layout.setContentsMargins(0, 0, 0, 0)
        
        self.excel_viewer = ExcelViewer()
        self.output_config = OutputConfig()
        
        # ใช้ QSplitter แนวนอนสำหรับ excel_viewer และ output_config
        horizontal_splitter = QSplitter(Qt.Orientation.Horizontal)
        horizontal_splitter.addWidget(self.excel_viewer)
        horizontal_splitter.addWidget(self.output_config)
        horizontal_splitter.setSizes([500, 500])  # ตั้งค่าขนาดเริ่มต้น
        
        bottom_layout.addWidget(horizontal_splitter)
        
        # เพิ่มวิดเจ็ตลงใน QSplitter หลัก
        splitter.addWidget(self.file_selector)
        splitter.addWidget(bottom_container)
        splitter.setSizes([200, 500])  # ตั้งค่าขนาดเริ่มต้น
        
        # เพิ่ม QSplitter ลงใน Layout หลัก
        main_layout.addWidget(splitter, 1)
        
        # สร้าง Control Panel พร้อมสไตล์
        control_panel = QFrame()
        control_panel.setFrameShape(QFrame.Shape.StyledPanel)
        control_panel.setStyleSheet("background-color: #f5f5f5; border-radius: 8px;")
        control_layout = QHBoxLayout(control_panel)
        
        # สร้างปุ่มสร้างเกียรติบัตร
        self.generate_btn = QPushButton("สร้างเกียรติบัตร")
        self.generate_btn.setMinimumHeight(45)
        self.generate_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                font-size: 14px;
                border-radius: 5px;
                padding: 8px 16px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        self.generate_btn.clicked.connect(self.generate_certificates)
        
        # เพิ่ม Progress Bar พร้อมสไตล์
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #bdbdbd;
                border-radius: 4px;
                text-align: center;
                height: 25px;
                font-weight: bold;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 3px;
            }
        """)
        
        # เพิ่มองค์ประกอบลงใน Control Panel
        control_layout.addWidget(self.progress_bar, 3)
        control_layout.addWidget(self.generate_btn, 1)
        
        # เพิ่ม Control Panel ลงใน Layout หลัก
        main_layout.addWidget(control_panel)
        
        # สร้าง Status Bar พร้อมสไตล์
        self.status_bar = QStatusBar()
        self.status_bar.setStyleSheet("""
            QStatusBar {
                background-color: #f5f5f5;
                color: #333333;
                min-height: 25px;
                font-size: 13px;
                padding: 2px 5px;
            }
        """)
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("พร้อมใช้งาน")
        
        # เชื่อมต่อสัญญาณระหว่างวิดเจ็ต
        self.connect_signals()
        
        # แสดงหน้าต่างเต็มจอ
        self.showMaximized()
    
    def setup_styles(self):
        # ตั้งค่าสไตล์ทั่วไปของแอปพลิเคชัน
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 13px;
                border: 1px solid #cccccc;
                border-radius: 5px;
                margin-top: 15px;
                padding-top: 15px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                background-color: white;
            }
            QLabel {
                font-size: 13px;
                }
            QComboBox, QLineEdit {
                padding: 5px;
                border: 1px solid #bdbdbd;
                border-radius: 3px;
                min-height: 25px;
            }
            QPushButton {
                padding: 5px 10px;
                border-radius: 3px;
            }
            QTableView {
                gridline-color: #d0d0d0;
                selection-background-color: #e0f0ff;
                selection-color: black;
                border: 1px solid #bdbdbd;
                border-radius: 3px;
            }
            QTableView::item {
                padding: 5px;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 5px;
                border: 1px solid #d0d0d0;
                font-weight: bold;
            }
        """)
        
        # ตั้งค่าฟอนต์ทั่วไป
        font = QFont("Segoe UI", 10)
        self.setFont(font)
    
    def connect_signals(self):
        # เชื่อมต่อสัญญาณจาก FileSelector
        self.file_selector.excel_file_selected.connect(self.on_excel_selected)
        
        # เชื่อมต่อสัญญาณจาก ExcelViewer
        self.excel_viewer.sheet_selected.connect(self.on_sheet_selected)
    
    def on_excel_selected(self, file_path):
        try:
            # โหลดข้อมูลชีทจากไฟล์ Excel
            excel_manager = ExcelManager(file_path)
            sheet_names = excel_manager.get_sheet_names()
            
            # อัปเดต ExcelViewer
            self.excel_viewer.set_sheet_names(sheet_names)
            self.status_bar.showMessage(f"โหลดไฟล์ Excel สำเร็จ: {file_path}")
        except Exception as e:
            QMessageBox.critical(self, "ข้อผิดพลาด", f"ไม่สามารถโหลดไฟล์ Excel: {str(e)}")
    
    def on_sheet_selected(self, sheet_name):
        try:
            # โหลดข้อมูลจากชีทที่เลือก
            excel_manager = ExcelManager(self.file_selector.get_excel_path())
            data = excel_manager.get_data(sheet_name)
            headers = excel_manager.get_headers(sheet_name)
            
            # อัปเดต ExcelViewer และ OutputConfig
            self.excel_viewer.set_data(data, headers)
            self.output_config.set_columns(headers)
            self.status_bar.showMessage(f"โหลดชีท '{sheet_name}' สำเร็จ")
        except Exception as e:
            QMessageBox.critical(self, "ข้อผิดพลาด", f"ไม่สามารถโหลดข้อมูลจากชีท: {str(e)}")
    
    def generate_certificates(self):
        # ตรวจสอบข้อมูลที่จำเป็น
        pdf_path = self.file_selector.get_pdf_path()
        excel_path = self.file_selector.get_excel_path()
        sheet_name = self.excel_viewer.get_selected_sheet()
        output_dir = self.output_config.get_output_path()
        name_format = self.output_config.get_name_format()
        column_values = self.output_config.get_selected_column_indices()
        
        if not all([pdf_path, excel_path, sheet_name, output_dir]):
            QMessageBox.warning(self, "ข้อมูลไม่ครบถ้วน", "กรุณาเลือกไฟล์ PDF, Excel, ชีท และโฟลเดอร์สำหรับไฟล์ output")
            return
        
        if not column_values:
            QMessageBox.warning(self, "ข้อมูลไม่ครบถ้วน", "กรุณากำหนดรูปแบบการตั้งชื่อไฟล์อย่างน้อย 1 รายการ")
            return
        
        # ตรวจสอบจำนวนหน้า PDF และจำนวนแถวข้อมูล
        try:
            pdf_manager = PDFManager(pdf_path)
            excel_manager = ExcelManager(excel_path)
            pdf_pages = pdf_manager.get_page_count()
            data_rows = len(excel_manager.get_data(sheet_name))
            
            if pdf_pages != data_rows:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Icon.Warning)
                msg.setText(f"จำนวนหน้า PDF ({pdf_pages}) ไม่ตรงกับจำนวนรายชื่อ ({data_rows})")
                msg.setInformativeText("คุณต้องการดำเนินการต่อหรือไม่?")
                msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if msg.exec() == QMessageBox.StandardButton.No:
                    return
        except Exception as e:
            QMessageBox.critical(self, "ข้อผิดพลาด", f"ไม่สามารถตรวจสอบไฟล์: {str(e)}")
            return
        
        # ปิดใช้งานปุ่มระหว่างดำเนินการ
        self.generate_btn.setEnabled(False)
        
        # สร้าง Worker Thread
        self.worker = CertificateGeneratorWorker(
            pdf_path, excel_path, sheet_name, output_dir, name_format, column_values
        )
        
        # เชื่อมต่อสัญญาณ
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.on_generation_finished)
        self.worker.error.connect(self.on_generation_error)
        
        # เริ่มการทำงาน
        self.worker.start()
        self.status_bar.showMessage("กำลังสร้างเกียรติบัตร...")
    
    def update_progress(self, value):
        self.progress_bar.setValue(value)
    
    def on_generation_finished(self, message):
        self.progress_bar.setValue(100)
        self.status_bar.showMessage(message)
        self.generate_btn.setEnabled(True)
        QMessageBox.information(self, "เสร็จสิ้น", message)
    
    def on_generation_error(self, error_message):
        self.progress_bar.setValue(0)
        self.status_bar.showMessage(f"เกิดข้อผิดพลาด: {error_message}")
        self.generate_btn.setEnabled(True)
        QMessageBox.critical(self, "ข้อผิดพลาด", error_message)