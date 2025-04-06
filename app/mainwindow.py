from PyQt6.QtWidgets import (QMainWindow, QWidget, QGridLayout, QVBoxLayout, 
                            QHBoxLayout, QPushButton, QLabel, QMessageBox,
                            QStatusBar, QProgressBar)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt6.QtGui import QIcon, QAction

from app.widgets.file_selector import FileSelector
from app.widgets.excel_viewer import ExcelViewer
from app.widgets.output_config import OutputConfig
from core.pdf_manager import PDFManager
from core.excel_manager import ExcelManager

class CertificateGeneratorWorker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    
    def __init__(self, pdf_path, excel_path, sheet_name, output_dir, name_format, column_indices):
        super().__init__()
        self.pdf_path = pdf_path
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.output_dir = output_dir
        self.name_format = name_format
        self.column_indices = column_indices
        
    def run(self):
        try:
            pdf_manager = PDFManager(self.pdf_path)
            excel_manager = ExcelManager(self.excel_path)
            data = excel_manager.get_data(self.sheet_name)
            
            total = len(data)
            for i, row in enumerate(data):
                # สร้างชื่อไฟล์จากคอลัมน์ที่เลือก
                filename_parts = []
                for idx in self.column_indices:
                    if idx == -1:  # กรณีเลือก index
                        filename_parts.append(str(i+1))
                    else:
                        filename_parts.append(str(row[idx]))
                
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

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Certificate Generator")
        self.setMinimumSize(900, 600)
        
        # สร้าง Central Widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # สร้าง Layout หลัก
        main_layout = QVBoxLayout(central_widget)
        
        # สร้างวิดเจ็ตต่างๆ
        self.file_selector = FileSelector()
        self.excel_viewer = ExcelViewer()
        self.output_config = OutputConfig()
        
        # เพิ่มวิดเจ็ตลงใน Layout
        main_layout.addWidget(self.file_selector)
        main_layout.addWidget(self.excel_viewer)
        main_layout.addWidget(self.output_config)
        
        # สร้าง Control Panel
        control_panel = QHBoxLayout()
        
        # สร้างปุ่มสร้างเกียรติบัตร
        self.generate_btn = QPushButton("สร้างเกียรติบัตร")
        self.generate_btn.setMinimumHeight(40)
        self.generate_btn.clicked.connect(self.generate_certificates)
        
        # เพิ่ม Progress Bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        
        # เพิ่มองค์ประกอบลงใน Control Panel
        control_panel.addWidget(self.progress_bar)
        control_panel.addWidget(self.generate_btn)
        
        # เพิ่ม Control Panel ลงใน Layout หลัก
        main_layout.addLayout(control_panel)
        
        # สร้าง Status Bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("พร้อมใช้งาน")
        
        # เชื่อมต่อสัญญาณระหว่างวิดเจ็ต
        self.connect_signals()
        
        # แสดงหน้าต่างเต็มจอ
        self.showMaximized()
    
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
        column_indices = self.output_config.get_selected_column_indices()
        
        if not all([pdf_path, excel_path, sheet_name, output_dir]):
            QMessageBox.warning(self, "ข้อมูลไม่ครบถ้วน", "กรุณาเลือกไฟล์ PDF, Excel, ชีท และโฟลเดอร์สำหรับไฟล์ output")
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
            pdf_path, excel_path, sheet_name, output_dir, name_format, column_indices
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