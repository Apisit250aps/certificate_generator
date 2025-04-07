from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                             QLineEdit, QPushButton, QComboBox, QGroupBox, 
                             QFileDialog, QGridLayout, QCheckBox, QFrame,
                             QSizePolicy)
from PyQt6.QtCore import pyqtSignal, Qt
from PyQt6.QtGui import QFont

class OutputConfig(QWidget):
    def __init__(self):
        super().__init__()
        
        # สร้าง Layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # สร้าง GroupBox
        self.group_box = QGroupBox("ตั้งค่าการสร้างไฟล์")
        group_layout = QGridLayout(self.group_box)
        group_layout.setVerticalSpacing(15)
        
        # สร้างวิดเจ็ตสำหรับเลือกโฟลเดอร์ output
        self.output_label = QLabel("โฟลเดอร์สำหรับบันทึกไฟล์:")
        
        # สร้าง container สำหรับ output path
        output_container = QFrame()
        output_container.setFrameShape(QFrame.Shape.NoFrame)
        output_layout = QHBoxLayout(output_container)
        output_layout.setContentsMargins(0, 0, 0, 0)
        
        self.output_path = QLineEdit()
        self.output_path.setReadOnly(True)
        self.output_path.setPlaceholderText("เลือกโฟลเดอร์...")
        
        self.output_browse_btn = QPushButton("เลือกโฟลเดอร์...")
        self.output_browse_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.output_browse_btn.setStyleSheet("""
            QPushButton {
                background-color: #9c27b0;
                color: white;
                padding: 5px 10px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #7b1fa2;
            }
            QPushButton:pressed {
                background-color: #6a1b9a;
            }
        """)
        self.output_browse_btn.clicked.connect(self.browse_output_dir)
        
        output_layout.addWidget(self.output_path, 5)
        output_layout.addWidget(self.output_browse_btn, 1)
        
        # สร้าง Frame สำหรับการตั้งค่ารูปแบบชื่อไฟล์
        filename_frame = QFrame()
        filename_frame.setFrameShape(QFrame.Shape.StyledPanel)
        filename_frame.setStyleSheet("background-color: #f3e5f5; border-radius: 5px;")
        filename_layout = QVBoxLayout(filename_frame)
        
        # สร้างวิดเจ็ตสำหรับเลือกคอลัมน์ที่ใช้ตั้งชื่อไฟล์
        self.filename_label = QLabel("ตัวเลือกตั้งชื่อไฟล์:")
        self.filename_label.setStyleSheet("background-color: transparent; font-weight: bold;")
        
        # สร้าง Container สำหรับตัวเลือกชื่อไฟล์
        self.columns_layout = QGridLayout()
        self.columns_layout.setVerticalSpacing(10)
        
        # เพิ่มตัวเลือกใช้ลำดับเป็นชื่อไฟล์
        self.use_index_checkbox = QCheckBox("ใช้ลำดับ")
        self.use_index_checkbox.setStyleSheet("background-color: transparent;")
        self.use_index_checkbox.setChecked(True)
        
        # สร้าง Combo Box สำหรับเลือกคอลัมน์
        col1_label = QLabel("คอลัมน์ที่ 1:")
        col1_label.setStyleSheet("background-color: transparent;")
        self.column1_combo = QComboBox()
        self.column1_combo.setStyleSheet("""
            QComboBox {
                padding: 5px;
                border: 1px solid #bdbdbd;
                border-radius: 3px;
                background-color: white;
            }
        """)
        
        col2_label = QLabel("คอลัมน์ที่ 2:")
        col2_label.setStyleSheet("background-color: transparent;")
        self.column2_combo = QComboBox()
        self.column2_combo.setStyleSheet("""
            QComboBox {
                padding: 5px;
                border: 1px solid #bdbdbd;
                border-radius: 3px;
                background-color: white;
            }
        """)
        
        # จัดเรียงในรูปแบบกริด
        self.columns_layout.addWidget(self.use_index_checkbox, 0, 0, 1, 2)
        self.columns_layout.addWidget(col1_label, 1, 0)
        self.columns_layout.addWidget(self.column1_combo, 1, 1)
        self.columns_layout.addWidget(col2_label, 2, 0)
        self.columns_layout.addWidget(self.column2_combo, 2, 1)
        
        # สร้างวิดเจ็ตสำหรับแสดงตัวอย่างรูปแบบชื่อไฟล์
        preview_frame = QFrame()
        preview_frame.setFrameShape(QFrame.Shape.StyledPanel)
        preview_frame.setStyleSheet("background-color: #e1bee7; border-radius: 5px; margin-top: 10px;")
        preview_layout = QVBoxLayout(preview_frame)
        
        self.format_label = QLabel("ตัวอย่างชื่อไฟล์:")
        self.format_label.setStyleSheet("background-color: transparent; font-weight: bold;")
        
        self.format_example = QLabel("ตัวอย่าง: 1_นายทดสอบ_ใจดี.pdf")
        self.format_example.setStyleSheet("background-color: transparent; font-style: italic;")
        
        preview_layout.addWidget(self.format_label)
        preview_layout.addWidget(self.format_example)
        
        # เพิ่ม layout ลงใน filename_frame
        filename_layout.addWidget(self.filename_label)
        filename_layout.addLayout(self.columns_layout)
        filename_layout.addWidget(preview_frame)
        
        # เพิ่มวิดเจ็ตลงใน GroupBox
        group_layout.addWidget(self.output_label, 0, 0)
        group_layout.addWidget(output_container, 0, 1)
        group_layout.addWidget(filename_frame, 1, 0, 1, 2)
        
        # เพิ่ม GroupBox ลงใน Layout หลัก
        layout.addWidget(self.group_box)
        
        # เชื่อมต่อสัญญาณ
        self.column1_combo.currentTextChanged.connect(self.update_example)
        self.column2_combo.currentTextChanged.connect(self.update_example)
        self.use_index_checkbox.toggled.connect(self.update_example)
        
        # ค่าเริ่มต้น
        self.column1_combo.addItem("-- เลือกคอลัมน์ --")
        self.column2_combo.addItem("-- เลือกคอลัมน์ --")
    
    def browse_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(
            self, "เลือกโฟลเดอร์สำหรับบันทึกไฟล์"
        )
        if dir_path:
            self.output_path.setText(dir_path)
            # อัปเดตสถานะว่าได้เลือกโฟลเดอร์แล้ว
            self.output_path.setStyleSheet("background-color: #e8f5e9; color: #1b5e20;")
    
    def set_columns(self, headers):
        self.column1_combo.clear()
        self.column2_combo.clear()
        
        self.column1_combo.addItem("-- เลือกคอลัมน์ --")
        self.column2_combo.addItem("-- เลือกคอลัมน์ --")
        
        for header in headers:
            self.column1_combo.addItem(header)
            self.column2_combo.addItem(header)
        
        # เลือกคอลัมน์แรกโดยอัตโนมัติ
        if len(headers) >= 1:
            self.column1_combo.setCurrentIndex(1)
        
        # เลือกคอลัมน์ที่สองโดยอัตโนมัติ (ถ้ามี)
        if len(headers) >= 2:
            self.column2_combo.setCurrentIndex(2)
        
        self.update_example()
    
    def update_example(self):
        use_index = self.use_index_checkbox.isChecked()
        col1 = self.column1_combo.currentText()
        col2 = self.column2_combo.currentText()
        
        example = ""
        if use_index:
            example += "1"
        
        if col1 and col1 != "-- เลือกคอลัมน์ --":
            if example:
                example += "_"
            example += "นายทดสอบ"
        
        if col2 and col2 != "-- เลือกคอลัมน์ --":
            if example:
                example += "_"
            example += "ใจดี"
        
        if not example:
            example = "ไม่มีข้อมูลที่เลือก"
        
        self.format_example.setText(f"ตัวอย่าง: {example}.pdf")
    
    def get_output_path(self):
        return self.output_path.text()
    
    def get_name_format(self):
        parts = []
        if self.use_index_checkbox.isChecked():
            parts.append("{}")
        
        col1 = self.column1_combo.currentText()
        if col1 and col1 != "-- เลือกคอลัมน์ --":
            parts.append("{}")
        
        col2 = self.column2_combo.currentText()
        if col2 and col2 != "-- เลือกคอลัมน์ --":
            parts.append("{}")
        
        return "_".join(parts)
    
    def get_selected_column_indices(self):
        indices = []
        
        if self.use_index_checkbox.isChecked():
            indices.append(-1)  # -1 หมายถึงใช้ index
        
        col1_idx = self.column1_combo.currentIndex() - 1  # -1 เพื่อข้ามตัวเลือก "-- เลือกคอลัมน์ --"
        if col1_idx >= 0:
            indices.append(col1_idx)
        
        col2_idx = self.column2_combo.currentIndex() - 1
        if col2_idx >= 0:
            indices.append(col2_idx)
        
        return indices