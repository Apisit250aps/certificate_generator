from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                             QLineEdit, QPushButton, QComboBox, QGroupBox, 
                             QFileDialog, QGridLayout, QCheckBox)
from PyQt6.QtCore import pyqtSignal

class OutputConfig(QWidget):
    def __init__(self):
        super().__init__()
        
        # สร้าง Layout
        layout = QVBoxLayout(self)
        
        # สร้าง GroupBox
        self.group_box = QGroupBox("ตั้งค่าการสร้างไฟล์")
        group_layout = QGridLayout(self.group_box)
        
        # สร้างวิดเจ็ตสำหรับเลือกโฟลเดอร์ output
        self.output_label = QLabel("โฟลเดอร์สำหรับบันทึกไฟล์:")
        self.output_path = QLineEdit()
        self.output_path.setReadOnly(True)
        self.output_browse_btn = QPushButton("เลือกโฟลเดอร์...")
        self.output_browse_btn.clicked.connect(self.browse_output_dir)
        
        # สร้างวิดเจ็ตสำหรับเลือกคอลัมน์ที่ใช้ตั้งชื่อไฟล์
        self.filename_label = QLabel("ตัวเลือกตั้งชื่อไฟล์:")
        self.columns_layout = QHBoxLayout()
        
        # เพิ่มตัวเลือกใช้ลำดับเป็นชื่อไฟล์
        self.use_index_checkbox = QCheckBox("ใช้ลำดับ")
        self.use_index_checkbox.setChecked(True)
        
        # สร้าง Combo Box สำหรับเลือกคอลัมน์
        self.column1_combo = QComboBox()
        self.column2_combo = QComboBox()
        
        self.columns_layout.addWidget(self.use_index_checkbox)
        self.columns_layout.addWidget(self.column1_combo)
        self.columns_layout.addWidget(self.column2_combo)
        self.columns_layout.addStretch()
        
        # สร้างวิดเจ็ตสำหรับแสดงตัวอย่างรูปแบบชื่อไฟล์
        self.format_label = QLabel("รูปแบบชื่อไฟล์:")
        self.format_example = QLabel("ตัวอย่าง: 1_นายทดสอบ_ใจดี.pdf")
        
        # เพิ่มวิดเจ็ตลงใน GroupBox
        group_layout.addWidget(self.output_label, 0, 0)
        group_layout.addWidget(self.output_path, 0, 1)
        group_layout.addWidget(self.output_browse_btn, 0, 2)
        
        group_layout.addWidget(self.filename_label, 1, 0)
        group_layout.addLayout(self.columns_layout, 1, 1, 1, 2)
        
        group_layout.addWidget(self.format_label, 2, 0)
        group_layout.addWidget(self.format_example, 2, 1, 1, 2)
        
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