from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                             QLineEdit, QPushButton, QComboBox, QGroupBox, 
                             QFileDialog, QGridLayout, QCheckBox, QFrame,
                             QSizePolicy, QListWidget, QAbstractItemView, 
                             QToolButton, QDialog, QDialogButtonBox, QSpinBox,
                             QListWidgetItem, QMessageBox)
from PyQt6.QtCore import pyqtSignal, Qt
from PyQt6.QtGui import QFont, QIcon, QColor, QBrush

class CustomTextDialog(QDialog):
    """ไดอะล็อกสำหรับกำหนดข้อความเอง"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("กำหนดข้อความเอง")
        self.resize(400, 150)
        
        layout = QVBoxLayout(self)
        
        # ส่วนกำหนดข้อความ
        form_layout = QGridLayout()
        self.text_label = QLabel("ข้อความ:")
        self.text_input = QLineEdit()
        self.text_input.setPlaceholderText("ใส่ข้อความที่ต้องการ...")
        form_layout.addWidget(self.text_label, 0, 0)
        form_layout.addWidget(self.text_input, 0, 1)
        
        layout.addLayout(form_layout)
        
        # ปุ่มยกเลิก/ตกลง
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def get_custom_text(self):
        """คืนค่าข้อความที่ผู้ใช้กำหนด"""
        return self.text_input.text()
    
    def set_custom_text(self, text):
        """ตั้งค่าข้อความเริ่มต้น"""
        self.text_input.setText(text)

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
        
        # สร้างวิดเจ็ตสำหรับส่วนของการตั้งค่าชื่อไฟล์
        self.filename_label = QLabel("ตั้งค่ารูปแบบชื่อไฟล์:")
        self.filename_label.setStyleSheet("background-color: transparent; font-weight: bold;")
        
        # สร้างคำอธิบายเพิ่มเติม
        self.help_label = QLabel("เพิ่มส่วนประกอบของชื่อไฟล์ด้านล่าง (สามารถลากเพื่อจัดลำดับใหม่):")
        self.help_label.setStyleSheet("background-color: transparent; font-style: italic;")
        
        # สร้าง ListWidget สำหรับแสดงส่วนประกอบของชื่อไฟล์
        self.filename_parts_list = QListWidget()
        self.filename_parts_list.setStyleSheet("""
            QListWidget {
                background-color: white;
                border: 1px solid #bdbdbd;
                border-radius: 3px;
                padding: 5px;
            }
            QListWidget::item:selected {
                background-color: #e0f7fa;
                color: #000000;
                border: 1px solid #80deea;
                border-radius: 2px;
            }
            QListWidget::item {
                padding: 5px;
                margin: 2px;
            }
        """)
        self.filename_parts_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.filename_parts_list.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.filename_parts_list.setMinimumHeight(120)
        self.filename_parts_list.model().rowsMoved.connect(self.update_parts_order)
        
        # ส่วนควบคุมการเพิ่ม/ลบส่วนประกอบของชื่อไฟล์
        controls_layout = QHBoxLayout()
        
        # ปุ่มเพิ่มลำดับ
        self.add_index_btn = QPushButton("เพิ่มลำดับ")
        self.add_index_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196f3;
                color: white;
                padding: 5px 10px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #1976d2;
            }
        """)
        self.add_index_btn.clicked.connect(self.add_index_part)
        
        # ปุ่มเพิ่มคอลัมน์
        self.add_column_btn = QPushButton("เพิ่มคอลัมน์")
        self.add_column_btn.setStyleSheet("""
            QPushButton {
                background-color: #4caf50;
                color: white;
                padding: 5px 10px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #388e3c;
            }
        """)
        self.add_column_btn.clicked.connect(self.show_column_dialog)
        
        # ปุ่มเพิ่มข้อความกำหนดเอง
        self.add_custom_btn = QPushButton("เพิ่มข้อความ")
        self.add_custom_btn.setStyleSheet("""
            QPushButton {
                background-color: #ff9800;
                color: white;
                padding: 5px 10px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #f57c00;
            }
        """)
        self.add_custom_btn.clicked.connect(self.add_custom_text)
        
        # ปุ่มลบรายการที่เลือก
        self.remove_btn = QPushButton("ลบรายการ")
        self.remove_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                padding: 5px 10px;
                border-radius: 3px;
            }
            QPushButton:hover {
                background-color: #d32f2f;
            }
        """)
        self.remove_btn.clicked.connect(self.remove_selected_part)
        
        controls_layout.addWidget(self.add_index_btn)
        controls_layout.addWidget(self.add_column_btn)
        controls_layout.addWidget(self.add_custom_btn)
        controls_layout.addWidget(self.remove_btn)
        
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
        filename_layout.addWidget(self.help_label)
        filename_layout.addWidget(self.filename_parts_list)
        filename_layout.addLayout(controls_layout)
        filename_layout.addWidget(preview_frame)
        
        # เพิ่มวิดเจ็ตลงใน GroupBox
        group_layout.addWidget(self.output_label, 0, 0)
        group_layout.addWidget(output_container, 0, 1)
        group_layout.addWidget(filename_frame, 1, 0, 1, 2)
        
        # เพิ่ม GroupBox ลงใน Layout หลัก
        layout.addWidget(self.group_box)
        
        # เก็บข้อมูลคอลัมน์
        self.headers = []
        
        # เตรียมข้อมูลสำหรับเก็บส่วนประกอบของชื่อไฟล์
        self.filename_parts = []  # เก็บข้อมูลในรูปแบบ (type, value)
    
    def browse_output_dir(self):
        dir_path = QFileDialog.getExistingDirectory(
            self, "เลือกโฟลเดอร์สำหรับบันทึกไฟล์"
        )
        if dir_path:
            self.output_path.setText(dir_path)
            # อัปเดตสถานะว่าได้เลือกโฟลเดอร์แล้ว
            self.output_path.setStyleSheet("background-color: #e8f5e9; color: #1b5e20;")
    
    def set_columns(self, headers):
        """อัปเดตรายการคอลัมน์จาก Excel"""
        self.headers = headers
        self.update_example()
        
        # เพิ่มคอลัมน์แรกและคอลัมน์ที่สองโดยอัตโนมัติถ้ามี และถ้ายังไม่มีรายการใดในลิสต์
        if len(self.filename_parts) == 0:
            # เพิ่มลำดับ
            self.add_index_part()
            
            # เพิ่มคอลัมน์แรก
            if len(headers) >= 1:
                # สร้าง item สำหรับคอลัมน์แรกโดยตรง โดยไม่ผ่านไดอะล็อก
                self.add_column_direct(0)
            
            # เพิ่มคอลัมน์ที่สอง
            if len(headers) >= 2:
                # สร้าง item สำหรับคอลัมน์ที่สองโดยตรง โดยไม่ผ่านไดอะล็อก
                self.add_column_direct(1)
    
    def add_index_part(self):
        """เพิ่มลำดับเป็นส่วนหนึ่งของชื่อไฟล์"""
        self.filename_parts.append(("index", None))
        item = QListWidgetItem("ลำดับ")
        item.setBackground(QBrush(QColor("#e3f2fd")))
        self.filename_parts_list.addItem(item)
        self.update_example()
    
    def show_column_dialog(self):
        """แสดงไดอะล็อกเลือกคอลัมน์"""
        if not self.headers:
            QMessageBox.warning(self, "ไม่พบข้อมูลคอลัมน์", "กรุณาโหลดไฟล์ Excel และเลือกชีทก่อน")
            return
        
        # สร้างและแสดงไดอะล็อก
        dialog = QDialog(self)
        dialog.setWindowTitle("เลือกคอลัมน์")
        dialog.resize(400, 200)
        
        layout = QVBoxLayout(dialog)
        
        # เพิ่มคำอธิบาย
        info_label = QLabel("เลือกคอลัมน์จาก Excel ที่ต้องการใช้ในชื่อไฟล์:")
        layout.addWidget(info_label)
        
        # เพิ่ม ComboBox สำหรับเลือกคอลัมน์
        column_combo = QComboBox()
        for header in self.headers:
            column_combo.addItem(header)
        
        layout.addWidget(column_combo)
        
        # ปุ่มยกเลิก/ตกลง
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)
        
        # แสดงไดอะล็อก
        result = dialog.exec()
        
        # ตรวจสอบและดำเนินการตามผลลัพธ์
        if result == QDialog.DialogCode.Accepted:
            column_index = column_combo.currentIndex()
            column_name = column_combo.currentText()
            self.add_column_direct(column_index)
    
    def add_column_direct(self, column_index):
        """เพิ่มคอลัมน์โดยตรงโดยใช้ column_index"""
        if 0 <= column_index < len(self.headers):
            column_name = self.headers[column_index]
            self.filename_parts.append(("column", column_index))
            item = QListWidgetItem(f"คอลัมน์: {column_name}")
            item.setBackground(QBrush(QColor("#e8f5e9")))
            self.filename_parts_list.addItem(item)
            self.update_example()
    
    def add_custom_text(self):
        """เพิ่มข้อความกำหนดเองเป็นส่วนหนึ่งของชื่อไฟล์"""
        dialog = CustomTextDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            text = dialog.get_custom_text()
            if text:
                self.filename_parts.append(("custom", text))
                item = QListWidgetItem(f"ข้อความ: {text}")
                item.setBackground(QBrush(QColor("#fff3e0")))
                self.filename_parts_list.addItem(item)
                self.update_example()
    
    def remove_selected_part(self):
        """ลบส่วนประกอบของชื่อไฟล์ที่เลือก"""
        current_row = self.filename_parts_list.currentRow()
        if current_row >= 0:
            self.filename_parts_list.takeItem(current_row)
            self.filename_parts.pop(current_row)
            self.update_example()
    
    def update_parts_order(self):
        """อัปเดตลำดับของส่วนประกอบชื่อไฟล์เมื่อมีการลากและวาง"""
        self.sync_parts_with_listwidget()
        self.update_example()
    
    def sync_parts_with_listwidget(self):
        """ซิงค์ข้อมูล filename_parts กับรายการใน ListWidget"""
        new_parts = []
        for i in range(self.filename_parts_list.count()):
            item_text = self.filename_parts_list.item(i).text()
            
            if item_text == "ลำดับ":
                new_parts.append(("index", None))
            elif item_text.startswith("คอลัมน์:"):
                header = item_text[len("คอลัมน์: "):]
                try:
                    column_index = self.headers.index(header)
                    new_parts.append(("column", column_index))
                except ValueError:
                    # กรณีที่ไม่พบคอลัมน์ในรายการคอลัมน์ปัจจุบัน
                    continue
            elif item_text.startswith("ข้อความ:"):
                text = item_text[len("ข้อความ: "):]
                new_parts.append(("custom", text))
        
        self.filename_parts = new_parts
    
    def update_example(self):
        """อัปเดตตัวอย่างชื่อไฟล์"""
        example = []
        for part_type, value in self.filename_parts:
            if part_type == "index":
                example.append("1")
            elif part_type == "column":
                if self.headers and 0 <= value < len(self.headers):
                    # แสดงตัวอย่างด้วยค่าจำลอง
                    header = self.headers[value].lower()
                    if "name" in header or "ชื่อ" in header:
                        example.append("นายทดสอบ")
                    elif "surname" in header or "นามสกุล" in header or "สกุล" in header:
                        example.append("ใจดี")
                    elif "id" in header or "รหัส" in header:
                        example.append("12345")
                    else:
                        example.append(f"ข้อมูล{self.headers[value]}")
            elif part_type == "custom":
                example.append(value)
        
        if example:
            self.format_example.setText(f"ตัวอย่าง: {'_'.join(example)}.pdf")
        else:
            self.format_example.setText("ตัวอย่าง: ไม่มีรูปแบบที่กำหนด.pdf")
    
    def get_output_path(self):
        return self.output_path.text()
    
    def get_name_format(self):
        """คืนค่ารูปแบบของชื่อไฟล์"""
        # สร้างรูปแบบสำหรับการ format
        parts = []
        for _ in self.filename_parts:
            parts.append("{}")
        
        return "_".join(parts)
    
    def get_selected_column_indices(self):
        """คืนค่าลิสต์ของค่าที่ใช้ในการตั้งชื่อไฟล์"""
        values = []
        for part_type, value in self.filename_parts:
            if part_type == "index":
                values.append(-1)  # -1 หมายถึงใช้ index
            elif part_type == "column":
                values.append(value)
            elif part_type == "custom":
                values.append(value)  # ส่งค่าข้อความกำหนดเอง
        
        return values