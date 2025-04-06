from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                            QComboBox, QTableView, QGroupBox, QHeaderView)
from PyQt6.QtCore import pyqtSignal, Qt, QAbstractTableModel

class ExcelTableModel(QAbstractTableModel):
    def __init__(self, data=None, headers=None):
        super().__init__()
        self._data = data or []
        self._headers = headers or []
    
    def rowCount(self, parent=None):
        return len(self._data)
    
    def columnCount(self, parent=None):
        return len(self._headers)
    
    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            return str(self._data[index.row()][index.column()])
        return None
    
    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return self._headers[section]
        return None

class ExcelViewer(QWidget):
    sheet_selected = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        
        # สร้าง Layout
        layout = QVBoxLayout(self)
        
        # สร้าง GroupBox
        self.group_box = QGroupBox("ข้อมูลในไฟล์ Excel")
        group_layout = QVBoxLayout(self.group_box)
        
        # สร้างวิดเจ็ตสำหรับเลือกชีท
        sheet_layout = QHBoxLayout()
        self.sheet_label = QLabel("เลือกชีท:")
        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self.on_sheet_changed)
        
        sheet_layout.addWidget(self.sheet_label)
        sheet_layout.addWidget(self.sheet_combo)
        sheet_layout.addStretch()
        
        # สร้างวิดเจ็ตสำหรับแสดงข้อมูล
        self.table_view = QTableView()
        self.table_model = ExcelTableModel()
        self.table_view.setModel(self.table_model)
        
        # เพิ่มวิดเจ็ตลงใน GroupBox
        group_layout.addLayout(sheet_layout)
        group_layout.addWidget(self.table_view)
        
        # เพิ่ม GroupBox ลงใน Layout หลัก
        layout.addWidget(self.group_box)
    
    def set_sheet_names(self, sheet_names):
        self.sheet_combo.clear()
        self.sheet_combo.addItems(sheet_names)
        if sheet_names:
            self.sheet_combo.setCurrentIndex(0)
    
    def on_sheet_changed(self, sheet_name):
        if sheet_name:
            self.sheet_selected.emit(sheet_name)
    
    def set_data(self, data, headers):
        self.table_model = ExcelTableModel(data, headers)
        self.table_view.setModel(self.table_model)
        
        # ปรับความกว้างของคอลัมน์
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
    
    def get_selected_sheet(self):
        return self.sheet_combo.currentText()