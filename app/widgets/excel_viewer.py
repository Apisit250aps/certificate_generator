from PyQt6.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QLabel, 
                            QComboBox, QTableView, QGroupBox, QHeaderView,
                            QFrame, QSizePolicy, QAbstractItemView)
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
        if not index.isValid():
            return None
            
        if role == Qt.ItemDataRole.DisplayRole:
            return str(self._data[index.row()][index.column()])
        elif role == Qt.ItemDataRole.TextAlignmentRole:
            # ตัวเลขชิดขวา ข้อความชิดซ้าย
            value = self._data[index.row()][index.column()]
            if isinstance(value, (int, float)):
                return Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
            return Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        elif role == Qt.ItemDataRole.BackgroundRole:
            # สลับสีแถวเพื่อให้อ่านง่าย
            if index.row() % 2 == 0:
                return Qt.GlobalColor.white
            else:
                return Qt.GlobalColor.gray
        return None
    
    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if orientation == Qt.Orientation.Horizontal and role == Qt.ItemDataRole.DisplayRole:
            return self._headers[section]
        elif orientation == Qt.Orientation.Vertical and role == Qt.ItemDataRole.DisplayRole:
            # แสดงลำดับแถว
            return str(section + 1)
        return None

class ExcelViewer(QWidget):
    sheet_selected = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        
        # สร้าง Layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # สร้าง GroupBox
        self.group_box = QGroupBox("ข้อมูลในไฟล์ Excel")
        group_layout = QVBoxLayout(self.group_box)
        
        # สร้าง Frame สำหรับส่วนเลือกชีท
        sheet_frame = QFrame()
        sheet_frame.setFrameShape(QFrame.Shape.StyledPanel)
        sheet_frame.setStyleSheet("background-color: #f0f8ff; border-radius: 5px;")
        sheet_layout = QHBoxLayout(sheet_frame)
        
        self.sheet_label = QLabel("เลือกชีท:")
        self.sheet_label.setStyleSheet("background-color: transparent;")
        
        self.sheet_combo = QComboBox()
        self.sheet_combo.setStyleSheet("""
            QComboBox {
                padding: 5px;
                border: 1px solid #bdbdbd;
                border-radius: 3px;
                background-color: white;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: center right;
                width: 20px;
                border-left: 1px solid #bdbdbd;
            }
            QComboBox QAbstractItemView {
                selection-background-color: #e0f0ff;
                selection-color: black;
            }
        """)
        self.sheet_combo.currentTextChanged.connect(self.on_sheet_changed)
        
        sheet_layout.addWidget(self.sheet_label)
        sheet_layout.addWidget(self.sheet_combo, 1)
        
        # สร้างวิดเจ็ตสำหรับแสดงข้อมูล
        self.table_view = QTableView()
        self.table_view.setStyleSheet("""
            QTableView {
                gridline-color: #d0d0d0;
                selection-background-color: #e0f0ff;
                selection-color: black;
                border: 1px solid #bdbdbd;
                border-radius: 3px;
            }
            QTableView::item:selected {
                background-color: #bbdefb;
            }
        """)
        self.table_view.setAlternatingRowColors(True)
        self.table_view.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table_view.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table_view.setSortingEnabled(True)
        
        self.table_model = ExcelTableModel()
        self.table_view.setModel(self.table_model)
        
        # ตั้งค่าเมื่อไม่มีข้อมูล
        no_data_label = QLabel("ยังไม่มีข้อมูล กรุณาเลือกไฟล์ Excel และชีท")
        no_data_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        no_data_label.setStyleSheet("font-style: italic; color: #777;")
        self.table_view.setModel(None)
        
        # เพิ่มวิดเจ็ตลงใน GroupBox
        group_layout.addWidget(sheet_frame)
        group_layout.addWidget(self.table_view)
        
        # เพิ่ม GroupBox ลงใน Layout หลัก
        layout.addWidget(self.group_box)
    
    def set_sheet_names(self, sheet_names):
        self.sheet_combo.clear()
        if not sheet_names:
            return
            
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
        
        # ตั้งค่าให้คอลัมน์สุดท้ายขยายเต็มพื้นที่
        self.table_view.horizontalHeader().setStretchLastSection(True)
    
    def get_selected_sheet(self):
        return self.sheet_combo.currentText()