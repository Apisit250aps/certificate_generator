# ใน main.py
import sys
import os
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QIcon
from app.mainwindow import MainWindow

def resource_path(relative_path):
    """รับพาธของไฟล์ที่อยู่ใน bundled exe หรือในโฟลเดอร์ปัจจุบัน"""
    if getattr(sys, 'frozen', False):
        # กรณีรันจาก frozen executable
        base_path = sys._MEIPASS
    else:
        # กรณีรันจากสคริปต์ปกติ
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # ตั้งค่าไอคอนระดับแอปพลิเคชัน
    icon_path = resource_path("resources/icons/app_icon.ico")
    app.setWindowIcon(QIcon(icon_path))
    
    # สร้างและแสดงหน้าต่างหลัก
    window = MainWindow()
    # ตั้งค่าไอคอนระดับหน้าต่าง
    window.setWindowIcon(QIcon(icon_path))
    window.show()
    
    sys.exit(app.exec())