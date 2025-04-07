import sys
import os
from PyQt6.QtWidgets import QApplication
from PyQt6.QtGui import QIcon
from app.mainwindow import MainWindow


def resource_path(relative_path):
    """รับพาธของไฟล์ที่อยู่ใน bundled exe หรือในโฟลเดอร์ปัจจุบัน"""
    if getattr(sys, "frozen", False):
        # กรณีรันจาก frozen executable
        base_path = sys._MEIPASS
    else:
        # กรณีรันจากสคริปต์ปกติ
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


if __name__ == "__main__":
    app = QApplication(sys.argv)

    # ตั้งค่าสไตล์ของแอปพลิเคชัน
    app.setStyle("Fusion")

    # ตั้งค่าไอคอนของแอปพลิเคชัน
    icon_path = resource_path("resources/icons/app_icon.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    # สร้างและแสดงหน้าต่างหลัก
    window = MainWindow()
    window.show()

    sys.exit(app.exec())