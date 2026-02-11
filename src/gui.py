import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QWidget, QHBoxLayout, QVBoxLayout, QListWidget
from PySide6.QtGui import QAction, QKeySequence, QKeyEvent, QPainter
from PySide6.QtCore import Qt, QFile
from PySide6 import QtUiTools

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        layout = QVBoxLayout(self)
        self.equipment_list = QListWidget()

        layout.addWidget(self.equipment_list)

        self.setLayout(layout)
        self.setCentralWidget(self.equipment_list)
 
    def keyPressEvent(self, event:QKeyEvent):
        if event.key() == Qt.Key.Key_Escape:
            self.close()
        return super().keyPressEvent(event)    

    def add_item(self, item:str):
        self.equipment_list.addItem(item)
        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow()
    w.showFullScreen()
    sys.exit(app.exec())
