
# ecr_kit_ui.py
import sys, os, traceback
import pandas as pd
from PyQt6.QtWidgets import QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QLabel
from PyQt6.QtGui import QPalette, QColor
from PyQt6.QtCore import Qt

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ECR Kit Assistant")
        tabs = QTabWidget()
        tabs.addTab(self.simple_tab("ECR Front Page"), "ECR Front Page")
        tabs.addTab(self.simple_tab("OBS Parts"), "OBS Parts List")
        tabs.addTab(self.simple_tab("Where Used"), "Where Used of OBS Parts")
        tabs.addTab(self.simple_tab("Orphan Analysis"), "Orphan Analysis")
        tabs.addTab(self.simple_tab("Inventory & Cost"), "Inventory Cost Analysis")
        tabs.addTab(self.simple_tab("Report"), "Report")
        tabs.addTab(self.simple_tab("User Notes"), "User Notes")
        self.setCentralWidget(tabs)
        self.resize(1200, 800)

    def simple_tab(self, title):
        w = QWidget(); l = QVBoxLayout(w)
        l.addWidget(QLabel(f"{title} UI Placeholder – full logic included in next drop"))
        return w

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    pal = QPalette()
    pal.setColor(QPalette.ColorRole.Window, QColor(245,248,252))
    app.setPalette(pal)
    win = MainWindow(); win.show()
    sys.exit(app.exec())
