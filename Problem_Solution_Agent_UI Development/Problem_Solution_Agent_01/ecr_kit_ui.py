
# ecr_kit_ui.py (Enhanced)
import sys
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QTextEdit, QTableWidget, QTableWidgetItem, QCheckBox,
    QPushButton, QProgressBar, QFormLayout, QSizePolicy
)
from PyQt6.QtGui import QPalette, QColor, QFont
from PyQt6.QtCore import Qt

APP_TITLE = "ECR Kit Assistant"

class ReadmeTab(QWidget):
    def __init__(self, readme_path: Path):
        super().__init__()
        layout = QVBoxLayout(self)
        title = QLabel("README")
        title.setFont(QFont("Segoe UI", 16, QFont.Weight.DemiBold))
        layout.addWidget(title)

        text = QTextEdit()
        text.setReadOnly(True)
        content = "README.txt not found."
        try:
            if readme_path.exists():
                content = readme_path.read_text(encoding="utf-8")
        except Exception as e:
            content = f"Error opening README: {e}"
        text.setPlainText(content)
        layout.addWidget(text)

class ECRFrontPageTab(QWidget):
    CHECKLIST_ROWS = [
        ("Project Association", "Does the ECR include a Project (PCR)?"),
        ("Safety & Compliance", "Provide PSER details if any safety incident occurred."),
        ("Part Release Status", "Are all PPRs and first EVAL released?"),
        ("Replacement Readiness", "If parent part is in production, is the replacement part (BTP) production released?"),
        ("Design Analysis", "PACE / DFMEA / DASH parts addressed?"),
        ("V&V", "Is V&V plan/halts completed (for new designs/parts only)?"),
        ("Watchlist & Spares", "Are new parts/designs MLO certified / parent/previous parts MLO certified?"),
        ("OBS Impact", "Have potential OBS/old parts been added to the watchlist?"),
        ("ABS Approval", "Do we have ABS approval for OBS parts (sparable) without replacement?"),
        ("Config Documents", "Is CR available with change matrix details for new options/reference designator updates?"),
        ("Strategies Identified", "Provide reason code, strategy, priority, and alignment across BUs/products."),
        ("Multi BU Alignment", "If scope impacts multiple BUs/products, verify and confirm all affected BUs/products are listed."),
        ("Interchangeability & Tags", "Are these complying with CRP (003-10 Interchangeability Policy)? Provide interchangeability details."),
        ("Testing & Reports", "Are test results/FQR available for all IFF impacted parts/critical parts?"),
        ("Cost & Savings", "If the project relates to DCR (cost reduction), provide cost-saving details."),
    ]

    def __init__(self):
        super().__init__()
        outer = QVBoxLayout(self)
        outer.setSpacing(12)

        # Header controls arranged in two rows
        row1 = QWidget(); r1 = QHBoxLayout(row1)
        self.ecr_no = QLineEdit(); self.ecr_no.setPlaceholderText("ECR#")
        self.eco_primer = QLineEdit(); self.eco_primer.setPlaceholderText("ECO Primer Refs#")
        self.ec_category = QLineEdit(); self.ec_category.setPlaceholderText("EC Category")
        self.bu = QLineEdit(); self.bu.setPlaceholderText("BU")
        self.tco = QLineEdit(); self.tco.setPlaceholderText("TCO")
        self.project_no = QLineEdit(); self.project_no.setPlaceholderText("Project#")
        self.product = QLineEdit(); self.product.setPlaceholderText("Product")
        for w in [self.ecr_no, self.eco_primer, self.ec_category, self.bu, self.tco, self.project_no, self.product]:
            w.setFixedHeight(28)
            r1.addWidget(w)
        row2 = QWidget(); r2 = QHBoxLayout(row2)
        self.affected_modules = QLineEdit(); self.affected_modules.setPlaceholderText("Affected Module(s)")
        self.place = QLineEdit(); self.place.setPlaceholderText("Place")
        for w in [self.affected_modules, self.place]:
            w.setFixedHeight(28)
            r2.addWidget(w)
        outer.addWidget(row1)
        outer.addWidget(row2)

        # Checklist label
        lbl = QLabel("ECR Creation Checklist")
        lbl.setFont(QFont("Segoe UI", 13, QFont.Weight.DemiBold))
        outer.addWidget(lbl)

        # Checklist table
        table = QTableWidget(len(self.CHECKLIST_ROWS), 7)
        table.setHorizontalHeaderLabels(["", "ID/ID", "Category", "Validation", "Comments", "Action Owner", "Due Date"])
        table.verticalHeader().setVisible(False)
        table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        for row, (cat, desc) in enumerate(self.CHECKLIST_ROWS):
            chk = QCheckBox(); table.setCellWidget(row, 0, chk)
            table.setItem(row, 1, QTableWidgetItem(str(row+1)))
            table.setItem(row, 2, QTableWidgetItem(cat))
            table.setItem(row, 3, QTableWidgetItem(desc))
            table.setItem(row, 4, QTableWidgetItem(""))
            table.setItem(row, 5, QTableWidgetItem(""))
            table.setItem(row, 6, QTableWidgetItem(""))
        table.resizeColumnsToContents()
        table.setAlternatingRowColors(True)
        outer.addWidget(table)

        # Progress row
        progress_row = QWidget(); pr_layout = QHBoxLayout(progress_row)
        self.progress_label = QLabel(f"Checklist Progress: 0 / {len(self.CHECKLIST_ROWS)}")
        self.progress_bar = QProgressBar(); self.progress_bar.setRange(0, len(self.CHECKLIST_ROWS))
        pr_layout.addWidget(self.progress_label)
        pr_layout.addWidget(QLabel("Progress Bar"))
        pr_layout.addWidget(self.progress_bar)
        outer.addWidget(progress_row)

        # Action row
        action_row = QWidget(); ar = QHBoxLayout(action_row)
        self.btn_generate = QPushButton("Generate Problem Statement")
        self.title_edit = QLineEdit(); self.title_edit.setPlaceholderText("Title:")
        ar.addWidget(self.btn_generate)
        ar.addWidget(self.title_edit)
        outer.addWidget(action_row)

        # Problem statement
        ps_label = QLabel("Problem Statement")
        ps_label.setFont(QFont("Segoe UI", 12, QFont.Weight.Medium))
        self.problem_edit = QTextEdit(); self.problem_edit.setPlaceholderText("Enter the problem statement here…")
        self.problem_edit.setFixedHeight(140)
        outer.addWidget(ps_label)
        outer.addWidget(self.problem_edit)

        def recalc():
            checked = 0
            for r in range(table.rowCount()):
                w = table.cellWidget(r, 0)
                if isinstance(w, QCheckBox) and w.isChecked():
                    checked += 1
            self.progress_label.setText(f"Checklist Progress: {checked} / {table.rowCount()}")
            self.progress_bar.setValue(checked)
        for r in range(table.rowCount()):
            w = table.cellWidget(r, 0)
            if isinstance(w, QCheckBox):
                w.stateChanged.connect(recalc)

class PlaceholderTab(QWidget):
    def __init__(self, title: str):
        super().__init__()
        l = QVBoxLayout(self)
        l.addWidget(QLabel(f"{title} – UI under development"))

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1280, 860)

        tabs = QTabWidget(); tabs.setDocumentMode(True); tabs.setMovable(True)
        tabs.setStyleSheet(
            "QTabBar::tab { background: #EAF2FB; color: #1F3B57; padding: 8px 14px; border: 1px solid #D5E3F6; border-bottom: none; border-top-left-radius:6px; border-top-right-radius:6px; }"
            "QTabBar::tab:selected { background: #FFFFFF; color: #0F2D46; font-weight: 600; }"
            "QTabWidget::pane { border: 1px solid #D5E3F6; top: -1px; }"
        )

        readme_path = Path(__file__).with_name('README.txt')
        tabs.addTab(ReadmeTab(readme_path), "README")
        tabs.addTab(ECRFrontPageTab(), "ECR Front Page")
        tabs.addTab(PlaceholderTab("OBS Parts List"), "OBS Parts")
        tabs.addTab(PlaceholderTab("Where Used of OBS Parts"), "Where Used")
        tabs.addTab(PlaceholderTab("Orphan Analysis"), "Orphan Analysis")
        tabs.addTab(PlaceholderTab("Inventory Cost Analysis"), "Inventory & Cost")
        tabs.addTab(PlaceholderTab("Report"), "Report")
        tabs.addTab(PlaceholderTab("User Notes"), "User Notes")
        self.setCentralWidget(tabs)


def run():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    pal = QPalette()
    pal.setColor(QPalette.ColorRole.Window, QColor(247, 250, 253))
    pal.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
    pal.setColor(QPalette.ColorRole.AlternateBase, QColor(241, 246, 252))
    pal.setColor(QPalette.ColorRole.Text, QColor(28, 41, 56))
    app.setPalette(pal)

    win = MainWindow(); win.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    run()
