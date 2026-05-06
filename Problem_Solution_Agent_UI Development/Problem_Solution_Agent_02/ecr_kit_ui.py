
# ecr_kit_ui.py (Enhanced v2)
import sys
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QTextEdit, QTableWidget, QTableWidgetItem, QCheckBox,
    QPushButton, QProgressBar, QSizePolicy, QScrollArea
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

        # --- Wrap all content in a scroll area to ensure horizontal+vertical scrolling ---
        top = QVBoxLayout(self)
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        top.addWidget(scroll)

        content = QWidget()
        scroll.setWidget(content)
        outer = QVBoxLayout(content)
        outer.setSpacing(10)
        outer.setContentsMargins(8, 8, 8, 8)

        # ---------- Header fields (two rows) ----------
        row1 = QWidget(); r1 = QHBoxLayout(row1); r1.setContentsMargins(12,0,0,0)
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
        row2 = QWidget(); r2 = QHBoxLayout(row2); r2.setContentsMargins(12,0,0,0)
        self.affected_modules = QLineEdit(); self.affected_modules.setPlaceholderText("Affected Module(s)")
        self.place = QLineEdit(); self.place.setPlaceholderText("Place")
        for w in [self.affected_modules, self.place]:
            w.setFixedHeight(28)
            r2.addWidget(w)
        outer.addWidget(row1)
        outer.addWidget(row2)

        # ---------- Checklist label ----------
        lbl = QLabel("ECR Creation Checklist")
        lbl.setFont(QFont("Segoe UI", 13, QFont.Weight.DemiBold))
        temp_lbl_row = QWidget(); tl = QHBoxLayout(temp_lbl_row); tl.setContentsMargins(12,0,0,0); tl.addWidget(lbl)
        outer.addWidget(temp_lbl_row)

        # ---------- Checklist table inside a row with a left gutter to move slightly right ----------
        table_row = QWidget(); tr = QHBoxLayout(table_row); tr.setContentsMargins(12,0,0,0)
        table = QTableWidget(len(self.CHECKLIST_ROWS), 7)
        table.setHorizontalHeaderLabels(["", "ID/ID", "Category", "Validation", "Comments", "Action Owner", "Due Date"])
        table.verticalHeader().setVisible(False)
        table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        table.setAlternatingRowColors(True)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        # Populate rows
        for row, (cat, desc) in enumerate(self.CHECKLIST_ROWS):
            chk = QCheckBox(); table.setCellWidget(row, 0, chk)
            table.setItem(row, 1, QTableWidgetItem(str(row+1)))
            table.setItem(row, 2, QTableWidgetItem(cat))
            table.setItem(row, 3, QTableWidgetItem(desc))
            table.setItem(row, 4, QTableWidgetItem("") )
            table.setItem(row, 5, QTableWidgetItem("") )
            table.setItem(row, 6, QTableWidgetItem("") )

        # Resize behavior
        try:
            from PyQt6.QtWidgets import QHeaderView
            header = table.horizontalHeader()
            header.setStretchLastSection(True)
            # Make Category and Validation columns wider
            header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        except Exception:
            pass

        tr.addWidget(table)
        outer.addWidget(table_row)

        # ---------- Progress (slim bar, aligned to table width) ----------
        progress_row = QWidget(); pr = QHBoxLayout(progress_row); pr.setContentsMargins(12,0,0,0)
        self.progress_label = QLabel(f"Checklist Progress: 0 / {len(self.CHECKLIST_ROWS)}")
        self.progress_bar = QProgressBar(); self.progress_bar.setRange(0, len(self.CHECKLIST_ROWS))
        self.progress_bar.setMaximumHeight(14)  # slimmer bar
        self.progress_bar.setTextVisible(True)
        pr.addWidget(self.progress_label)
        pr.addWidget(self.progress_bar, 1)
        outer.addWidget(progress_row)

        # ---------- Action row (button + title) with 75% title width and 75 char cap ----------
        action_row = QWidget(); ar = QHBoxLayout(action_row); ar.setContentsMargins(12,0,0,0)
        self.btn_generate = QPushButton("Generate Problem Statement")
        self.title_edit = QLineEdit(); self.title_edit.setPlaceholderText("Title (max 75 chars)")
        self.title_edit.setMaxLength(75)
        # Layout stretching: button : title : spacer = 1 : 3 : 1  (title ~75% of remaining width)
        ar.addWidget(self.btn_generate, 1)
        ar.addWidget(self.title_edit, 3)
        ar.addStretch(1)
        outer.addWidget(action_row)

        # ---------- Problem Statement (75% width, 2000 char limit) ----------
        ps_label_row = QWidget(); psl = QHBoxLayout(ps_label_row); psl.setContentsMargins(12,0,0,0)
        ps_label = QLabel("Problem Statement")
        ps_label.setFont(QFont("Segoe UI", 12, QFont.Weight.Medium))
        psl.addWidget(ps_label)
        outer.addWidget(ps_label_row)

        ps_row = QWidget(); psr = QHBoxLayout(ps_row); psr.setContentsMargins(12,0,0,0)
        self.problem_edit = QTextEdit(); self.problem_edit.setPlaceholderText("Write the problem statement here (max 2000 characters)…")
        self.problem_edit.setFixedHeight(140)
        psr.addWidget(self.problem_edit, 3)
        psr.addStretch(1)  # ensures editor ~75% of row width
        outer.addWidget(ps_row)

        # Limit to 2000 characters
        self.problem_edit.textChanged.connect(lambda: self._limit_text(self.problem_edit, 2000))

        # ---------- Solution Statement (same as Problem, with colors) ----------
        ss_label_row = QWidget(); ssl = QHBoxLayout(ss_label_row); ssl.setContentsMargins(12,0,0,0)
        ss_label = QLabel("Solution Statement")
        ss_label.setFont(QFont("Segoe UI", 12, QFont.Weight.Medium))
        ssl.addWidget(ss_label)
        outer.addWidget(ss_label_row)

        ss_row = QWidget(); ssr = QHBoxLayout(ss_row); ssr.setContentsMargins(12,0,0,0)
        self.solution_edit = QTextEdit(); self.solution_edit.setPlaceholderText("Write the proposed solution here (max 2000 characters)…")
        self.solution_edit.setFixedHeight(140)
        ssr.addWidget(self.solution_edit, 3)
        ssr.addStretch(1)
        outer.addWidget(ss_row)

        self.solution_edit.textChanged.connect(lambda: self._limit_text(self.solution_edit, 2000))

        # ---------- Colors / styling ----------
        self.setStyleSheet('''
            QLabel { color: #12324A; }
            QLineEdit { background:#FFFFFF; border:1px solid #BBD3EA; border-radius:4px; padding:4px; }
            QLineEdit:focus { border-color:#639AD2; }
            QTextEdit { background:#FFFFFF; border:1px solid #BBD3EA; border-radius:4px; padding:6px; }
            QTextEdit:focus { border-color:#639AD2; }
            QPushButton { background-color:#3BAFDA; color:#ffffff; border:1px solid #2C9CC8; border-radius:5px; padding:6px 10px; }
            QPushButton:hover { background-color:#35A0C9; }
            QProgressBar { border:1px solid #BBD3EA; border-radius:3px; background:#ECF4FF; text-align:center; color:#12324A; }
            QProgressBar::chunk { background-color:#5CC0FF; }
            QTableWidget { gridline-color:#D5E3F6; background:#FFFFFF; alternate-background-color:#F4F8FD; }
            QHeaderView::section { background-color:#DCE8F7; color:#1F3B57; padding:4px; border:1px solid #BBD3EA; }
        ''')

        # Connect checkboxes to progress calculation
        self._table = table
        def recalc():
            checked = 0
            for r in range(self._table.rowCount()):
                w = self._table.cellWidget(r, 0)
                if isinstance(w, QCheckBox) and w.isChecked():
                    checked += 1
            self.progress_label.setText(f"Checklist Progress: {checked} / {self._table.rowCount()}")
            self.progress_bar.setValue(checked)
        for r in range(self._table.rowCount()):
            w = self._table.cellWidget(r, 0)
            if isinstance(w, QCheckBox):
                w.stateChanged.connect(recalc)

    def _limit_text(self, editor: QTextEdit, max_chars: int):
        # Enforce a hard character cap by trimming extra characters
        doc = editor.toPlainText()
        if len(doc) > max_chars:
            cursor = editor.textCursor()
            pos = cursor.position()
            editor.blockSignals(True)
            editor.setPlainText(doc[:max_chars])
            # restore cursor at the end if previous position is beyond cap
            cursor.setPosition(min(pos, max_chars))
            editor.setTextCursor(cursor)
            editor.blockSignals(False)

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

        tabs = QTabWidget()
        tabs.setDocumentMode(True)
        tabs.setMovable(True)

        # Light, attractive palette for tabs
        tabs.setStyleSheet('''
            QTabBar::tab { background: #EAF2FB; color: #1F3B57; padding: 8px 14px; border: 1px solid #D5E3F6; border-bottom: none; border-top-left-radius:6px; border-top-right-radius:6px; }
            QTabBar::tab:selected { background: #FFFFFF; color: #0F2D46; font-weight: 600; }
            QTabWidget::pane { border: 1px solid #D5E3F6; top: -1px; }
        ''')

        # Create tabs (README first)
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

    # Global pastel background
    pal = QPalette()
    pal.setColor(QPalette.ColorRole.Window, QColor(247, 250, 253))
    pal.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
    pal.setColor(QPalette.ColorRole.AlternateBase, QColor(241, 246, 252))
    pal.setColor(QPalette.ColorRole.Text, QColor(28, 41, 56))
    app.setPalette(pal)

    win = MainWindow()
    win.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    run()
