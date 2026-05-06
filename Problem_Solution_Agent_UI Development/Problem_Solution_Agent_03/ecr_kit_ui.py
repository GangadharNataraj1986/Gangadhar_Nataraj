
# ecr_kit_ui.py (Enhanced v3)
import sys
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QTextEdit, QTableWidget, QTableWidgetItem, QCheckBox,
    QPushButton, QProgressBar, QSizePolicy, QScrollArea
)
from PyQt6.QtGui import QPalette, QColor, QFont, QGuiApplication
from PyQt6.QtCore import Qt

APP_TITLE = "ECR Kit Assistant"

class ReadmeTab(QWidget):
    def __init__(self, readme_path: Path):
        super().__init__()
        layout = QVBoxLayout(self)
        title = QLabel("README")
        title.setFont(QFont("Segoe UI", 16, QFont.Weight.DemiBold))
        layout.addWidget(title)

        text = QTextEdit(); text.setReadOnly(True)
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
        top = QVBoxLayout(self)
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        top.addWidget(scroll)

        content = QWidget(); outer = QVBoxLayout(content)
        outer.setSpacing(10); outer.setContentsMargins(8,8,8,8)
        scroll.setWidget(content)

        # ---------- Header ----------
        row1 = QWidget(); r1 = QHBoxLayout(row1); r1.setContentsMargins(12,0,0,0)
        self.ecr_no = QLineEdit(); self.ecr_no.setPlaceholderText("ECR#")
        self.eco_primer = QLineEdit(); self.eco_primer.setPlaceholderText("ECO Primer Refs#")
        self.ec_category = QLineEdit(); self.ec_category.setPlaceholderText("EC Category")
        self.bu = QLineEdit(); self.bu.setPlaceholderText("BU")
        self.tco = QLineEdit(); self.tco.setPlaceholderText("TCO")
        self.project_no = QLineEdit(); self.project_no.setPlaceholderText("Project#")
        self.product = QLineEdit(); self.product.setPlaceholderText("Product")
        for w in [self.ecr_no, self.eco_primer, self.ec_category, self.bu, self.tco, self.project_no, self.product]:
            w.setFixedHeight(28); r1.addWidget(w)
        row2 = QWidget(); r2 = QHBoxLayout(row2); r2.setContentsMargins(12,0,0,0)
        self.affected_modules = QLineEdit(); self.affected_modules.setPlaceholderText("Affected Module(s)")
        self.place = QLineEdit(); self.place.setPlaceholderText("Place")
        for w in [self.affected_modules, self.place]:
            w.setFixedHeight(28); r2.addWidget(w)
        outer.addWidget(row1); outer.addWidget(row2)

        # ---------- Checklist label ----------
        lbl = QLabel("ECR Creation Checklist"); lbl.setFont(QFont("Segoe UI", 13, QFont.Weight.DemiBold))
        temp_lbl_row = QWidget(); tl = QHBoxLayout(temp_lbl_row); tl.setContentsMargins(12,0,0,0); tl.addWidget(lbl)
        outer.addWidget(temp_lbl_row)

        # ---------- Checklist table ----------
        table_row = QWidget(); tr = QHBoxLayout(table_row); tr.setContentsMargins(12,0,0,0)
        table = QTableWidget(len(self.CHECKLIST_ROWS), 7)
        table.setHorizontalHeaderLabels(["Sl.No", "", "Category", "Validation", "Comments", "Action Owner", "Due Date"])
        table.verticalHeader().setVisible(False)
        table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        table.setAlternatingRowColors(True)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        # Populate rows with Sl.No first, then checkbox
        for row, (cat, desc) in enumerate(self.CHECKLIST_ROWS):
            # Sl.No non-editable
            sl_item = QTableWidgetItem(str(row+1))
            sl_item.setFlags(sl_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(row, 0, sl_item)
            # Checkbox
            chk = QCheckBox(); table.setCellWidget(row, 1, chk)
            # Category (locked)
            cat_item = QTableWidgetItem(cat)
            cat_item.setFlags(cat_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(row, 2, cat_item)
            # Validation (locked)
            val_item = QTableWidgetItem(desc)
            val_item.setFlags(val_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            table.setItem(row, 3, val_item)
            # Editable text cells
            table.setItem(row, 4, QTableWidgetItem(""))
            table.setItem(row, 5, QTableWidgetItem(""))
            table.setItem(row, 6, QTableWidgetItem(""))

        try:
            from PyQt6.QtWidgets import QHeaderView
            header = table.horizontalHeader()
            header.setStretchLastSection(True)
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)  # Sl.No
            header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)  # Checkbox
            header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)  # Category
            header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)           # Validation
        except Exception:
            pass

        tr.addWidget(table); outer.addWidget(table_row)

        # ---------- Progress (slightly taller; ~75% width of the row) ----------
        progress_row = QWidget(); pr = QHBoxLayout(progress_row); pr.setContentsMargins(12,0,0,0)
        self.progress_label = QLabel(f"Checklist Progress: 0 / {len(self.CHECKLIST_ROWS)}")
        self.progress_bar = QProgressBar(); self.progress_bar.setRange(0, len(self.CHECKLIST_ROWS))
        self.progress_bar.setMaximumHeight(18)
        self.progress_bar.setTextVisible(True)
        pr.addWidget(self.progress_label)
        pr.addWidget(self.progress_bar, 3)
        pr.addStretch(1)  # leaves ~25% spacer => bar ~75% of remaining row
        outer.addWidget(progress_row)

        # ---------- Action row (button + title) ----------
        action_row = QWidget(); ar = QHBoxLayout(action_row); ar.setContentsMargins(12,0,0,0)
        self.btn_generate = QPushButton("Generate Problem Statement")
        self.title_edit = QLineEdit(); self.title_edit.setPlaceholderText("Title (max 75 chars)")
        self.title_edit.setMaxLength(75)
        ar.addWidget(self.btn_generate, 1)
        ar.addWidget(self.title_edit, 3)
        ar.addStretch(1)
        outer.addWidget(action_row)

        # ---------- Problem Statement ----------
        ps_label_row = QWidget(); psl = QHBoxLayout(ps_label_row); psl.setContentsMargins(12,0,0,0)
        ps_label = QLabel("Problem Statement"); ps_label.setFont(QFont("Segoe UI", 12, QFont.Weight.Medium))
        psl.addWidget(ps_label); outer.addWidget(ps_label_row)
        ps_row = QWidget(); psr = QHBoxLayout(ps_row); psr.setContentsMargins(12,0,0,0)
        self.problem_edit = QTextEdit(); self.problem_edit.setPlaceholderText("Write the problem statement here (max 2000 characters)…")
        self.problem_edit.setFixedHeight(140)
        psr.addWidget(self.problem_edit, 3); psr.addStretch(1)
        outer.addWidget(ps_row)
        self.problem_edit.textChanged.connect(lambda: self._limit_text(self.problem_edit, 2000))

        # ---------- Solution Statement ----------
        ss_label_row = QWidget(); ssl = QHBoxLayout(ss_label_row); ssl.setContentsMargins(12,0,0,0)
        ss_label = QLabel("Solution Statement"); ss_label.setFont(QFont("Segoe UI", 12, QFont.Weight.Medium))
        ssl.addWidget(ss_label); outer.addWidget(ss_label_row)
        ss_row = QWidget(); ssr = QHBoxLayout(ss_row); ssr.setContentsMargins(12,0,0,0)
        self.solution_edit = QTextEdit(); self.solution_edit.setPlaceholderText("Write the proposed solution here (max 2000 characters)…")
        self.solution_edit.setFixedHeight(140)
        ssr.addWidget(self.solution_edit, 3); ssr.addStretch(1)
        outer.addWidget(ss_row)
        self.solution_edit.textChanged.connect(lambda: self._limit_text(self.solution_edit, 2000))

        # ---------- Colors ----------
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

        # progress calc
        self._table = table
        def recalc():
            checked = 0
            for r in range(self._table.rowCount()):
                w = self._table.cellWidget(r, 1)
                if isinstance(w, QCheckBox) and w.isChecked():
                    checked += 1
            self.progress_label.setText(f"Checklist Progress: {checked} / {self._table.rowCount()}")
            self.progress_bar.setValue(checked)
        for r in range(self._table.rowCount()):
            w = self._table.cellWidget(r, 1)
            if isinstance(w, QCheckBox):
                w.stateChanged.connect(recalc)

    def _limit_text(self, editor: QTextEdit, max_chars: int):
        doc = editor.toPlainText()
        if len(doc) > max_chars:
            cursor = editor.textCursor(); pos = cursor.position()
            editor.blockSignals(True)
            editor.setPlainText(doc[:max_chars])
            cursor.setPosition(min(pos, max_chars))
            editor.setTextCursor(cursor)
            editor.blockSignals(False)

# ---------------- OBS Parts Tab ----------------
class OBSTable(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(10, 4, parent)  # start with 10 rows
        self.setHorizontalHeaderLabels(["Select", "OBS Parts", "Change", "Replacement"])
        self.verticalHeader().setVisible(False)
        self.setAlternatingRowColors(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        try:
            from PyQt6.QtWidgets import QHeaderView
            header = self.horizontalHeader()
            header.setStretchLastSection(True)
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
            header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        except Exception:
            pass
        self._init_rows(0, self.rowCount())
        self.apply_style()

    def apply_style(self):
        self.setStyleSheet('''
            QTableWidget { gridline-color:#D5E3F6; background:#FFFFFF; alternate-background-color:#F8FBFF; }
            QHeaderView::section { background-color:#F0F5FF; color:#243B53; padding:6px; border:1px solid #D5E3F6; font-weight:600; }
        ''')

    def _init_rows(self, start, end):
        yellow = QColor('#FFF59D')
        for r in range(start, end):
            # Checkbox
            chk = QCheckBox(); self.setCellWidget(r, 0, chk)
            # OBS Parts (editable)
            if not self.item(r,1):
                self.setItem(r, 1, QTableWidgetItem(""))
            # Change (default Obsolete, yellow, locked)
            change_item = QTableWidgetItem("Obsolete")
            change_item.setBackground(yellow)
            change_item.setFlags(change_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.setItem(r, 2, change_item)
            # Replacement (editable)
            if not self.item(r,3):
                self.setItem(r, 3, QTableWidgetItem(""))

    # Allow pasting multiple part numbers; auto-expand rows
    def keyPressEvent(self, event):
        if (event.matches(event.StandardKey.Paste)):
            text = QGuiApplication.clipboard().text()
            if text:
                lines = [ln for ln in text.splitlines() if ln.strip()]
                if lines:
                    # find first target row: current row or first empty
                    row = self.currentRow()
                    if row < 0:
                        row = self._first_empty_row()
                        if row < 0:
                            row = self.rowCount()
                    needed = row + len(lines) - self.rowCount()
                    if needed > 0:
                        old = self.rowCount()
                        self.setRowCount(self.rowCount()+needed)
                        self._init_rows(old, self.rowCount())
                    # fill rows (support pasting with tabs: part [	] change [	] replacement)
                    for i, ln in enumerate(lines):
                        parts = [p.strip() for p in ln.split('	')]
                        pr = parts[0] if len(parts) > 0 else ""
                        ch = parts[1] if len(parts) > 1 else None
                        rp = parts[2] if len(parts) > 2 else None
                        self.setItem(row+i, 1, QTableWidgetItem(pr))
                        if ch is not None and ch:
                            itm = QTableWidgetItem(ch); itm.setFlags(itm.flags() & ~Qt.ItemFlag.ItemIsEditable)
                            itm.setBackground(QColor('#FFF59D'))
                            self.setItem(row+i, 2, itm)
                        if rp is not None:
                            self.setItem(row+i, 3, QTableWidgetItem(rp))
                    return
        super().keyPressEvent(event)

    def _first_empty_row(self):
        for r in range(self.rowCount()):
            it = self.item(r, 1)
            if it is None or not it.text().strip():
                return r
        return -1

    def delete_selected_rows(self):
        rows_to_delete = []
        for r in range(self.rowCount()):
            w = self.cellWidget(r, 0)
            if isinstance(w, QCheckBox) and w.isChecked():
                rows_to_delete.append(r)
        for r in reversed(rows_to_delete):
            self.removeRow(r)
        # ensure there is at least one empty row
        if self.rowCount() == 0:
            self.setRowCount(1)
            self._init_rows(0,1)

class OBSPartsTab(QWidget):
    def __init__(self):
        super().__init__()
        outer = QVBoxLayout(self)
        title_row = QHBoxLayout()
        title = QLabel("Final OBS List")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.DemiBold))
        delete_btn = QPushButton("Delete Selected")
        title_row.addWidget(title)
        title_row.addStretch(1)
        title_row.addWidget(delete_btn)
        outer.addLayout(title_row)

        self.table = OBSTable(self)
        outer.addWidget(self.table)

        delete_btn.clicked.connect(self.table.delete_selected_rows)

        # Colors for tab controls
        self.setStyleSheet('''
            QLabel { color:#0F2D46; }
            QPushButton { background:#E25563; color:#fff; border:1px solid #C94855; border-radius:5px; padding:6px 10px; }
            QPushButton:hover { background:#D04A58; }
        ''')

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
        tabs.setStyleSheet('''
            QTabBar::tab { background: #EAF2FB; color: #1F3B57; padding: 8px 14px; border: 1px solid #D5E3F6; border-bottom: none; border-top-left-radius:6px; border-top-right-radius:6px; }
            QTabBar::tab:selected { background: #FFFFFF; color: #0F2D46; font-weight: 600; }
            QTabWidget::pane { border: 1px solid #D5E3F6; top: -1px; }
        ''')

        readme_path = Path(__file__).with_name('README.txt')
        tabs.addTab(ReadmeTab(readme_path), "README")
        tabs.addTab(ECRFrontPageTab(), "ECR Front Page")
        tabs.addTab(OBSPartsTab(), "OBS Parts")
        tabs.addTab(PlaceholderTab("Where Used of OBS Parts"), "Where Used")
        tabs.addTab(PlaceholderTab("Orphan Analysis"), "Orphan Analysis")
        tabs.addTab(PlaceholderTab("Structure sheet"), "Structure sheet")
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
