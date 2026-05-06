# ecr_kit_ui.py (Enhanced v6 per latest Front Page requirements)
import sys
import json
from pathlib import Path
from typing import Any, Dict, List

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QTextEdit, QTableWidget, QTableWidgetItem, QCheckBox,
    QPushButton, QProgressBar, QSizePolicy, QScrollArea, QToolBar,
    QMessageBox, QComboBox, QFileDialog
)
from PyQt6.QtGui import QPalette, QColor, QFont, QGuiApplication, QAction
from PyQt6.QtCore import Qt

APP_TITLE = "ECR Kit Assistant"
DATA_FILE = Path(__file__).with_name('ecr_kit_data.json')
TEMPLATE_FILE = Path(__file__).with_name('obs_parts_template.xlsx')

import traceback

def _handle_exception(exc_type, exc_value, exc_tb):
    try:
        log_path = Path(__file__).with_name('error_log.txt')
        msg = ''.join(traceback.format_exception(exc_type, exc_value, exc_tb))
        log_path.write_text(msg, encoding='utf-8')
        m = QMessageBox()
        m.setIcon(QMessageBox.Icon.Critical)
        m.setWindowTitle('Unexpected Error')
        m.setText('An unexpected error occurred. A log was written to error_log.txt')
        m.setDetailedText(msg)
        m.setStandardButtons(QMessageBox.StandardButton.Ok)
        m.show()
        app = QApplication.instance()
        if app is not None:
            if not hasattr(app, '_error_dialogs'):
                app._error_dialogs = []
            app._error_dialogs.append(m)
    except Exception:
        print('Unhandled exception:', file=sys.stderr)
        traceback.print_exception(exc_type, exc_value, exc_tb)

sys.excepthook = _handle_exception


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
    # EXACTLY 10 categories and 16 questions in total
    # Each tuple: (Category, [list of question strings])
    CHECKLIST_DEF = [
        ("Project Association", [
            "Does the ECR include a Project (PCR)?"
        ]),
        ("Safety & Compliance", [
            "Provide PSER details if any safety incident occurred.",
            "Is a PCN required and approved by ROW and CE?"
        ]),
        ("Part Release Status", [
            "Are all BTPs and kits EVAL released?",
            "If the parent part is in production, is the replacement part (BTP) production released?"
        ]),
        ("Design Analysis", [
            "Is VA/VB analysis completed (for new designs/parts only)?",
            "PACE / DASH parts addressed?"
        ]),
        ("Watchlist & Spares", [
            "Are new parts/designs MLO certified / were parent/previous parts MLO certified?",
            "Do we have AGS approval for OBS parts (sparable) without replacement?"
        ]),
        ("Config Documents", [
            "Is CCR available with change matrix details for new options/reference designator updates?"
        ]),
        ("ECR Strategies Identified", [
            "Provide reason code, effective strategy, priority, and disposition."
        ]),
        ("Multi BU Alignment", [
            "If scope impacts multiple BUs/products, verify and confirm all affected BUs/products are listed in the project and approved."
        ]),
        ("Interchangeability & Tags", [
            "Are the changes FFF compatible (per 03-3-10 Interchangeability Policy)? Provide system details.",
            "Provide interchangeability/tagging details as per CRP."
        ]),
        ("Testing, Reports & Cost", [
            "Are test results/FDR available for all IFF impacted parts/critical parts?",
            "If the project relates to DCR (cost reduction), provide cost-saving details."
        ])
    ]
    # Count check (10 categories, 16 questions)
    _QUESTIONS_TOTAL = sum(len(qs) for _, qs in CHECKLIST_DEF)
    assert len(CHECKLIST_DEF) == 10, "Checklist must contain exactly 10 categories"
    assert _QUESTIONS_TOTAL == 16, "Checklist must contain exactly 16 questions"

    def __init__(self):
        super().__init__()
        top = QVBoxLayout(self)

        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        top.addWidget(scroll)

        content = QWidget(); outer = QVBoxLayout(content)
        outer.setSpacing(10); outer.setContentsMargins(8, 8, 8, 8)
        scroll.setWidget(content)

        # ---------- Header ----------
        row1 = QWidget(); r1 = QHBoxLayout(row1); r1.setContentsMargins(12, 0, 0, 0)
        self.ecr_no = QLineEdit(); self.ecr_no.setPlaceholderText("ECR#")
        self.eco_primer = QLineEdit(); self.eco_primer.setPlaceholderText("ECO Primer Refs#")
        self.ec_category = QLineEdit(); self.ec_category.setPlaceholderText("EC Category")
        self.bu = QLineEdit(); self.bu.setPlaceholderText("BU")
        self.tco = QLineEdit(); self.tco.setPlaceholderText("TCO")
        self.project_no = QLineEdit(); self.project_no.setPlaceholderText("Project#")
        self.product = QLineEdit(); self.product.setPlaceholderText("Product")
        for w in [self.ecr_no, self.eco_primer, self.ec_category, self.bu, self.tco, self.project_no, self.product]:
            w.setFixedHeight(28); r1.addWidget(w)

        row2 = QWidget(); r2 = QHBoxLayout(row2); r2.setContentsMargins(12, 0, 0, 0)
        self.affected_modules = QLineEdit(); self.affected_modules.setPlaceholderText("Affected Module(s)")
        self.place = QLineEdit(); self.place.setPlaceholderText("Place")
        for w in [self.affected_modules, self.place]:
            w.setFixedHeight(28); r2.addWidget(w)
        outer.addWidget(row1); outer.addWidget(row2)

        # ---------- Checklist label ----------
        lbl = QLabel("ECR Creation Checklist"); lbl.setFont(QFont("Segoe UI", 13, QFont.Weight.DemiBold))
        temp_lbl_row = QWidget(); tl = QHBoxLayout(temp_lbl_row); tl.setContentsMargins(12, 0, 0, 0); tl.addWidget(lbl)
        outer.addWidget(temp_lbl_row)

        # ---------- Checklist table ----------
        table_row = QWidget(); tr = QHBoxLayout(table_row); tr.setContentsMargins(12, 0, 0, 0)
        # Column order: Sl.No | Completion | Category | Validation | Actioner Owner | Due Date | Comments
        self.table = QTableWidget(len(self.CHECKLIST_DEF), 7)
        self.table.setHorizontalHeaderLabels(["Sl.No", "Completion", "Category", "Validation", "Actioner Owner", "Due Date", "Comments"])
        self.table.verticalHeader().setVisible(False)
        self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.table.setWordWrap(True)

        # Remove cell coloring (use plain backgrounds)
        self.table.setStyleSheet("QTableWidget { background:#FFFFFF; } QHeaderView::section { background:#FFFFFF; color:#1F3B57; }")

        # Populate rows
        for row, (cat, qlist) in enumerate(self.CHECKLIST_DEF):
            # Sl.No non-editable
            sl_item = QTableWidgetItem(str(row + 1))
            sl_item.setFlags(sl_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 0, sl_item)
            # Checkbox
            chk = QCheckBox(); self.table.setCellWidget(row, 1, chk)
            # Category (locked)
            cat_item = QTableWidgetItem(cat)
            cat_item.setFlags(cat_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 2, cat_item)
            # Validation as plain, multi-line text (no colors)
            text = "\n".join(qlist)
            lab = QLabel(text)
            lab.setTextFormat(Qt.TextFormat.PlainText)
            lab.setWordWrap(True)
            self.table.setCellWidget(row, 3, lab)
            # Editable cells for user typing
            self.table.setItem(row, 4, QTableWidgetItem(""))
            self.table.setItem(row, 5, QTableWidgetItem(""))
            self.table.setItem(row, 6, QTableWidgetItem(""))

        try:
            from PyQt6.QtWidgets import QHeaderView
            header = self.table.horizontalHeader()
            header.setStretchLastSection(False)
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
            header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
            header.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)
            self.table.resizeRowsToContents()
        except Exception:
            pass

        tr.addWidget(self.table); outer.addWidget(table_row)

        # ---------- Progress row ----------
        progress_row = QWidget(); pr = QHBoxLayout(progress_row); pr.setContentsMargins(12, 0, 0, 0)
        self.progress_label = QLabel(f"Checklist Progress: 0 / {len(self.CHECKLIST_DEF)}")
        self.progress_bar = QProgressBar(); self.progress_bar.setRange(0, len(self.CHECKLIST_DEF))
        self.progress_bar.setMaximumHeight(18)
        self.progress_bar.setTextVisible(True)
        pr.addWidget(self.progress_label)
        pr.addWidget(self.progress_bar, 3)
        pr.addStretch(1)
        outer.addWidget(progress_row)

        # ---------- Action row ----------
        action_row = QWidget(); ar = QHBoxLayout(action_row); ar.setContentsMargins(12, 0, 0, 0)
        self.btn_generate = QPushButton("Generate Problem Statement")
        self.title_edit = QLineEdit(); self.title_edit.setPlaceholderText("Title (max 75 chars)")
        self.title_edit.setMaxLength(75)
        ar.addWidget(self.btn_generate, 1)
        ar.addWidget(self.title_edit, 3)
        ar.addStretch(1)
        outer.addWidget(action_row)

        # ---------- Problem/Solution ----------
        ps_row = QWidget(); psr = QHBoxLayout(ps_row); psr.setContentsMargins(12, 0, 0, 0)
        self.problem_edit = QTextEdit(); self.problem_edit.setPlaceholderText("Write the problem statement here (max 2000 characters)…")
        self.problem_edit.setFixedHeight(140)
        psr.addWidget(self.problem_edit, 3); psr.addStretch(1)
        outer.addWidget(ps_row)
        self.problem_edit.textChanged.connect(lambda: self._limit_text(self.problem_edit, 2000))

        ss_row = QWidget(); ssr = QHBoxLayout(ss_row); ssr.setContentsMargins(12, 0, 0, 0)
        self.solution_edit = QTextEdit(); self.solution_edit.setPlaceholderText("Write the proposed solution here (max 2000 characters)…")
        self.solution_edit.setFixedHeight(140)
        ssr.addWidget(self.solution_edit, 3); ssr.addStretch(1)
        outer.addWidget(ss_row)
        self.solution_edit.textChanged.connect(lambda: self._limit_text(self.solution_edit, 2000))

        # Colors (keep neutral, no cell coloring on table)
        self.setStyleSheet("""
            QLabel { color: #12324A; }
            QLineEdit, QTextEdit { background:#FFFFFF; border:1px solid #BBD3EA; border-radius:4px; padding:4px; }
            QLineEdit:focus, QTextEdit:focus { border-color:#639AD2; }
            QPushButton { background-color:#3BAFDA; color:#ffffff; border:1px solid #2C9CC8; border-radius:5px; padding:6px 10px; }
            QPushButton:hover { background-color:#35A0C9; }
            QProgressBar { border:1px solid #BBD3EA; border-radius:3px; background:#ECF4FF; text-align:center; color:#12324A; }
            QProgressBar::chunk { background-color:#5CC0FF; }
        """)

        # progress calc
        def recalc():
            checked = 0
            for r in range(self.table.rowCount()):
                w = self.table.cellWidget(r, 1)
                if isinstance(w, QCheckBox) and w.isChecked():
                    checked += 1
            self.progress_label.setText(f"Checklist Progress: {checked} / {self.table.rowCount()}")
            self.progress_bar.setValue(checked)
        for r in range(self.table.rowCount()):
            w = self.table.cellWidget(r, 1)
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

    def to_dict(self) -> Dict[str, Any]:
        data = {
            'fields': {
                'ecr_no': self.ecr_no.text(),
                'eco_primer': self.eco_primer.text(),
                'ec_category': self.ec_category.text(),
                'bu': self.bu.text(),
                'tco': self.tco.text(),
                'project_no': self.project_no.text(),
                'product': self.product.text(),
                'affected_modules': self.affected_modules.text(),
                'place': self.place.text(),
                'title': self.title_edit.text(),
                'problem': self.problem_edit.toPlainText(),
                'solution': self.solution_edit.toPlainText(),
            },
            'checklist': []
        }
        t = self.table
        for r in range(t.rowCount()):
            completion = False
            w = t.cellWidget(r, 1)
            if isinstance(w, QCheckBox):
                completion = w.isChecked()
            row = {
                'slno': t.item(r, 0).text() if t.item(r, 0) else str(r+1),
                'completion': completion,
                'category': t.item(r, 2).text() if t.item(r, 2) else '',
                'validation_text': t.cellWidget(r, 3).text() if isinstance(t.cellWidget(r,3), QLabel) else '',
                'actioner_owner': t.item(r, 4).text() if t.item(r, 4) else '',
                'due_date': t.item(r, 5).text() if t.item(r, 5) else '',
                'comments': t.item(r, 6).text() if t.item(r, 6) else '',
            }
            data['checklist'].append(row)
        return data

    def from_dict(self, data: Dict[str, Any]):
        f = data.get('fields', {})
        self.ecr_no.setText(f.get('ecr_no', ''))
        self.eco_primer.setText(f.get('eco_primer', ''))
        self.ec_category.setText(f.get('ec_category', ''))
        self.bu.setText(f.get('bu', ''))
        self.tco.setText(f.get('tco', ''))
        self.project_no.setText(f.get('project_no', ''))
        self.product.setText(f.get('product', ''))
        self.affected_modules.setText(f.get('affected_modules', ''))
        self.place.setText(f.get('place', ''))
        self.title_edit.setText(f.get('title', ''))
        self.problem_edit.setPlainText(f.get('problem', ''))
        self.solution_edit.setPlainText(f.get('solution', ''))

        checklist = data.get('checklist', [])
        t = self.table
        rows = min(len(checklist), t.rowCount())
        for r in range(rows):
            row = checklist[r]
            w = t.cellWidget(r, 1)
            if isinstance(w, QCheckBox):
                w.setChecked(bool(row.get('completion', False)))
            def _set(col, text):
                it = t.item(r, col)
                if it is None:
                    it = QTableWidgetItem("")
                    t.setItem(r, col, it)
                it.setText(text)
            _set(4, row.get('actioner_owner', ''))
            _set(5, row.get('due_date', ''))
            _set(6, row.get('comments', ''))
        try:
            t.resizeRowsToContents()
        except Exception:
            pass


# ---------------- OBS Parts Tab (unchanged from v5) ----------------
class OBSTable(QTableWidget):
    def __init__(self, parent=None, initial_rows: int = 10):
        super().__init__(initial_rows, 4, parent)
        self.setHorizontalHeaderLabels(["Select", "OBS Parts", "Change", "Replacement"])
        self.verticalHeader().setVisible(False)
        try:
            from PyQt6.QtWidgets import QHeaderView
            header = self.horizontalHeader()
            header.setStretchLastSection(False)
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
            header.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
            header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        except Exception:
            pass
        self._init_rows(0, self.rowCount())
        self._apply_excel_widths()

    def _apply_excel_widths(self, chars: int = 14):
        fm = self.fontMetrics()
        px = fm.horizontalAdvance('0' * chars) + 16
        for col in (1, 2, 3):
            try:
                self.setColumnWidth(col, px)
            except Exception:
                pass

    def _make_change_combo(self) -> QComboBox:
        combo = QComboBox()
        combo.addItems(["Obsolete", "Inactivate"])
        combo.setCurrentText("Obsolete")
        return combo

    def _init_rows(self, start, end):
        for r in range(start, end):
            chk = QCheckBox(); self.setCellWidget(r, 0, chk)
            if not self.item(r, 1):
                self.setItem(r, 1, QTableWidgetItem(""))
            combo = self._make_change_combo()
            self.setCellWidget(r, 2, combo)
            if not self.item(r, 3):
                self.setItem(r, 3, QTableWidgetItem(""))

    def keyPressEvent(self, event):
        try:
            if event.matches(event.StandardKey.Copy):
                self._copy_selection_to_clipboard(); return
            if event.matches(event.StandardKey.Paste):
                self._paste_from_clipboard(); return
        except Exception:
            traceback.print_exc()
        super().keyPressEvent(event)

    def _copy_selection_to_clipboard(self):
        sel = self.selectedRanges()
        if not sel:
            return
        r = sel[0]
        rows = []
        for i in range(r.rowCount()):
            cols = []
            for j in range(r.columnCount()):
                row_i = r.topRow() + i
                col_j = r.leftColumn() + j
                if col_j == 2:
                    w = self.cellWidget(row_i, 2)
                    txt = w.currentText() if isinstance(w, QComboBox) else ''
                else:
                    it = self.item(row_i, col_j)
                    txt = it.text() if it else ''
                cols.append(txt)
            rows.append('\t'.join(cols))
        QGuiApplication.clipboard().setText('\n'.join(rows))

    def _ensure_rows(self, upto_row_inclusive: int):
        if upto_row_inclusive >= self.rowCount():
            old = self.rowCount()
            self.setRowCount(upto_row_inclusive + 1)
            self._init_rows(old, self.rowCount())

    def _paste_from_clipboard(self):
        text = QGuiApplication.clipboard().text()
        if not text:
            return
        start_row = self.currentRow()
        start_col = self.currentColumn()
        if start_row < 0:
            start_row = self._first_empty_row()
        if start_row < 0:
            start_row = self.rowCount()
        lines = [ln for ln in text.splitlines() if ln.strip()]
        for r_offset, line in enumerate(lines):
            parts = [p.strip() for p in line.split('\t')]
            row = start_row + r_offset
            self._ensure_rows(row)
            for c_offset, val in enumerate(parts):
                col = start_col + c_offset
                if col == 0:
                    continue
                if col == 2:
                    w = self.cellWidget(row, 2)
                    if isinstance(w, QComboBox):
                        idx = w.findText(val)
                        if idx >= 0:
                            w.setCurrentIndex(idx)
                else:
                    self.setItem(row, col, QTableWidgetItem(val))
        self._apply_excel_widths()

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
        if self.rowCount() == 0:
            self.setRowCount(10)
            self._init_rows(0, 10)


class OBSPartsTab(QWidget):
    def __init__(self):
        super().__init__()
        outer = QVBoxLayout(self)

        title_row = QHBoxLayout()
        title = QLabel("Final OBS List")
        title.setFont(QFont("Segoe UI", 14, QFont.Weight.DemiBold))
        title.setStyleSheet("color:#C1272D;")
        btn_template = QPushButton("Download Template (.xlsx)")
        btn_upload = QPushButton("Upload from Excel")
        delete_btn = QPushButton("Delete Selected")

        title_row.addWidget(title)
        title_row.addStretch(1)
        title_row.addWidget(btn_template)
        title_row.addWidget(btn_upload)
        title_row.addWidget(delete_btn)
        outer.addLayout(title_row)

        self.table = OBSTable(self, initial_rows=10)
        outer.addWidget(self.table)

        delete_btn.clicked.connect(self.table.delete_selected_rows)
        btn_template.clicked.connect(self.download_template)
        btn_upload.clicked.connect(self.upload_from_excel)

    def download_template(self):
        try:
            path, _ = QFileDialog.getSaveFileName(self, 'Save Template', str(TEMPLATE_FILE), 'Excel Files (*.xlsx)')
            if not path:
                return
            import pandas as pd
            df = pd.DataFrame({'OBS Parts': [''], 'Change': ['Obsolete'], 'Replacement': ['']})
            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='OBS_Template')
            QMessageBox.information(self, 'Template Saved', f'Template saved to:\n{path}')
        except Exception as e:
            QMessageBox.warning(self, 'Template Error', str(e))

    def upload_from_excel(self):
        try:
            path, _ = QFileDialog.getOpenFileName(self, 'Open OBS Parts Excel', '', 'Excel Files (*.xlsx *.xls)')
            if not path:
                return
            import pandas as pd
            if path.lower().endswith('.xls'):
                df = pd.read_excel(path, engine='xlrd')
            else:
                df = pd.read_excel(path, engine='openpyxl')
            cols = {c.strip().lower(): c for c in df.columns}
            def pick(name):
                for key in cols:
                    if key == name:
                        return cols[key]
                return None
            c_obs = pick('obs parts') or pick('obs part') or pick('part')
            c_change = pick('change')
            c_rep = pick('replacement') or pick('replace') or pick('new part')
            if not c_obs:
                raise ValueError('Column "OBS Parts" is required in the Excel file.')
            rows = []
            for _, r in df.iterrows():
                obs = str(r.get(c_obs, '')).strip()
                if not obs:
                    continue
                change_val = str(r.get(c_change, 'Obsolete')).strip() if c_change else 'Obsolete'
                if change_val not in ['Obsolete', 'Inactivate']:
                    change_val = 'Obsolete'
                rep = str(r.get(c_rep, '')).strip() if c_rep else ''
                rows.append((obs, change_val, rep))
            if not rows:
                QMessageBox.information(self, 'No Data', 'No valid rows found in the Excel file.'); return
            t = self.table
            t.setRowCount(len(rows))
            t._init_rows(0, len(rows))
            for r, (obs, change, rep) in enumerate(rows):
                t.setItem(r, 1, QTableWidgetItem(obs))
                w = t.cellWidget(r, 2)
                if isinstance(w, QComboBox):
                    idx = w.findText(change)
                    w.setCurrentIndex(idx if idx >= 0 else 0)
                t.setItem(r, 3, QTableWidgetItem(rep))
            t._apply_excel_widths()
            QMessageBox.information(self, 'Upload Complete', f'Loaded {len(rows)} rows from Excel.')
        except Exception as e:
            QMessageBox.warning(self, 'Upload Error', str(e))

    def to_dict(self) -> Dict[str, Any]:
        t = self.table
        rows: List[Dict[str, Any]] = []
        for r in range(t.rowCount()):
            obs = t.item(r, 1).text() if t.item(r, 1) else ''
            rep = t.item(r, 3).text() if t.item(r, 3) else ''
            w = t.cellWidget(r, 2)
            change = w.currentText() if isinstance(w, QComboBox) else 'Obsolete'
            if any([obs.strip(), rep.strip()]):
                rows.append({'obs_part': obs, 'change': change, 'replacement': rep})
        return {'rows': rows}

    def from_dict(self, data: Dict[str, Any]):
        rows = data.get('rows', [])
        t = self.table
        needed = max(10, len(rows))
        t.setRowCount(needed)
        t._init_rows(0, needed)
        for r, row in enumerate(rows):
            t.setItem(r, 1, QTableWidgetItem(row.get('obs_part', '')))
            w = t.cellWidget(r, 2)
            if isinstance(w, QComboBox):
                idx = w.findText(row.get('change', 'Obsolete'))
                if idx >= 0:
                    w.setCurrentIndex(idx)
            t.setItem(r, 3, QTableWidgetItem(row.get('replacement', '')))
        t._apply_excel_widths()


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

        self.tabs = QTabWidget(); self.tabs.setDocumentMode(True); self.tabs.setMovable(True)
        self.tabs.setStyleSheet("""
            QTabBar::tab { background: #EAF2FB; color: #1F3B57; padding: 8px 14px; border: 1px solid #D5E3F6; border-bottom: none; border-top-left-radius:6px; border-top-right-radius:6px; }
            QTabBar::tab:selected { background: #FFFFFF; color: #0F2D46; font-weight: 600; }
            QTabWidget::pane { border: 1px solid #D5E3F6; top: -1px; }
        """)

        readme_path = Path(__file__).with_name('README.txt')
        self.readme_tab = ReadmeTab(readme_path)
        self.front_tab = ECRFrontPageTab()
        self.obs_tab = OBSPartsTab()

        self.tabs.addTab(self.readme_tab, "README")
        self.tabs.addTab(self.front_tab, "ECR Front Page")
        self.tabs.addTab(self.obs_tab, "OBS Parts")
        self.tabs.addTab(PlaceholderTab("Where Used of OBS Parts"), "Where Used")
        self.tabs.addTab(PlaceholderTab("Orphan Analysis"), "Orphan Analysis")
        self.tabs.addTab(PlaceholderTab("Structure sheet"), "Structure sheet")
        self.tabs.addTab(PlaceholderTab("Inventory Cost Analysis"), "Inventory & Cost")
        self.tabs.addTab(PlaceholderTab("Report"), "Report")
        self.tabs.addTab(PlaceholderTab("User Notes"), "User Notes")

        self.setCentralWidget(self.tabs)

        tb = QToolBar("File"); self.addToolBar(tb)
        act_save = QAction("Save", self); act_save.triggered.connect(self.save_data); tb.addAction(act_save)
        act_load = QAction("Open", self); act_load.triggered.connect(self.load_data_dialog); tb.addAction(act_load)

        self.load_data_if_exists()

    def aggregate_data(self) -> Dict[str, Any]:
        return {
            'front_page': self.front_tab.to_dict(),
            'obs_parts': self.obs_tab.to_dict(),
        }

    def apply_data(self, data: Dict[str, Any]):
        if not data:
            return
        if 'front_page' in data:
            self.front_tab.from_dict(data['front_page'])
        if 'obs_parts' in data:
            self.obs_tab.from_dict(data['obs_parts'])

    def save_data(self):
        data = self.aggregate_data()
        try:
            DATA_FILE.write_text(json.dumps(data, indent=2), encoding='utf-8')
            QMessageBox.information(self, "Saved", f"Data saved to {DATA_FILE.name}")
        except Exception as e:
            QMessageBox.warning(self, "Save Failed", str(e))

    def load_data_if_exists(self):
        try:
            if DATA_FILE.exists():
                data = json.loads(DATA_FILE.read_text(encoding='utf-8'))
                self.apply_data(data)
        except Exception as e:
            QMessageBox.warning(self, "Load Failed", str(e))

    def load_data_dialog(self):
        if DATA_FILE.exists():
            try:
                data = json.loads(DATA_FILE.read_text(encoding='utf-8'))
                self.apply_data(data)
                QMessageBox.information(self, "Loaded", f"Data loaded from {DATA_FILE.name}")
            except Exception as e:
                QMessageBox.warning(self, "Load Failed", str(e))
        else:
            QMessageBox.information(self, "No Save Found", "No saved data file found yet. Click Save to create one.")


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
