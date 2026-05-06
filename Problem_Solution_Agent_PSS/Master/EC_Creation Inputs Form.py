from PyQt6.QtWidgets import QRadioButton
import sys
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget,
    QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QFrame,
    QButtonGroup, QTextEdit, QScrollArea, QPushButton, QFileDialog
)
from PyQt6.QtGui import QFont

APP_TITLE = "ECR Kit Assistant"

EC_CATEGORY_DESC = {
    "A1": "SMBoM Options as revised items and having CDW",
    "A2": "SMBoM Options as revised items, and No CDW",
    "B1": "No SMBoM Options as revised items and having CDW",
    "B2": "No SMBoM Options as revised items, No CDW, and revised item status at Eval and/or moving to Eval or adding already released parts to Proto buckets",
    "B3": "No SMBoM Option as revised items, No CDW, and revised Item status at Production and/or moving to Production",
}


def header(text):
    lbl = QLabel(text)
    lbl.setFont(QFont("Segoe UI", 12, QFont.Weight.DemiBold))
    return lbl


def highlight_label(text):
    lbl = QLabel(text)
    lbl.setFont(QFont("Segoe UI", 11, QFont.Weight.DemiBold))
    lbl.setStyleSheet("color:#1F4E79")
    return lbl


class LimitedTextEdit(QTextEdit):
    def __init__(self, limit=2000):
        super().__init__()
        self.limit = limit
        self.textChanged.connect(self._limit)

    def _limit(self):
        text = self.toPlainText()
        if len(text) > self.limit:
            self.blockSignals(True)
            self.setPlainText(text[:self.limit])
            self.moveCursor(self.textCursor().End)
            self.blockSignals(False)

    def _update_counter(self):
        if not self.counter_label:
            return
        count = len(self.toPlainText())
        self.counter_label.setText(f"{count} / {self.limit} characters")
        if count >= self.limit:
            self.counter_label.setStyleSheet("color:red;font-size:10px")
        else:
            self.counter_label.setStyleSheet("color:gray;font-size:10px")

class ECCreationInputsFormTab(QWidget):
    def __init__(self):
        super().__init__()
        root = QVBoxLayout(self)

        scroll = QScrollArea(); scroll.setWidgetResizable(True)
        root.addWidget(scroll)
        container = QWidget(); scroll.setWidget(container)
        outer = QVBoxLayout(container)

        # ---- Section A ----
        outer.addWidget(header("Section A: EC Category Form"))
        secA = QFrame(); a = QVBoxLayout(secA)

        scope_row = QHBoxLayout(); scope_grp = QButtonGroup(self)
        for scope in ["Up Revision", "Status Roll", "Production Release", "OBS / Inactivate", "Product Release"]:
            rb = QRadioButton(scope)
            scope_grp.addButton(rb); rb.toggled.connect(self.start_flow)
            scope_row.addWidget(rb)
        a.addLayout(scope_row)

        self.flow_area = QVBoxLayout(); a.addLayout(self.flow_area)
        self.ec_result_lbl = QLabel(""); self.ec_result_lbl.setWordWrap(True)
        self.ec_result_lbl.setStyleSheet("color:green;font-weight:600;font-size:14px")
        a.addWidget(self.ec_result_lbl)

        self.ec_divider = QFrame(); self.ec_divider.setFrameShape(QFrame.Shape.HLine)
        self.ec_divider.setVisible(False); a.addWidget(self.ec_divider)
        outer.addWidget(secA)

        outer.addSpacing(12)

        # ---- Section B ----
        self.secB_header = header("Section B: ECR Change Assessment")
        self.secB_header.setVisible(False); outer.addWidget(self.secB_header)
        self.secB = QFrame(); self.secB.setVisible(False)
        b = QVBoxLayout(self.secB)

        # Reference fields
        self.ref_boxes = {}
        def yes_no_with_box(key, label):
            row = QHBoxLayout(); row.addWidget(QLabel(label))
            grp = QButtonGroup(self)
            rb_y = QRadioButton("Yes"); rb_n = QRadioButton("No")
            grp.addButton(rb_y); grp.addButton(rb_n)
            row.addWidget(rb_y); row.addWidget(rb_n)
            txt = QTextEdit(); txt.setFixedHeight(30); txt.setMaximumWidth(500)
            txt.setStyleSheet("background:#FFF2CC; border:none"); txt.setVisible(False)
            row.addWidget(txt); row.addStretch(1)
            rb_y.toggled.connect(lambda c: txt.setVisible(c))
            rb_n.toggled.connect(lambda c: txt.setVisible(False))
            self.ref_boxes[key] = txt
            b.addLayout(row)

        yes_no_with_box("PCR_PCN", "1.  Does this Project include Product Change Request (PCR)")
        yes_no_with_box("PSN", "2.  Is there any Associated Product Safety Note (PSN)")
        yes_no_with_box("SPS", "3.  Associated Open SPSs")
        yes_no_with_box("ESW", "4.  Associated ESWs")
        yes_no_with_box("REF_ECR", "5.  Reference ECR Number(s)")

        # Reference Email Attachments
        rowm = QHBoxLayout(); rowm.addWidget(QLabel("6. Reference e-mail attachments?"))
        grp_m = QButtonGroup(self)
        rb_my = QRadioButton("Yes"); rb_mn = QRadioButton("No")
        grp_m.addButton(rb_my); grp_m.addButton(rb_mn)
        rowm.addWidget(rb_my); rowm.addWidget(rb_mn)
        browse = QPushButton("Browse E-mail"); browse.setFixedHeight(24); browse.setVisible(False)
        rowm.addWidget(browse); rowm.addStretch(1)
        b.addLayout(rowm)
        rb_my.toggled.connect(lambda c: browse.setVisible(c))
        browse.clicked.connect(lambda: QFileDialog.getOpenFileName(self, "Select Email", "", "Email Files (*.msg *.eml)"))

        # Impact caused by
        row = QHBoxLayout(); row.addWidget(QLabel("7.  Impact caused by:"))
        grp_sc = QButtonGroup(self)
        rb_sup = QRadioButton("Supplier"); rb_cust = QRadioButton("Customer")
        grp_sc.addButton(rb_sup); grp_sc.addButton(rb_cust)
        row.addWidget(rb_sup); row.addWidget(rb_cust); row.addStretch(1)
        b.addLayout(row)

        sub_row = QHBoxLayout(); sub_row.setContentsMargins(20, 0, 0, 0)
        grp_c = QButtonGroup(self)
        rb_int = QRadioButton("Internal"); rb_ext = QRadioButton("External")
        grp_c.addButton(rb_int); grp_c.addButton(rb_ext)
        sub_row.addWidget(rb_int); sub_row.addWidget(rb_ext); sub_row.addStretch(1)
        b.addLayout(sub_row)

        rb_sup.toggled.connect(lambda c: (rb_int.setEnabled(False), rb_ext.setEnabled(False), rb_int.setChecked(False), rb_ext.setChecked(False)))
        rb_cust.toggled.connect(lambda c: (rb_int.setEnabled(c), rb_ext.setEnabled(c)))

        # Reason Code
        rc_row = QHBoxLayout(); rc_row.addWidget(QLabel("8.  ECR Reason Code:"))
        self.reason_cb = QComboBox(); self.reason_cb.addItems([
            "Beyond Spec Request","Cap Code Management","CES","Cost Reduction",
            "Design Correction","Document Correction","Manufacturing Improvement",
            "Obsolescence","Option Reduction and Product End of Life","Order BOM Change",
            "Product Improvement","Product Release","Safety Event"])
        rc_row.addWidget(self.reason_cb); rc_row.addStretch(1)
        b.addLayout(rc_row)

        # Problem Summary
        ps_row = QHBoxLayout()
        ps_row.addWidget(highlight_label("AI Assisted Problem Summary from PCR, PCN, SPS and ESW"))
        btn_ps = QPushButton("Generate Problem Summary"); btn_ps.setFixedSize(180,26)
        ps_row.addWidget(btn_ps); ps_row.addStretch(1)
        b.addLayout(ps_row)

        self.problem_txt = LimitedTextEdit(2000); self.problem_txt.setMinimumHeight(140); self.problem_txt.setMaximumWidth(900)
        b.addWidget(self.problem_txt)
        btn_ps.clicked.connect(lambda: self.problem_txt.setPlainText("Will be enabled in Future"))

        b.addSpacing(6)

        # Solution
        sol_row = QHBoxLayout()
        sol_row.addWidget(highlight_label("AI Assisted Proposed Solution"))
        btn_sol = QPushButton("Generate Proposed Solution"); btn_sol.setFixedSize(180,26)
        sol_row.addWidget(btn_sol); sol_row.addStretch(1)
        b.addLayout(sol_row)

        self.solution_txt = LimitedTextEdit(2000); self.solution_txt.setMinimumHeight(140); self.solution_txt.setMaximumWidth(900)
        b.addWidget(self.solution_txt)
        btn_sol.clicked.connect(lambda: self.solution_txt.setPlainText("Will be enabled in Future"))

        outer.addWidget(self.secB); outer.addStretch(1)


    # ---- Flow Logic ----
    def clear_flow(self):
        while self.flow_area.count():
            w = self.flow_area.takeAt(0).widget()
            if w: w.deleteLater()
        self.ec_result_lbl.setText(""); self.ec_divider.setVisible(False)

    def add_q(self,q,opts,cb):
        row = QHBoxLayout(); row.addWidget(QLabel(q))
        grp = QButtonGroup(self)
        for o in opts:
            rb = QRadioButton(o); grp.addButton(rb)
            rb.toggled.connect(lambda c,v=o: c and cb(v))
            row.addWidget(rb)
        row.addStretch(1)
        w = QWidget(); w.setLayout(row); self.flow_area.addWidget(w)

    def start_flow(self):
        if not self.sender().isChecked(): return
        self.clear_flow(); s=self.sender().text()
        if s=="Production Release": return self.finish("B3")
        if s=="OBS / Inactivate": return self.finish("B1")
        if s=="Up Revision": return self.add_q("SmBOM Impacted?",["Yes","No"],self.up1)
        if s=="Status Roll": return self.add_q("Transition Type?",["EVAL → Production","Concept → BOM List"],self.status1)
        if s=="Product Release": return self.add_q("SmBOM Impacted?",["Yes","No"],self.prod1)

    def up1(self,v): self.trim(1); self.finish("A1" if v=="Yes" else "A2") if v in ["Yes","No"] and False else (self.add_q("OBS Kit / Pieces?",["Yes","No"],lambda x:self.finish("A1" if x=="Yes" else "A2")) if v=="Yes" else self.add_q("Part Status?",["EVAL","PROD"],lambda x:self.finish("B2" if x=="EVAL" else "B3")))
    def status1(self,v): self.trim(1); self.finish("B3") if v.startswith("EVAL") else self.add_q("OBS Kit / Pieces?",["Yes","No"],lambda x:self.finish("A1" if x=="Yes" else "A2"))
    def prod1(self,v): self.trim(1); self.finish("A2") if v=="Yes" else self.add_q("Part Status?",["EVAL","PROD"],lambda x:self.finish("B2" if x=="EVAL" else "B3"))

    def trim(self,k):
        while self.flow_area.count()>k:
            w=self.flow_area.takeAt(self.flow_area.count()-1).widget()
            if w: w.deleteLater()

    def finish(self,c):
        self.ec_result_lbl.setText(f"EC Category: {c} – {EC_CATEGORY_DESC[c]}")
        self.ec_divider.setVisible(True); self.secB_header.setVisible(True); self.secB.setVisible(True)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__(); self.setWindowTitle(APP_TITLE); self.resize(1250,900)
        tabs=QTabWidget(); tabs.addTab(ECCreationInputsFormTab(),"EC Creation Inputs Form"); self.setCentralWidget(tabs)


def run(): app=QApplication(sys.argv); w=MainWindow(); w.show(); sys.exit(app.exec())

if __name__=='__main__': run()
