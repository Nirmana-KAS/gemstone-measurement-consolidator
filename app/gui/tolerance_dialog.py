from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QDialogButtonBox, QScrollArea, QWidget, QMessageBox
)
from PyQt5.QtCore import Qt

class ToleranceDialog(QDialog):
    def __init__(self, columns, parent=None, previous_nominals=None):
        super().__init__(parent)
        self.setWindowTitle("Set Tolerance for Each Column")
        self.resize(700, 400)
        self.inputs = {}

        main_layout = QVBoxLayout(self)
        info = QLabel("Set Nominal value, Tolerance + (upper), and Tolerance - (lower) for each column:")
        info.setWordWrap(True)
        main_layout.addWidget(info)

        # --- Scrollable Container ---
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_widget = QWidget()
        layout = QVBoxLayout(scroll_widget)

        # --- Header ---
        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("Column"), 1)
        header_layout.addWidget(QLabel("Nominal Value"), 1)
        header_layout.addWidget(QLabel("Tolerance + (Upper)"), 1)
        header_layout.addWidget(QLabel("Tolerance - (Lower)"), 1)
        layout.addLayout(header_layout)

        # --- Rows for Each Column ---
        for col in columns:
            row = QHBoxLayout()
            lbl = QLabel(col)
            lbl.setFixedWidth(200)
            nominal = QLineEdit()
            nominal.setPlaceholderText("Nominal value")
            nominal.setText("")
            plus = QLineEdit()
            plus.setPlaceholderText("+Tol (default: 0.05)")
            plus.setText("0.05")
            minus = QLineEdit()
            minus.setPlaceholderText("-Tol (default: 0.05)")
            minus.setText("0.05")
            # Restore/or set previous values
            if previous_nominals and col in previous_nominals:
                if previous_nominals[col][0] is not None:
                    nominal.setText(str(previous_nominals[col][0]))
                if previous_nominals[col][1] is not None:
                    plus.setText(str(previous_nominals[col][1]))
                if previous_nominals[col][2] is not None:
                    minus.setText(str(previous_nominals[col][2]))
            row.addWidget(lbl, 1)
            row.addWidget(nominal, 1)
            row.addWidget(plus, 1)
            row.addWidget(minus, 1)
            layout.addLayout(row)
            self.inputs[col] = (nominal, plus, minus)

        layout.addStretch()
        scroll_area.setWidget(scroll_widget)
        main_layout.addWidget(scroll_area)

        # --- Clear Nominals Button ---
        clear_button = QPushButton("Clear Nominal Values")
        clear_button.setStyleSheet("color:#366092; font-size:14px; padding:4px 18px; background:#eee;")
        clear_button.clicked.connect(self.clear_nominals)
        main_layout.addWidget(clear_button)

        # --- Dialog Buttons (OK/Cancel) ---
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.handle_accept_ok)
        buttons.rejected.connect(self.reject)
        main_layout.addWidget(buttons)

    def clear_nominals(self):
        for nominal, _, _ in self.inputs.values():
            nominal.clear()

    def handle_accept_ok(self):
        # Require ALL nominal values filled
        missing = []
        for col, (nom, plus, minus) in self.inputs.items():
            if not nom.text().strip():
                missing.append(col)
        if missing:
            QMessageBox.critical(
                self,
                "Missing Value",
                "Please enter a nominal value for:\n" + "\n".join(missing)
            )
            return
        self.accept()

    def get_tolerances(self):
        result = {}
        for col, (nom, plus, minus) in self.inputs.items():
            try:
                result[col] = (
                    float(nom.text().strip()) if nom.text().strip() else None,
                    float(plus.text().strip()) if plus.text().strip() else 0.05,
                    float(minus.text().strip()) if minus.text().strip() else 0.05,
                )
            except Exception:
                result[col] = (None, 0.05, 0.05)
        return result
