from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog, QLabel, QMessageBox,
    QSpacerItem, QSizePolicy, QFrame, QListWidget, QListWidgetItem, QListView,
    QAbstractItemView, QStackedLayout, QLineEdit
)
from PyQt5.QtGui import QPixmap, QFont, QIcon
from PyQt5.QtCore import Qt

from app.core.parser import build_master_row
from app.core.parser import extract_types_and_values
from app.gui.tolerance_dialog import ToleranceDialog
from app.io.excel_writer import export_master_report
from app.core.validator import is_pass

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Orava Gemstone Master Reporter")
        self.resize(900, 600)
        self.setStyleSheet("background-color:#f7f9fa;")

        # App state
        self.uploadedfiles = []
        self.tolerancedict = {}
        self.allheaders = {}
        self.alldata = {}
        self.reportcreatorinput = None
        self.reporttitleinput = None
        self.lastnominals = None
        self.master_colnames = []
        self.master_rows = []

        self.stacked = QStackedLayout()
        self.buildwelcomescreen()
        self.buildworkflowscreen()
        self.buildexportscreen()
        self.setLayout(self.stacked)

    def buildwelcomescreen(self):
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignTop)
        accent = QFrame()
        accent.setFixedHeight(6)
        accent.setStyleSheet("background-color:#366092; border:none;")
        layout.addWidget(accent)
        block = QVBoxLayout()
        block.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        block.setContentsMargins(0, 28, 0, 0)
        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        try:
            pixmap = QPixmap("resources/logo.png")
            logo_label.setPixmap(pixmap.scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        except Exception:
            logo_label.setText("Logo")
            logo_label.setStyleSheet("color:#bbb;font-size:12px;")
        block.addWidget(logo_label)
        name_label = QLabel("Orava Gemstone Master Reporter")
        name_font = QFont("Segoe UI", 28, QFont.Bold)
        name_label.setFont(name_font)
        name_label.setAlignment(Qt.AlignHCenter)
        name_label.setStyleSheet("color:#366092;letter-spacing:1px;margin-top:10px;")
        block.addWidget(name_label)
        subtext = QLabel("Welcome! Upload gemstone Excel reports to consolidate and validate into a master sheet.<br>Set tolerance after file selection. Export a color-coded, professional Excel with one click.")
        subtext.setWordWrap(True)
        subtext.setAlignment(Qt.AlignHCenter)
        subtext.setStyleSheet("font-size:15px;color:#222;margin-top:18px;margin-bottom:12px;")
        block.addWidget(subtext)
        layout.addLayout(block, stretch=0)
        layout.addStretch(1)
        infolabel = QLabel()
        infolabel.setStyleSheet("font-size:13px;color:#444;margin:6px 0 0 0;")
        infolabel.setWordWrap(True)
        infolabel.setAlignment(Qt.AlignHCenter)
        layout.addWidget(infolabel)
        self.welcomeinfolabel = infolabel
        layout.addStretch(2)
        buttonlayout = QHBoxLayout()
        buttonlayout.setContentsMargins(0, 0, 0, 30)
        buttonlayout.setSpacing(24)
        buttonlayout.addSpacerItem(QSpacerItem(100, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        upload_btn = QPushButton("Upload Files (Excel)")
        upload_btn.setIcon(QIcon.fromTheme("document-open"))
        upload_btn.setStyleSheet(
            "background-color:#366092; color:white; padding:11px 32px; border-radius:8px; font-size:17px; font-weight:bold;"
        )
        upload_btn.clicked.connect(self.handleuploadclicked)
        buttonlayout.addWidget(upload_btn)
        buttonlayout.addSpacerItem(QSpacerItem(100, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        layout.addLayout(buttonlayout)
        footer = QLabel("© 2025 Orava Solutions|Developed by Shehan Nirmana")
        footer.setStyleSheet("color:#aaa;background:transparent;font-size:12px;letter-spacing:0.6px;padding-bottom:6px;")
        footer.setAlignment(Qt.AlignCenter)
        layout.addWidget(footer)
        welcomewidget = QWidget()
        welcomewidget.setLayout(layout)
        self.stacked.addWidget(welcomewidget)

    def buildworkflowscreen(self):
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignTop)
        accent = QFrame()
        accent.setFixedHeight(6)
        accent.setStyleSheet("background-color:#366092; border:none;")
        layout.addWidget(accent)
        block = QVBoxLayout()
        block.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        block.setContentsMargins(0, 28, 0, 0)
        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        try:
            pixmap = QPixmap("resources/logo.png")
            logo_label.setPixmap(pixmap.scaled(200, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        except Exception:
            logo_label.setText("Logo")
            logo_label.setStyleSheet("color:#bbb;font-size:12px;")
        block.addWidget(logo_label)
        name_label = QLabel("Orava Gemstone Master Reporter")
        name_font = QFont("Segoe UI", 28, QFont.Bold)
        name_label.setFont(name_font)
        name_label.setAlignment(Qt.AlignHCenter)
        name_label.setStyleSheet("color:#366092;letter-spacing:1px;margin-top:10px;")
        block.addWidget(name_label)
        subtext = QLabel("Set tolerance after file selection. Export a color-coded, professional Excel with one click.")
        subtext.setWordWrap(True)
        subtext.setAlignment(Qt.AlignHCenter)
        subtext.setStyleSheet("font-size:15px;color:#222;margin-top:18px;margin-bottom:12px;")
        block.addWidget(subtext)
        layout.addLayout(block, stretch=0)
        self.filecountlabel = QLabel("Total uploaded files: 0")
        self.filecountlabel.setAlignment(Qt.AlignHCenter)
        self.filecountlabel.setStyleSheet("font-size:14px;color:#237346;margin-bottom:4px;")
        layout.addWidget(self.filecountlabel)
        self.filelist = QListWidget()
        self.filelist.setViewMode(QListView.ListMode)
        self.filelist.setSelectionMode(QAbstractItemView.NoSelection)
        self.filelist.setFixedHeight(180)
        self.filelist.setStyleSheet(
            "background:white; font-size:14px; border:1px solid #ddd; margin:16px 80px; padding:4px 10px;"
        )
        layout.addWidget(self.filelist)
        btnslayout = QHBoxLayout()
        btnslayout.setContentsMargins(0, 0, 0, 0)
        btnslayout.setSpacing(18)
        btnslayout.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        add_btn = QPushButton("Add more Excel")
        add_btn.setIcon(QIcon.fromTheme("document-open"))
        add_btn.setStyleSheet(
            "background-color:#366092; color:white; padding:8px 20px; border-radius:8px; font-size:15px; font-weight:bold;"
        )
        add_btn.clicked.connect(self.addmorefiles)
        btnslayout.addWidget(add_btn)
        clear_btn = QPushButton("Clear All")
        clear_btn.setIcon(QIcon.fromTheme("edit-clear"))
        clear_btn.setStyleSheet(
            "background-color:#d9534f; color:white; padding:8px 20px; border-radius:8px; font-size:15px; font-weight:bold;"
        )
        clear_btn.clicked.connect(self.clearallfiles)
        btnslayout.addWidget(clear_btn)
        tolerance_btn = QPushButton("Add Tolerance")
        tolerance_btn.setIcon(QIcon.fromTheme("dialog-ok"))
        tolerance_btn.setStyleSheet(
            "background-color:#237346; color:white; padding:8px 20px; border-radius:8px; font-size:15px; font-weight:bold;"
        )
        tolerance_btn.clicked.connect(self.addtolerance)
        tolerance_btn.setEnabled(True)
        self.tolerancebtn = tolerance_btn
        btnslayout.addWidget(tolerance_btn)
        btnslayout.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        layout.addLayout(btnslayout)
        infolabel = QLabel()
        infolabel.setStyleSheet("font-size:13px;color:#444;margin:12px 80px 0 80px;")
        infolabel.setWordWrap(True)
        infolabel.setAlignment(Qt.AlignHCenter)
        layout.addWidget(infolabel)
        self.workflowinfolabel = infolabel
        footer = QLabel("© 2025 Orava Solutions|Developed by Shehan Nirmana")
        footer.setStyleSheet("color:#aaa;background:transparent;font-size:12px;letter-spacing:0.6px;padding-bottom:6px;")
        footer.setAlignment(Qt.AlignCenter)
        layout.addWidget(footer)
        workflowwidget = QWidget()
        workflowwidget.setLayout(layout)
        self.stacked.addWidget(workflowwidget)

    def buildexportscreen(self):
        outer = QVBoxLayout()
        outer.setAlignment(Qt.AlignTop)
        outer.setContentsMargins(0, 0, 0, 0)
        accent = QFrame()
        accent.setFixedHeight(6)
        accent.setStyleSheet("background-color:#366092; border:none;")
        outer.addWidget(accent)
        headerblock = QVBoxLayout()
        headerblock.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        headerblock.setContentsMargins(0, 32, 0, 0)
        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        try:
            pixmap = QPixmap("resources/logo.png")
            logo_label.setPixmap(pixmap.scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        except Exception:
            logo_label.setText("Logo")
            logo_label.setStyleSheet("color:#bbb;font-size:18px;")
        headerblock.addWidget(logo_label)
        namelabel = QLabel("Ready to Export Master Excel Sheet")
        name_font = QFont("Segoe UI", 23, QFont.Bold)
        namelabel.setFont(name_font)
        namelabel.setAlignment(Qt.AlignHCenter)
        namelabel.setStyleSheet("color:#366092;letter-spacing:1px;margin-top:10px;")
        headerblock.addWidget(namelabel)
        slogan = QLabel("Click the button below to generate and save your color-coded master report.")
        slogan.setWordWrap(True)
        slogan.setAlignment(Qt.AlignHCenter)
        slogan.setStyleSheet("font-size:16px;color:#444;margin-top:8px;margin-bottom:20px;")
        headerblock.addWidget(slogan)
        headerblock.addSpacing(25)
        outer.addLayout(headerblock, stretch=0)
        outer.addStretch(1)
        center = QVBoxLayout()
        center.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        center.setSpacing(12)
        self.reporttitleinput = QLineEdit()
        self.reporttitleinput.setPlaceholderText("Master Excel report title (file name)")
        self.reporttitleinput.setFixedWidth(340)
        self.reporttitleinput.setStyleSheet("margin-bottom:9px;font-size:15px;padding:8px 12px;")
        center.addWidget(self.reporttitleinput)
        self.reportcreatorinput = QLineEdit()
        self.reportcreatorinput.setPlaceholderText("Report creator name will appear in Excel")
        self.reportcreatorinput.setFixedWidth(340)
        self.reportcreatorinput.setStyleSheet("margin-bottom:7px;font-size:15px;padding:8px 12px;")
        center.addWidget(self.reportcreatorinput)
        center.addSpacing(16)
        self.exportbutton = QPushButton("Export Master Excel Sheet")
        self.exportbutton.setStyleSheet(
            "background-color:#237346; color:white; padding:14px 36px; border-radius:8px; font-size:18px; font-weight:bold;"
        )
        self.exportbutton.clicked.connect(self.exportmasterreport)
        center.addWidget(self.exportbutton, alignment=Qt.AlignHCenter)
        backbtn = QPushButton("Back")
        backbtn.setStyleSheet(
            "background-color:#366092; color:white; padding:13px 34px; border-radius:8px; font-size:15px; margin-top:14px; font-weight:bold;"
        )
        backbtn.clicked.connect(self.gobacktoworkflow)
        home_btn = QPushButton("Home")
        home_btn.setIcon(QIcon.fromTheme("go-home"))
        home_btn.setStyleSheet(
            "background-color:#4b85c5; color:white; padding:13px 34px; border-radius:8px; font-size:15px; margin-top:14px; font-weight:bold;"
        )
        home_btn.clicked.connect(self.go_home_reset)
        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)
        btn_row.setAlignment(Qt.AlignHCenter)
        btn_row.addWidget(backbtn)
        btn_row.addWidget(home_btn)
        center.addLayout(btn_row)
        outer.addLayout(center)
        outer.addStretch(3)
        footer = QLabel("© 2025 Orava Solutions|Developed by Shehan Nirmana")
        footer.setStyleSheet("color:#aaa;background:transparent;font-size:12px;letter-spacing:0.6px;padding-bottom:6px;")
        footer.setAlignment(Qt.AlignCenter)
        outer.addWidget(footer)
        exportwidget = QWidget()
        exportwidget.setLayout(outer)
        self.stacked.addWidget(exportwidget)

    def handleuploadclicked(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Excel Files", "", "Excel Files (*.xlsx)")
        if not files:
            return
        self.onuploadfiles(files)
        self.updatefilelist()
        self.stacked.setCurrentIndex(1)

    def onuploadfiles(self, filelist):
        self.uploadedfiles = []
        for file in filelist:
            try:
                import openpyxl
                wb = openpyxl.load_workbook(file, data_only=True)
                wb.close()
                self.uploadedfiles.append(file)
            except Exception as e:
                print(f"Failed to validate file: {e}")
                QMessageBox.warning(self, "Parse Error", f"Failed to read file\n{str(e)}")

    def updatefilelist(self):
        self.filelist.clear()
        for filename in self.uploadedfiles:
            item = QListWidgetItem(filename)
            widget = QWidget()
            hlayout = QHBoxLayout()
            hlayout.setContentsMargins(0, 0, 0, 0)
            label = QLabel(filename)
            label.setStyleSheet("font-size:14px;color:#444;")
            removebtn = QPushButton("X")
            removebtn.setFixedWidth(26)
            removebtn.setStyleSheet(
                "background-color:#d9534f; color:white; border:none; border-radius:13px; font-size:13px;"
            )
            def makeremover(fname):
                def inner():
                    if fname in self.uploadedfiles:
                        self.uploadedfiles.remove(fname)
                    if fname in self.allheaders:
                        del self.allheaders[fname]
                    if fname in self.alldata:
                        del self.alldata[fname]
                    self.updatefilelist()
                return inner
            removebtn.clicked.connect(makeremover(filename))
            hlayout.addWidget(label)
            hlayout.addStretch()
            hlayout.addWidget(removebtn)
            widget.setLayout(hlayout)
            item.setSizeHint(widget.sizeHint())
            self.filelist.addItem(item)
            self.filelist.setItemWidget(item, widget)
        self.tolerancebtn.setEnabled(bool(self.uploadedfiles))
        try:
            hasfiles = bool(self.uploadedfiles)
            hastolerances = bool(self.tolerancedict)
            self.exportbutton.setEnabled(hasfiles and hastolerances)
        except Exception:
            try:
                self.exportbutton.setEnabled(False)
            except Exception:
                pass
        self.filecountlabel.setText(f"Total uploaded files: {len(self.uploadedfiles)}")

    def addmorefiles(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Excel Files", "", "Excel Files (*.xlsx)")
        if files:
            for f in files:
                if f not in self.uploadedfiles:
                    try:
                        import openpyxl
                        wb = openpyxl.load_workbook(f, data_only=True)
                        wb.close()
                        self.uploadedfiles.append(f)
                    except Exception as e:
                        print(f"Failed to validate: {f} {e}")
                        QMessageBox.warning(self, "Parse Error", f"Failed to read {f}\n{str(e)}")
            self.updatefilelist()
            self.workflowinfolabel.clear()

    def clearallfiles(self):
        self.uploadedfiles = []
        self.tolerancedict = {}
        self.lastnominals = None
        self.updatefilelist()
        self.workflowinfolabel.clear()

    def addtolerance(self):
        if not self.uploadedfiles:
            QMessageBox.warning(self, "Error", "Please add Excel files first!")
            return
        columns = None
        try:
            firstfile = self.uploadedfiles[0]
            cols, _ = extract_types_and_values(firstfile)
            columns = [self.map_symbol(c) for c in cols] if cols else []
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to read columns from file\n{str(e)}")
            self.tolerancedict = {}
            self.workflowinfolabel.setText(f"<b>Error:</b> Failed to read columns {e}")
            return
        if columns:
            dlg = ToleranceDialog(columns, self, previous_nominals=self.lastnominals)
            if dlg.exec_():
                self.tolerancedict = dlg.get_tolerances()
                try:
                    self.lastnominals = self.tolerancedict.copy()
                except Exception:
                    self.lastnominals = None
                self.updatefilelist()
                self.stacked.setCurrentIndex(2)
            else:
                try:
                    self.tolerancedict = dlg.get_tolerances()
                    self.lastnominals = self.tolerancedict.copy()
                except Exception:
                    self.tolerancedict = {}
                    self.lastnominals = None
                self.updatefilelist()
                summary = "<br><b style='color:red'>Tolerance setup canceled.</b>"
                self.workflowinfolabel.setText(summary)
        else:
            summary = "<br><b style='color:red'>No measurement columns detected.</b>"
            self.tolerancedict = {}
            self.workflowinfolabel.setText(summary)

    def gobacktoworkflow(self):
        self.stacked.setCurrentIndex(1)

    def go_home_reset(self):
        self.uploadedfiles = []
        self.tolerancedict = {}
        self.allheaders = {}
        self.alldata = {}
        self.master_colnames = []
        self.master_rows = []
        try:
            self.lastnominals = None
        except Exception:
            pass
        try:
            self.updatefilelist()
        except Exception:
            pass
        try:
            self.workflowinfolabel.clear()
        except Exception:
            pass
        self.stacked.setCurrentIndex(0)

    def _extract_filename(self, path):
        import os
        return os.path.basename(path)

    def map_symbol(self, name):
        if "(mm)" in name or "(⟳)" in name or "(°)" in name:
            return name
        if "Distance" in name or "Diameter" in name:
            return f"{name} (mm)"
        elif "Concentricity" in name:
            return f"{name} (⟳)"
        elif "Angle" in name:
            return f"{name} (°)"
        return name

    def process_all_files_for_report(self):
        all_file_columns = []
        all_data_rows = []
        raw_col_set = []
        for filepath in self.uploadedfiles:
            source_name = self._extract_filename(filepath)
            cols, row = build_master_row(filepath, source_name)
            source_id = source_name.replace(".xlsx", "").replace(".xls", "")
            row[0] = source_id
            all_file_columns.append(cols)
            all_data_rows.append((source_id, row, cols))
            for c in cols:
                if c not in raw_col_set:
                    raw_col_set.append(c)
        master_cols = [self.map_symbol(c) for c in raw_col_set]
        self.master_colnames = ["Source_File", "Report_Runtime"] + master_cols + ["Final Status"]
        def file_id_key(item):
            try:
                return int(item[0])
            except:
                return item[0]
        all_data_rows.sort(key=file_id_key)
        self.master_rows = []
        for source_id, row, cols in all_data_rows:
            output_row = [source_id]
            output_row.append(row[1] if len(row) > 1 else "")
            col2val = {self.map_symbol(c): v for c, v in zip(cols, row[2:])}
            for h in master_cols:
                output_row.append(col2val.get(h, ""))
            output_row.append("")
            self.master_rows.append(output_row)

    def exportmasterreport(self):
        creator = self.reportcreatorinput.text().strip()
        reporttitle = self.reporttitleinput.text().strip()
        if not creator or not reporttitle:
            QMessageBox.warning(self, "Missing Info", "Please enter both the report creator and a report title!")
            return
        if not self.uploadedfiles:
            QMessageBox.warning(self, "Error", "No files to export.")
            return
        savename = f"{reporttitle.replace(' ', '_')}.xlsx"
        defaultpath = savename
        path, _ = QFileDialog.getSaveFileName(self, "Save Master Report", defaultpath, "Excel Files (*.xlsx)")
        if not path:
            return
        try:
            self.process_all_files_for_report()
            synthetic_key = "__master__"
            transformed_headers = {synthetic_key: self.master_colnames}
            transformed_data = {synthetic_key: self.master_rows}
            outpath = export_master_report(
                files=[synthetic_key],
                all_headers=transformed_headers,
                all_data=transformed_data,
                tolerance_dict=self.tolerancedict,
                col_names=self.master_colnames,
                output_path=path,
                creator=creator,
                report_title=reporttitle,
            )
            QMessageBox.information(
                self, "Export Complete",
                f"Master report saved to:\n{outpath}\nCreator: {creator}\nTitle: {reporttitle}"
            )
            self.lastnominals = None
        except Exception as e:
            QMessageBox.critical(self, "Export Failed", f"Failed to export master report\n{str(e)}")
