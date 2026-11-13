from pathlib import Path

from PySide6.QtCore import Qt, QThread
from PySide6.QtGui import QIcon, QPixmap
from PySide6.QtWidgets import QApplication, QStyle
from PySide6.QtWidgets import (
    QMainWindow,
    QWidget,
    QLabel,
    QPushButton,
    QLineEdit,
    QFileDialog,
    QVBoxLayout,
    QHBoxLayout,
    QGroupBox,
    QCheckBox,
    QProgressBar,
    QTextEdit,
    QMessageBox,
)

from src.workers.converter_worker import ConverterWorker


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        
        self.setWindowIcon(QIcon("assets/icon.ico"))
        self.setWindowTitle("Chainage PDF Converter")
        self.resize(900, 650)
        self.setStyleSheet("""
        QMainWindow{
            background:#f4f8fc;
    }

        QGroupBox{
            border:1px solid #c7d7e6;
            border-radius:8px;
            margin-top:10px;
            font-weight:bold;
            padding-top:10px;
            background:white;
        }

        QGroupBox::title{
            subcontrol-origin: margin;
            left:10px;
            padding:0 5px;
        }

        QPushButton{
            background:#1976d2;
            color:white;
            border:none;
            border-radius:6px;
            padding:8px;
            font-weight:bold;
        }

        QPushButton:hover{
            background:#1565c0;
        }

        QPushButton:disabled{
            background:#b0bec5;
        }

        QLineEdit{
            border:1px solid #b0bec5;
            border-radius:5px;
            padding:6px;
            background:white;
        }

        QTextEdit{
            border:1px solid #b0bec5;
            border-radius:5px;
            background:white;
        }

        QProgressBar{
            border:1px solid #90caf9;
            border-radius:5px;
            text-align:center;
        }

        QProgressBar::chunk{
            background:#1976d2;
        }
        """)

        self.thread = None
        self.worker = None

        self.setup_ui()

    # -----------------------------------------------------
    # UI
    # -----------------------------------------------------

    def setup_ui(self):

        central = QWidget()
        self.setCentralWidget(central)

        layout = QVBoxLayout(central)

        # -----------------------------
        # Company Logo
        # -----------------------------
        from pathlib import Path

        logo = QLabel()

        logo_path = Path(__file__).resolve().parents[2] / "assets" / "amx_logo.png"

        print("Logo Path:", logo_path)
        print("Exists:", logo_path.exists())

        pixmap = QPixmap(str(logo_path))

        logo.setPixmap(
            pixmap.scaled(
                100,
                100,
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )
        )

        logo.setAlignment(Qt.AlignCenter)

        layout.addWidget(logo)

        # -----------------------------
        # Title
        # -----------------------------

        title = QLabel("Chainage PDF Converter")

        title.setAlignment(Qt.AlignCenter)

        title.setStyleSheet("""
            font-size:28px;
            font-weight:bold;
            color:#0D47A1;
        """)

        layout.addWidget(title)
        created_by = QLabel("Developed by\nAMX Aerial Survey Team")

        created_by.setAlignment(Qt.AlignCenter)

        created_by.setStyleSheet("""
            font-size:12px;
            color:#666666;
            font-style:italic;
        """)

        layout.addWidget(created_by)

        # -----------------------------
        # Input
        # -----------------------------

        # ===================================================
        # Excel Workbook
        # ===================================================

        excel_group = QGroupBox("📗 Excel Workbook")

        excel_layout = QHBoxLayout()

        self.input_path = QLineEdit()
        self.input_path.setPlaceholderText(
            "Select Excel Workbook (*.xlsx)"
        )

        self.btn_input = QPushButton("Browse")
        self.btn_input.setIcon(QIcon("assets/excel.png"))
        self.btn_input.setMinimumWidth(120)

        excel_layout.addWidget(self.input_path)
        excel_layout.addWidget(self.btn_input)

        excel_group.setLayout(excel_layout)

        layout.addWidget(excel_group)

        # ===================================================
        # Output Folder
        # ===================================================

        output_group = QGroupBox("📁 Output Folder")

        output_layout = QHBoxLayout()

        self.output_path = QLineEdit()
        self.output_path.setPlaceholderText(
            "Select Output Folder"
        )

        self.btn_output = QPushButton("Browse")
        self.btn_output.setIcon(QIcon("assets/folder.png"))
        self.btn_output.setMinimumWidth(120)

        output_layout.addWidget(self.output_path)
        output_layout.addWidget(self.btn_output)

        output_group.setLayout(output_layout)

        layout.addWidget(output_group)

        # -----------------------------
        # Options
        # -----------------------------

        group = QGroupBox("Options")

        g_layout = QVBoxLayout()

        self.chk_a3 = QCheckBox("Force A3 Landscape")
        self.chk_crop = QCheckBox("Crop White Space")
        self.chk_merge = QCheckBox("Merge to Single Page")
        self.chk_open = QCheckBox("Open Output Folder")

        self.chk_a3.setChecked(True)
        self.chk_crop.setChecked(True)
        self.chk_merge.setChecked(True)
        self.chk_open.setChecked(True)

        g_layout.addWidget(self.chk_a3)
        g_layout.addWidget(self.chk_crop)
        g_layout.addWidget(self.chk_merge)
        g_layout.addWidget(self.chk_open)

        group.setLayout(g_layout)

        layout.addWidget(group)
        # -----------------------------
        # Status
        # -----------------------------

        self.lbl_status = QLabel("Status : Ready")
        self.lbl_status.setStyleSheet("""
            font-size: 12px;
            font-weight: bold;
            color: #0066CC;
        """)

        layout.addWidget(self.lbl_status)
        # -----------------------------
        # Progress
        # -----------------------------

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)

        layout.addWidget(self.progress)

        # -----------------------------
        # Log
        # -----------------------------

        self.log = QTextEdit()
        self.log.setReadOnly(True)

        layout.addWidget(self.log)

        # -----------------------------
        # Convert Button
        # -----------------------------

        # -----------------------------
        # Convert Button
        # -----------------------------

        self.btn_convert = QPushButton("Convert")

        self.btn_convert.setIcon(
            QApplication.style().standardIcon(QStyle.SP_MediaPlay)
        )

        self.btn_convert.setMinimumHeight(45)

        layout.addWidget(self.btn_convert)

        # -----------------------------
        # Connections
        # -----------------------------

        self.btn_input.clicked.connect(self.select_input)
        self.btn_output.clicked.connect(self.select_output)
        self.btn_convert.clicked.connect(self.start_conversion)

    # -----------------------------------------------------
    # Log
    # -----------------------------------------------------

    def add_log(self, text):
        self.log.append(text)

    # -----------------------------------------------------
    # Browse Excel
    # -----------------------------------------------------

    def select_input(self):

        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Select Excel File",
            "",
            "Excel Files (*.xlsx *.xlsm)"
        )

        if filename:
            self.input_path.setText(filename)

    # -----------------------------------------------------
    # Browse Output Folder
    # -----------------------------------------------------

    def select_output(self):

        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Output Folder"
        )

        if folder:
            self.output_path.setText(folder)
    def set_controls_enabled(self, enabled):

        self.btn_input.setEnabled(enabled)
        self.btn_output.setEnabled(enabled)
        self.btn_convert.setEnabled(enabled)
    def set_status(self, text):
        self.lbl_status.setText(f"Status : {text}")   
    # -----------------------------------------------------
    # Convert
    # -----------------------------------------------------
    
    def start_conversion(self):

        excel = self.input_path.text().strip()
        output = self.output_path.text().strip()

        if not excel:
            QMessageBox.warning(
                self,
                "Missing File",
                "Please select an Excel file."
            )
            return

        if not output:
            QMessageBox.warning(
                self,
                "Missing Folder",
                "Please select an output folder."
            )
            return

        if not Path(excel).exists():
            QMessageBox.warning(
                self,
                "Error",
                "Excel file does not exist."
            )
            return

        self.log.clear()
        self.progress.setValue(0)

        self.set_controls_enabled(False)
        
        self.set_status("Starting...")

        self.thread = QThread()

        self.worker = ConverterWorker(
            excel,
            output
        )

        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)

        self.worker.log.connect(self.add_log)
        self.worker.progress.connect(self.progress.setValue)

        self.worker.finished.connect(self.conversion_finished)
        self.worker.error.connect(self.conversion_failed)

        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)

        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    # -----------------------------------------------------
    # Finished
    # -----------------------------------------------------

    def conversion_finished(self):

        self.set_controls_enabled(True)
        self.progress.setValue(100)
        self.set_status("Completed")
        QMessageBox.information(
            self,
            "Completed",
            "Conversion completed successfully."
        )

    # -----------------------------------------------------
    # Error
    # -----------------------------------------------------

    def conversion_failed(self, message):

        self.set_controls_enabled(True)
        self.set_status("Failed")

        QMessageBox.critical(
            self,
            "Conversion Failed",
            message
        )
    