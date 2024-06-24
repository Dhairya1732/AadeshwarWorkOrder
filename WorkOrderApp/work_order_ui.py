import os
from PyQt5.QtWidgets import QWidget, QPushButton, QLabel, QFileDialog, QLineEdit
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSettings, QDir
from work_order_backend import WorkOrderAppBackend
from utils import center_widget, set_button_style

COMPANY_NAME = "Aadeshwar"
APP_NAME = "WorkOrderGenerator"

class WorkOrderAppUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setGeometry(200, 200, 550, 400)
        self.setWindowTitle("Work Order Generator")
        self.setWindowIcon(QIcon(r"D:\AadeshwarWorkOrder\card_image.png"))
        self.setStyleSheet("background-color:lightblue")
        self.backend = WorkOrderAppBackend()
        self.settings = QSettings(COMPANY_NAME, APP_NAME)
        self.initUI()

    def initUI(self):
       
        self.set_path_btn = QPushButton('Set Download Path',self)
        self.set_path_btn.clicked.connect(self.set_download_path)
        self.set_path_btn.setGeometry(30, 20, 200, 60)
        set_button_style(self.set_path_btn)

        self.excel_btn = QPushButton('Upload Excel File',self)
        self.excel_btn.clicked.connect(self.upload_excel)
        self.excel_btn.setGeometry(300, 20, 200, 60)
        set_button_style(self.excel_btn)
        
        self.template_btn = QPushButton('Upload Word Template',self)
        self.template_btn.clicked.connect(self.upload_template)
        self.template_btn.setGeometry(30, 110, 230, 60)  
        set_button_style(self.template_btn)
        
        self.generate_btn = QPushButton('Generate Work Order',self)
        self.generate_btn.clicked.connect(self.generate_work_order)
        self.generate_btn.setGeometry(150, 200, 230, 60)
        set_button_style(self.generate_btn)
        self.generate_btn.setEnabled(False)
        
        self.status_label = QLabel('',self)
        self.status_label.setGeometry(20, 280, 550, 50)
        self.status_label.setStyleSheet("Qlabel { font-size: 15px; }")

        self.order_no_input = QLineEdit(self)
        self.order_no_input.setPlaceholderText("Enter Starting Order Number")
        self.order_no_input.setGeometry(300, 110, 200, 30)

        self.save_btn = QPushButton('Save',self)
        self.save_btn.clicked.connect(self.set_order_no)
        self.save_btn.setGeometry(300, 141, 60, 30)
        self.save_btn.setStyleSheet('QPushButton { padding: 5px; font-size: 14px; background-color:white; }')

        self.setWindowTitle('Work Order Generator')
        self.setWindowIcon(QIcon('app_icon.ico'))  
        self.excel_path = ''
        self.template_path = ''

        template_path = self.settings.value("template_path")
        if template_path and os.path.isfile(template_path):
            self.template_path = template_path
            self.status_label.setText('Word Template Loaded')
        else:
            self.status_label.setText('Please upload Word Template')

        center_widget(self)  

    def set_download_path(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        download_path = QFileDialog.getExistingDirectory(self, "Select Download Path", QDir.homePath(), options=options)
        if download_path:
            self.backend.set_download_path(download_path)


    def upload_excel(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(self, 'Upload Excel File', '', 'Excel Files (*.xlsx)', options=options)
        if file_path:
            self.backend.set_excel_path(file_path)
            self.status_label.setText('Excel File Uploaded')
    
    def upload_template(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(self, 'Upload Word Template', '', 'Word Files (*.docx)', options=options)
        if file_path:
            self.backend.set_template_path(file_path)
            self.status_label.setText('Word Template Uploaded')

    def set_order_no(self):
        try:
            if not self.order_no_input.text():
                raise ValueError("Order number cannot be empty.")
            starting_order_no = int(self.order_no_input.text())  
            if starting_order_no <= 0:
                raise ValueError("Starting order number must be a positive integer")
            self.backend.current_order_no = starting_order_no  
            self.status_label.setText("Order number saved.")
            self.generate_btn.setEnabled(True)
        except ValueError:
            self.status_label.setText("Invalid starting order number. Please enter a positive integer.")
            self.generate_btn.setEnabled(False)
    
    def generate_work_order(self):
        if not self.backend.excel_path or not self.backend.template_path:
            self.status_label.setText('Please upload both files')
            return
        
        if not self.generate_btn.isEnabled():
            self.status_label.setText("Invalid order number. Please enter a valid order number.")
            return
        
        self.backend.generate_work_order()
        self.status_label.setText('Work Order Generated')