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
        self.setWindowIcon(QIcon(r"D:\AadeshwarWorkOrder\Icon\card_image_bAz_icon.ico"))
        self.setStyleSheet("background-color:lightblue")
        self.backend = WorkOrderAppBackend()
        self.settings = QSettings(COMPANY_NAME, APP_NAME)
        self.initUI()

    def initUI(self):
       
        self.set_path_btn = QPushButton('Set Download Path',self)
        self.set_path_btn.clicked.connect(self.set_download_path)
        self.set_path_btn.setGeometry(30, 20, 200, 60)
        set_button_style(self.set_path_btn)

        self.csv_btn = QPushButton('Upload CSV File',self)
        self.csv_btn.clicked.connect(self.upload_csv)
        self.csv_btn.setGeometry(300, 20, 200, 60)
        set_button_style(self.csv_btn)
        
        self.foaming_template_btn = QPushButton('Upload Foaming',self)
        self.foaming_template_btn.clicked.connect(self.upload_foaming_template)
        self.foaming_template_btn.setGeometry(30, 110, 200, 60)  
        set_button_style(self.foaming_template_btn)

        self.carpenter_template_btn = QPushButton('Upload Carpenter', self)
        self.carpenter_template_btn.clicked.connect(self.upload_carpenter_template)
        self.carpenter_template_btn.setGeometry(300, 110, 200, 60)
        set_button_style(self.carpenter_template_btn)

        self.sales_template_btn = QPushButton('Upload Sales', self)
        self.sales_template_btn.clicked.connect(self.upload_sales_template)
        self.sales_template_btn.setGeometry(30, 200, 200, 60)
        set_button_style(self.sales_template_btn)
        
        self.generate_btn = QPushButton('Generate Work Order',self)
        self.generate_btn.clicked.connect(self.generate_work_order)
        self.generate_btn.setGeometry(150, 280, 230, 60)
        set_button_style(self.generate_btn)
        self.generate_btn.setEnabled(False)
        
        self.status_label = QLabel('',self)
        self.status_label.setGeometry(20, 360, 550, 50)
        self.status_label.setStyleSheet("Qlabel { font-size: 15px; }")

        self.order_no_input = QLineEdit(self)
        self.order_no_input.setPlaceholderText("Enter Starting Order Number")
        self.order_no_input.setGeometry(300, 200, 200, 30)

        self.save_btn = QPushButton('Save',self)
        self.save_btn.clicked.connect(self.set_order_no)
        self.save_btn.setGeometry(300, 231, 60, 30)
        self.save_btn.setStyleSheet('QPushButton { padding: 5px; font-size: 14px; background-color:white; }')

        self.setWindowTitle('Work Order Generator')
        self.setWindowIcon(QIcon('app_icon.ico'))  
        self.csv_path = ''
        self.template_path = ''

        center_widget(self)  

    def set_download_path(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        download_path = QFileDialog.getExistingDirectory(self, "Select Download Path", QDir.homePath(), options=options)
        if download_path:
            self.backend.set_download_path(download_path)


    def upload_csv(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(self, 'Upload CSV File', '', 'CSV Files (*.csv)', options=options)
        if file_path:
            self.backend.set_csv_path(file_path)
            self.status_label.setText('CSV File Uploaded')
    
    def upload_foaming_template(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(self, 'Upload Foaming Template', '', 'Word Files (*.docx)', options=options)
        if file_path:
            self.backend.set_foaming_template_path(file_path)
            self.status_label.setText('Foaming Template Uploaded')

    def upload_carpenter_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Carpenter Template", "", "Word Files (*.docx)")
        if path:
            self.backend.set_carpenter_template_path(path)
            self.status_label.setText(f"Carpenter Template Uploaded")

    def upload_sales_template(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Sales Template", "", "Word Files (*.docx)")
        if path:
            self.backend.set_sales_template_path(path)
            self.status_label.setText(f"Sales Template Uploaded")

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
        if not self.backend.csv_path:
            self.status_label.setText('Please upload the CSV file')
            return
        elif not self.backend.foaming_template_path:
            self.status_label.setText('Please upload the Foaming Template file')
            return
        elif not self.backend.carpenter_template_path:
            self.status_label.setText('Please upload the Carpenter Template file')
            return
        elif not self.backend.sales_template_path:
            self.status_label.setText('Please upload the Sales Template file')
            return
        
        if not self.generate_btn.isEnabled():
            self.status_label.setText("Invalid order number. Please enter a valid order number.")
            return
        
        self.backend.generate_work_order()
        self.status_label.setText('Work Order Generated')