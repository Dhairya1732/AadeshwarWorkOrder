from PyQt5.QtWidgets import QWidget, QPushButton, QLabel, QLineEdit
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSettings
from work_order_backend import WorkOrderAppBackend
from functools import partial
from utils import center_widget, set_button_style, upload_file
from config import COMPANY_NAME, APP_NAME

class WorkOrderAppUI(QWidget):
    def __init__(self):
        super().__init__()
        self.setGeometry(200, 200, 820, 400)
        self.setWindowTitle("Work Order Generator")
        self.setWindowIcon(QIcon(r"D:\AadeshwarWorkOrder\Icon\card_image_bAz_icon.ico"))
        self.setStyleSheet("background-color:lightblue")
        self.backend = WorkOrderAppBackend()
        self.settings = QSettings(COMPANY_NAME, APP_NAME)
        self.initUI()

    def initUI(self):

        self.set_path_btn = QPushButton('Set Download Path',self)
        self.set_path_btn.clicked.connect(partial(upload_file, self, 'Download Path', 'None'))
        self.set_path_btn.setGeometry(30, 20, 200, 60)
        set_button_style(self.set_path_btn)

        self.csv_btn = QPushButton('Upload CSV File',self)
        self.csv_btn.clicked.connect(partial(upload_file, self, 'CSV file', 'CSV Files (*.csv)'))
        self.csv_btn.setGeometry(300, 20, 200, 60)
        set_button_style(self.csv_btn)
        
        self.foaming_template_btn = QPushButton('Upload Foaming',self)
        self.foaming_template_btn.clicked.connect(partial(upload_file, self, 'Foaming Template', 'Word Files (*.docx)'))
        self.foaming_template_btn.setGeometry(570, 20, 200, 60)  
        set_button_style(self.foaming_template_btn)

        self.carpenter_template_btn = QPushButton('Upload Carpenter', self)
        self.carpenter_template_btn.clicked.connect(partial(upload_file, self, 'Carpenter Template', 'Word Files (*.docx)'))
        self.carpenter_template_btn.setGeometry(30, 110, 200, 60)
        set_button_style(self.carpenter_template_btn)

        self.sales_template_btn = QPushButton('Upload Sales', self)
        self.sales_template_btn.clicked.connect(partial(upload_file, self, 'Sales Template', 'Word Files (*.docx)'))
        self.sales_template_btn.setGeometry(300, 110, 200, 60)
        set_button_style(self.sales_template_btn)

        self.database_btn = QPushButton('Upload Database', self)
        self.database_btn.clicked.connect(partial(upload_file, self, 'Database file', 'Excel Files (*.xlsx)'))
        self.database_btn.setGeometry(570, 110, 200, 60)
        set_button_style(self.database_btn)
        
        self.generate_btn = QPushButton('Generate Work Order',self)
        self.generate_btn.clicked.connect(self.generate_work_order)
        self.generate_btn.setGeometry(555, 200, 230, 60)
        set_button_style(self.generate_btn)
        self.generate_btn.setEnabled(False)
        
        self.status_label = QLabel('',self)
        self.status_label.setGeometry(20, 280, 400, 50)
        self.status_label.setStyleSheet("Qlabel { font-size: 15px; }")

        self.order_no_input = QLineEdit(self)
        self.order_no_input.setPlaceholderText("Enter Starting Order Number")
        self.order_no_input.setGeometry(30, 200, 200, 30)
        self.order_no_input.setStyleSheet("background-color: white;")

        self.sales_no_input = QLineEdit(self)
        self.sales_no_input.setPlaceholderText("Enter Starting Sales Summary Number")
        self.sales_no_input.setGeometry(30, 232, 230, 30)
        self.sales_no_input.setStyleSheet("background-color: white;")

        self.save_btn = QPushButton('Save',self)
        self.save_btn.clicked.connect(self.set_numbers)
        self.save_btn.setGeometry(300, 200, 80, 40)
        set_button_style(self.save_btn)

        self.setWindowTitle('Work Order Generator')
        self.setWindowIcon(QIcon('app_icon.ico'))  

        center_widget(self)  

    def set_numbers(self):
        try:
            if not self.order_no_input.text() or not self.sales_no_input.text():
                raise ValueError("Input field cannot be empty.")
            starting_order_no = int(self.order_no_input.text()) 
            starting_sales_no = int(self.sales_no_input.text()) 
            if starting_order_no <= 0 or starting_sales_no<=0:
                raise ValueError("Input values must be a positive integer")
            self.backend.current_order_no = starting_order_no  
            self.backend.current_sales_no = starting_sales_no
            self.status_label.setText("Order number and sales number saved.")
            self.generate_btn.setEnabled(True)
        except ValueError:
            self.status_label.setText("Invalid input value. Please enter a positive integer.")
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
        elif not self.backend.database_path:
            self.status_label.setText('Please upload the Database file')
            return
        
        if not self.generate_btn.isEnabled():
            self.status_label.setText("Invalid input value. Please enter a positive integer.")
            return
        
        self.backend.generate_work_order()
        self.status_label.setText('Work Order Generated')