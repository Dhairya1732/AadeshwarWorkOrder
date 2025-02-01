from PyQt5.QtWidgets import QMainWindow, QLabel, QApplication, QWidget, QVBoxLayout, QHBoxLayout
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import QSettings, Qt
from work_order_backend import WorkOrderAppBackend
from utils import create_button, upload_file, create_line_edit
from config import COMPANY_NAME, APP_NAME
import sys

class WorkOrderAppUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.backend = WorkOrderAppBackend()
        self.settings = QSettings(COMPANY_NAME, APP_NAME)
        self.initUI()

    def initUI(self):
        self.setWindowState(Qt.WindowState.WindowMaximized)
        self.setWindowTitle("Aadeshwar Work Order Generator")
        self.setWindowIcon(QIcon(r"assets\images\card_image.png"))
        self.setStyleSheet("background-color: #2c3e50; color: #ecf0f1;")

        # Container Widget
        container = QWidget()
        self.setCentralWidget(container)

        # Main Vertical Layout
        main_layout = QVBoxLayout()
        container.setLayout(main_layout)

        # Title Label
        title_label = QLabel("Aadeshwar Work Order Generator")
        title_label.setFont(QFont("Arial", 26, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: #ecf0f1; margin-bottom: 20px;")
        main_layout.addWidget(title_label)

        # Grid Layout for Buttons
        button_layout = QVBoxLayout()
        button_layout.setSpacing(20)
        main_layout.addLayout(button_layout)

        # Row 1: Upload Templates
        upload_layout = QHBoxLayout()
        upload_layout.setSpacing(30)
        button_layout.addLayout(upload_layout)
        button_layout.setSpacing(50)

        # Row 2: Set Paths and Upload Files
        path_layout = QHBoxLayout()
        path_layout.setSpacing(30)
        button_layout.addLayout(path_layout)
        button_layout.setSpacing(50)

        # Row 3: Inputs and Generate Button
        input_generate_layout = QHBoxLayout()
        input_generate_layout.setSpacing(20)
        button_layout.addLayout(input_generate_layout)

        # Button to upload foaming template
        foaming_btn = create_button('Upload Foaming Template')
        foaming_btn.clicked.connect(lambda: upload_file(self, 'Foaming Template', 'Word Files (*.docx)'))
        upload_layout.addWidget(foaming_btn)

        # Button to upload carpenter template
        carpenter_btn = create_button('Upload Carpenter Template')
        carpenter_btn.clicked.connect(lambda: upload_file(self, 'Carpenter Template', 'Word Files (*.docx)'))
        upload_layout.addWidget(carpenter_btn)

        # Button to upload sales template
        sales_btn = create_button('Upload Sales Template')
        sales_btn.clicked.connect(lambda: upload_file(self, 'Sales Template', 'Word Files (*.docx)'))
        upload_layout.addWidget(sales_btn)

        # Button to set download path
        path_btn = create_button('Set Download Path')
        path_btn.clicked.connect(lambda: upload_file(self, 'Download Path', 'None'))
        path_layout.addWidget(path_btn)

        # Button to upload pending orders
        csv_btn = create_button('Upload Pending Orders')
        csv_btn.clicked.connect(lambda: upload_file(self, 'Pending Orders', 'CSV Files (*.csv)'))
        path_layout.addWidget(csv_btn)

        # Button to upload database file
        database_btn = create_button('Upload Database')
        database_btn.clicked.connect(lambda: upload_file(self, 'Database file', 'Excel Files (*.xlsx)'))
        path_layout.addWidget(database_btn)

        # Layout for inputs
        input_fields_layout = QVBoxLayout()
        input_fields_layout.setSpacing(10)
        input_generate_layout.addLayout(input_fields_layout)

        # Order no. input
        self.order_no_input = create_line_edit("Enter Starting Order Number")
        input_fields_layout.addWidget(self.order_no_input)

        # Sales no. input
        self.sales_no_input = create_line_edit("Enter Starting Sales Summary Number")
        input_fields_layout.addWidget(self.sales_no_input)

        # Button to order no. and sales no.
        save_btn = create_button('Save')
        save_btn.clicked.connect(lambda: self.set_numbers()) 
        save_btn.setMinimumWidth(200) 
        input_generate_layout.addWidget(save_btn)

        # Button to generate work orders
        self.generate_btn = create_button('Generate Work Orders')
        self.generate_btn.clicked.connect(lambda: self.generate_work_order())
        self.generate_btn.setEnabled(False)
        input_generate_layout.addWidget(self.generate_btn)

        # Label display various messages
        self.status_label = QLabel('')
        self.status_label.setFont(QFont("Arial", 16))
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("font-family: 'Roboto', sans-serif; font-size: 28px; color: #ecf0f1; margin-top: 20px;")
        main_layout.addWidget(self.status_label)

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
            self.status_label.setText('Please upload the Pending Orders')
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
        
        self.status_label.setText('Generating Work Orders. Please Wait...')
        QApplication.processEvents()
        self.backend.generate_work_order()
        self.status_label.setText('Work Order Generated')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = WorkOrderAppUI()
    ex.show()
    sys.exit(app.exec_())