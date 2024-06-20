import os
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import QSettings, QDir
from work_order_backend import WorkOrderAppBackend
from utils import center_widget, set_button_font_size

COMPANY_NAME = "Aadeshwar"
APP_NAME = "WorkOrderGenerator"

class WorkOrderAppUI(QWidget):
    def __init__(self):
        super().__init__()
        self.backend = WorkOrderAppBackend()
        self.settings = QSettings(COMPANY_NAME, APP_NAME)
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        
        self.set_path_btn = QPushButton('Set Download Path')
        self.set_path_btn.clicked.connect(self.set_download_path)
        set_button_font_size(self.set_path_btn, 12)
        self.set_path_btn.setStyleSheet('QPushButton { padding: 15px; }')
        layout.addWidget(self.set_path_btn)

        self.excel_btn = QPushButton('Upload Excel File')
        self.excel_btn.clicked.connect(self.upload_excel)
        set_button_font_size(self.excel_btn, 12)
        self.excel_btn.setStyleSheet('QPushButton { padding: 15px; }') 
        layout.addWidget(self.excel_btn)
        
        self.template_btn = QPushButton('Upload Word Template')
        self.template_btn.clicked.connect(self.upload_template)
        set_button_font_size(self.template_btn, 12)
        self.template_btn.setStyleSheet('QPushButton { padding: 15px; }') 
        layout.addWidget(self.template_btn)
        
        self.generate_btn = QPushButton('Generate Work Order')
        self.generate_btn.clicked.connect(self.generate_work_order)
        set_button_font_size(self.generate_btn, 12)
        self.generate_btn.setStyleSheet('QPushButton { padding: 15px; }') 
        layout.addWidget(self.generate_btn)
        
        self.status_label = QLabel('')
        layout.addWidget(self.status_label)
        
        self.setLayout(layout)
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
    
    def generate_work_order(self):
        if not self.backend.excel_path or not self.backend.template_path:
            self.status_label.setText('Please upload both files')
            return
        
        self.backend.generate_work_order()
        self.status_label.setText('Work Order Generated')