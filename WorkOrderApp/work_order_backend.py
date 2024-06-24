from docx import Document
from docx.shared import Pt
import pandas as pd
import requests
from io import BytesIO
from PyQt5.QtCore import QSettings, QDir
from utils import set_run_font
from docx2pdf import convert 
import os
from datetime import timedelta, datetime

COMPANY_NAME = "Aadeshwar"
APP_NAME = "WorkOrderGenerator"

class WorkOrderAppBackend:
    def __init__(self):
        self.excel_path = ''
        self.template_path = ''
        self.download_path = QDir.homePath()  
        self.settings = QSettings(COMPANY_NAME, APP_NAME)
        self.load_settings()

    def load_settings(self):
        self.download_path = self.settings.value("download_path", QDir.homePath())
        self.template_path = self.settings.value("template_path", '')

    def save_settings(self):
        self.settings.setValue("download_path", self.download_path)

    def set_excel_path(self, path):
        self.excel_path = path
    
    def set_template_path(self, path):
        self.template_path = path
        self.settings.setValue("template_path", self.template_path)

    def set_download_path(self, path):
        self.download_path = path
        self.save_settings()

    def generate_work_order(self):
        orders = pd.read_excel(self.excel_path)
        orders = orders.head(1)
        columns_to_ignore = ['Unit Price', 'TOTAL', 'SKU ID', 'Shipping Address', 'status', 'Promised Delivery Date']
        
        for index, row in orders.iterrows():
            order_data = row.to_dict()
            for column in columns_to_ignore:
                order_data.pop(column, None)

            if 'Order Confirmed Date' in order_data:
                order_data['Order Confirmed Date'] = order_data['Order Confirmed Date'].date()

            if 'To be shipped Before' in order_data:
                delivery_date = order_data['To be shipped Before']
                if pd.notnull(delivery_date):
                    adjusted_date = (delivery_date - timedelta(days=2)).date()
                    order_data['To be shipped Before'] = adjusted_date

            current_month = datetime.now().strftime("%B")
            order_no = f"G1/{current_month}/{self.current_order_no}"
            order_data['OrderNo'] = order_no
            
            docx_output_path = os.path.join(self.download_path, f"Work_Order_{index + 1}.docx")
            pdf_output_path = os.path.join(self.download_path, f"Work_Order_{index + 1}.pdf")
            
            self.process_work_order(order_data, self.template_path, docx_output_path, image_width=Pt(100), image_height=Pt(100))
            convert(docx_output_path, pdf_output_path)
            os.remove(docx_output_path)
            self.current_order_no += 1
    
    def process_work_order(self, data, template_path, output_path, image_width=None, image_height=None):
        doc = Document(template_path)
        for key, value in data.items():
            if isinstance(value, str):
                value = value.strip()
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            if f'[{key}]' in paragraph.text:
                                if key == 'Image url':
                                    paragraph.text = paragraph.text.replace(f'[{key}]', '')
                                    self.insert_image_from_url(cell, value, width=image_width, height=image_height)
                                else:
                                    paragraph.text = paragraph.text.replace(f'[{key}]', str(value))
                                    for run in paragraph.runs:
                                        set_run_font(run)
        doc.save(output_path)
    
    def insert_image_from_url(self, cell, image_url, width=None, height=None):
        response = requests.get(image_url)
        image_stream = BytesIO(response.content)
        paragraph = cell.paragraphs[0]
        run = paragraph.add_run()
        run.add_picture(image_stream, width=width, height=height)