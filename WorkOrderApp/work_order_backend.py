import pandas as pd
from PyQt5.QtCore import QSettings, QDir
from datetime import timedelta, datetime
from collections import defaultdict

COMPANY_NAME = "Aadeshwar"
APP_NAME = "WorkOrderGenerator"

class WorkOrderAppBackend:
    def __init__(self):
        self.csv_path = ''
        self.foaming_template_path = ''
        self.carpenter_template_path = ''
        self.sales_template_path = ''
        self.download_path = QDir.homePath()  
        self.settings = QSettings(COMPANY_NAME, APP_NAME)
        self.load_settings()

    def load_settings(self):
        self.download_path = self.settings.value("download_path", QDir.homePath())
        self.foaming_template_path = self.settings.value("foaming_template_path", '')
        self.carpenter_template_path = self.settings.value("carpenter_template_path", '')
        self.sales_template_path = self.settings.value("sales_template_path", '')
        self.database_path = self.settings.value("database_path", '')

    def set_csv_path(self, path):
        self.csv_path = path
    
    def set_foaming_template_path(self, path):
        self.foaming_template_path = path
        self.settings.setValue("foaming_template_path", self.foaming_template_path)

    def set_carpenter_template_path(self, path):
        self.carpenter_template_path = path
        self.settings.setValue("carpenter_template_path", self.carpenter_template_path)

    def set_sales_template_path(self, path):
        self.sales_template_path = path
        self.settings.setValue("sales_template_path", self.sales_template_path)

    def set_database_path(self, path):
        self.database_path = path
        self.settings.setValue("database_path", self.database_path)

    def set_download_path(self, path):
        self.download_path = path
        self.settings.setValue("download_path", self.download_path)

    def generate_work_order(self):
        orders = pd.read_csv(self.csv_path)
        fabric_sheet = pd.read_excel(self.database_path ,sheet_name="Fabric")
        orders = orders.head(3)
        columns_to_ignore = ['Unit Price', 'TOTAL', 'Shipping Address', 'status', 'Promised Delivery Date']
        sales_summary_data = defaultdict(list)

        relevant_columns = ['SKU_ID', 'Legs', 'Legs Quantity', 'Cushion Qty', 'Cushion Fabric', 'Sofa Fabric', 'Legs Finish', 'Legs Assembly', 'Cushions', 'Cushion Size', 'Dimensions']
        fabric_sheet = fabric_sheet[relevant_columns]
        fabric_sheet = fabric_sheet.fillna("") 
        
        for index, row in orders.iterrows():
            order_data = row.to_dict()
            for column in columns_to_ignore:
                order_data.pop(column, None)

            if 'Order Confirmed Date' in order_data:
                order_data['Order Confirmed Date'] = pd.to_datetime(order_data['Order Confirmed Date'],dayfirst=True).date()

            if 'To be shipped Before' in order_data:
                delivery_date = pd.to_datetime(order_data['To be shipped Before'],dayfirst=True)
                if pd.notnull(delivery_date):
                    adjusted_date = (delivery_date - timedelta(days=2)).date()
                    order_data['To be shipped Before'] = adjusted_date.strftime(f"%d-%m-%Y")

            current_month = datetime.now().strftime("%B")
            order_no = f"G1/{current_month}/{self.current_order_no}"
            order_data['OrderNo'] = order_no

            sku_id = order_data['SKU ID']
            if fabric_sheet['SKU_ID'].isin([sku_id]).any():
                additional_data = fabric_sheet[fabric_sheet['SKU_ID'] == sku_id].to_dict('records')[0]
            else:
                additional_data= {col: "" for col in relevant_columns if col!="SKU_ID"}

            order_data.update(additional_data)

            from foaming import FoamingWorkOrder
            foaming = FoamingWorkOrder()
            foaming.create_work_order(order_data, self.foaming_template_path, f"Foaming_{index+1}.pdf")
            #self.create_work_order(order_data, self.carpenter_template_path, f"carpenter_{index+1}.pdf")
            sales_summary_data[order_data['To be shipped Before']].append(order_data)
            self.current_order_no += 1

        from sales import SalesSummary
        for shipping_date, orders_data in sales_summary_data.items():
            sales = SalesSummary()
            sales.create_sales_summary(orders_data, self.sales_template_path, f"sales_{shipping_date}.pdf")