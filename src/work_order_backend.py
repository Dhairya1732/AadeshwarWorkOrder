import pandas as pd
from PyQt5.QtCore import QSettings, QDir
from datetime import timedelta, datetime
from collections import defaultdict
from config import COMPANY_NAME, APP_NAME
from foaming import FoamingWorkOrder
from carpenter import CarpenterWorkOrder
from sales import SalesSummary

class WorkOrderAppBackend:
    def __init__(self):
        # Default Values for the variables
        self.csv_path = ''
        self.foaming_template_path = ''
        self.carpenter_template_path = ''
        self.sales_template_path = ''
        self.database_path = ''
        self.download_path = QDir.homePath()  
        self.current_order_no = 1
        self.current_sales_no = 1
        self.settings = QSettings(COMPANY_NAME, APP_NAME)
        self.load_settings()

    def load_settings(self):
        # Load all the previously saved files 
        self.download_path = self.settings.value("download_path", QDir.homePath())
        self.foaming_template_path = self.settings.value("foaming_template_path", '')
        self.carpenter_template_path = self.settings.value("carpenter_template_path", '')
        self.sales_template_path = self.settings.value("sales_template_path", '')
        self.database_path = self.settings.value("database_path", '')

    def generate_work_order(self):
        # Load the CSV and Excel files once at the start
        orders = pd.read_csv(self.csv_path).fillna("")
        fabric_sheet = pd.read_excel(self.database_path ,sheet_name="Fabric").fillna("")

        # Create a dictionary for fast lookup of fabric sheet by SKU_ID
        fabric_dict = fabric_sheet.set_index('SKU_ID').to_dict(orient='index')

        orders = orders.head(5)

        # Columns to ignore
        columns_to_ignore = ['Unit Price', 'TOTAL', 'Shipping Address', 'status', 'Promised Delivery Date' , 'Product_Name' , 'Merchant_SKU_ID']
        
        # Data containers
        sales_summary_data = defaultdict(list)
        carpenter_work_orders = defaultdict(list)
         
        # Loop through orders and process
        for index, row in orders.iterrows():
            order_data = row.to_dict()

            # Get matching fabric data using fast dictionary lookup
            sku_id = order_data.get('SKU ID')
            if sku_id in fabric_dict:
                fabric_data = fabric_dict[sku_id]
                for key, value in fabric_data.items():
                    if order_data.get(key, "") == "":
                        order_data[key] = value

            # Remove ignored columns
            for column in columns_to_ignore:
                order_data.pop(column, None)

            # Convert dates
            if 'Order Confirmed Date' in order_data:
                order_data['Order Confirmed Date'] = pd.to_datetime(order_data['Order Confirmed Date'],dayfirst=True).date()

            if 'To be shipped Before' in order_data:
                delivery_date = pd.to_datetime(order_data['To be shipped Before'],dayfirst=True)
                if pd.notnull(delivery_date):
                    adjusted_date = (delivery_date - timedelta(days=2)).date()
                    order_data['To be shipped Before'] = adjusted_date.strftime(f"%d-%m-%Y")

            # Generate order number
            current_month = datetime.now().strftime("%B")
            order_no = f"G1/{current_month}/{self.current_order_no}"
            order_data['OrderNo'] = order_no

            # Calculate TotalInches
            foaming_inches = str(order_data.get('Foaming Inches', '')).strip()
            qty = order_data.get('QTY', '')
            try:
                order_data['TotalInches'] = str(int(foaming_inches) * int(qty)) if foaming_inches and qty else '0'
            except ValueError:
                order_data['TotalInches'] = '0'

            # Process foaming
            self.foaming = FoamingWorkOrder(self)
            self.foaming.create_work_order(order_data, self.foaming_template_path, f"Foaming_{index+1}.pdf")
            
            # Store orders for sales and carpenter
            sales_summary_data[order_data['To be shipped Before']].append(order_data)
            carpenter_work_orders[order_data['To be shipped Before']].append(order_data)
            self.current_order_no += 1

        for shipping_date,orders_data in carpenter_work_orders.items():
            self.carpenter = CarpenterWorkOrder(self)
            self.carpenter.create_carpenter_order(orders_data, self.carpenter_template_path, f"Carpenter_{shipping_date}.pdf")

        for shipping_date, orders_data in sales_summary_data.items():
            self.sales = SalesSummary(self)
            self.sales.create_sales_summary(orders_data, self.sales_template_path, f"Sales_{shipping_date}.pdf")

    def set_file_path(self, label, path):
        match label:
            case 'Pending Orders':
                self.csv_path = path
            case 'Foaming Template':
                self.foaming_template_path = path
                self.settings.setValue("foaming_template_path", self.foaming_template_path)
            case 'Carpenter Template':
                self.carpenter_template_path = path
                self.settings.setValue("carpenter_template_path", self.carpenter_template_path)
            case 'Sales Template':
                self.sales_template_path = path
                self.settings.setValue("sales_template_path", self.sales_template_path)
            case 'Database file':
                self.database_path = path
                self.settings.setValue("database_path", self.database_path)
            case 'Download Path':
                self.download_path = path
                self.settings.setValue("download_path", self.download_path)