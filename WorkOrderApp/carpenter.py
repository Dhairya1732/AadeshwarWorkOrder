from docx import Document
import os
from docx2pdf import convert 
from utils import set_run_font

class CarpenterWorkOrder:
    def create_work_order(self, order_data, template_path, output_path):
        try:
            doc = Document(template_path)
        except Exception as e:
            print(f"Error opening template: {e}")
            return
        
        try:
            # Process the document with order data
            doc.save(output_path)
            print(f"Carpenter work order saved to {output_path}")
        except Exception as e:
            print(f"Error saving document: {e}")

    def convert_to_pdf(self, docx_path, pdf_path):
        try:
            convert(docx_path, pdf_path)
            print(f"Converted to PDF: {pdf_path}")
        except Exception as e:
            print(f"Failed to convert to PDF: {e}")