from PyQt5.QtWidgets import QFileDialog, QPushButton, QLineEdit
from PyQt5.QtCore import QDir
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def create_button(text):
    button = QPushButton(text)
    button.setStyleSheet("""QPushButton {
                                background-color: #4CAF50;
                                color: white;
                                font-size: 18px;
                                border-radius: 8px;
                                padding: 10px 20px;
                                        }
                            QPushButton:hover {
                                background-color: #45a049;
                                                }""")
    return button

def create_line_edit(text):
    line_edit = QLineEdit()
    line_edit.setPlaceholderText(text)
    line_edit.setMaximumWidth(350)
    line_edit.setStyleSheet("""QLineEdit {
        background-color: #f0f0f0;
        border-radius: 5px;
        padding: 10px;
        font-size: 16px;
    }
    QLineEdit:focus {
        background-color: #e8f5e9;
    }""")
    return line_edit

def set_run_font(run, font_size, font_name='Times New Roman'):
    run.font.name = font_name
    run.font.size = Pt(font_size)

def set_cell_border(cell, border_color="000000", border_size="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    
    for border_name in ["top", "left", "bottom", "right"]:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), border_size)
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), border_color)
        tcPr.append(border)

def upload_file(parent, label, file_type):
    options = QFileDialog.Options()
    if label == 'Download Path':
        options |= QFileDialog.DontUseNativeDialog
        path = QFileDialog.getExistingDirectory(parent, f"Select {label}", QDir.homePath(), options=options)
    else:
        options |= QFileDialog.ReadOnly
        path, _ = QFileDialog.getOpenFileName(parent, f'Upload {label}', '', file_type, options=options)
    if path:
        parent.backend.set_file_path(label, path)
        parent.status_label.setText(f'{label} Uploaded')