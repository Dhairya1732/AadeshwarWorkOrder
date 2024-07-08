from PyQt5.QtWidgets import QDesktopWidget, QFileDialog
from PyQt5.QtCore import QDir
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def center_widget(widget):
    qr = widget.frameGeometry()
    cp = QDesktopWidget().availableGeometry().center()
    qr.moveCenter(cp)
    widget.move(qr.topLeft())

def set_button_style(button):
    button.setStyleSheet('QPushButton { padding: 10px; font-size: 20px; background-color: white; }')

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