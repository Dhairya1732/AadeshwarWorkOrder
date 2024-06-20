from PyQt5.QtWidgets import QDesktopWidget
from docx.shared import Pt

def center_widget(widget):
    # Calculate center position based on screen geometry
    qr = widget.frameGeometry()
    cp = QDesktopWidget().availableGeometry().center()
    qr.moveCenter(cp)
    widget.move(qr.topLeft())

def set_button_font_size(button, font_size):
    font = button.font()
    font.setPointSize(font_size)
    button.setFont(font)

def set_run_font(run, font_name='Times New Roman', font_size=9):
    run.font.name = font_name
    run.font.size = Pt(font_size)