from PyQt5.QtWidgets import QDesktopWidget
from docx.shared import Pt

def center_widget(widget):
    qr = widget.frameGeometry()
    cp = QDesktopWidget().availableGeometry().center()
    qr.moveCenter(cp)
    widget.move(qr.topLeft())

def set_button_style(button):
    button.setStyleSheet('QPushButton { padding: 10px; font-size: 20px; background-color: white; }')

def set_run_font(run, font_name='Times New Roman', font_size=9):
    run.font.name = font_name
    run.font.size = Pt(font_size)