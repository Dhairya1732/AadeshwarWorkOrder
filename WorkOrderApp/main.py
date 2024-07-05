import sys
from PyQt5.QtWidgets import QApplication
from work_order_ui import WorkOrderAppUI

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = WorkOrderAppUI()
    ex.show()
    sys.exit(app.exec_())