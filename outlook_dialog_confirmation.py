import sys

from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QDialog, QMainWindow
from layout.outlook_window_dialog_confirmation import DialogUI


class OutlookConfirmationDialog(QDialog, QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = DialogUI()
        self.ui.setupUi(self)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = OutlookConfirmationDialog()
    w.show()
    sys.exit(app.exec_())
