import sys

from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QDialog,QMainWindow
from layout.separator import DialogSeparatorUI


class OutlookSeparator(QDialog, QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = DialogSeparatorUI()
        self.ui.setupUi(self)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = OutlookSeparator()
    w.show()
    sys.exit(app.exec_())
