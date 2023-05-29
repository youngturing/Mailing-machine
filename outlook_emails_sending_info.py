import sys

from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QDialog, QMainWindow
from layout.outlook_emails_sending_info import SendingDialogUI


class OutlookSendingInfo(QDialog, QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = SendingDialogUI()
        self.ui.setupUi(self)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = OutlookSendingInfo()
    w.show()
    sys.exit(app.exec_())
