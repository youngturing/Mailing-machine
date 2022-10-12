from PyQt5.QtWidgets import QDialog, QMessageBox,QMainWindow, QInputDialog
from layout.outlook_emails_sending_info import *
from outlook import *

class OutlookSendingInfo(QDialog, QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Dialog_Sending()
        self.ui.setupUi(self)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = OutlookSendingInfo()
    w.show()
    sys.exit(app.exec_())