from PyQt5.QtWidgets import QDialog, QMessageBox,QMainWindow, QInputDialog
from layout.outlook_window_dialog_confirmation import *
from outlook import *

class OutlookConfirmationDialog(QDialog, QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Dialog()
        self.ui.setupUi(self)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = OutlookConfirmationDialog()
    w.show()
    sys.exit(app.exec_())