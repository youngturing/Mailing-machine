from PyQt5.QtWidgets import QDialog, QMessageBox,QMainWindow, QInputDialog
from layout.separator import *
from outlook import *

class OutlookSeparator(QDialog, QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Dialog_Separator()
        self.ui.setupUi(self)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = OutlookSeparator()
    w.show()
    sys.exit(app.exec_())