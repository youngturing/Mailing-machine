import re
import sys
import traceback
from typing import Any

import pandas as pd
import win32com.client as win32
from PyQt5.QtWidgets import QFileDialog, QTableWidget, QApplication, QTableWidgetItem

from layout.outlook_window import Ui_MainWindow
from outlook_dialog_confirmation import *
from outlook_emails_sending_info import *
from separator import *


class OutlookForm(QMainWindow, QTableWidget):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.actionOpen_data_file.triggered.connect(self.load_data)
        self.ui.list_widget_columns.itemClicked.connect(self.get_clicked_item_from_list)
        self.ui.list_selected_variables.itemClicked.connect(self.get_clicked_item_from_list_of_variables)
        self.ui.push_button_add_variable.clicked.connect(self.add_data_to_listed_variables)
        self.ui.push_button_remove_variable.clicked.connect(self.remove_item_from_selected_variables)
        self.ui.push_button_clean_list.clicked.connect(self.clear_list_of_selected_items)
        self.ui.push_button_send.clicked.connect(self.open_confirmation_dialog)
        # self.ui.push_button_test_send.clicked.connect(self.send_email)
        self.ui.push_button_refresh.clicked.connect(self.load_columns_to_list_of_variables)
        self.ui.push_button_change_separator.clicked.connect(self.open_separator_dialog)
        self.ui.push_button_copy_selected.clicked.connect(self.copy_selected_value_from_list_of_variables)
        self.ui.push_button_copy_addresses.clicked.connect(self.copy_addresses)
        """"""
        self.confirmation_dialog = OutlookConfirmationDialog()
        self.sending_email_dialog = OutlookSendingInfo()
        self.separator_dialog = OutlookSeparator()
        self.confirmation_dialog.ui.push_buttton_ok.clicked.connect(self.send_email)
        self.confirmation_dialog.ui.push_button_cancel.clicked.connect(self.cancel_sending_email)
        self.separator_dialog.ui.push_button_ok_sep.clicked.connect(self.change_separator)
        self.separator_dialog.ui.push_button_cancel_sep.clicked.connect(self.cancel_changing_separator)
        """"""
        self.data = None
        self.separator: str = ';'
        self.show()

    def open_confirmation_dialog(self) -> None:
        self.confirmation_dialog.show()

    def open_sending_emails_info_dialog(self) -> None:
        self.sending_email_dialog.show()

    def open_separator_dialog(self) -> None:
        self.separator_dialog.show()

    def change_separator(self) -> None:
        new_separator = self.separator_dialog.ui.line_edit_separator.text()
        self.separator = new_separator
        self.separator_dialog.close()

    def cancel_changing_separator(self) -> None:
        self.separator_dialog.close()

    def copy_selected_value_from_list_of_variables(self) -> None:
        item = self.ui.list_selected_variables.currentItem().text()
        QApplication.clipboard().setText(item)

    def copy_addresses(self, item) -> None:
        addresses_column_name = self.get_clicked_item_from_list(item)
        self.ui.line_edit_addresses.setText(addresses_column_name)

    def load_data(self) -> None:
        try:
            file, _ = QFileDialog.getOpenFileName(self, "Open file", "", "All files (*);;CSV files (*.csv)")
            if file:
                if len(self.separator) != 0:
                    self.data = pd.read_csv(file, sep=str(self.separator))
                else:
                    self.data = pd.read_csv(file)
                self.clean_data_from_data_frame()
                self.ui.table_widget_data_from_data_frame.setColumnCount(self.data.shape[1])
                self.ui.table_widget_data_from_data_frame.setRowCount(self.data.shape[0])
                for column, key in enumerate(self.data.columns):
                    for row, item in enumerate(self.data[key]):
                        new_item = QTableWidgetItem(item)
                        self.ui.table_widget_data_from_data_frame.setItem(row, column, new_item)
                self.ui.table_widget_data_from_data_frame.setHorizontalHeaderLabels(self.data.columns)
                self.load_columns_to_list_of_variables()
                QMessageBox.information(self, 'Info', 'Database successfully loaded!')
        except FileNotFoundError:
            QMessageBox.critical(self, 'Error', f'Something went wrong: {traceback.format_exc()}')

    def clean_data_from_data_frame(self) -> None:
        self.data = self.data.dropna(axis=1)
        self.data.columns = self.data.columns.str.rstrip()
        for column, data_type in zip(self.data.columns, self.data.dtypes):
            if data_type == 'object' or data_type == 'str':
                self.data[column] = self.data[column].str.strip()

    def load_columns_to_list_of_variables(self) -> None:
        if isinstance(self.data, pd.DataFrame):
            self.ui.list_widget_columns.clear()
            for column in self.data.columns:
                self.ui.list_widget_columns.addItem(column)

    def get_clicked_item_from_list(self, item) -> Any:
        item_from_list = self.ui.list_widget_columns.currentItem().text()
        return item_from_list

    def get_clicked_item_from_list_of_variables(self, item) -> Any:
        item_from_list = self.ui.list_selected_variables.currentRow()
        return item_from_list

    def add_data_to_listed_variables(self, item) -> None:
        if len(self.ui.list_widget_columns.selectedItems()) > 0:
            item_from_list = self.get_clicked_item_from_list(item)
            self.ui.list_selected_variables.addItem(item_from_list)

    def remove_item_from_selected_variables(self, item) -> None:
        item_from_list = self.get_clicked_item_from_list_of_variables(item)
        self.ui.list_selected_variables.takeItem(item_from_list)

    def clear_list_of_selected_items(self) -> None:
        self.ui.list_selected_variables.clear()

    def get_variables_from_list(self) -> list:
        variables_from_list = []
        for i in range(self.ui.list_selected_variables.count()):
            variables_from_list.append(self.ui.list_selected_variables.item(i).text())
        return variables_from_list

    def get_data_from_dataframe(self) -> pd.DataFrame:
        variables_from_list = self.get_variables_from_list()
        columns_to_slice_from_df = []
        for variable in variables_from_list:
            if variable in self.data.columns:
                columns_to_slice_from_df.append(variable)
        sliced_df = self.data[columns_to_slice_from_df]
        return sliced_df

    @property
    def find_matching_patterns_from_text(self) -> Any:
        sequence = r'<<(.*?)>>'
        pattern = re.compile(pattern=sequence)
        email_body = self.ui.text_edit_email_body.toPlainText()
        variables = re.findall(pattern, email_body)
        positions = pattern.finditer(email_body)
        return variables, positions

    @staticmethod
    def find_sending_account() -> Any:
        send_account = None
        outlook = win32.Dispatch('Outlook.Application')
        for account in outlook.Session.Accounts:
            if account.DisplayName == 'bblab@uj.edu.pl':
                send_account = account
                break
        return send_account, outlook

    def create_list_of_mails_messages(self) -> list:
        list_of_mails = []
        sliced_data_frame = self.get_data_from_dataframe()
        variables_from_list, positions = self.find_matching_patterns_from_text
        email_body = self.ui.text_edit_email_body.toPlainText()
        email_body_list = list(email_body)
        for row in range(len(sliced_data_frame)):
            for num, (position, variable) in enumerate(zip(positions, variables_from_list)):
                if variable in sliced_data_frame.columns:
                    email_body_list[position.start():position.end()] = str(sliced_data_frame[variable][row])
                    email_composed = ''.join(email_body_list)
                    if num + 1 == len(variables_from_list):
                        list_of_mails.append(email_composed)
                        email_body_list = list(email_body)
        return list_of_mails

    def get_email_addresses(self) -> list:
        column_with_addresses = self.ui.line_edit_addresses.text()
        list_of_addresses = self.data[column_with_addresses].to_list()
        return list_of_addresses

    def send_email(self):
        list_of_emails = self.create_list_of_mails_messages()
        list_of_addresses = self.get_email_addresses()

        # send_account, outlook = self.find_sending_account()
        self.confirmation_dialog.close()
        self.open_sending_emails_info_dialog()
        for address in list_of_addresses:
            self.sending_email_dialog.ui.label_email_sending_info.setText(f'Email send to: {address}\n'
                                                                          f'{"="*50}')
        for x in list_of_addresses:
            print(x)

        for x in list_of_emails:
            print(x)
        # for mail, address in zip(list_of_emails, list_of_addresses):
        #     mail_object = outlook.CreateItem(0)
        #     mail_object.To = address
        #     mail_object.Subject = mail_subject
        #     mail_object.Body = mail
        #     mail_object._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))
        #     mail_object.Send()    # Sending emails to to the list of users
        #     print(f'Email send to: {name} at {mail} with project name: {project_name}')
        #     print('=' * 40)
        #     time.sleep(3)

    def cancel_sending_email(self):
        self.confirmation_dialog.close()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = OutlookForm()
    w.show()
    sys.exit(app.exec_())
