import re
import sys
import traceback
import time
from typing import Any, List
from enum import Enum

import pandas as pd
import win32com.client as win32
from PyQt5.QtWidgets import QFileDialog, QApplication, QTableWidgetItem, QLabel

from layout.outlook_window import Ui_MainWindow
from outlook_dialog_confirmation import *
from outlook_emails_sending_info import *
from separator import *


class OutlookForm(QMainWindow):
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
        self.ui.push_button_test_send.clicked.connect(self.test_send)
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

    @property
    def find_matching_patterns_from_text(self) -> Any:
        sequence = r'<<(.*?)>>'
        pattern = re.compile(pattern=sequence)
        email_body = self.ui.text_edit_email_body.toPlainText()
        variables = re.findall(pattern, email_body)
        return variables

    @staticmethod
    def find_sending_account() -> Any:
        send_account = None
        outlook = win32.Dispatch('Outlook.Application')
        for account in outlook.Session.Accounts:
            if account.DisplayName == 'bblab@uj.edu.pl':
                send_account = account
                break
        return send_account, outlook

    def open_confirmation_dialog(self) -> None:
        self.confirmation_dialog.show()

    def open_sending_emails_info_dialog(self) -> None:
        self.sending_email_dialog.show()
        self.sending_email_dialog.ui.text_edit_mail_info.clear()

    def open_separator_dialog(self) -> None:
        self.separator_dialog.show()

    def change_separator(self) -> None:
        new_separator = self.separator_dialog.ui.line_edit_separator.text()
        self.separator = new_separator
        self.separator_dialog.close()

    def cancel_changing_separator(self) -> None:
        self.separator_dialog.close()

    def cancel_sending_email(self):
        self.confirmation_dialog.close()

    def copy_selected_value_from_list_of_variables(self) -> None:
        try:
            item = self.ui.list_selected_variables.currentItem().text()
            QApplication.clipboard().setText(item)
        except:
            QMessageBox.critical(self, 'Error', f'Something went wrong: {traceback.format_exc()}')

    def copy_addresses(self, item) -> None:
        try:
            addresses_column_name = self.get_clicked_item_from_list(item)
            self.ui.line_edit_addresses.setText(addresses_column_name)
        except:
            QMessageBox.critical(self, 'Error', f'Something went wrong: {traceback.format_exc()}')

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
                self.ui.table_widget_data_from_data_frame.setHorizontalHeaderLabels(self.data.columns)
                for column, key in enumerate(self.data.columns):
                    for row, item in enumerate(self.data[key]):
                        new_item = QTableWidgetItem(item)
                        self.ui.table_widget_data_from_data_frame.setItem(row, column, new_item)
                self.load_columns_to_list_of_variables()
                QMessageBox.information(self, 'Info', 'Database successfully loaded!')
        except Exception:
            QMessageBox.critical(self, 'Error', f'Something went wrong: {traceback.format_exc()}')

    def clean_data_from_data_frame(self) -> None:
        self.data = self.data.dropna(axis=1)
        self.data.columns = self.data.columns.str.strip()
        for column, data_type in zip(self.data.columns, self.data.dtypes):
            if data_type in ['object', 'str']:
                self.data[column] = self.data[column].str.strip()
            elif data_type in ['int64', 'float64', 'int32', 'float32']:
                self.data[column] = self.data[column].astype('str')

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

    def get_variables_from_list(self) -> List[str]:
        if self.ui.list_selected_variables.count() > 0:
            variables_from_list = []
            for i in range(self.ui.list_selected_variables.count()):
                variables_from_list.append(self.ui.list_selected_variables.item(i).text())
            return variables_from_list

    def get_data_from_dataframe(self) -> pd.DataFrame:
        if isinstance(self.data, pd.DataFrame):
            variables_from_list = self.get_variables_from_list()
            columns_to_slice_from_df = []
            for variable in variables_from_list:
                if variable in self.data.columns:
                    columns_to_slice_from_df.append(variable)
            sliced_df = self.data[columns_to_slice_from_df]
            return sliced_df

    def create_list_of_mails_messages(self) -> List[str]:
        list_of_mails = []
        sliced_data_frame = self.get_data_from_dataframe()
        variables_from_list = self.find_matching_patterns_from_text
        email_body = self.ui.text_edit_email_body.toPlainText()
        email_body_dict = {'Body': email_body}
        for row in range(len(sliced_data_frame)):
            for num, variable in enumerate(variables_from_list):
                if variable in sliced_data_frame.columns:
                    email_body_dict['Body'] = email_body_dict['Body'] \
                        .replace(f'<<{variable}>>', str(sliced_data_frame[variable][row]))
                    if num + 1 == len(variables_from_list):
                        list_of_mails.append(email_body_dict['Body'])
                        email_body_dict['Body'] = email_body
        return list_of_mails

    def get_email_addresses(self) -> List[str]:
        column_with_addresses = self.ui.line_edit_addresses.text()
        list_of_addresses = self.data[column_with_addresses].to_list()
        return list_of_addresses

    def compose_sending_operation(self, sending_type: str) -> Any:
        list_of_emails = self.create_list_of_mails_messages()
        list_of_addresses = self.get_email_addresses()
        self.confirmation_dialog.close()
        self.open_sending_emails_info_dialog()
        if sending_type == 'test_send':
            return list_of_emails, list_of_addresses
        elif sending_type == 'normal_send':
            send_account, outlook = self.find_sending_account()
            return list_of_emails, list_of_addresses, send_account, outlook

    def test_send(self):
        try:
            list_of_emails, list_of_addresses = self.compose_sending_operation(sending_type=SendingType.TEST_SEND.value)
            for address, mail in zip(list_of_addresses, list_of_emails):
                self.sending_email_dialog.ui.text_edit_mail_info.insertPlainText(f'Email send to: {address}\n'
                                                                                 f'Body:\n'
                                                                                 f'{mail}\n'
                                                                                 f'{"=" * 60}\n')
        except:
            QMessageBox.critical(self, 'Error', f'No data: \n{traceback.format_exc()}')
        finally:
            self.confirmation_dialog.close()

    def send_email(self):
        list_of_emails, list_of_addresses, send_account, outlook = self.compose_sending_operation(
            sending_type=SendingType.NORMAL_SEND.value)
        mail_subject = self.ui.line_edit_subject.text()
        try:
            for address, mail in zip(list_of_addresses, list_of_emails):
                mail_object = outlook.CreateItem(0)
                mail_object.To = address
                mail_object.Subject = mail_subject
                mail_object.Body = mail
                mail_object._oleobj_.Invoke(*(64209, 0, 8, 0, send_account))
                mail_object.Send()  # Sending emails to to the list of users
                self.sending_email_dialog.ui.text_edit_mail_info.insertPlainText(f'Email send to: {address}\n'
                                                                                 f'Body:\n'
                                                                                 f'{mail}\n'
                                                                                 f'{"=" * 60}\n')
                time.sleep(3)
        except:
            QMessageBox.critical(self, 'Error', f'No data: \n{traceback.format_exc()}')
        finally:
            self.confirmation_dialog.close()


class SendingType(Enum):
    TEST_SEND = 'test_send'
    NORMAL_SEND = 'normal_send'


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = OutlookForm()
    w.show()
    sys.exit(app.exec_())
