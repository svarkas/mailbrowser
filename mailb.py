from __future__ import annotations
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox 
from PyQt5 import QtCore, QtWidgets, QtWebEngineWidgets, uic
from PyQt5.QtGui import QStandardItemModel, QStandardItem, QColor
import sys
import os
import re
from typing import Callable
from typing import cast
from typing import Dict
from typing import List
from typing import Literal
from typing import Optional
from email.utils import parsedate_to_datetime

folders = []

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = uic.loadUi(f"mailb.ui", self)
        self.menuBar().setNativeMenuBar(False)
        self.ui.actionQuit.triggered.connect(QApplication.quit)
        self.email_dir = None
        self.ui.actionSelect_email_folder.triggered.connect(self.set_email_dir)
        self.ui.loadButton.clicked.connect(self.load_root_folder)
        self.folderListView.clicked.connect(self.on_folder_clicked)
        self.ui.messagesTableView.clicked.connect(self.on_message_clicked)

    def quit_clicked(self) -> None:
        sys.exit(app.exec_())
    
    def load_root_folder(self):
        if self.email_dir is None or self.email_dir == '':
            QMessageBox.warning(self, "No e-mail folder Selected", "Please select an e-mail folder first!!!")
        else:
            folders_found = os.listdir(self.email_dir)
            for f in folders_found:
                folder = Folder(f, self.email_dir)
                folders.append(folder)
                self.folder_list()


    def set_email_dir(self):
        email_path = QFileDialog.getExistingDirectory(self, "Choose E-mail Folder")
        if email_path:
            self.email_dir = email_path

    def folder_list(self) -> None:
        self.model = QStandardItemModel()
        self.model.clear()
        self.model.setHorizontalHeaderLabels(['Folder'])
        for f in folders:
            if f.mailfolder:
                item = QStandardItem(f.folder_name)
                self.model.appendRow(item)
        self.folderListView.setModel(self.model)

    def on_folder_clicked(self, index):
        try:
            item = self.model.itemFromIndex(index)
            folder_name_selected = item.text()
            folder = next(f for f in folders if f.folder_name == folder_name_selected)
            self.messages_list(folder)
        except AttributeError as e:
            print("Folder is Empty")

    def on_message_clicked(self, index) -> None:
        row = index.row()
        item_selected = self.model.item(row, 3) # the subject column
        self.messageContent.setText(item_selected.data())

    def messages_list(self, folder):
        self.model = QStandardItemModel()
        self.model.setHorizontalHeaderLabels(['Date Sent', 'From', 'To', 'Subject'])
        for m in folder.messages_list:
            item_date = QStandardItem(m.date)
            item_date.setBackground(QColor("lightgreen"))
            item_sender = QStandardItem(m.sender)
            item_sender.setBackground(QColor("lightblue"))
            item_recipient = QStandardItem(m.recipient)
            item_recipient.setBackground(QColor("lightgreen"))
            item_subject = QStandardItem(m.subject)
            item_subject.setBackground(QColor("lightblue"))
            item_subject.setData(m.content)
            self.model.appendRow([item_date, item_sender, item_recipient, item_subject])
        self.messagesTableView.setModel(self.model)
        self.messagesTableView.setColumnWidth(0, 250)
        self.messagesTableView.setColumnWidth(3, 450)

class Folder():
    def __init__(self, folder_name: str, parent: str):
        self._folder_name = folder_name
        self._parent = parent
        self._mailfolder = self.is_mailfolder() 
        self._messages_list = self.get_messages()

    @property
    def folder_name(self) -> str:
        return self._folder_name

    @property
    def mailfolder(self):
        return self._mailfolder
    
    @property
    def messages_list(self):
        return self._messages_list

    def is_mailfolder(self) -> bool:
        pattern = 'Message[0-9]{5}'
        try:
            self._folder_content = os.listdir(f'{self._parent}/{self._folder_name}')
            for sf in self._folder_content:
                if re.match(pattern, sf):
                    print({self._folder_name})
                    return True
        except NotADirectoryError as e: 
            print(f"{self._folder_name} is not a dir ERROR:{e}")
        return False
    
    def extract_mail_from_string(self, line: str) -> str:
        parts = []
        parts = line.split('<')
        email = parts[-1]
        email = email.replace("To: ", "")
        email = email.replace("From: ", "")
        email = email.replace("Date: ", "")
        email = email.replace("Subject: ", "")
        email = email.strip('<')
        email = email.strip('>')
        email = email.strip('"')
        return email 

    def get_messages(self) -> Optional(list[Message]) :
        message_list = []
        try:
            for message_folder in os.listdir(f'{self._parent}/{self._folder_name}'):
                subject = ''
                date = ''
                sender = ''
                recipient = ''
                message_html = ''
                if os.path.exists(f"{self._parent}/{self._folder_name}/{message_folder}/InternetHeaders.txt"):
                    with open(f"{self._parent}/{self._folder_name}/{message_folder}/InternetHeaders.txt", 'r') as fih:
                        header_content=fih.read()
                        for line in header_content.splitlines():
                            if re.search(r'^To:', line):
                                recipient = self.extract_mail_from_string(line)
                            elif re.search(r'^From:', line):
                                sender = self.extract_mail_from_string(line)
                            elif re.search(r'^Subject:', line):
                                subject = self.extract_mail_from_string(line)
                            elif re.search(r'^Date:', line):
                                date = self.extract_mail_from_string(line)

                            if subject != '' and date != '' and sender != '' and recipient != '':
                                break
                elif os.path.exists(f"{self._parent}/{self._folder_name}/{message_folder}/OutlookHeaders.txt"):
                    with open(f"{self._parent}/{self._folder_name}/{message_folder}/OutlookHeaders.txt", 'r') as fih:
                        header_content=fih.read()
                        for line in header_content.splitlines():
                            if re.search(r'^Sender name:', line):
                                sender = self.extract_mail_from_string(line)
                            elif re.search(r'^Subject:', line):
                                subject = self.extract_mail_from_string(line)
                            elif re.search(r'^Client submit time:', line):
                                line = line.replace("Client submit time:", "")
                                line = line.strip(" ")
                                date = self.extract_mail_from_string(date)
                            
                            if subject != '' and date != '' and sender != '' :
                                break 
                else:
                    print("No headers file found")

                if os.path.exists(f"{self._parent}/{self._folder_name}/{message_folder}/Message.txt"):
                    with open(f"{self._parent}/{self._folder_name}/{message_folder}/Message.txt", 'r', encoding='latin1') as fm:
                        message_html = fm.read()
                elif os.path.exists(f"{self._parent}/{self._folder_name}/{message_folder}/Message.html"):
                    with open(f"{self._parent}/{self._folder_name}/{message_folder}/Message.html", 'r', encoding='latin1') as fm:
                        message_html = fm.read()
                else:
                    print(f"No message body file found")
                    message_html = '' 
                message = Message(message_folder, date, subject, sender, recipient, message_html)
                message_list.append(message)
        
            def by_date(e):
                return e.int_date
            message_list.sort(key=by_date, reverse=True)
            return message_list
        except NotADirectoryError as e: 
            print(f"{self._folder_name} is not a dir ERROR:{e}")

class Message:
    def __init__(self, folder_name: str, date: str, subject: str, sender: str, recipient: str, content: str) ->None:
        self._folder_name = folder_name
        self._date = date
        self._subject = subject
        self._sender = sender
        self._recipient = recipient
        self._content = content

        if self._date != '':
            self._dt = parsedate_to_datetime(self._date)
            self._int_date = int(self._dt.timestamp())
        else:
            self._int_date = -1 
    
    @property
    def folder_name(self):
        return self._folder_name
    
    @property
    def date(self):
        return self._date
    
    @property
    def int_date(self):
        return self._int_date
    
    @property
    def subject(self):
        return self._subject
    
    @property
    def sender(self):
        return self._sender
    
    @property
    def recipient(self):
        return self._recipient
    
    @property
    def content(self):
        return self._content

if __name__ == "__main__":
    mailb_app = QApplication(sys.argv)
    mailb_app.setApplicationName("Mail Browser")
    mwindows = MainWindow()
    mwindows.show()
    sys.exit(mailb_app.exec_())
