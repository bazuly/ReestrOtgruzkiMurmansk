import sys
import json
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton, QListWidget, QMessageBox, QTabWidget

space_index = 0


class AddressList(QWidget):

    def __init__(self):
        super().__init__()

        try:
            with open('addresses.json', 'r', encoding='utf-8') as f:
                self.addresses = json.load(f)
        except FileNotFoundError:
            self.addresses = []

        self.setWindowTitle('Адреса ТТ Останкино/Мир Колбас/Черкизово')
        self.address_line_edit = QLineEdit()
        self.tab_widget = QTabWidget()
        self.add_address_button = QPushButton('Добавить адрес')
        self.address_list = QListWidget()
        self.main_layout = QVBoxLayout()
        self.address_layout = QHBoxLayout()
        self.address_layout.addWidget(self.address_line_edit)
        self.address_layout.addWidget(self.add_address_button)
        self.main_layout.addLayout(self.address_layout)
        self.main_layout.addWidget(self.address_list)
        self.setLayout(self.main_layout)
        self.setGeometry(750, 200, 500, 500)
        self.add_address_button.clicked.connect(self.add_address)
        self.address_list.itemDoubleClicked.connect(self.remove_address)
        
        
        
        self.update_list_widget()

    def add_address(self):
        address = self.address_line_edit.text()

        while address[-1] == ' ':
            address = address[:-1]

        space_index = 0
        while space_index < len(address) - 1:
            if address[space_index] == ' ' and address[space_index + 1] == ' ':
                address = address[:space_index] + address[space_index + 1:]
            else:
                space_index += 1

        while address[0] == ' ':
            address = address[1:]

        if not address:
            return

        if address in self.addresses:
            QMessageBox.warning(self, 'Ошибка', 'Такой адрес уже существует')
            return

        self.addresses.append(address)
        self.address_line_edit.setText('')
        self.update_list_widget()

    def remove_address(self, item):
        address = item.text()
        reply = QMessageBox.question(
            self, 'Подтверждение удаления',
            'Вы действительно хотите удалить данный адрес?',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.addresses.remove(address)
            self.update_list_widget()

    def update_list_widget(self):
        self.address_list.clear()
        self.address_list.addItems(self.addresses)

    def closeEvent(self, event):
        with open('addresses.json', 'w', encoding='utf-8') as f:
            json.dump(self.addresses, f, indent= 2, ensure_ascii=False)

        event.accept()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    address_list = AddressList()
    address_list.show()
    sys.exit(app.exec_())
