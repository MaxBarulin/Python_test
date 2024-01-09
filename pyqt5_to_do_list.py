import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QListWidget, QListWidgetItem
from PyQt5.QtCore import Qt

class ToDoList(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # заголовок
        label = QLabel("To Do List")
        label.setAlignment(Qt.AlignCenter)
        label.setStyleSheet("font-size: 18px; font-weight: bold;")

        # поле для ввода задачи
        self.task_input = QLineEdit()
        self.task_input.setPlaceholderText("Введите задачу")

        # кнопка добавления задачи
        self.add_button = QPushButton("Добавить")
        self.add_button.clicked.connect(self.add_task)

        # список задач
        self.task_list = QListWidget()
        self.task_list.itemChanged.connect(self.check_task)

        # кнопка удаления задачи
        self.remove_button = QPushButton("Удалить")
        self.remove_button.clicked.connect(self.remove_task)

        # расположение элементов
        h_layout = QHBoxLayout()
        h_layout.addWidget(self.task_input)
        h_layout.addWidget(self.add_button)
        layout.addLayout(h_layout)
        layout.addWidget(self.task_list)
        layout.addWidget(self.remove_button)

        self.setLayout(layout)
        self.setWindowTitle("To Do List")
        self.show()

    def add_task(self):
        task_text = self.task_input.text().strip()
        if task_text:
            task_item = QListWidgetItem(task_text)
            task_item.setFlags(task_item.flags() | Qt.ItemIsUserCheckable)
            task_item.setCheckState(Qt.Unchecked)
            self.task_list.addItem(task_item)
            self.task_input.clear()

    def remove_task(self):
        for item in self.task_list.selectedItems():
            self.task_list.takeItem(self.task_list.row(item))

    def check_task(self, item):
        if item.checkState() == Qt.Checked:
            item.setFlags(item.flags() & ~Qt.ItemIsEditable)
        else:
            item.setFlags(item.flags() | Qt.ItemIsEditable)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    to_do_list = ToDoList()
    sys.exit(app.exec_())
