import sys
import requests  # Для отправки HTTP-запросов к Airflow
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget,
                             QFileDialog, QLabel, QListWidget, QMessageBox, QProgressBar)
from PyQt5.QtCore import Qt


class ExcelCombinerGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #212124;
                color: #ffffff;
            }
            QPushButton {
                color: #ffffff;
                padding-top: 12px;
                padding-bottom: 12px;
                margin-top: 2.5px;
                margin-bottom: 2.5px;
                border-radius: 10px;
                background-color: #161618;
            }
            QListWidget {
                color: #ffffff;
                border-radius: 10px;
                background-color: #000000;
            }
            QLabel {
                color: #ffffff;
            }
            QProgressBar {
                color: #ffffff;
            }
            QMessageBox {
                color: #ffffff;
                background-color: #161618;
            }
        """)

        self.setWindowTitle('Excel Combiner')
        self.setGeometry(100, 100, 800, 600)
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        layout = QVBoxLayout(self.central_widget)

        # Элементы интерфейса
        self.file_list = QListWidget()
        self.add_file_button = QPushButton('Add File')
        self.remove_file_button = QPushButton('Remove File')
        self.combine_button = QPushButton('Merge Files')
        self.status_label = QLabel('Ready for Assembly')

        # Добавление элементов в layout
        layout.addWidget(self.file_list)
        layout.addWidget(self.add_file_button)
        layout.addWidget(self.remove_file_button)
        layout.addWidget(self.combine_button)
        layout.addWidget(self.status_label, alignment=Qt.AlignBottom)

        # Привязка событий к кнопкам
        self.add_file_button.clicked.connect(self.add_files)
        self.remove_file_button.clicked.connect(self.remove_files)
        self.combine_button.clicked.connect(self.combine_files)

    def add_files(self):
        file_names, _ = QFileDialog.getOpenFileNames(self, 'Open File', '', 'Excel Files (*.xlsx)')
        self.file_list.addItems(file_names)

    def remove_files(self):
        list_items = self.file_list.selectedItems()
        if not list_items: return
        for item in list_items:
            self.file_list.takeItem(self.file_list.row(item))

    def combine_files(self):
        # Получаем пути ко всем выбранным файлам
        file_paths = [self.file_list.item(i).text() for i in range(self.file_list.count())]

        # Указываем путь для сохранения результата
        output_file, _ = QFileDialog.getSaveFileName(self, 'Save File', '', 'Excel Files (*.xlsx)')

        if not output_file:
            QMessageBox.warning(self, 'Error', 'You must specify the file to save.')
            return

        # Настройка для REST API Airflow
        dag_id = "combine_excel_sheets"
        url = f"http://localhost:8080/api/v1/dags/{dag_id}/dagRuns"  # URL Airflow для запуска DAG
        headers = {"Content-Type": "application/json"}
        data = {
            "conf": {
                "file_paths": file_paths,
                "output_file": output_file
            }
        }

        try:
            # Запуск DAG через REST API Airflow
            response = requests.post(url, json=data, headers=headers,
                                     auth=('airflow', 'airflow'))  # Airflow логин/пароль
            if response.status_code == 200:
                self.status_label.setText('DAG has been triggered successfully')
                QMessageBox.information(self, 'Success', 'The DAG has been triggered successfully.')
            else:
                raise Exception(f"Error triggering DAG: {response.text}")
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'An error occurred while triggering the DAG: {e}')
            self.status_label.setText('Error')


def main():
    app = QApplication(sys.argv)
    ex = ExcelCombinerGUI()
    ex.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
