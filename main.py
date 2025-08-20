import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QAction, QFileDialog, QMessageBox
from PyQt5.QtGui import QIcon
import pandas as pd
from analysis import run_analysis
from about import show_about

class SADApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SAD - Статистичний аналіз даних")
        self.setGeometry(100, 100, 900, 600)
        self.setWindowIcon(QIcon("icon.ico"))

        # Таблиця з початковими даними
        self.table = QTableWidget(self)
        self.table.setRowCount(5)
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels([f"Колонка {i+1}" for i in range(5)])
        self.setCentralWidget(self.table)

        # Меню
        menubar = self.menuBar()

        # Файл
        file_menu = menubar.addMenu("Файл")
        load_action = QAction("Завантажити з Excel", self)
        load_action.triggered.connect(self.load_excel)
        file_menu.addAction(load_action)

        save_action = QAction("Зберегти в Excel", self)
        save_action.triggered.connect(self.save_excel)
        file_menu.addAction(save_action)

        exit_action = QAction("Вихід", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Редагування
        edit_menu = menubar.addMenu("Редагування")
        add_row_action = QAction("Додати рядок", self)
        add_row_action.triggered.connect(self.add_row)
        edit_menu.addAction(add_row_action)

        del_row_action = QAction("Видалити рядок", self)
        del_row_action.triggered.connect(self.del_row)
        edit_menu.addAction(del_row_action)

        add_col_action = QAction("Додати стовпчик", self)
        add_col_action.triggered.connect(self.add_col)
        edit_menu.addAction(add_col_action)

        del_col_action = QAction("Видалити стовпчик", self)
        del_col_action.triggered.connect(self.del_col)
        edit_menu.addAction(del_col_action)

        # Аналіз
        analysis_menu = menubar.addMenu("Аналіз даних")
        run_analysis_action = QAction("Виконати аналіз", self)
        run_analysis_action.triggered.connect(self.run_analysis)
        analysis_menu.addAction(run_analysis_action)

        # Про програму
        about_menu = menubar.addMenu("Про програму")
        about_action = QAction("Інформація", self)
        about_action.triggered.connect(show_about)
        about_menu.addAction(about_action)

    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Відкрити Excel", "", "Excel Files (*.xlsx *.xls)")
        if path:
            df = pd.read_excel(path)
            self.table.setRowCount(len(df))
            self.table.setColumnCount(len(df.columns))
            self.table.setHorizontalHeaderLabels(df.columns)
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    self.table.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))

    def save_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Зберегти Excel", "", "Excel Files (*.xlsx)")
        if path:
            data = []
            headers = [self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())]
            for i in range(self.table.rowCount()):
                row_data = []
                for j in range(self.table.columnCount()):
                    item = self.table.item(i, j)
                    row_data.append(item.text() if item else "")
                data.append(row_data)
            df = pd.DataFrame(data, columns=headers)
            df.to_excel(path, index=False)

    def add_row(self):
        self.table.insertRow(self.table.rowCount())

    def del_row(self):
        if self.table.rowCount() > 0:
            self.table.removeRow(self.table.currentRow())

    def add_col(self):
        col = self.table.columnCount()
        self.table.insertColumn(col)
        self.table.setHorizontalHeaderItem(col, QTableWidgetItem(f"Колонка {col+1}"))

    def del_col(self):
        if self.table.columnCount() > 0:
            self.table.removeColumn(self.table.currentColumn())

    def run_analysis(self):
        run_analysis(self.table)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SADApp()
    window.show()
    sys.exit(app.exec_())
