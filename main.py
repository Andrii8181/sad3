import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QAction, QFileDialog,
    QMessageBox, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget
)
import pandas as pd
from analysis import AnalysisWindow
from about import AboutWindow


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SAD — Статистичний аналіз даних")
        self.resize(900, 600)

        # Початкова таблиця (5x5)
        self.table = QTableWidget(5, 5)
        self.setCentralWidget(self.table)

        # Меню
        self.create_menu()

    def create_menu(self):
        menubar = self.menuBar()

        # Файл
        file_menu = menubar.addMenu("Файл")

        load_action = QAction("Завантажити з Excel", self)
        load_action.triggered.connect(self.load_from_excel)
        file_menu.addAction(load_action)

        save_action = QAction("Зберегти у Excel", self)
        save_action.triggered.connect(self.save_to_excel)
        file_menu.addAction(save_action)

        exit_action = QAction("Вихід", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Редагування
        edit_menu = menubar.addMenu("Редагування")

        add_row_action = QAction("Додати рядок", self)
        add_row_action.triggered.connect(self.add_row)
        edit_menu.addAction(add_row_action)

        remove_row_action = QAction("Видалити рядок", self)
        remove_row_action.triggered.connect(self.remove_row)
        edit_menu.addAction(remove_row_action)

        add_col_action = QAction("Додати стовпчик", self)
        add_col_action.triggered.connect(self.add_column)
        edit_menu.addAction(add_col_action)

        remove_col_action = QAction("Видалити стовпчик", self)
        remove_col_action.triggered.connect(self.remove_column)
        edit_menu.addAction(remove_col_action)

        # Аналіз
        analysis_menu = menubar.addMenu("Аналіз")
        analyze_action = QAction("Аналіз даних", self)
        analyze_action.triggered.connect(self.open_analysis)
        analysis_menu.addAction(analyze_action)

        # Допомога
        help_menu = menubar.addMenu("Допомога")
        about_action = QAction("Про програму", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    # === Операції з таблицею ===
    def add_row(self):
        self.table.insertRow(self.table.rowCount())

    def remove_row(self):
        if self.table.rowCount() > 0:
            self.table.removeRow(self.table.rowCount() - 1)

    def add_column(self):
        self.table.insertColumn(self.table.columnCount())

    def remove_column(self):
        if self.table.columnCount() > 0:
            self.table.removeColumn(self.table.columnCount() - 1)

    # === Робота з Excel ===
    def load_from_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Відкрити Excel", "", "Excel Files (*.xlsx *.xls)")
        if path:
            df = pd.read_excel(path)
            self.table.setRowCount(df.shape[0])
            self.table.setColumnCount(df.shape[1])
            for i in range(df.shape[0]):
                for j in range(df.shape[1]):
                    self.table.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))

    def save_to_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Зберегти Excel", "", "Excel Files (*.xlsx)")
        if path:
            data = []
            for i in range(self.table.rowCount()):
                row = []
                for j in range(self.table.columnCount()):
                    item = self.table.item(i, j)
                    row.append(item.text() if item else "")
                data.append(row)
            df = pd.DataFrame(data)
            df.to_excel(path, index=False)

    # === Аналіз ===
    def open_analysis(self):
        data = []
        for i in range(self.table.rowCount()):
            row = []
            for j in range(self.table.columnCount()):
                item = self.table.item(i, j)
                row.append(item.text() if item else "")
            data.append(row)
        df = pd.DataFrame(data)
        self.analysis_window = AnalysisWindow(df)
        self.analysis_window.show()

    def show_about(self):
        self.about_window = AboutWindow()
        self.about_window.exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
