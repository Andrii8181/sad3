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
        self.setWindowTitle("SAD - Статистичний аналіз даних")
        self.setGeometry(100, 100, 900, 600)

        self.table = QTableWidget()
        self.setCentralWidget(self.table)

        self.create_menu()

    def create_menu(self):
        # --- Головне меню ---
        menubar = self.menuBar()

        # Файл
        file_menu = menubar.addMenu("Файл")

        open_action = QAction("Відкрити (Excel/CSV)", self)
        open_action.triggered.connect(self.load_data)
        file_menu.addAction(open_action)

        save_action = QAction("Зберегти (Excel)", self)
        save_action.triggered.connect(self.save_data)
        file_menu.addAction(save_action)

        exit_action = QAction("Вихід", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Редагувати
        edit_menu = menubar.addMenu("Редагувати")

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
        analysis_menu = menubar.addMenu("Аналіз")
        analysis_action = QAction("Аналіз даних", self)
        analysis_action.triggered.connect(self.open_analysis_window)
        analysis_menu.addAction(analysis_action)

        # Довідка
        help_menu = menubar.addMenu("Довідка")
        about_action = QAction("Про програму", self)
        about_action.triggered.connect(self.open_about_window)
        help_menu.addAction(about_action)

    # --- Робота з таблицею ---
    def add_row(self):
        self.table.insertRow(self.table.rowCount())

    def del_row(self):
        if self.table.rowCount() > 0:
            self.table.removeRow(self.table.rowCount() - 1)

    def add_col(self):
        self.table.insertColumn(self.table.columnCount())

    def del_col(self):
        if self.table.columnCount() > 0:
            self.table.removeColumn(self.table.columnCount() - 1)

    def load_data(self):
        fname, _ = QFileDialog.getOpenFileName(self, "Відкрити файл", "", "Excel Files (*.xlsx);;CSV Files (*.csv)")
        if fname:
            if fname.endswith(".csv"):
                df = pd.read_csv(fname)
            else:
                df = pd.read_excel(fname)

            self.table.setRowCount(df.shape[0])
            self.table.setColumnCount(df.shape[1])
            self.table.setHorizontalHeaderLabels(df.columns)

            for i in range(df.shape[0]):
                for j in range(df.shape[1]):
                    self.table.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))

    def save_data(self):
        fname, _ = QFileDialog.getSaveFileName(self, "Зберегти файл", "", "Excel Files (*.xlsx)")
        if fname:
            data = []
            for i in range(self.table.rowCount()):
                row_data = []
                for j in range(self.table.columnCount()):
                    item = self.table.item(i, j)
                    row_data.append(item.text() if item else "")
                data.append(row_data)

            df = pd.DataFrame(data)
            df.to_excel(fname, index=False)

    # --- Аналіз ---
    def open_analysis_window(self):
        self.analysis_window = AnalysisWindow(self.table)
        self.analysis_window.show()

    # --- Про програму ---
    def open_about_window(self):
        self.about_window = AboutWindow()
        self.about_window.show()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
