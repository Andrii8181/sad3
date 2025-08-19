import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QAction, QFileDialog,
    QMessageBox, QTableWidget, QTableWidgetItem, QMenu, QVBoxLayout, QWidget
)
from analysis import AnalysisDialog


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SAD – Статистичний аналіз даних")
        self.setGeometry(200, 100, 1000, 600)

        # таблиця для введення
        self.table = QTableWidget(5, 5)  # початкова таблиця 5х5
        self.setCentralWidget(self.table)

        # меню
        menubar = self.menuBar()

        file_menu = menubar.addMenu("Файл")
        analysis_menu = menubar.addMenu("Аналіз даних")
        about_menu = menubar.addMenu("Про програму")

        # дії для Файл
        open_action = QAction("Відкрити Excel", self)
        open_action.triggered.connect(self.load_excel)
        save_action = QAction("Зберегти у Excel", self)
        save_action.triggered.connect(self.save_excel)
        file_menu.addAction(open_action)
        file_menu.addAction(save_action)

        # дії для редагування таблиці
        edit_menu = menubar.addMenu("Редагувати таблицю")
        add_row = QAction("Вставити рядок", self)
        add_row.triggered.connect(self.add_row)
        del_row = QAction("Видалити рядок", self)
        del_row.triggered.connect(self.delete_row)
        add_col = QAction("Вставити стовпчик", self)
        add_col.triggered.connect(self.add_col)
        del_col = QAction("Видалити стовпчик", self)
        del_col.triggered.connect(self.delete_col)
        edit_menu.addActions([add_row, del_row, add_col, del_col])

        # дія аналіз даних
        run_analysis = QAction("Вибір аналізу", self)
        run_analysis.triggered.connect(self.open_analysis_dialog)
        analysis_menu.addAction(run_analysis)

        # дія про програму
        about_action = QAction("Інформація", self)
        about_action.triggered.connect(self.show_about)
        about_menu.addAction(about_action)

    # функції редагування
    def add_row(self):
        self.table.insertRow(self.table.rowCount())

    def delete_row(self):
        if self.table.rowCount() > 0:
            self.table.removeRow(self.table.currentRow())

    def add_col(self):
        self.table.insertColumn(self.table.columnCount())

    def delete_col(self):
        if self.table.columnCount() > 0:
            self.table.removeColumn(self.table.currentColumn())

    # робота з Excel
    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Відкрити файл", "", "Excel Files (*.xlsx)")
        if path:
            df = pd.read_excel(path)
            self.set_table_from_df(df)

    def save_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Зберегти файл", "", "Excel Files (*.xlsx)")
        if path:
            df = self.get_df_from_table()
            df.to_excel(path, index=False)

    def set_table_from_df(self, df):
        self.table.setRowCount(df.shape[0])
        self.table.setColumnCount(df.shape[1])
        self.table.setHorizontalHeaderLabels(df.columns)
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                self.table.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))

    def get_df_from_table(self):
        rows = self.table.rowCount()
        cols = self.table.columnCount()
        data = []
        headers = []
        for j in range(cols):
            header = self.table.horizontalHeaderItem(j)
            headers.append(header.text() if header else f"Колонка {j+1}")
        for i in range(rows):
            row = []
            for j in range(cols):
                item = self.table.item(i, j)
                row.append(item.text() if item else "")
            data.append(row)
        return pd.DataFrame(data, columns=headers)

    # відкриття діалогу аналізу
    def open_analysis_dialog(self):
        df = self.get_df_from_table()
        if df.empty:
            QMessageBox.warning(self, "Помилка", "Таблиця порожня!")
            return
        dlg = AnalysisDialog(df, self)
        dlg.exec_()

    def show_about(self):
        QMessageBox.information(
            self,
            "Про програму",
            "SAD – Статистичний аналіз даних\n\nРозробник: Чаплоуцький А.М.\n"
            "Кафедра плодівництва і виноградарства УНУ\n"
            "© Усі права захищено, 2025"
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
