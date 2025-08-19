import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QAction,
    QFileDialog, QMessageBox
)
from PyQt5.QtCore import Qt
import pandas as pd
from analysis import AnalysisWindow
from about import AboutDialog


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SAD – Статистичний аналіз даних")
        self.setGeometry(200, 100, 900, 600)

        # Таблиця для введення даних
        self.table = QTableWidget(self)
        self.table.setColumnCount(3)
        self.table.setRowCount(5)
        self.table.setHorizontalHeaderLabels(["Фактор A", "Фактор B", "Результат"])
        self.setCentralWidget(self.table)

        # Меню
        menubar = self.menuBar()

        file_menu = menubar.addMenu("Файл")
        open_action = QAction("Відкрити з Excel", self)
        open_action.triggered.connect(self.load_from_excel)
        save_action = QAction("Зберегти в Excel", self)
        save_action.triggered.connect(self.save_to_excel)
        file_menu.addAction(open_action)
        file_menu.addAction(save_action)

        edit_menu = menubar.addMenu("Редагувати")
        add_row_action = QAction("Додати рядок", self)
        add_row_action.triggered.connect(self.add_row)
        del_row_action = QAction("Видалити рядок", self)
        del_row_action.triggered.connect(self.delete_row)
        add_col_action = QAction("Додати стовпець", self)
        add_col_action.triggered.connect(self.add_col)
        del_col_action = QAction("Видалити стовпець", self)
        del_col_action.triggered.connect(self.delete_col)
        edit_menu.addAction(add_row_action)
        edit_menu.addAction(del_row_action)
        edit_menu.addAction(add_col_action)
        edit_menu.addAction(del_col_action)

        analysis_menu = menubar.addMenu("Аналіз даних")
        analysis_action = QAction("Запустити аналіз", self)
        analysis_action.triggered.connect(self.open_analysis)
        analysis_menu.addAction(analysis_action)

        about_menu = menubar.addMenu("Про програму")
        about_action = QAction("Інформація", self)
        about_action.triggered.connect(self.show_about)
        about_menu.addAction(about_action)

        # Підтримка копіювання / вставки
        self.table.installEventFilter(self)

    # Додавання і видалення рядків / стовпців
    def add_row(self):
        self.table.insertRow(self.table.rowCount())

    def delete_row(self):
        if self.table.rowCount() > 0:
            self.table.removeRow(self.table.currentRow())

    def add_col(self):
        col = self.table.columnCount()
        self.table.insertColumn(col)
        self.table.setHorizontalHeaderItem(col, QTableWidgetItem(f"Стовпець {col+1}"))

    def delete_col(self):
        if self.table.columnCount() > 0:
            self.table.removeColumn(self.table.currentColumn())

    # Збереження в Excel
    def save_to_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Зберегти файл", "", "Excel Files (*.xlsx)")
        if path:
            data = []
            for row in range(self.table.rowCount()):
                data.append([self.table.item(row, col).text() if self.table.item(row, col) else "" for col in range(self.table.columnCount())])
            df = pd.DataFrame(data, columns=[self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())])
            df.to_excel(path, index=False)

    # Завантаження з Excel
    def load_from_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Відкрити файл", "", "Excel Files (*.xlsx)")
        if path:
            df = pd.read_excel(path)
            self.table.setColumnCount(len(df.columns))
            self.table.setRowCount(len(df))
            self.table.setHorizontalHeaderLabels(list(df.columns))
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    self.table.setItem(i, j, QTableWidgetItem(str(df.iat[i, j])))

    # Запуск аналізу
    def open_analysis(self):
        df = self.table_to_dataframe()
        if df is None or df.empty:
            QMessageBox.warning(self, "Помилка", "Таблиця пуста!")
            return
        self.analysis_window = AnalysisWindow(df)
        self.analysis_window.show()

    def table_to_dataframe(self):
        data = []
        for row in range(self.table.rowCount()):
            row_data = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                row_data.append(item.text() if item else "")
            data.append(row_data)
        df = pd.DataFrame(data, columns=[self.table.horizontalHeaderItem(i).text() for i in range(self.table.columnCount())])
        return df

    def show_about(self):
        dlg = AboutDialog()
        dlg.exec_()

    # Події для Ctrl+C / Ctrl+V
    def eventFilter(self, source, event):
        if event.type() == event.KeyPress:
            if event.matches(event.Copy):
                self.copy_selection()
                return True
            elif event.matches(event.Paste):
                self.paste_selection()
                return True
        return super().eventFilter(source, event)

    def copy_selection(self):
        selection = self.table.selectedRanges()
        if selection:
            s = ""
            for r in range(selection[0].topRow(), selection[0].bottomRow() + 1):
                row_data = []
                for c in range(selection[0].leftColumn(), selection[0].rightColumn() + 1):
                    item = self.table.item(r, c)
                    row_data.append("" if item is None else item.text())
                s += "\t".join(row_data) + "\n"
            QApplication.clipboard().setText(s)

    def paste_selection(self):
        selection = self.table.selectedRanges()
        if selection:
            r = selection[0].topRow()
            c = selection[0].leftColumn()
            text = QApplication.clipboard().text()
            rows = text.split("\n")
            for i, row_data in enumerate(rows):
                if row_data.strip() == "":
                    continue
                cols = row_data.split("\t")
                for j, val in enumerate(cols):
                    if r + i < self.table.rowCount() and c + j < self.table.columnCount():
                        self.table.setItem(r + i, c + j, QTableWidgetItem(val))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
