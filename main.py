import sys
from PyQt5.QtCore import Qt, QEvent
from PyQt5.QtGui import QKeySequence
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QAction,
    QFileDialog, QMessageBox, QInputDialog, QAbstractItemView
)
import pandas as pd
from analysis import AnalysisWindow
from about import AboutDialog


class EditableTable(QTableWidget):
    """Таблиця з коректним редагуванням, копію/вставкою і переходом Enter."""
    def __init__(self, rows=5, cols=5, parent=None):
        super().__init__(rows, cols, parent)
        self.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setEditTriggers(
            QAbstractItemView.DoubleClicked
            | QAbstractItemView.SelectedClicked
            | QAbstractItemView.EditKeyPressed
            | QAbstractItemView.AnyKeyPressed
        )
        # Заповнюємо порожні елементи, щоб редагування завжди було можливе
        for r in range(rows):
            for c in range(cols):
                if not self.item(r, c):
                    self.setItem(r, c, QTableWidgetItem(""))

        # Дозволяємо перейменування заголовків по дабл-кліку
        header = self.horizontalHeader()
        header.setSectionsClickable(True)
        header.sectionDoubleClicked.connect(self.rename_header)

        # Вішаємо фільтр для Ctrl+C / Ctrl+V
        self.installEventFilter(self)

    def rename_header(self, section: int):
        current = self.horizontalHeaderItem(section).text() if self.horizontalHeaderItem(section) else f"Стовпець {section+1}"
        text, ok = QInputDialog.getText(self, "Перейменувати стовпчик", "Нова назва:", text=current)
        if ok and text.strip():
            self.setHorizontalHeaderItem(section, QTableWidgetItem(text.strip()))
        elif ok:
            # Якщо ввели порожньо — лишаємо як було
            pass

    def ensure_item(self, row, col):
        if not self.item(row, col):
            self.setItem(row, col, QTableWidgetItem(""))

    def keyPressEvent(self, event):
        # Enter / Return → перехід до наступної клітинки (вниз, як в Excel)
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            r = self.currentRow()
            c = self.currentColumn()
            # Якщо стоїмо на останньому рядку — додаємо новий
            if r == self.rowCount() - 1:
                self.insertRow(self.rowCount())
                for cc in range(self.columnCount()):
                    self.ensure_item(self.rowCount() - 1, cc)
            # Перейти вниз у тому ж стовпчику
            self.setCurrentCell(r + 1, c)
            event.accept()
            return
        super().keyPressEvent(event)

    def eventFilter(self, source, event):
        if source is self and event.type() == QEvent.KeyPress:
            # Копіювання
            if QKeySequence(event.modifiers() | event.key()) == QKeySequence.Copy:
                self.copy_selection()
                return True
            # Вставка
            if QKeySequence(event.modifiers() | event.key()) == QKeySequence.Paste:
                self.paste_selection()
                return True
        return super().eventFilter(source, event)

    def copy_selection(self):
        ranges = self.selectedRanges()
        if not ranges:
            return
        r = ranges[0]
        lines = []
        for row in range(r.topRow(), r.bottomRow() + 1):
            vals = []
            for col in range(r.leftColumn(), r.rightColumn() + 1):
                item = self.item(row, col)
                vals.append("" if item is None else item.text())
            lines.append("\t".join(vals))
        QApplication.clipboard().setText("\n".join(lines))

    def paste_selection(self):
        """Вставка з буфера (розширює таблицю, якщо не вистачає місця)."""
        start_ranges = self.selectedRanges()
        if not start_ranges:
            return
        start = start_ranges[0]
        r0, c0 = start.topRow(), start.leftColumn()
        text = QApplication.clipboard().text()
        rows = [row for row in text.splitlines() if row.strip() != ""]
        # Розширюємо таблицю за потреби
        needed_rows = r0 + len(rows)
        if needed_rows > self.rowCount():
            for _ in range(needed_rows - self.rowCount()):
                self.insertRow(self.rowCount())
        for i, line in enumerate(rows):
            cols = line.split("\t")
            needed_cols = c0 + len(cols)
            if needed_cols > self.columnCount():
                for _ in range(needed_cols - self.columnCount()):
                    self.insertColumn(self.columnCount())
            for j, val in enumerate(cols):
                self.ensure_item(r0 + i, c0 + j)
                self.item(r0 + i, c0 + j).setText(val)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SAD – Статистичний аналіз даних")
        self.resize(1000, 640)

        # Стартова таблиця 5×5 з трьома заголовками як приклад
        self.table = EditableTable(5, 5, self)
        self.table.setHorizontalHeaderLabels(["Фактор A", "Фактор B", "Результат", "Стовпець 4", "Стовпець 5"])
        self.setCentralWidget(self.table)

        self.create_menu()

    def create_menu(self):
        menubar = self.menuBar()

        # Файл
        m_file = menubar.addMenu("Файл")
        act_open = QAction("Відкрити з Excel", self)
        act_open.triggered.connect(self.load_from_excel)
        act_save = QAction("Зберегти в Excel", self)
        act_save.triggered.connect(self.save_to_excel)
        m_file.addAction(act_open)
        m_file.addAction(act_save)

        # Редагувати
        m_edit = menubar.addMenu("Редагувати")
        act_add_row = QAction("Додати рядок", self)
        act_add_row.triggered.connect(self.add_row)
        act_del_row = QAction("Видалити рядок", self)
        act_del_row.triggered.connect(self.delete_row)
        act_add_col = QAction("Додати стовпець", self)
        act_add_col.triggered.connect(self.add_col)
        act_del_col = QAction("Видалити стовпець", self)
        act_del_col.triggered.connect(self.delete_col)
        m_edit.addAction(act_add_row)
        m_edit.addAction(act_del_row)
        m_edit.addAction(act_add_col)
        m_edit.addAction(act_del_col)

        # Аналіз даних (без проміжної кнопки)
        m_analysis = menubar.addMenu("Аналіз даних")
        act_run = QAction("Запустити аналіз…", self)
        act_run.triggered.connect(self.open_analysis)
        m_analysis.addAction(act_run)

        # Про програму
        m_about = menubar.addMenu("Про програму")
        act_about = QAction("Інформація", self)
        act_about.triggered.connect(self.show_about)
        m_about.addAction(act_about)

    # ==== Операції з таблицею ====
    def add_row(self):
        r = self.table.rowCount()
        self.table.insertRow(r)
        for c in range(self.table.columnCount()):
            self.table.setItem(r, c, QTableWidgetItem(""))

    def delete_row(self):
        if self.table.rowCount() > 0:
            self.table.removeRow(self.table.currentRow())

    def add_col(self):
        c = self.table.columnCount()
        self.table.insertColumn(c)
        self.table.setHorizontalHeaderItem(c, QTableWidgetItem(f"Стовпець {c+1}"))
        for r in range(self.table.rowCount()):
            self.table.setItem(r, c, QTableWidgetItem(""))

    def delete_col(self):
        if self.table.columnCount() > 0:
            self.table.removeColumn(self.table.currentColumn())

    # ==== Excel ====
    def save_to_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Зберегти файл", "", "Excel (*.xlsx)")
        if not path:
            return
        data = []
        for r in range(self.table.rowCount()):
            row = []
            for c in range(self.table.columnCount()):
                item = self.table.item(r, c)
                row.append("" if item is None else item.text())
            data.append(row)
        cols = []
        for i in range(self.table.columnCount()):
            hi = self.table.horizontalHeaderItem(i)
            cols.append(hi.text() if hi else f"Стовпець {i+1}")
        df = pd.DataFrame(data, columns=cols)
        try:
            df.to_excel(path, index=False)
            QMessageBox.information(self, "Готово", "Файл збережено.")
        except Exception as e:
            QMessageBox.critical(self, "Помилка збереження", str(e))

    def load_from_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Відкрити файл", "", "Excel (*.xlsx *.xls)")
        if not path:
            return
        try:
            df = pd.read_excel(path, dtype=object)
        except Exception as e:
            QMessageBox.critical(self, "Помилка відкриття", str(e))
            return

        self.table.clear()
        self.table.setRowCount(len(df.index))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels([str(c) for c in df.columns])

        for r in range(len(df.index)):
            for c in range(len(df.columns)):
                self.table.setItem(r, c, QTableWidgetItem("" if pd.isna(df.iat[r, c]) else str(df.iat[r, c])))

    # ==== Аналіз ====
    def open_analysis(self):
        df = self.table_to_dataframe()
        if df is None or df.empty:
            QMessageBox.warning(self, "Помилка", "Таблиця порожня.")
            return
        dlg = AnalysisWindow(df)
        dlg.exec_()

    def table_to_dataframe(self) -> pd.DataFrame:
        rows = self.table.rowCount()
        cols = self.table.columnCount()
        if cols == 0:
            return pd.DataFrame()

        headers = []
        for i in range(cols):
            hi = self.table.horizontalHeaderItem(i)
            headers.append(hi.text() if hi else f"Стовпець {i+1}")

        data = []
        for r in range(rows):
            row = []
            for c in range(cols):
                it = self.table.item(r, c)
                row.append("" if it is None else it.text())
            data.append(row)

        df = pd.DataFrame(data, columns=headers)
        return df

    def show_about(self):
        AboutDialog(self).exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())
