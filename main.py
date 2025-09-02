import sys
from PyQt5.QtCore import Qt, QRectF
from PyQt5.QtGui import QIcon, QPixmap, QPainter, QLinearGradient, QColor, QBrush, QPainterPath, QFont, QPen
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QAction, QFileDialog,
    QMessageBox, QMenu, QInputDialog, QToolBar
)


def create_app_icon() -> QIcon:
    """Створює яскраву сучасну іконку 'SAD' динамічно (без зовнішніх файлів)."""
    size = 256
    px = QPixmap(size, size)
    px.fill(Qt.transparent)
    p = QPainter(px)

    grad = QLinearGradient(0, 0, size, size)
    grad.setColorAt(0.0, QColor(0, 122, 255))
    grad.setColorAt(0.5, QColor(88, 86, 214))
    grad.setColorAt(1.0, QColor(255, 45, 85))
    p.setBrush(QBrush(grad))
    p.setPen(Qt.NoPen)

    path = QPainterPath()
    path.addRoundedRect(QRectF(18, 18, size - 36, size - 36), 44, 44)
    p.drawPath(path)

    p.setPen(QPen(Qt.white))
    font = QFont("Segoe UI", 92, QFont.Bold)
    p.setFont(font)
    p.drawText(px.rect(), Qt.AlignCenter, "SAD")
    p.end()
    return QIcon(px)


class TableWidget(QTableWidget):
    def __init__(self, rows=10, cols=6, parent=None):
        super().__init__(rows, cols, parent)
        self.setHorizontalHeaderLabels([f"Колонка {i+1}" for i in range(cols)])
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.open_context_menu)
        self.setAlternatingRowColors(True)
        self.setCornerButtonEnabled(True)
        self.resizeColumnsToContents()

    def keyPressEvent(self, event):
        if event.matches(event.StandardKey.Copy):
            self.copy_selection_to_clipboard()
            return
        if event.matches(event.StandardKey.Paste):
            self.paste_from_clipboard()
            return
        super().keyPressEvent(event)

    def copy_selection_to_clipboard(self):
        ranges = self.selectedRanges()
        if not ranges:
            return
        r = ranges[0]
        rows = []
        for i in range(r.topRow(), r.bottomRow() + 1):
            cells = []
            for j in range(r.leftColumn(), r.rightColumn() + 1):
                it = self.item(i, j)
                cells.append("" if it is None else it.text())
            rows.append("\t".join(cells))
        text = "\n".join(rows)
        QApplication.clipboard().setText(text)

    def paste_from_clipboard(self):
        text = QApplication.clipboard().text()
        if not text:
            return
        start_row = self.currentRow() if self.currentRow() >= 0 else 0
        start_col = self.currentColumn() if self.currentColumn() >= 0 else 0
        lines = [l for l in text.splitlines() if l.strip() != ""]
        for r_offset, line in enumerate(lines):
            cells = line.split("\t")
            while start_row + r_offset >= self.rowCount():
                self.insertRow(self.rowCount())
            while start_col + len(cells) - 1 >= self.columnCount():
                self.insertColumn(self.columnCount())
                self.setHorizontalHeaderItem(self.columnCount()-1, QTableWidgetItem(f"Колонка {self.columnCount()}"))
            for c_offset, value in enumerate(cells):
                item = QTableWidgetItem(value)
                self.setItem(start_row + r_offset, start_col + c_offset, item)
        self.resizeColumnsToContents()

    def open_context_menu(self, pos):
        menu = QMenu(self)
        act_copy = menu.addAction("Копіювати")
        act_paste = menu.addAction("Вставити")
        menu.addSeparator()
        act_add_row = menu.addAction("Вставити рядок")
        act_del_row = menu.addAction("Видалити рядок")
        act_add_col = menu.addAction("Вставити стовпчик")
        act_del_col = menu.addAction("Видалити стовпчик")
        menu.addSeparator()
        act_rename = menu.addAction("Перейменувати стовпчик…")

        act = menu.exec_(self.viewport().mapToGlobal(pos))
        if act == act_copy:
            self.copy_selection_to_clipboard()
        elif act == act_paste:
            self.paste_from_clipboard()
        elif act == act_add_row:
            self.insertRow(self.currentRow() if self.currentRow() >= 0 else self.rowCount())
        elif act == act_del_row:
            if self.rowCount() > 0 and self.currentRow() >= 0:
                self.removeRow(self.currentRow())
        elif act == act_add_col:
            col = self.currentColumn()
            if col < 0:
                col = self.columnCount()
            self.insertColumn(col + 1)
            self.setHorizontalHeaderItem(col + 1, QTableWidgetItem(f"Колонка {col + 2}"))
        elif act == act_del_col:
            if self.columnCount() > 0 and self.currentColumn() >= 0:
                self.removeColumn(self.currentColumn())
        elif act == act_rename:
            self.rename_current_column()

    def rename_current_column(self):
        col = self.currentColumn()
        if col < 0:
            QMessageBox.information(self, "Перейменування", "Оберіть стовпчик для перейменування.")
            return
        current = self.horizontalHeaderItem(col).text() if self.horizontalHeaderItem(col) else f"Колонка {col+1}"
        new_name, ok = QInputDialog.getText(self, "Перейменувати стовпчик", "Нова назва:", text=current)
        if ok and new_name.strip():
            self.setHorizontalHeaderItem(col, QTableWidgetItem(new_name.strip()))
            self.resizeColumnsToContents()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SAD – Статистичний аналіз даних")
        self.setMinimumSize(1100, 650)
        self.setWindowIcon(create_app_icon())
        self.table = TableWidget(10, 6, self)
        self.setCentralWidget(self.table)
        self.statusBar().showMessage("Готово")
        self._build_menus()
        self._build_toolbar()

    def _build_menus(self):
        menubar = self.menuBar()
        m_file = menubar.addMenu("Файл")
        act_new = QAction("Новий", self)
        act_new.triggered.connect(self.new_table)
        m_file.addAction(act_new)
        act_exit = QAction("Вихід", self)
        act_exit.triggered.connect(self.close)
        m_file.addAction(act_exit)

        m_edit = menubar.addMenu("Редагування таблиці")
        act_add_row = QAction("Вставити рядок", self)
        act_add_row.triggered.connect(lambda: self.table.insertRow(self.table.currentRow() if self.table.currentRow() >= 0 else self.table.rowCount()))
        m_edit.addAction(act_add_row)
        act_del_row = QAction("Видалити рядок", self)
        act_del_row.triggered.connect(lambda: (self.table.rowCount() and self.table.currentRow() >= 0 and self.table.removeRow(self.table.currentRow())))
        m_edit.addAction(act_del_row)
        act_add_col = QAction("Вставити стовпчик", self)
        act_add_col.triggered.connect(self.add_column_after_current)
        m_edit.addAction(act_add_col)
        act_del_col = QAction("Видалити стовпчик", self)
        act_del_col.triggered.connect(lambda: (self.table.columnCount() and self.table.currentColumn() >= 0 and self.table.removeColumn(self.table.currentColumn())))
        m_edit.addAction(act_del_col)
        act_rename = QAction("Перейменувати стовпчик…", self)
        act_rename.triggered.connect(self.table.rename_current_column)
        m_edit.addAction(act_rename)

        m_help = menubar.addMenu("Довідка")
        act_about = QAction("Про програму", self)
        act_about.triggered.connect(self.show_about)
        m_help.addAction(act_about)

    def _build_toolbar(self):
        tb = QToolBar("Швидкі дії", self)
        tb.setMovable(False)
        self.addToolBar(tb)
        a_add_row = QAction("＋ рядок", self)
        a_add_row.triggered.connect(lambda: self.table.insertRow(self.table.currentRow() if self.table.currentRow() >= 0 else self.table.rowCount()))
        tb.addAction(a_add_row)
        a_del_row = QAction("− рядок", self)
        a_del_row.triggered.connect(lambda: (self.table.rowCount() and self.table.currentRow() >= 0 and self.table.removeRow(self.table.currentRow())))
        tb.addAction(a_del_row)

    def new_table(self):
        self.table.clear()
        self.table.setRowCount(10)
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels([f"Колонка {i+1}" for i in range(6)])
        self.table.resizeColumnsToContents()
        self.statusBar().showMessage("Створено нову таблицю 10×6", 3000)

    def add_column_after_current(self):
        col = self.table.currentColumn()
        if col < 0:
            col = self.table.columnCount() - 1
        self.table.insertColumn(col + 1)
        self.table.setHorizontalHeaderItem(col + 1, QTableWidgetItem(f"Колонка {col + 2}"))
        self.table.resizeColumnsToContents()

    def show_about(self):
        QMessageBox.information(
            self,
            "Про програму",
            "SAD – Статистичний аналіз даних\n\n"
            "Розробник: Чаплоуцький А.М.\n"
            "Кафедра плодівництва і виноградарства УНУ\n\n"
            "© Усі права захищено, 2025"
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())
