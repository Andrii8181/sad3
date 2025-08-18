import sys, os, tempfile
import pandas as pd
from PyQt5 import QtGui
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem,
    QPushButton, QFileDialog, QMessageBox, QInputDialog, QAction, QWidget, QVBoxLayout, QHBoxLayout
)
from analysis import (
    detect_columns, check_normality, explain_non_normal,
    run_selected_analyses, run_selected_nonparametric_analyses
)
from charts import build_default_charts_from_results
from export_word import export_to_word
from data_input import read_table_from_widget, write_df_to_widget, import_dataframe_from_file


APP_TITLE = "SAD — Статистичний Аналіз Даних"
ABOUT_TEXT = (
    "SAD — Статистичний Аналіз Даних\n\n"
    "Розробник: Чаплоуцький А.М.\n"
    "Кафедра плодівництва і виноградарства УНУ"
)


class SADWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.setWindowIcon(QtGui.QIcon("icon.ico"))
        self.resize(1100, 700)

        # центральний віджет + лейаути
        central = QWidget(self)
        self.setCentralWidget(central)
        main_v = QVBoxLayout(central)

        # Таблиця
        self.table = QTableWidget(self)
        self.table.setRowCount(8)
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Фактор A", "Фактор B", "Фактор C", "Значення"])
        main_v.addWidget(self.table)

        # Кнопки під таблицею
        row = QHBoxLayout()
        self.btn_add_row = QPushButton("Додати рядок")
        self.btn_del_row = QPushButton("Видалити рядок")
        self.btn_add_col = QPushButton("Додати стовпець")
        self.btn_del_col = QPushButton("Видалити стовпець")
        self.btn_clear = QPushButton("Очистити таблицю")
        self.btn_analyze = QPushButton("Аналіз даних")
        self.btn_export = QPushButton("Експорт у Word")
        self.btn_about = QPushButton("Про програму")

        for b in [self.btn_add_row, self.btn_del_row, self.btn_add_col, self.btn_del_col, self.btn_clear,
                  self.btn_analyze, self.btn_export, self.btn_about]:
            row.addWidget(b)
        main_v.addLayout(row)

        # Події
        self.btn_add_row.clicked.connect(lambda: self.table.insertRow(self.table.rowCount()))
        self.btn_del_row.clicked.connect(self.delete_current_row)
        self.btn_add_col.clicked.connect(self.add_column)
        self.btn_del_col.clicked.connect(self.delete_current_column)
        self.btn_clear.clicked.connect(self.clear_table)
        self.btn_analyze.clicked.connect(self.analyze_flow)
        self.btn_export.clicked.connect(self.export_flow)
        self.btn_about.clicked.connect(lambda: QMessageBox.information(self, "Про розробника", ABOUT_TEXT))

        # Меню: Файл → Імпорт
        menubar = self.menuBar()
        file_menu = menubar.addMenu("Файл")
        act_import = QAction("Імпорт з Excel/CSV…", self)
        act_import.triggered.connect(self.import_file)
        act_exit = QAction("Вихід", self)
        act_exit.triggered.connect(self.close)
        file_menu.addAction(act_import)
        file_menu.addSeparator()
        file_menu.addAction(act_exit)

        # змінні стану
        self.last_results = {}   # dict: назва_аналізу -> результати/таблиці/тексти
        self.last_fig_paths = [] # шляхи до збережених графіків
        self.last_df = None
        self.title_text = ""
        self.units_text = ""

    # ---- Табличні утиліти ----
    def add_column(self):
        c = self.table.columnCount()
        self.table.insertColumn(c)
        self.table.setHorizontalHeaderItem(c, QTableWidgetItem(f"Колонка {c+1}"))

    def delete_current_row(self):
        r = self.table.currentRow()
        if r >= 0:
            self.table.removeRow(r)

    def delete_current_column(self):
        c = self.table.currentColumn()
        if c >= 0:
            self.table.removeColumn(c)

    def clear_table(self):
        ok = QMessageBox.question(self, "Підтвердження", "Очистити всю таблицю?",
                                  QMessageBox.Yes | QMessageBox.No)
        if ok == QMessageBox.Yes:
            self.table.clear()
            self.table.setRowCount(8)
            self.table.setColumnCount(4)
            self.table.setHorizontalHeaderLabels(["Фактор A", "Фактор B", "Фактор C", "Значення"])

    # ---- Імпорт ----
    def import_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Обрати файл", "", "Файли даних (*.xlsx *.xls *.csv)")
        if not path:
            return
        try:
            df = import_dataframe_from_file(path)
            write_df_to_widget(df, self.table)
            QMessageBox.information(self, "Імпорт", "Дані імпортовано успішно.")
        except Exception as e:
            QMessageBox.critical(self, "Помилка імпорту", str(e))

    # ---- Аналіз ----
    def analyze_flow(self):
        df = read_table_from_widget(self.table)
        if df.empty:
            QMessageBox.warning(self, "Помилка", "Таблиця порожня.")
            return

        # Назва показника / одиниці виміру (для звіту)
        self.title_text, ok = QInputDialog.getText(self, "Назва показника", "Введіть назву показника:")
        if not ok: self.title_text = "Показник"
        self.units_text, ok = QInputDialog.getText(self, "Одиниці виміру", "Введіть одиниці виміру (необов'язково):")
        if not ok: self.units_text = ""

        # Автовизначення колонок (фактори та числові)
        meta = detect_columns(df)  # {'numeric': [...], 'factors': [...]}
        if not meta["numeric"]:
            QMessageBox.warning(self, "Помилка", "Не знайдено числових стовпців.")
            return

        # Перевірка нормальності (Шапіро–Вілка) для кожної numeric-колонки
        pvals = {}
        try:
            for col in meta["numeric"]:
                pvals[col] = check_normality(pd.to_numeric(df[col], errors="coerce").dropna())
        except Exception as e:
            QMessageBox.critical(self, "Помилка перевірки нормальності", f"Деталі: {e}")
            return

        # якщо хоч один ключовий показник не нормальний → гілка непараметричних методів
        non_normal = [c for c, p in pvals.items() if p < 0.05]

        if non_normal:
            # Пояснення та вибір непараметричних методів
            QMessageBox.information(self, "Дані не нормальні", explain_non_normal(non_normal, pvals))
            methods = self.ask_methods(nonparametric=True)
            if not methods:
                return
            results = run_selected_nonparametric_analyses(df, meta, methods)
        else:
            # Вибір параметричних методів
            methods = self.ask_methods(nonparametric=False)
            if not methods:
                return
            results = run_selected_analyses(df, meta, methods)

        self.last_results = results
        self.last_df = df

        # Побудова графіків на основі результатів
        try:
            self.last_fig_paths = build_default_charts_from_results(df, meta, results, out_dir=tempfile.gettempdir())
        except Exception as e:
            self.last_fig_paths = []
            QMessageBox.warning(self, "Графіки", f"Не вдалося побудувати частину графіків: {e}")

        QMessageBox.information(self, "Готово", "Аналіз завершено. Можете експортувати звіт у Word.")

    def ask_methods(self, nonparametric: bool):
        """
        Діалог вибору методів: повертає список internal-ключів методів.
        """
        # перелік методів (ключ -> відображення)
        parametric = [
            ("anova_oneway", "Однофакторний ANOVA"),
            ("anova_twoway", "Двофакторний ANOVA"),
            ("anova_multi", "Багатофакторний ANOVA"),
            ("anova_rm", "Дисперсійний з повторними вимірюваннями"),
            ("hsd_lsd", "НІР (Fisher's LSD)"),
            ("tukey", "Тест Тьюкі"),
            ("duncan", "Тест Данкана"),
            ("bonferroni", "Тест Бонфероні"),
            ("variance", "Квадратичне відхилення"),
            ("effect_size", "Сила впливу факторів (η²)")
        ]
        nonparam = [
            ("kruskal", "Kruskal–Wallis (1-фактор)"),
            ("friedman", "Friedman (повторні вимірювання)"),
            ("mannwhitney", "Mann–Whitney (2 групи)"),
            ("dunn", "Dunn (пост-хок)"),
            ("variance", "Квадратичне відхилення"),
            ("effect_size", "Сила впливу (ε²/KW)"),
        ]

        items = nonparam if nonparametric else parametric

        # простий чекбокс-діалог без додаткового файлу
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QCheckBox, QDialogButtonBox
        dlg = QDialog(self)
        dlg.setWindowTitle("Виберіть методи аналізу")
        lay = QVBoxLayout(dlg)
        checks = []
        for key, label in items:
            cb = QCheckBox(label)
            lay.addWidget(cb)
            checks.append((key, cb))

        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept)
        bb.rejected.connect(dlg.reject)
        lay.addWidget(bb)

        if dlg.exec_() != QDialog.Accepted:
            return []

        selected = [k for k, cb in checks if cb.isChecked()]
        return selected

    # ---- Експорт ----
    def export_flow(self):
        if not self.last_results or self.last_df is None:
            QMessageBox.warning(self, "Експорт", "Немає готових результатів. Спочатку виконайте аналіз.")
            return
        save_path, _ = QFileDialog.getSaveFileName(self, "Зберегти звіт у Word", "", "Word (*.docx)")
        if not save_path:
            return
        try:
            export_to_word(
                filename=save_path,
                title=self.title_text or "Звіт SAD",
                units=self.units_text or "",
                raw_data=self.last_df,
                results=self.last_results,
                figure_paths=self.last_fig_paths,
                author_footer="Розробник: Чаплоуцький А.М., кафедра плодівництва і виноградарства УНУ"
            )
            QMessageBox.information(self, "Експорт", "Звіт збережено.")
        except Exception as e:
            QMessageBox.critical(self, "Помилка експорту", str(e))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = SADWindow()
    w.show()
    sys.exit(app.exec_())
