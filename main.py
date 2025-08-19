import sys
import os
import tempfile
from datetime import datetime

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon, QKeySequence
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QTableWidget, QTableWidgetItem,
    QFileDialog, QAction, QToolBar, QMessageBox, QInputDialog, QDialog, QDialogButtonBox,
    QLabel, QFormLayout, QLineEdit, QCheckBox, QGroupBox, QGridLayout
)
import pandas as pd

from data_input import df_from_table, df_to_table, load_excel_or_csv, infer_roles
from analysis import shapiro_check, run_parametric_suite, run_nonparametric_suite
from charts import select_and_build_charts
from export_word import export_report


# ---------- допоміжні діалоги ----------
class MetaDialog(QDialog):
    """Ввід назви показника та одиниць виміру."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Параметри показника")
        form = QFormLayout(self)

        self.name_edit = QLineEdit()
        self.unit_edit = QLineEdit()
        form.addRow("Назва показника:", self.name_edit)
        form.addRow("Одиниці виміру:", self.unit_edit)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=self)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        form.addRow(btns)

    def values(self):
        return self.name_edit.text().strip(), self.unit_edit.text().strip()


class ParametricDialog(QDialog):
    """Вибір параметричних аналізів."""
    def __init__(self, factors, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Вибір параметричних аналізів")
        layout = QVBoxLayout(self)

        info = QLabel("Обрано параметричні методи (дані нормальні). Відмітьте потрібні тести:")
        info.setWordWrap(True)
        layout.addWidget(info)

        self.chk_lsd = QCheckBox("НІР (LSD)")
        self.chk_anova1 = QCheckBox("Однофакторний ANOVA")
        self.chk_anova2 = QCheckBox("Двофакторний ANOVA")
        self.chk_anovaN = QCheckBox("Багатофакторний ANOVA")
        self.chk_rm = QCheckBox("Повторні вимірювання (RM-ANOVA)")
        self.chk_tukey = QCheckBox("Пост-хок Тьюкі")
        self.chk_duncan = QCheckBox("Пост-хок Данкан")
        self.chk_bonf = QCheckBox("Пост-хок Бонферроні")
        self.chk_effect = QCheckBox("Сила впливу факторів (η²/partial η²)")

        for w in [self.chk_lsd, self.chk_anova1, self.chk_anova2, self.chk_anovaN,
                  self.chk_rm, self.chk_tukey, self.chk_duncan, self.chk_bonf, self.chk_effect]:
            w.setChecked(False)
            layout.addWidget(w)

        note = QLabel(f"Виявлені фактори: {', '.join(factors) if factors else 'не виявлено'}")
        note.setStyleSheet("color: #555")
        layout.addWidget(note)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=self)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def selection(self):
        return {
            "lsd": self.chk_lsd.isChecked(),
            "anova1": self.chk_anova1.isChecked(),
            "anova2": self.chk_anova2.isChecked(),
            "anovaN": self.chk_anovaN.isChecked(),
            "rm": self.chk_rm.isChecked(),
            "tukey": self.chk_tukey.isChecked(),
            "duncan": self.chk_duncan.isChecked(),
            "bonf": self.chk_bonf.isChecked(),
            "effect": self.chk_effect.isChecked(),
        }


class NonParamDialog(QDialog):
    """Вибір непараметричних аналізів та довідка."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Дані не нормальні — непараметричні методи")
        layout = QVBoxLayout(self)

        msg = QLabel(
            "Перевірка нормальності виявила відхилення.\n"
            "Параметричні методи (ANOVA тощо) можуть бути некоректними.\n"
            "Оберіть непараметричні аналізи для подальшої роботи:"
        )
        msg.setWordWrap(True)
        layout.addWidget(msg)

        self.chk_mw = QCheckBox("Манна–Вітні (2 групи)")
        self.chk_kw = QCheckBox("Краскела–Уолліса (k груп)")
        self.chk_friedman = QCheckBox("Фрідмана (повторні вимірювання)")
        self.chk_dunn = QCheckBox("Пост-хок Данна (для KW)")
        for w in [self.chk_mw, self.chk_kw, self.chk_friedman, self.chk_dunn]:
            layout.addWidget(w)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=self)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def selection(self):
        return {
            "mw": self.chk_mw.isChecked(),
            "kw": self.chk_kw.isChecked(),
            "friedman": self.chk_friedman.isChecked(),
            "dunn": self.chk_dunn.isChecked(),
        }


class ChartsDialog(QDialog):
    """Вибір типів діаграм для звіту."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Вибір діаграм")
        layout = QVBoxLayout(self)

        g = QGroupBox("Оберіть один або кілька типів графіків")
        grid = QGridLayout(g)
        self.chk_hist = QCheckBox("Гістограма/щільність")
        self.chk_box = QCheckBox("Box-plot за факторами")
        self.chk_bar = QCheckBox("Стовпчикова (середні ± SE)")
        self.chk_line = QCheckBox("Лінія тренду (регресія)")
        self.chk_pie = QCheckBox("Кругова (частки за групами)")
        for i, w in enumerate([self.chk_hist, self.chk_box, self.chk_bar, self.chk_line, self.chk_pie]):
            grid.addWidget(w, i // 2, i % 2)
        layout.addWidget(g)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, parent=self)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def selection(self):
        return {
            "hist": self.chk_hist.isChecked(),
            "box": self.chk_box.isChecked(),
            "bar": self.chk_bar.isChecked(),
            "line": self.chk_line.isChecked(),
            "pie": self.chk_pie.isChecked(),
        }


# ---------- головне вікно ----------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("SAD — Статистичний аналіз даних")
        if os.path.exists("icon.ico"):
            self.setWindowIcon(QIcon("icon.ico"))

        central = QWidget()
        self.setCentralWidget(central)
        self.table = QTableWidget(10, 6)

        lay = QVBoxLayout(central)
        lay.addWidget(self.table)

        self.create_menu()
        self.create_toolbar()

        # клавіатурні шорткати копію/вставки
        copy_act = QAction(self)
        copy_act.setShortcut(QKeySequence.Copy)
        copy_act.triggered.connect(self.copy_selection)
        self.addAction(copy_act)

        paste_act = QAction(self)
        paste_act.setShortcut(QKeySequence.Paste)
        paste_act.triggered.connect(self.paste_from_clipboard)
        self.addAction(paste_act)

    # ----- меню -----
    def create_menu(self):
        menubar = self.menuBar()

        m_file = menubar.addMenu("Файл")
        act_new = QAction("Новий", self, triggered=self.file_new)
        act_open = QAction("Відкрити… (Excel/CSV)", self, triggered=self.file_open)
        act_save = QAction("Зберегти звіт у Word…", self, triggered=self.export_word)
        act_exit = QAction("Вихід", self, triggered=self.close)
        m_file.addActions([act_new, act_open, act_save])
        m_file.addSeparator()
        m_file.addAction(act_exit)

        m_edit = menubar.addMenu("Редагувати")
        act_add_row = QAction("Додати рядок", self, triggered=lambda: self.table.insertRow(self.table.rowCount()))
        act_del_row = QAction("Видалити рядок", self, triggered=lambda: self.table.removeRow(max(self.table.currentRow(), 0)))
        act_add_col = QAction("Додати стовпець", self, triggered=lambda: self.table.insertColumn(self.table.columnCount()))
        act_del_col = QAction("Видалити стовпець", self, triggered=lambda: self.table.removeColumn(max(self.table.currentColumn(), 0)))
        act_copy = QAction("Копіювати", self, triggered=self.copy_selection)
        act_paste = QAction("Вставити", self, triggered=self.paste_from_clipboard)
        for a in [act_add_row, act_del_row, act_add_col, act_del_col]:
            a.setIconVisibleInMenu(False)
        m_edit.addActions([act_add_row, act_del_row, act_add_col, act_del_col])
        m_edit.addSeparator()
        m_edit.addActions([act_copy, act_paste])

        m_analysis = menubar.addMenu("Аналіз")
        act_analyze = QAction("Аналіз даних…", self, triggered=self.full_analysis_workflow)
        act_norm = QAction("Перевірка нормальності (Шапіро–Вілка)", self, triggered=self.check_normality_only)
        m_analysis.addActions([act_analyze, act_norm])

        m_help = menubar.addMenu("Довідка")
        act_about = QAction("Про програму", self, triggered=self.about_box)
        m_help.addAction(act_about)

    def create_toolbar(self):
        tb = QToolBar("Інструменти", self)
        self.addToolBar(tb)

        btn_add_row = QAction("＋ Рядок", self, triggered=lambda: self.table.insertRow(self.table.rowCount()))
        btn_del_row = QAction("－ Рядок", self, triggered=lambda: self.table.removeRow(max(self.table.currentRow(), 0)))
        btn_add_col = QAction("＋ Стовпець", self, triggered=lambda: self.table.insertColumn(self.table.columnCount()))
        btn_del_col = QAction("－ Стовпець", self, triggered=lambda: self.table.removeColumn(max(self.table.currentColumn(), 0)))
        btn_open = QAction("Відкрити", self, triggered=self.file_open)
        btn_an = QAction("Аналіз", self, triggered=self.full_analysis_workflow)

        for a in [btn_open, btn_add_row, btn_del_row, btn_add_col, btn_del_col, btn_an]:
            tb.addAction(a)

    # ----- дії меню/кнопок -----
    def file_new(self):
        self.table.clear()
        self.table.setRowCount(10)
        self.table.setColumnCount(6)

    def file_open(self):
        path, _ = QFileDialog.getOpenFileName(self, "Відкрити файл", "", "Таблиці (*.xlsx *.xls *.csv)")
        if not path:
            return
        try:
            df = load_excel_or_csv(path)
            df_to_table(self.table, df)
        except Exception as e:
            QMessageBox.critical(self, "Помилка відкриття", str(e))

    def export_word(self):
        df = df_from_table(self.table)
        title, unit = "Показник", ""
        path, _ = QFileDialog.getSaveFileName(self, "Зберегти звіт", "", "Word (*.docx)")
        if not path:
            return
        # порожній звіт (лише таблиця) — без аналізів і графіків
        export_report(df, title, unit, [], {}, path)

    def about_box(self):
        QMessageBox.information(
            self, "Про програму",
            "SAD — Статистичний аналіз даних\n"
            "Автор: Чаплоуцький А.М.\n"
            "Кафедра плодівництва і виноградарства УНУ\n"
            "© Усі права захищені. Ліцензія: MIT."
        )

    # ----- копіювання/вставка -----
    def copy_selection(self):
        sel = self.table.selectedRanges()
        if not sel:
            return
        r = sel[0]
        rows = range(r.topRow(), r.bottomRow() + 1)
        cols = range(r.leftColumn(), r.rightColumn() + 1)
        buf = []
        for i in rows:
            row_vals = []
            for j in cols:
                it = self.table.item(i, j)
                row_vals.append("" if it is None else it.text())
            buf.append("\t".join(row_vals))
        QApplication.clipboard().setText("\n".join(buf))

    def paste_from_clipboard(self):
        text = QApplication.clipboard().text()
        if not text:
            return
        start_row = max(self.table.currentRow(), 0)
        start_col = max(self.table.currentColumn(), 0)
        rows = text.splitlines()
        for i, line in enumerate(rows):
            cells = line.split("\t")
            r = start_row + i
            if r >= self.table.rowCount():
                self.table.insertRow(self.table.rowCount())
            for j, val in enumerate(cells):
                c = start_col + j
                if c >= self.table.columnCount():
                    self.table.insertColumn(self.table.columnCount())
                self.table.setItem(r, c, QTableWidgetItem(val))

    # ----- окремий запуск лише Шапіро -----
    def check_normality_only(self):
        df = df_from_table(self.table)
        roles = infer_roles(df)
        y = roles["response"]
        if y is None:
            QMessageBox.warning(self, "Немає показника", "Не вдалося ідентифікувати числовий стовпець як показник.")
            return
        ok, info = shapiro_check(df, y)
        if ok:
            QMessageBox.information(self, "Нормальність підтверджено", info)
        else:
            QMessageBox.warning(self, "Порушення нормальності", info)

    # ----- повний робочий процес -----
    def full_analysis_workflow(self):
        # 1) зчитати дані
        df = df_from_table(self.table)
        if df.empty:
            QMessageBox.warning(self, "Немає даних", "Заповніть таблицю або імпортуйте дані з Excel/CSV.")
            return

        # 2) назва показника та одиниці
        meta = MetaDialog(self)
        if meta.exec_() != QDialog.Accepted:
            return
        title, unit = meta.values()

        # 3) визначити ролі (показник/фактори)
        roles = infer_roles(df)
        y = roles["response"]
        factors = roles["factors"]
        subject = roles.get("subject")

        if y is None:
            QMessageBox.warning(self, "Немає показника", "Не вдалося ідентифікувати числовий стовпець як показник.")
            return

        # 4) автоматична перевірка нормальності
        normal, info = shapiro_check(df, y)
        if normal:
            QMessageBox.information(self, "Нормальність підтверджено", info)
            # 5a) параметричні аналізи
            dlg = ParametricDialog(factors, self)
            if dlg.exec_() != QDialog.Accepted:
                return
            sel = dlg.selection()
            results = run_parametric_suite(df, y, factors, subject, sel)
        else:
            # вікно з рекомендаціями та вибором непараметричних методів
            QMessageBox.information(
                self, "Дані не нормальні",
                info + "\n\nРекомендація: використовуйте непараметричні методи."
            )
            dlg = NonParamDialog(self)
            if dlg.exec_() != QDialog.Accepted:
                return
            sel = dlg.selection()
            results = run_nonparametric_suite(df, y, roles["group_for_np"], subject, sel)

        # 6) вибір графіків і побудова
        chart_dlg = ChartsDialog(self)
        figs = []
        if chart_dlg.exec_() == QDialog.Accepted:
            chart_sel = chart_dlg.selection()
            figs = select_and_build_charts(df, y, factors, results, chart_sel)

        # 7) формування Word-звіту (автоматично запропонувати збереження)
        default_name = f"SAD_звіт_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
        path, _ = QFileDialog.getSaveFileName(self, "Зберегти звіт", default_name, "Word (*.docx)")
        if not path:
            # якщо користувач не обрав — збережемо у тимчасову папку і повідомимо де
            tmp = os.path.join(tempfile.gettempdir(), default_name)
            export_report(df, title, unit, figs, results, tmp)
            QMessageBox.information(self, "Звіт збережено", f"Файл збережено: {tmp}")
        else:
            export_report(df, title, unit, figs, results, path)
            QMessageBox.information(self, "Готово", "Звіт збережено.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = MainWindow()
    w.resize(1100, 700)
    w.show()
    sys.exit(app.exec_())
