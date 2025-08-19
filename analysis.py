import sys
import scipy.stats as stats
import statsmodels.api as sm
from statsmodels.formula.api import ols
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QPushButton, QCheckBox, QMessageBox
)
from export import export_to_word


class AnalysisDialog(QDialog):
    def __init__(self, df, parent=None):
        super().__init__(parent)
        self.df = df
        self.setWindowTitle("Вибір аналізу")
        self.setMinimumWidth(400)

        layout = QVBoxLayout()
        self.label = QLabel("Автоматична перевірка даних на нормальність...")
        layout.addWidget(self.label)

        # перевірка нормальності Шапіро
        try:
            values = self.df.apply(pd.to_numeric, errors="coerce").dropna().values.flatten()
            stat, p = stats.shapiro(values)
            if p < 0.05:
                QMessageBox.warning(self, "Ненормальний розподіл",
                                    "Дані не відповідають нормальному розподілу!\n"
                                    "Використовуйте непараметричні аналізи.")
        except Exception as e:
            QMessageBox.warning(self, "Помилка", f"Неможливо виконати перевірку: {e}")

        # список аналізів
        self.checks = []
        analyses = [
            "Однофакторний дисперсійний",
            "Двофакторний дисперсійний",
            "Багатофакторний дисперсійний",
            "Дисперсійний з повторними вимірюваннями",
            "НІР05",
            "Тест Данканa",
            "Тест Тьюкі",
            "Тест Бонфероні",
            "Сила впливу факторів",
            "Квадратичне відхилення",
            "Візуалізація результатів"
        ]
        for a in analyses:
            cb = QCheckBox(a)
            layout.addWidget(cb)
            self.checks.append(cb)

        # кнопка виконання
        run_btn = QPushButton("Виконати аналіз")
        run_btn.clicked.connect(self.run_analysis)
        layout.addWidget(run_btn)

        self.setLayout(layout)

    def run_analysis(self):
        selected = [cb.text() for cb in self.checks if cb.isChecked()]
        if not selected:
            QMessageBox.warning(self, "Помилка", "Оберіть хоча б один аналіз")
            return

        # тут вставимо виклики розрахунків (зараз простий шаблон)
        results = {}
        for method in selected:
            results[method] = f"Результати методу: {method} (тут буде формула/таблиця)"

        export_to_word(self.df, results)
        QMessageBox.information(self, "Готово", "Результати збережено у Word!")
        self.close()
