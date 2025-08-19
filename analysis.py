import pandas as pd
import numpy as np
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QPushButton, QListWidget, QMessageBox, QFileDialog
)
from scipy.stats import shapiro
from statsmodels.stats.multicomp import pairwise_tukeyhsd
import matplotlib.pyplot as plt
from docx import Document
import datetime


class AnalysisWindow(QDialog):
    def __init__(self, data: pd.DataFrame):
        super().__init__()
        self.data = data
        self.setWindowTitle("Аналіз даних")

        layout = QVBoxLayout()

        self.label = QLabel("Оберіть метод(и) аналізу:")
        layout.addWidget(self.label)

        self.listWidget = QListWidget()
        self.listWidget.addItems([
            "Перевірка нормальності (Шапіро-Вілка)",
            "Однофакторний дисперсійний аналіз",
            "Двофакторний дисперсійний аналіз",
            "Багатофакторний дисперсійний аналіз",
            "Дисперсійний аналіз з повторними вимірюваннями",
            "НІР05",
            "Тест Данканa",
            "Тест Тьюкі",
            "Тест Бонфероні",
            "Квадратичне відхилення",
            "Сила впливу факторів"
        ])
        self.listWidget.setSelectionMode(self.listWidget.MultiSelection)
        layout.addWidget(self.listWidget)

        self.runButton = QPushButton("Виконати аналіз")
        self.runButton.clicked.connect(self.run_analysis)
        layout.addWidget(self.runButton)

        self.setLayout(layout)

    def run_analysis(self):
        selected = [item.text() for item in self.listWidget.selectedItems()]
        if not selected:
            QMessageBox.warning(self, "Помилка", "Оберіть хоча б один метод аналізу")
            return

        # Створення Word-документу
        doc = Document()
        doc.add_heading("Результати статистичного аналізу", level=1)

        # Додаємо початкові дані
        doc.add_heading("Початкові дані", level=2)
        table = doc.add_table(rows=self.data.shape[0] + 1, cols=self.data.shape[1])
        table.style = 'LightShading-Accent1'
        # заголовки
        for j, col in enumerate(self.data.columns):
            table.cell(0, j).text = str(col)
        # дані
        for i in range(self.data.shape[0]):
            for j in range(self.data.shape[1]):
                table.cell(i + 1, j).text = str(self.data.iloc[i, j])

        doc.add_paragraph("Дата аналізу: " + datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        doc.add_paragraph("Програма: SAD — статистичний аналіз даних")

        # Приклад: якщо вибрали Шапіро
        if "Перевірка нормальності (Шапіро-Вілка)" in selected:
            doc.add_heading("Тест Шапіро-Вілка", level=2)
            for col in self.data.select_dtypes(include=[np.number]).columns:
                stat, p = shapiro(self.data[col].dropna())
                doc.add_paragraph(f"{col}: W={stat:.3f}, p={p:.3f}")

        # Збереження документа
        save_path, _ = QFileDialog.getSaveFileName(self, "Зберегти результати", "", "Word Files (*.docx)")
        if save_path:
            doc.save(save_path)
            QMessageBox.information(self, "Готово", f"Результати збережено у {save_path}")
