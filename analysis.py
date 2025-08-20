import pandas as pd
from PyQt5.QtWidgets import QMessageBox
from scipy import stats
import matplotlib.pyplot as plt
import seaborn as sns
from export_word import export_to_word

def run_analysis(table_widget):
    # Зчитуємо дані з таблиці
    rows = table_widget.rowCount()
    cols = table_widget.columnCount()
    headers = [table_widget.horizontalHeaderItem(i).text() for i in range(cols)]
    data = []

    for i in range(rows):
        row_data = []
        for j in range(cols):
            item = table_widget.item(i, j)
            try:
                row_data.append(float(item.text()) if item else None)
            except:
                row_data.append(None)
        data.append(row_data)

    df = pd.DataFrame(data, columns=headers)

    # Автоматична перевірка на нормальність
    normal_results = {}
    for col in df.columns:
        col_data = df[col].dropna()
        if len(col_data) > 3:
            stat, p = stats.shapiro(col_data)
            normal_results[col] = (stat, p)

    # Якщо дані не нормальні → повідомлення
    for col, (stat, p) in normal_results.items():
        if p < 0.05:
            QMessageBox.warning(None, "Перевірка нормальності", 
                f"Дані у стовпчику '{col}' не відповідають нормальному розподілу (p={p:.3f}).\n"
                f"Рекомендується використати непараметричні методи.")
            return

    # Якщо нормальні → будуємо графік і експортуємо
    plt.figure(figsize=(6,4))
    sns.boxplot(data=df)
    plt.title("Розподіл даних")
    plt.savefig("graph.png")
    plt.close()

    # Експорт у Word
    export_to_word(df, "Результати аналізу", "од.", {"Нормальність": normal_results}, "graph.png")

    QMessageBox.information(None, "Аналіз завершено", "Результати збережено у Word.")
