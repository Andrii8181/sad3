from docx import Document
from docx.shared import Inches
import matplotlib.pyplot as plt
import pandas as pd
import datetime


def export_to_word(df, results: dict, filename="результати_аналізу.docx"):
    doc = Document()
    doc.add_heading("Результати статистичного аналізу", level=1)

    # сирі дані
    doc.add_heading("Сирі дані", level=2)
    table = doc.add_table(rows=df.shape[0] + 1, cols=df.shape[1])
    table.style = "Light List Accent 1"
    for j, col in enumerate(df.columns):
        table.cell(0, j).text = str(col)
    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            table.cell(i + 1, j).text = str(df.iat[i, j])

    # результати аналізів
    doc.add_heading("Результати аналізів", level=2)
    for name, res in results.items():
        doc.add_paragraph(f"{name}: {res}")

    # графік (приклад)
    plt.figure()
    df.apply(pd.to_numeric, errors="coerce").dropna().plot(kind="box")
    plt.title("Boxplot даних")
    plt.savefig("plot.png")
    doc.add_picture("plot.png", width=Inches(5))

    # дата та підпис
    doc.add_paragraph(f"\nДата: {datetime.date.today()}")
    doc.add_paragraph("SAD – Статистичний аналіз даних")

    doc.save(filename)
