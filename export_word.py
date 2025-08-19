from docx import Document
from docx.shared import Inches
from datetime import datetime
import pandas as pd

def _append_dataframe(doc: Document, df: pd.DataFrame, title=None):
    if title:
        doc.add_heading(title, level=2)
    if df is None or df.empty:
        p = doc.add_paragraph("— немає даних —")
        return
    tbl = doc.add_table(rows=len(df) + 1, cols=len(df.columns))
    hdr = tbl.rows[0].cells
    for j, col in enumerate(df.columns):
        hdr[j].text = str(col)
    for i in range(len(df)):
        row = tbl.rows[i + 1].cells
        for j, col in enumerate(df.columns):
            row[j].text = "" if pd.isna(df.iloc[i, j]) else str(df.iloc[i, j])

def export_report(df_raw: pd.DataFrame, indicator_name: str, unit: str,
                  figure_paths: list, results: dict, out_path: str):
    doc = Document()
    # заголовок
    h = doc.add_heading("SAD — Статистичний аналіз даних", 0)
    doc.add_paragraph(f"Показник: {indicator_name}" + (f" [{unit}]" if unit else ""))

    # сирі дані
    doc.add_heading("Початкові дані", level=1)
    _append_dataframe(doc, df_raw)

    # результати
    doc.add_heading("Результати аналізів", level=1)
    order = [
        ("means", "Узагальнення (середнє/SD/n)"),
        ("anova1", "Однофакторний ANOVA"),
        ("anova2", "Двофакторний ANOVA"),
        ("anovaN", "Багатофакторний ANOVA"),
        ("rm", "Повторні вимірювання (RM-ANOVA)"),
        ("lsd", "НІР (LSD, парні порівняння)"),
        ("tukey", "Пост-хок Тьюкі"),
        ("duncan", "Пост-хок Данкан"),
        ("bonf", "Пост-хок Бонферроні"),
        ("effect", "Сила впливу факторів (η²)"),
        ("mannwhitney", "Манна–Вітні"),
        ("kruskal", "Краскела–Уолліса"),
        ("friedman", "Фрідмана"),
        ("dunn", "Пост-хок Данна"),
        ("info", "Додаткова інформація")
    ]
    for key, title in order:
        if key in results and results[key] is not None:
            val = results[key]
            if isinstance(val, pd.DataFrame):
                _append_dataframe(doc, val.reset_index(drop=True), title)
            else:
                # statsmodels table — конвертуємо
                try:
                    df = val if isinstance(val, pd.DataFrame) else pd.DataFrame(val)
                    _append_dataframe(doc, df, title)
                except Exception:
                    doc.add_paragraph(f"{title}: {str(val)}")

    # графіки
    if figure_paths:
        doc.add_heading("Графічна візуалізація", level=1)
        for p in figure_paths:
            try:
                doc.add_picture(p, width=Inches(5.8))
            except Exception:
                pass

    # футер
    doc.add_paragraph()
    doc.add_paragraph(f"Виконано: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    doc.add_paragraph("Програма SAD — Статистичний аналіз даних")

    doc.save(out_path)
