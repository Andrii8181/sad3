from docx import Document
from docx.shared import Inches
import pandas as pd

def _add_df_table(doc: Document, df: pd.DataFrame):
    table = doc.add_table(rows=1, cols=len(df.columns))
    hdr = table.rows[0].cells
    for i, c in enumerate(df.columns):
        hdr[i].text = str(c)
    for _, row in df.iterrows():
        cells = table.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = "" if pd.isna(v) else str(v)

def export_to_word(filename, title, units, raw_data, results: dict, figure_paths, author_footer=""):
    doc = Document()
    doc.add_heading(title, level=1)
    if units:
        doc.add_paragraph(f"Одиниці виміру: {units}")

    # Сирі дані
    doc.add_heading("Сирі дані", level=2)
    _add_df_table(doc, raw_data)

    # Результати аналізів
    doc.add_heading("Результати аналізів", level=2)
    for name, obj in results.items():
        doc.add_heading(str(name), level=3)
        if isinstance(obj, pd.DataFrame):
            _add_df_table(doc, obj.reset_index() if obj.index.name or obj.index.names else obj)
        else:
            doc.add_paragraph(str(obj))

    # Графіки
    if figure_paths:
        doc.add_heading("Графіки", level=2)
        for p in figure_paths:
            try:
                doc.add_picture(p, width=Inches(5.5))
            except Exception:
                pass

    if author_footer:
        doc.add_paragraph("")
        doc.add_paragraph(author_footer)

    doc.save(filename)
