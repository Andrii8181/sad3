import pandas as pd
import numpy as np
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem

def load_excel_or_csv(path: str) -> pd.DataFrame:
    if path.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(path)
    elif path.lower().endswith(".csv"):
        return pd.read_csv(path)
    else:
        raise ValueError("Підтримуються лише файли Excel (.xlsx/.xls) або CSV.")

def df_to_table(table: QTableWidget, df: pd.DataFrame):
    table.clear()
    table.setRowCount(max(1, df.shape[0]))
    table.setColumnCount(max(1, df.shape[1]))
    # встановити заголовки, якщо є
    if df.columns is not None:
        for j, c in enumerate(df.columns):
            table.setHorizontalHeaderItem(j, QTableWidgetItem(str(c)))
    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            table.setItem(i, j, QTableWidgetItem("" if pd.isna(df.iat[i, j]) else str(df.iat[i, j])))

def df_from_table(table: QTableWidget) -> pd.DataFrame:
    rows, cols = table.rowCount(), table.columnCount()
    data = []
    headers = [table.horizontalHeaderItem(c).text() if table.horizontalHeaderItem(c) else f"col{c+1}" for c in range(cols)]
    for r in range(rows):
        row = []
        empty_row = True
        for c in range(cols):
            it = table.item(r, c)
            val = "" if it is None else it.text()
            if val != "":
                empty_row = False
            row.append(val)
        if not empty_row:
            data.append(row)
    if not data:
        return pd.DataFrame(columns=headers)
    df = pd.DataFrame(data, columns=headers)
    # спробуємо перетворити числові стовпці
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="ignore")
    return df

def infer_roles(df: pd.DataFrame) -> dict:
    """
    Евристика:
      - 'Subject' (якщо є) — суб'єкт для RM-ANOVA
      - 'response' — перший числовий стовпець
      - 'factors' — нечислові стовпці + числові з ≤10 унікальних значень
      - 'group_for_np' — перший фактор для непараметричних тестів
    """
    subject = None
    for name in df.columns:
        if str(name).lower() in ["subject", "id", "subj", "повтор", "повторення"]:
            subject = name
            break

    numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    response = numeric_cols[0] if numeric_cols else None

    factors = []
    for c in df.columns:
        if c == response:
            continue
        series = df[c]
        if not pd.api.types.is_numeric_dtype(series):
            factors.append(c)
        else:
            uniq = series.dropna().unique()
            if len(uniq) <= 10:
                factors.append(c)

    group_for_np = factors[0] if factors else None

    return {"subject": subject, "response": response, "factors": factors, "group_for_np": group_for_np}
