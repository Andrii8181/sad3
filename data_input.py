import pandas as pd
from PyQt5.QtWidgets import QTableWidgetItem

def import_dataframe_from_file(path: str) -> pd.DataFrame:
    if path.lower().endswith(".csv"):
        return pd.read_csv(path)
    return pd.read_excel(path)  # xlsx/xls (потрібен openpyxl)

def read_table_from_widget(table) -> pd.DataFrame:
    rows = table.rowCount()
    cols = table.columnCount()
    headers = []
    for c in range(cols):
        it = table.horizontalHeaderItem(c)
        headers.append(it.text() if it else f"Col{c+1}")
    data = []
    for r in range(rows):
        row = []
        empty = True
        for c in range(cols):
            it = table.item(r, c)
            txt = it.text().strip() if it else ""
            if txt != "":
                empty = False
            row.append(txt)
        if not empty:
            data.append(row)
    df = pd.DataFrame(data, columns=headers)
    return df

def write_df_to_widget(df: pd.DataFrame, table):
    table.clear()
    table.setRowCount(max(8, len(df)))
    table.setColumnCount(max(4, len(df.columns)))
    table.setHorizontalHeaderLabels([str(c) for c in df.columns] + [f"Колонка {i+1}" for i in range(len(df.columns), table.columnCount())])
    for r in range(len(df)):
        for c in range(len(df.columns)):
            table.setItem(r, c, QTableWidgetItem(str(df.iloc[r, c])))
