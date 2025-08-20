from PyQt5.QtWidgets import QMessageBox

def show_about():
    QMessageBox.information(None, "Про програму", 
        "SAD - Статистичний аналіз даних\n\n"
        "Розробник: Чаплоуцький А.М.\n"
        "Кафедра плодівництва і виноградарства УНУ\n\n"
        "© Усі права захищені. Ліцензія лише для наукового використання.")
