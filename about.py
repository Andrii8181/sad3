from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton

class AboutDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Про програму")
        self.setGeometry(300, 200, 400, 200)

        layout = QVBoxLayout()
        text = QLabel(
            "SAD – Статистичний аналіз даних\n\n"
            "Розробник: Чаплоуцький А.М.\n"
            "Кафедра плодівництва і виноградарства УНУ\n"
            "Всі права захищені. Програма поширюється під академічною ліцензією.\n"
        )
        layout.addWidget(text)

        close_btn = QPushButton("Закрити")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)

        self.setLayout(layout)
