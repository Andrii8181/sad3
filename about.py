from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton

class AboutWindow(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Про програму")

        layout = QVBoxLayout()

        text = QLabel(
            "<h2>SAD — Статистичний аналіз даних</h2>"
            "<p>Програма для статистичного аналізу експериментальних даних "
            "(нормальність, дисперсійний аналіз, порівняння середніх, "
            "візуалізація результатів).</p>"
            "<p><b>Розробник:</b> Чаплоуцький А.М.<br>"
            "Кафедра плодівництва і виноградарства УНУ</p>"
            "<p>© 2025 Усі права захищені.</p>"
        )
        layout.addWidget(text)

        close_button = QPushButton("Закрити")
        close_button.clicked.connect(self.close)
        layout.addWidget(close_button)

        self.setLayout(layout)
