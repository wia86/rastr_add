from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow

import sys

def add_label():
    print("нажал")
    print("нажал")
    print("нажал")
    print("нажал")

def run ():
    app = QApplication(sys.argv)
    window = QMainWindow()
    window.setWindowTitle("Макрос Иванович")
    window.setGeometry(300,250,350,200)

    text = QtWidgets.QLabel(window)
    text.setText("пошла")
    text.move(100,100)
    text.adjustSize()

    btn = QtWidgets.QPushButton(window)
    btn.move(70,150)
    btn.setText("Нажми")
    btn.setFixedWidth(200)
    btn.clicked.connect(add_label)



    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    run()