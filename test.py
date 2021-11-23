from PySide2 import QtWidgets
from calc_ui import Ui_MainWindow
import sys
import math

class Calculator(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        # Создание формы и Ui (наш дизайн)
        self.setupUi(self)
        self.show()
        self.lineEdit.setText('0')

        self.first_value = None
        self.second_value = None
        self.result = None
        self.example = ""
        self.equal = ""

        # pressed
        self.pushButton.clicked.connect(self.digit_pressed)  # 1
        self.pushButton_2.clicked.connect(self.digit_pressed)  # 2
        self.pushButton_3.clicked.connect(self.digit_pressed)  # 3
        self.pushButton_4.clicked.connect(self.digit_pressed)  # 4
        self.pushButton_5.clicked.connect(self.digit_pressed)  # 5
        self.pushButton_6.clicked.connect(self.digit_pressed)  # 6
        self.pushButton_7.clicked.connect(self.digit_pressed)  # 7
        self.pushButton_8.clicked.connect(self.digit_pressed)  # 8
        self.pushButton_9.clicked.connect(self.digit_pressed)  # 9
        self.pushButton_10.clicked.connect(self.digit_pressed)  # 0
        self.pushButton_add.clicked.connect(self.pressed_equal)  # +
        self.pushButton_ded.clicked.connect(self.pressed_equal)  # -
        self.pushButton_div.clicked.connect(self.pressed_equal)  # /
        self.pushButton_mul.clicked.connect(self.pressed_equal)  # *
        self.pushButton_exp.clicked.connect(self.pressed_equal)  # **
        self.pushButton_log.clicked.connect(self.pressed_equal)  # log
        self.pushButton_procent.clicked.connect(self.pressed_equal)  # %
        self.pushButton_ENTER.clicked.connect(self.function_result)  # =
        self.pushButton_C.clicked.connect(self.function_clear)  # C
        self.pushButton_point.clicked.connect(self.make_fractional)  # .
        self.pushButton_delete.clicked.connect(self.function_delete)  # <
        self.pushButton_open_skob.clicked.connect(self.create_big_example)  # (

    def digit_pressed(self):
        # sender - функция, которая возвращает отправителя сигнала (какая кнопка была нажата, от какой идет сигнал)
        button = self.sender()
        if self.lineEdit.text() == '0':
            self.lineEdit.setText(button.text())
        else:
            if self.result == self.lineEdit.text():
                self.lineEdit.setText(button.text())
            else:
                self.lineEdit.setText(self.lineEdit.text() + button.text())
        self.result = 0

    def form_result(self):
        self.result = str(self.result)
        if self.result[-2:] == '.0':
            self.result = self.result[:-2]
        self.lineEdit.setText(str(self.result))
        self.label.clear()

    def make_fractional(self):
        value = self.lineEdit.text()
        if '.' not in value:
            self.lineEdit.setText(value + '.')

    def function_delete(self):
        value = self.lineEdit.text()
        self.lineEdit.setText(value[:-1])

    def function_clear(self):
        self.lineEdit.setText('0')

    def pressed_equal(self):
        button = self.sender()
        self.first_value = float(self.lineEdit.text())
        self.lineEdit.clear()
        self.label.setText(str(self.first_value) + button.text())
        self.equal = button.text()

    def function_addition(self):
        self.determinate_second_value()
        self.result = float(self.first_value + self.second_value)
        self.form_result()

    def function_subtraction(self):
        self.determinate_second_value()
        self.result = float(self.first_value - self.second_value)
        self.form_result()

    def function_divison(self):
        self.determinate_second_value()
        self.result = float(self.first_value / self.second_value)
        self.form_result()

    def function_multiply(self):
        self.determinate_second_value()
        self.result = float(self.first_value * self.second_value)
        self.form_result()

    def function_exponentiation(self):
        self.determinate_second_value()
        self.result = float(self.first_value ** self.second_value)
        self.form_result()

    def function_percent(self):
        self.determinate_second_value()
        self.result = float(self.first_value * (self.second_value / 100))
        self.form_result()

    def function_log(self):
        self.determinate_second_value()
        self.result = float(math.log(self.first_value, self.second_value))
        self.form_result()

    def determinate_second_value(self):
        self.second_value = float(self.lineEdit.text())
        self.lineEdit.clear()
        self.label.setText(str(self.first_value) + self.equal + str(self.second_value))

    def function_result(self):
        if self.equal == '+':
            self.function_addition()
        elif self.equal == '-':
            self.function_subtraction()
        elif self.equal == "/":
            self.function_divison()
        elif self.equal == '*':
            self.function_multiply()
        elif self.equal == "^":
            self.exponentiation()
        elif self.equal == "%":
            self.function_percent()
        elif self.equal == "log":
            self.function_log()

if __name__ == '__main__':
    # Новый экземпляр QApplication
    app = QtWidgets.QApplication(sys.argv)
    # Сздание инстанса класса
    calc = Calculator()
    # Запуск
    sys.exit(app.exec_())