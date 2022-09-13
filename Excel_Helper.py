# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'Excel_Helper.ui'
#
# Created by: PyQt5 UI code generator 5.15.7
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QFileDialog

class Ui_Excel_Helper(object):
    def setupUi(self, Excel_Helper):
        Excel_Helper.setObjectName("Excel_Helper")
        Excel_Helper.setEnabled(True)
        Excel_Helper.setFixedSize(485, 371)
        Excel_Helper.setBaseSize(QtCore.QSize(500, 500))
        Excel_Helper.setStyleSheet("background-color: rgb(23, 23, 23)")
        self.centralwidget = QtWidgets.QWidget(Excel_Helper)
        self.centralwidget.setObjectName("centralwidget")
        self.label_1 = QtWidgets.QLabel(self.centralwidget)
        self.label_1.setGeometry(QtCore.QRect(10, 30, 461, 41))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(18)
        self.label_1.setFont(font)
        self.label_1.setStyleSheet("color: rgb(255, 255, 255)")
        self.label_1.setObjectName("label_1")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(10, 90, 461, 41))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(18)
        self.label_2.setFont(font)
        self.label_2.setStyleSheet("color: rgb(255, 255, 255)")
        self.label_2.setObjectName("label_2")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(10, 90, 461, 41))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(18)
        self.pushButton_2.setFont(font)
        self.pushButton_2.setStyleSheet("background-color: rgb(181, 255, 166)")
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_1 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_1.setGeometry(QtCore.QRect(10, 40, 461, 41))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(18)
        self.pushButton_1.setFont(font)
        self.pushButton_1.setStyleSheet("background-color: rgb(181, 255, 166)")
        self.pushButton_1.setObjectName("pushButton_1")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(10, 260, 461, 31))
        font = QtGui.QFont()
        font.setPointSize(18)
        font.setBold(True)
        font.setWeight(75)
        self.label_5.setFont(font)
        self.label_5.setStyleSheet("color: rgb(107, 255, 107);")
        self.label_5.setAlignment(QtCore.Qt.AlignCenter)
        self.label_5.setObjectName("label_5")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(10, 150, 461, 31))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.label_6.setFont(font)
        self.label_6.setStyleSheet("color: rgb(255, 255, 255);")
        self.label_6.setAlignment(QtCore.Qt.AlignCenter)
        self.label_6.setObjectName("label_6")
        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setGeometry(QtCore.QRect(10, 150, 461, 41))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(18)
        self.pushButton_4.setFont(font)
        self.pushButton_4.setStyleSheet("background-color: rgb(255, 214, 149)")
        self.pushButton_4.setObjectName("pushButton_4")
        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        self.label_7.setGeometry(QtCore.QRect(400, 331, 71, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.label_7.setFont(font)
        self.label_7.setStyleSheet("color: rgb(57, 57, 57);")
        self.label_7.setAlignment(QtCore.Qt.AlignCenter)
        self.label_7.setObjectName("label_7")
        self.label_8 = QtWidgets.QLabel(self.centralwidget)
        self.label_8.setGeometry(QtCore.QRect(10, 290, 461, 31))
        font = QtGui.QFont()
        font.setPointSize(16)
        font.setBold(False)
        font.setWeight(50)
        self.label_8.setFont(font)
        self.label_8.setStyleSheet("color: rgb(255, 255, 255);")
        self.label_8.setAlignment(QtCore.Qt.AlignCenter)
        self.label_8.setObjectName("label_8")
        Excel_Helper.setCentralWidget(self.centralwidget)

        self.retranslateUi(Excel_Helper)
        QtCore.QMetaObject.connectSlotsByName(Excel_Helper)

    def retranslateUi(self, Excel_Helper):
        _translate = QtCore.QCoreApplication.translate
        Excel_Helper.setWindowTitle(_translate("Excel_Helper", "Excel Helper"))
        self.label_1.setText(_translate("Excel_Helper", "Выбранный файл: отсутствует"))
        self.label_2.setText(_translate("Excel_Helper", "Выбранный файл: отсутствует"))
        self.pushButton_2.setText(_translate("Excel_Helper", "Выберите файл \'Мониторинг\'"))
        self.pushButton_1.setText(_translate("Excel_Helper", "Выберите файл \'ПОС\'"))
        self.label_5.setText(_translate("Excel_Helper", "Файлы успешно сохранены!"))
        self.label_6.setText(_translate("Excel_Helper", "Программа закончила работу!"))
        self.pushButton_4.setText(_translate("Excel_Helper", "Запустить"))
        self.label_7.setText(_translate("Excel_Helper", "By Serezha M"))
        self.label_8.setText(_translate("Excel_Helper", "Время выполнения программы: %s минут"))


if __name__ == "__main__":
    import sys
    import os
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    app = QtWidgets.QApplication(sys.argv)
    Excel_Helper = QtWidgets.QMainWindow()
    ui = Ui_Excel_Helper()
    ui.setupUi(Excel_Helper)
    Excel_Helper.show()
    sys.exit(app.exec_())