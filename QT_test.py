import main_console
from Excel_Helper import *
from main_console import *
from Excel_Helper_names import *
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtCore import QStandardPaths

import sys
import os

path_file_1 = None
path_file_2 = None
path = QStandardPaths.writableLocation(QStandardPaths.DesktopLocation)
os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
app = QtWidgets.QApplication(sys.argv)
Excel_Helper_names = QtWidgets.QMainWindow()
ui = Ui_Excel_Helper_names()
ui.setupUi(Excel_Helper_names)
Excel_Helper_names.show()

f = open('Names.txt', 'r')
names_1 = f.read()

def click():
    text = ui.textEdit.toPlainText()
    f = open('Names.txt', 'w')
    f.write(f'{text}')
    f.close()
    OpenOtherWindow()

ui.pushButton.clicked.connect(click)
ui.textEdit.setText(f"{names_1}")

def OpenOtherWindow():
    global MainWindow
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_Excel_Helper()
    ui.setupUi(MainWindow)
    Excel_Helper_names.close()
    MainWindow.show()


    def on_click_1():
        global path_file_1
        fileName_choose_1 = QFileDialog.getOpenFileName(None,
                                    "Выбрать файл 'ПОС'",
                                    f'{path}', # Начальный путь
                                    "Excel files (*.xlsx);;Excel files (*.xlsm);; Excel files (*.xltx);; Excel files (*.xltm)")[0]

        if fileName_choose_1 == "":
            print("\n Отменить выбор")
            return

        path_file_1 = fileName_choose_1

        print("\n Вы выбрали файл:")
        print(fileName_choose_1)

        file_name_1 = os.path.basename(f"{fileName_choose_1}").split("'")[0]

        ui.pushButton_1.hide()
        ui.label_1.setText(f"Выбран файл: {file_name_1}")

    def on_click_2():
        global path_file_2
        fileName_choose_2 = QFileDialog.getOpenFileName(None,
                                    "Выбрать файл 'Мониторинг'",
                                    f'{path}', # Начальный путь
                                    "Excel files (*.xlsx);;Excel files (*.xlsm);; Excel files (*.xltx);; Excel files (*.xltm)")[0]

        if fileName_choose_2 == "":
            print("\n Отменить выбор")
            return

        path_file_2 = fileName_choose_2

        print("\n Вы выбрали файл:")
        print(fileName_choose_2)

        file_name_2 = os.path.basename(f"{fileName_choose_2}").split("'")[0]

        ui.pushButton_2.hide()
        ui.label_2.setText(f"Выбран файл: {file_name_2}")


    def on_click_go():
        global path_file_1, path_file_2

        if path_file_1 is not None or path_file_2 is not None:
            ui.pushButton_4.hide()
            main_go()
        else:
            return


    def main_go():
        start_time = time.time()
        ui.pushButton_4.hide()
        main_console.main(path_file_1, path_file_2)
        ui.label_5.show()
        ui.label_8.show()
        ui.label_8.setText("Время выполнения программы: %s минут" % round(((time.time() - start_time) / 60), 2))

    cwd = os.getcwd()

    ui.label_5.hide()
    ui.label_8.hide()

    ui.pushButton_1.clicked.connect(on_click_1)
    ui.pushButton_2.clicked.connect(on_click_2)
    ui.pushButton_4.clicked.connect(on_click_go)


sys.exit(app.exec_())