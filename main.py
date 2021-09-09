from PyQt5 import QtWidgets
from PyQt5 import QtCore
from PyQt5 import QtGui
from PyQt5.QtGui import QIcon
import time

from qtable import Ui_MainWindow  # Наше окно
from help import Ui_help_ui  # Наше ресурсcomboBox_list
from line_config import Ui_MainWindowList
from procedures import *
import sys

import xlwings as xw
import xlrd
import xlwt
import math
import sys
import gc
from PyQt5.QtWidgets import QMessageBox
import os
import re
import traceback
import ast
import win32api

class line_config_window(QtWidgets.QMainWindow):
    return_value = QtCore.pyqtSignal()
    ret = []
    names_condig = []
    weight = ["Нет"]

    def __init__(self, massiv = [""]):
        self.ret = [massiv [0]]
        self.names_condig = massiv[1]

        super(line_config_window, self).__init__()
        self.setWindowFlag(QtCore.Qt.WindowStaysOnTopHint)  # Поверх других окон

        self.setWindowFlags(QtCore.Qt.WindowCloseButtonHint)  # Кпока закрыть
        self.ui = Ui_MainWindowList()  # инициализация
        self.ui.setupUi(self)  # инициализация
        self.ui.save_config_button.clicked.connect(self.close1)

    def error_msg(self, message, title = "Ошибка", icon = QMessageBox.Warning):
        msg = QMessageBox()
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.setIcon(icon)
        msg.exec_()
    #
    def close1(self):
        self.ret.append(self.ui.lineConfig.text())
        if self.ret[-1] == '':
            self.error_msg("Пустое поле не допустимо!")
            return

        i = 1

        while self.ret[-1] in self.names_condig:
            if "(" in self.ret[-1]:
                self.ret[-1] = "".join(self.ret[-1].split("(")[:-1]) + "(" + str(i) + ")"
            else:
                self.ret[-1] = self.ret[-1] + "(" + str(i) + ")"
            i += 1
        self.return_value.emit()
        self.close()

        # print(self.ui.weight)
        # print("KOKOKOKO")


# наш класс главного окна
class mywindow(QtWidgets.QMainWindow):
    switch_window = QtCore.pyqtSignal()
    switch_window_list = QtCore.pyqtSignal()

    array_export = []  # Не используется
    QT_moduls = []
    id_modul = 0

    # Файл 1
    table_1 = []
    path_file_1 = [""]
    name_file_1 = [""]
    files_size = []
    config_files_1 = []
    list_of_keys = []  # Левая колонка 1
    list_of_keys_flags = []  # Правая колонка 1(позиционирование ключей в таблице)
    list_of_keys_selected = []  # Выбранные поля справа(имена ключей)
    sheets_1 = ""

    # Файл 2
    table_2 = []  # Вроде не юзал
    path_file_2 = [""]
    name_file_2 = [""]
    files_size_2 = []
    config_files_2 = []
    list_of_keys_2 = []  # Левая колонка 1
    list_of_keys_flags_2 = []  # Правая колонка 1, Выбраные ключи !!!! Устарела!! Надо удалять
    list_of_keys_selected_2 = []  # Выбранные поля справа  (имена ключей)
    sheets_2 = ""

    list_of_keys_selected_3 = []  # Выбранные поля справа (имена ключей)

    # Доп блоки для Вычислительного режима
    list_of_keys_selected_4 = []
    list_of_keys_selected_5 = []

    # Фильтер 1-2
    filter = [{}, {}]  # ВЫбранные элкменты фильтра

    # Шаг
    step = 1  # Номер шага(заюзаной функции)

    #Конфигурации
    config = []

    def __init__(self):

        # something default
        super(mywindow, self).__init__()
        self.setWindowFlags(QtCore.Qt.Window)  #
        # self.setWindowFlags(QtCore.Qt.WindowCloseButtonHint)  # Кпока закрыть
        self.ui = Ui_MainWindow()  # инициализация
        self.ui.setupUi(self)  # инициализация

        self.ui.help_button.clicked.connect(self.help)  # Нажатие кнопки help

        self.QT_moduls = [
            {"list_table_selected":self.ui.list_table_selected,
            "button_open_file": self.ui.button_open_file,
             "list_keys": self.ui.list_keys,
             "list_selected_keys": self.ui.list_selected_keys,
             "pushButton_add": self.ui.pushButton_add,
             "pushButton_delete": self.ui.pushButton_delete,
             "pushButton_up": self.ui.pushButton_up,
             "pushButton_down": self.ui.pushButton_down,
             # "label_file_name": self.ui.label_file_name,
             "comboBox_list": self.ui.comboBox_list_1,
             "spinBox_list": self.ui.spinBox_list,
             "comboBox_head": self.ui.comboBox_head,
             "comboBox_config": self.ui.comboBox_config
             },
            {"list_table_selected":self.ui.list_table_selected_2,
             "button_open_file": self.ui.button_open_file_2,
             "list_keys": self.ui.list_keys_2,
             "list_selected_keys": self.ui.list_selected_keys_2,
             "pushButton_add": self.ui.pushButton_add_2,
             "pushButton_delete": self.ui.pushButton_delete_2,
             "pushButton_up": self.ui.pushButton_up_2,
             "pushButton_down": self.ui.pushButton_down_2,
             # "label_file_name": self.ui.label_file_name_2,
             "comboBox_list": self.ui.comboBox_list_2,
             "spinBox_list": self.ui.spinBox_list_2,
             "comboBox_head": self.ui.comboBox_head_2,
             "comboBox_config": self.ui.comboBox_config_2
             },
            {
                "list_keys": self.ui.list_keys_2,
                "list_selected_keys": self.ui.list_selected_keys_3
            },
            {
                "list_keys": self.ui.list_keys,
                "list_selected_keys": self.ui.list_selected_keys_4
            },
            {
                "list_keys": self.ui.list_keys_2,
                "list_selected_keys": self.ui.list_selected_keys_5
            }

        ]

        self.QT_filter_moduls = [
            {
                "comboBox_keys": self.ui.comboBox_keys,
                "list_unique_keys": self.ui.list_unique_keys,
                "lineEdit_filter": self.ui.lineEdit_filter,
                "list_selected_unique_keys": self.ui.list_selected_unique_keys,
                "checkBox_filter": self.ui.checkBox_filter
            },
            {
                "comboBox_keys": self.ui.comboBox_keys_2,
                "list_unique_keys": self.ui.list_unique_keys_2,
                "lineEdit_filter": self.ui.lineEdit_filter_2,
                "list_selected_unique_keys": self.ui.list_selected_unique_keys_2,
                "checkBox_filter": self.ui.checkBox_filter_2
            }
        ]

        # Файл 1
        # button to open files
        self.ui.button_open_file.clicked.connect(lambda: self.startToListen(0))  # Нажатие кнопки открыть файл
        # События колонки 1-2 (Файл 1)
        self.ui.pushButton_add.clicked.connect(lambda: self.Add_key(0))
        self.ui.list_keys.doubleClicked.connect(lambda: self.Add_key(0))
        # действие по нажатию кнопки удалить
        self.ui.pushButton_delete.clicked.connect(lambda: self.Delete_key(0))
        self.ui.list_selected_keys.doubleClicked.connect(lambda: self.Delete_key(0))
        # Стрелочки
        self.ui.pushButton_up.clicked.connect(lambda: self.pushButton_up(0))
        self.ui.pushButton_down.clicked.connect(lambda: self.pushButton_down(0))
        self.ui.comboBox_list_1.currentIndexChanged.connect(lambda: self.choice_case_list(0))

        # Спинбокс
        self.ui.spinBox_list.valueChanged.connect(lambda: self.spinBox_list(0))
        # Параметр шапки
        self.ui.comboBox_head.currentIndexChanged.connect(lambda: self.choice_case_list(0))

        # Допкопки
        self.ui.pushButton_add_4.clicked.connect(lambda: self.Add_key(3))
        self.ui.pushButton_delete_4.clicked.connect(lambda: self.Delete_key(3))
        self.ui.pushButton_up_4.clicked.connect(lambda: self.pushButton_up(3))
        self.ui.pushButton_down_4.clicked.connect(lambda: self.pushButton_down(3))
        self.ui.list_selected_keys_4.doubleClicked.connect(lambda: self.Delete_key(3))

        # Автободбор начала таблицы
        self.ui.button_auto_select_str.clicked.connect(lambda: self.Auto_select_str(0))
        #Выбор конфигурации file 1
        self.ui.comboBox_config.currentIndexChanged.connect(lambda: self.config_selected(0))
        #Сохранение конфигурации file 1
        self.ui.pushButton_add_config.clicked.connect(lambda: self.go_comboBox_config(0))
        #Удаление конфигурации file 1
        self.ui.pushButton_filter_del_config.clicked.connect(lambda: self.del_config(0))

        # Файл 2
        # button to open files
        self.ui.button_open_file_2.clicked.connect(lambda: self.startToListen(1))  # Нажатие кнопки открыть файл
        # События колонки 1-2 (Файл 1)
        self.ui.pushButton_add_2.clicked.connect(lambda: self.Add_key(1))
        self.ui.list_keys_2.doubleClicked.connect(lambda: self.Add_key(1))
        # действие по нажатию кнопки удалить
        self.ui.pushButton_delete_2.clicked.connect(lambda: self.Delete_key(1))
        self.ui.list_selected_keys_2.doubleClicked.connect(lambda: self.Delete_key(1))

        # Стрелочки
        self.ui.pushButton_up_2.clicked.connect(lambda: self.pushButton_up(1))
        self.ui.pushButton_down_2.clicked.connect(lambda: self.pushButton_down(1))
        self.ui.comboBox_list_2.currentIndexChanged.connect(lambda: self.choice_case_list(1))

        # Спинбокс file 2
        self.ui.spinBox_list_2.valueChanged.connect(lambda: self.spinBox_list(1))
        # Параметр шапки file 2
        self.ui.comboBox_head_2.currentIndexChanged.connect(lambda: self.choice_case_list(1))

        # Допкопки file 2
        self.ui.pushButton_add_5.clicked.connect(lambda: self.Add_key(4))
        self.ui.pushButton_delete_5.clicked.connect(lambda: self.Delete_key(4))
        self.ui.pushButton_up_5.clicked.connect(lambda: self.pushButton_up(4))
        self.ui.pushButton_down_5.clicked.connect(lambda: self.pushButton_down(4))
        self.ui.list_selected_keys_5.doubleClicked.connect(lambda: self.Delete_key(4))

        # Автободбор начала таблицы file 2
        self.ui.button_auto_select_str_2.clicked.connect(lambda: self.Auto_select_str(1))
        # Выбор конфигурации file 2
        self.ui.comboBox_config_2.currentIndexChanged.connect(lambda: self.config_selected(1))
        # Сохранение конфигурации file 2
        self.ui.pushButton_add_config_2.clicked.connect(lambda: self.go_comboBox_config(1))
        # Удаление конфигурации file 1
        self.ui.pushButton_filter_del_config_2.clicked.connect(lambda: self.del_config(1))

        # действие по нажатию кнопок 3 серии !!!
        self.ui.list_selected_keys_3.doubleClicked.connect(lambda: self.Delete_key(2))
        self.ui.pushButton_add_3.clicked.connect(lambda: self.Add_key(2))
        self.ui.pushButton_delete_3.clicked.connect(lambda: self.Delete_key(2))
        self.ui.pushButton_up_3.clicked.connect(lambda: self.pushButton_up(2))
        self.ui.pushButton_down_3.clicked.connect(lambda: self.pushButton_down(2))

        # Поле фильтра 1
        # Клик по Комбобоксу (filter 1)
        self.ui.comboBox_keys.currentIndexChanged.connect(lambda: self.choice_case_filter(0))

        self.ui.pushButton_add_unique_keys.clicked.connect(lambda: self.Add_key_filter(0))
        self.ui.list_unique_keys.doubleClicked.connect(lambda: self.Add_key_filter(0))

        self.ui.pushButton_delete_unique_keys.clicked.connect(lambda: self.Delete_key_filter(0))
        self.ui.list_selected_unique_keys.doubleClicked.connect(lambda: self.Delete_key_filter(0))

        # Подгрузить (filter 1)
        self.ui.pushButton_filter_download_keys.clicked.connect(lambda: self.filter_download_keys(0))
        # Удалить (filter 1)
        self.ui.pushButton_filter_clear_keys.clicked.connect(lambda: self.filter_delete_keys(0))

        # Поле фильтра 2
        # Клик по Комбобоксу (filter 2)
        self.ui.comboBox_keys_2.currentIndexChanged.connect(lambda: self.choice_case_filter(1))

        self.ui.pushButton_add_unique_keys_2.clicked.connect(lambda: self.Add_key_filter(1))
        self.ui.list_unique_keys_2.doubleClicked.connect(lambda: self.Add_key_filter(1))

        self.ui.pushButton_delete_unique_keys_2.clicked.connect(lambda: self.Delete_key_filter(1))
        self.ui.list_selected_unique_keys_2.doubleClicked.connect(lambda: self.Delete_key_filter(1))

        # Подгрузить (filter 2)
        self.ui.pushButton_filter_download_keys_2.clicked.connect(lambda: self.filter_download_keys(1))
        # Удалить (filter 2)
        self.ui.pushButton_filter_clear_keys_2.clicked.connect(lambda: self.filter_delete_keys(1))

        # result
        self.ui.result_button.clicked.connect(lambda: self.progress())

        # Поле старта
        self.ui.comboBox_function.currentIndexChanged.connect(self.choice_comboBox_function)

        # Скрываем модули разрабочкика
        self.ui.keys_label_calculations.hide()
        self.ui.keys_label_calculations_2.hide()
        self.ui.pushButton_delete_4.hide()
        self.ui.pushButton_add_4.hide()
        self.ui.pushButton_up_4.hide()
        self.ui.pushButton_down_4.hide()
        self.ui.list_selected_keys_4.hide()
        self.ui.pushButton_delete_5.hide()
        self.ui.pushButton_add_5.hide()
        self.ui.pushButton_up_5.hide()
        self.ui.pushButton_down_5.hide()
        self.ui.list_selected_keys_5.hide()

        # Удалять помеченные на удаление по умолчанию
        self.ui.comboBox_del.setCurrentIndex(2)

        # Правильная загрузка окон
        self.choice_comboBox_function()

        #Закачка файла конфигурации
        self.download_config()

    def logger(self, msg = "", level = 0):
        log_file = open('log.txt', 'a')
        if level == 0:
            log_file.write(msg.replace("\u0306","") + '\n')
        else:
            try:
                msg += "\nФайл(ы) 1: '" + str(self.name_file_1) + "'  Размер: " + str(sum(self.files_size)) + " MB\n"
                msg += "Файл(ы) 2: '" + str(self.name_file_2) + "'  Размер: " + str(sum(self.files_size_2)) + " MB\n"
                write_conf1 = "\n".join(self.write_format_conf(self.create_config(0, "Параметры первого файла"))[:-2].split("\n")[:-1])
                write_conf2 = "\n".join(
                    self.write_format_conf(self.create_config(1, "Параметры второго файла"))[:-2].split("\n")[:-1])
                msg += '_' * 103 + "\n" + write_conf1 + "\n" + '_' * 103 + "\n" + write_conf2 + "\n" + '_' * 103
                msg += "\nХод работы:\n"
            except BaseException:
                msg = "Ошибка логирования!\n"
            log_file.write(msg.replace("\u0306", "") + '\n')

        log_file.close()

        try:
            global_logger_name = os.environ.get("USERNAME")
            try:
                global_logger_name += " (" + win32api.GetUserNameEx(3) + ")"
            except BaseException:
                pass
            global_logger_path = "V:\\Обмен МБУ ФК\\23_Отдел НСИ\\Обработчик\\logger\\"
            global_log_file = open(global_logger_path + global_logger_name + ".txt", 'a')
            global_log_file.write(msg.replace("\u0306", "") + '\n')
            global_log_file.close()
        except BaseException:
            pass


    def progressBarMsg(self, value, msg = ""):
        t1 = str(time.strftime("%H:%M:%S", time.localtime(time.time())))
        msg_write = t1 + " - " + msg
        print(msg_write)

        self.ui.progressBar.setValue(value)
        self.ui.progressBarMsg.setText(msg)
        self.logger(msg_write)
        QtWidgets.qApp.processEvents()

    def error_msg(self, message, title = "Ошибка", icon = QMessageBox.Warning):
        self.progressBarMsg(0, 'Готов в анализу')
        self.logger(message)
        msg = QMessageBox()
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.setIcon(icon)
        msg.exec_()


    def warning_select_msg(self, msg = "", title = "Предупреждение"):
        buttonReply = QMessageBox.question(self, title,  msg,
                                           QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel, QMessageBox.Cancel)
        if buttonReply == QMessageBox.Yes:
            return 'Yes'
        if buttonReply == QMessageBox.No:
            return 'No'
        if buttonReply == QMessageBox.Cancel:
            return 'Cancel'

    def startToListen(self, id_modul):  # Нажатие кнопки открыть файл
        self.step = self.step + 1
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self,
            'Open File',
            './',
             'Excel files (*.xm*;*.xl*);;All files (*)')

        if len(files) == 0:
            self.progressBarMsg(0, 'Файл не выбран')
            return

        names_files = [file.split("/")[-1] for file in files]

        self.ui.progressBar.setValue(0)
        self.ui.progressBarMsg.setText('Выбор файла')

        self.QT_moduls[id_modul]["list_keys"].clear()
        self.QT_moduls[id_modul]["list_selected_keys"].clear()


        self.QT_filter_moduls[id_modul]["comboBox_keys"].clear()
        self.QT_filter_moduls[id_modul]["comboBox_keys"].addItem("Нет")
        self.QT_filter_moduls[id_modul]["list_unique_keys"].clear()
        self.QT_filter_moduls[id_modul]["list_selected_unique_keys"].clear()

        if id_modul == 0:
            self.path_file_1 = files
            self.name_file_1 = names_files
        else:
            self.path_file_2 = files
            self.name_file_2 = names_files

        # если файл был открыт
        # меняем строку label на название файла
        self.restart_list(self.QT_moduls[id_modul]["list_table_selected"], names_files)
        self.progressBarMsg(0, 'Загрузка..')
        self.id_modul = id_modul

        # filesize_ratio = 0.85
        filesize_ratio = 1.80
        files_size = []
        for file in files:
            file_stats = os.stat(file)
            files_size.append(int(file_stats.st_size / (1024 * 1024) * 100) / 100)

        t1 = time.strftime("%M мин %S сек", time.localtime(sum(files_size) * filesize_ratio))
        t2 = time.strftime("%H:%M:%S", time.localtime(time.time() + sum(files_size) * filesize_ratio))

        if(len(files) == 1):
            self.progressBarMsg(0, f'Размер выбранного файла {sum(files_size)} МВ, '
                               f'примерное время загрузки {t1}({t2})')
        else:
            self.progressBarMsg(0, f'Размер выбранных файлов {sum(files_size)} МВ, '
                                   f'примерное время загрузки {t1}({t2})')

        if (id_modul == 0 and self.ui.checkBox_group_str.isChecked() or
                id_modul == 1 and self.ui.checkBox_group_str_2.isChecked()):
            self.progressBarMsg(0, f'Используется анализ группированных строк. '
                                f'Время выполнения будет выше заявленного!')

        if (id_modul == 0):
            self.table_1 = []
            self.files_size = []
            for i in range(len(files)):
                self.progressBarMsg(0, "Загрузка файла '" + names_files[i] + "'" + f". Примерное время до завершения всех загрузок {t1}({t2})")
                self.files_size.append(files_size[i])

                if (files[i].split(".")[-1] == "xlsx"):
                    if self.ui.checkBox_group_str.isChecked():
                        res = creat_level_tables(files[i])
                        file_path = write_table_defolt(res[0], res[1], ".".join(names_files[i].split(".")[:-1]))
                        self.table_1.append(xlrd.open_workbook(file_path))
                    else:
                        self.table_1.append(xlrd.open_workbook(files[i]))

                else:
                    self.table_1.append(xlrd.open_workbook(files[i], formatting_info=True))
                self.progressBarMsg(0, "Файл "+ names_files[i] + " закачан")
            self.sheets_1 = self.table_1[0].sheet_names()

            self.ui.comboBox_list_1.currentIndexChanged.disconnect()
            self.ui.comboBox_list_1.clear()
            if (len(self.sheets_1) > 1):
                self.ui.comboBox_list_1.addItem("Лист не выбран")
                for el in self.sheets_1:
                    self.ui.comboBox_list_1.addItem(el)
            else:
                self.ui.comboBox_list_1.addItem(self.sheets_1[0])
                self.choice_case_list(0)

            self.ui.comboBox_list_1.currentIndexChanged.connect(lambda: self.choice_case_list(0))

            self.search_config(0)
        else:
            self.table_2 = []
            self.files_size_2 = []
            for i in range(len(files)):
                self.progressBarMsg(0, "Загрузка файла '" + names_files[i] + "'" + f". Примерное время до завершения всех загрузок {t1}({t2})")
                self.files_size_2.append(files_size[i])
                if (files[i].split(".")[-1] == "xlsx"):
                    if self.ui.checkBox_group_str_2.isChecked():
                        res = creat_level_tables(files[i])
                        file_path = write_table_defolt(res[0], res[1], ".".join(names_files[i].split(".")[:-1]))
                        self.table_2.append(xlrd.open_workbook(file_path))
                    else:
                        self.table_2.append(xlrd.open_workbook(files[i]))
                else:
                    self.table_2.append(xlrd.open_workbook(files[i], formatting_info=True))
                self.table_2.append(xlrd.open_workbook(files[i]))
            self.progressBarMsg(0, "Файл "+ names_files[i] +" закачан")
            self.sheets_2 = self.table_2[0].sheet_names()

            self.ui.comboBox_list_2.currentIndexChanged.disconnect()
            self.ui.comboBox_list_2.clear()
            if (len(self.sheets_2) > 1):
                self.ui.comboBox_list_2.addItem("Лист не выбран")
                for el in self.sheets_2:
                    self.ui.comboBox_list_2.addItem(el)
            else:
                self.ui.comboBox_list_2.addItem(self.sheets_2[0])
                self.choice_case_list(1)

            self.ui.comboBox_list_2.currentIndexChanged.connect(lambda: self.choice_case_list(1))

            self.search_config(1)
        # self.switch_window_list.emit()

    #Проверка таблиц на соответвии конфигурации
    def proverka_tables(self, id_modul, list, start, keys_config):
        # if self.QT_moduls[id_modul]["comboBox_head"].currentText() == "Нет":
        #     return True
        # print(keys_config)
        try:
            head = False
            for i in range(len(keys_config)):
                if str(i + 1) != keys_config[i]:
                    head = True
                    break

            if id_modul == 0:
                for i in range(len(self.table_1)):
                    keys_table = Analis_Table_keys(self.table_1[i], self.sheets_1.index(list),
                                                   start)
                    if head:
                        err = 0
                        for j in range(len(keys_config)):
                            if keys_config[j] in keys_table and j == keys_table.index(keys_config[j]):
                                pass
                            else:
                                err += 1
                        if err > 3:
                            return False
                    else:
                        if abs(len(keys_table) - len(keys_config) - 1) > 1:
                            return False

            else:
                for i in range(len(self.table_2)):
                    keys_table = Analis_Table_keys(self.table_2[i], self.sheets_2.index(list),
                                                   start)
                    if head:
                        err = 0
                        for j in range(len(keys_config)):
                            if keys_config[j] in keys_table and j == keys_table.index(keys_config[j]):
                                pass
                            else:
                                err += 1
                        if err > 2:
                            return False
                    else:
                        if abs(len(keys_table) - len(keys_config)) > 2:
                            return False
        except:
            return False
        return True

    #Запускает всплывающее окно
    def go_comboBox_config(self, id_modul):
        names_config = []
        for conf in self.config:
            try:
                names_config.append(conf['Имя конфигурации'])
            except:
                continue
        self.array_export = [id_modul, names_config]
        self.switch_window_list.emit()

    #Сбор информации и создание конфигурации
    def create_config(self, id_modul, name_config):
        conf = {}
        conf.update({'Имя конфигурации': name_config})
        conf.update({'Режим работы': self.ui.comboBox_function.currentText()})
        conf.update({'Открыть файл по завершению': self.ui.checkBox_open.isChecked()})
        conf.update({'Аналитическая раскраска столбцов': self.ui.checkBox_color.isChecked()})
        conf.update({'Убрать не участвующие в анализе столбцы': self.ui.checkBox_del_stb.isChecked()})
        conf.update({'Склеить выбранные ключи': self.ui.checkBox_join_keys.isChecked()})
        if id_modul == 0:
            conf.update({'Состояние фильтра': self.ui.checkBox_filter.isChecked()})
            conf.update({'Выбранный лист': self.ui.comboBox_list_1.currentText()})

        else:
            conf.update({'Состояние фильтра': self.ui.checkBox_filter_2.isChecked()})
            conf.update({'Выбранный лист': self.ui.comboBox_list_2.currentText()})
        conf.update({'Начало таблицы': self.QT_moduls[id_modul]["spinBox_list"].value()})
        conf.update({'Наличие шапки': self.QT_moduls[id_modul]["comboBox_head"].currentText()})
        if id_modul == 0:
            conf.update({'Ключи': self.list_of_keys})
            conf.update({'Выбранные ключи': self.list_of_keys_selected})
            conf.update({'Выбранные ключи индексы': self.list_of_keys_flags})
            conf.update({'Выбранные ключи для подсчета': self.list_of_keys_selected_4})
        else:
            conf.update({'Ключи': self.list_of_keys_2})
            conf.update({'Выбранные ключи': self.list_of_keys_selected_2})
            conf.update({'Выбранные ключи индексы': self.list_of_keys_flags_2})
            conf.update({'Выбранные ключи для подсчета': self.list_of_keys_selected_5})
        conf.update({'Выбранные столбцы для подгрузки': self.list_of_keys_selected_3})
        conf.update({'Фильтер': self.filter[id_modul]})
        return conf

    def write_format_conf(self, conf):
        write_format_conf = "{\n"
        for key in conf:
            write_value = conf[key]
            if (type(write_value) == str):
                write_value = "'" + str(write_value) + "'"
            else:
                write_value = str(write_value)
            write_format_conf += "  '" + key + "': " + write_value + ",\n"
        write_format_conf = write_format_conf[:-2] + "\n}\n"
        write_format_conf = write_format_conf.replace("set()", "{}")
        write_format_conf += "_______________________________________________________________________________________________________\n\n"
        return write_format_conf

    #Сохранение новой конфигурации
    def save_config(self, value):
        id_modul, name_config = value[0], value[1]
        conf = self.create_config(id_modul, name_config)
        write_format_conf = self.write_format_conf(conf)

        self.config.append(conf)
        if id_modul == 0:
            self.config_files_1.append(conf)
        else:
            self.config_files_2.append(conf)

        config_file = open('config.txt', 'a', encoding='utf-8')
        config_file.write(write_format_conf)
        config_file.close()
        self.QT_moduls[id_modul]["comboBox_config"].addItem(conf["Имя конфигурации"])
        self.QT_moduls[id_modul]["comboBox_config"].setCurrentText(conf["Имя конфигурации"])
    #Удаление выбранной конфигурации
    def del_config(self, id_modul):
        ind_config = self.QT_moduls[id_modul]["comboBox_config"].currentIndex()
        name_config = self.QT_moduls[id_modul]["comboBox_config"].currentText()
        if self.QT_moduls[id_modul]["comboBox_config"].currentText() != "Нет":
            for i in range(len(self.config)):
                if self.config[i]['Имя конфигурации'] == name_config:
                    del self.config[i]
                    break
            file = open("config.txt", "w", encoding='utf-8')
            for i in range(len(self.config)):
                file.write(self.write_format_conf(self.config[i]))
            if id_modul == 0:
                self.search_config(0)
            else:
                self.search_config(1)

    #Загрузка всех конфигураций
    def download_config(self):
        try:
            file = open("config.txt", "r", encoding='utf-8')
            text = file.read()
            config = text.split(
                "_______________________________________________________________________________________________________")
            for i in range(len(config)):
                try:
                    config[i] = ast.literal_eval(config[i])
                    for j in config[i]['Фильтер']:
                        config[i]['Фильтер'][j] = set(config[i]['Фильтер'][j])

                except:
                    continue
            self.config = []
            for i in range(len(config)):
                if(type(config[i]) == dict):
                    self.config.append(config[i])
        except:
            self.logger("Готовых конфигураций не обнаружено")

    #Поиск подходящих кофигураций и отображение
    def search_config(self, id_modul):
        if id_modul == 0:
            self.config_files_1 = []
        else:
            self.config_files_2 = []
        for conf in self.config:
            try:
                list = conf["Выбранный лист"]
                start = int(conf["Начало таблицы"])
                keys = conf["Ключи"]
                flag = self.proverka_tables(id_modul, list, start, keys)
                if flag:
                    if id_modul == 0:
                        self.config_files_1.append(conf)
                    else:
                        self.config_files_2.append(conf)
            except BaseException:
                continue


        self.QT_moduls[id_modul]["comboBox_config"].currentIndexChanged.disconnect()
        self.QT_moduls[id_modul]["comboBox_config"].clear()
        self.QT_moduls[id_modul]["comboBox_config"].addItem("Нет")
        if id_modul == 0:
            for conf in self.config_files_1:
                self.QT_moduls[id_modul]["comboBox_config"].addItem(conf["Имя конфигурации"])
        else:
            for conf in self.config_files_2:
                self.QT_moduls[id_modul]["comboBox_config"].addItem(conf["Имя конфигурации"])
        self.QT_moduls[id_modul]["comboBox_config"].currentIndexChanged.connect(lambda: self.config_selected(id_modul))

    #Нажатие выбора конфигурации
    def config_selected(self, id_modul):
        ind_config = self.QT_moduls[id_modul]["comboBox_config"].currentIndex()
        if self.QT_moduls[id_modul]["comboBox_config"].currentText() != "Нет":
            ind_config -= 1
            if id_modul == 0:
                conf = self.config_files_1[ind_config]
            else:
                conf = self.config_files_2[ind_config]
            try:
                if id_modul == 0:
                    if 'Открыть файл по завершению' in conf:
                        self.ui.checkBox_open.setChecked(conf['Открыть файл по завершению'])
                    if 'Аналитическая раскраска столбцов' in conf:
                        self.ui.checkBox_color.setChecked(conf['Аналитическая раскраска столбцов'])
                    if 'Убрать не участвующие в анализе столбцы' in conf:
                        self.ui.checkBox_del_stb.setChecked(conf['Убрать не участвующие в анализе столбцы'])
                    if 'Склеить выбранные ключи' in conf:
                        self.ui.checkBox_join_keys.setChecked(conf['Склеить выбранные ключи'])

                    if 'Действия над строками с пометкой Удалить' in conf:
                        self.ui.comboBox_del.setCurrentText(conf['Действия над строками с пометкой Удалить'])

                if 'Режим работы' in conf:
                    self.ui.comboBox_function.setCurrentText(conf['Режим работы'])

                if 'Состояние фильтра' in conf:
                    self.QT_filter_moduls[id_modul]["checkBox_filter"].setChecked(conf['Состояние фильтра'])

                if 'Выбранный лист' in conf:
                    self.QT_moduls[id_modul]["comboBox_list"].setCurrentText(conf['Выбранный лист'])
                    if 'Начало таблицы' in conf and 'Наличие шапки' in conf:
                        self.QT_moduls[id_modul]["spinBox_list"].setValue(conf['Начало таблицы'])
                        self.QT_moduls[id_modul]["comboBox_head"].setCurrentText(conf['Наличие шапки'])
                        if 'Выбранные ключи' in conf and 'Выбранные ключи индексы' in conf:
                            # self.QT_moduls[id_modul]["list_selected_keys"].currentIndexChanged.disconnect()
                            if id_modul == 0:
                                self.list_of_keys_selected = conf['Выбранные ключи'].copy()
                                self.list_of_keys_flags = conf['Выбранные ключи индексы'].copy()
                            else:
                                self.list_of_keys_selected_2 = conf['Выбранные ключи'].copy()
                                self.list_of_keys_flags_2 = conf['Выбранные ключи индексы'].copy()
                            self.restart_list(self.QT_moduls[id_modul]["list_selected_keys"], conf['Выбранные ключи'])

                        if 'Выбранные ключи для подсчета' in conf:
                            if id_modul == 0:
                                self.list_of_keys_selected_4 = conf['Выбранные ключи для подсчета'].copy()
                            else:
                                self.list_of_keys_selected_5 = conf['Выбранные ключи для подсчета'].copy()
                            self.restart_list(self.QT_moduls[id_modul + 3]["list_selected_keys"], conf['Выбранные ключи для подсчета'])

                        if id_modul == 1:
                            if 'Выбранные столбцы для подгрузки' in conf:
                                self.list_of_keys_selected_3 = conf['Выбранные столбцы для подгрузки'].copy()
                                self.restart_list(self.QT_moduls[2]["list_selected_keys"],
                                                  conf['Выбранные столбцы для подгрузки'])

                        if 'Фильтер' in conf:
                            self.filter[id_modul] = conf['Фильтер'].copy()
                            self.Apdate_comboBox(id_modul)

            except BaseException:
                pass

    # Фильтер "Очистить"
    def filter_delete_keys(self, id_modul):
        stb = self.QT_filter_moduls[id_modul]["comboBox_keys"].currentText()
        if stb not in "Нет":
            self.filter[id_modul][stb] = set()
            self.QT_filter_moduls[id_modul]["list_selected_unique_keys"].clear()

            list_unique_keys_id_select = self.QT_filter_moduls[id_modul]["comboBox_keys"].currentIndex()
            self.QT_filter_moduls[id_modul]["comboBox_keys"].disconnect()
            self.Apdate_comboBox(id_modul)
            self.QT_filter_moduls[id_modul]["comboBox_keys"].setCurrentIndex(list_unique_keys_id_select)
            self.QT_filter_moduls[id_modul]["comboBox_keys"].currentIndexChanged.connect(
                lambda: self.choice_case_filter(id_modul))

    # Фильтер "Подгрузка"
    # Подгружает значения фильтра по формуле
    def filter_download_keys(self, id_modul):
        formula = self.QT_filter_moduls[id_modul]["lineEdit_filter"].text()
        if (formula == ""):
            return
        elif (formula == "rfrrfrrfr1"):
            print("Автор программы: Валерий Кобзев(НСИ)")
            self.error_msg("Автор программы: Валерий Кобзев(НСИ)", title="Информация", icon=QMessageBox.Information)
        stb = self.QT_filter_moduls[id_modul]["comboBox_keys"].currentText()
        if stb not in "Нет":
            flag_head = self.QT_moduls[id_modul]["comboBox_head"].currentText() == "Да"
            formula = formula.replace("*", ".{0,}").replace("?", ".")
            if id_modul == 0:
                elems = set()
                for i in range(len(self.table_1)):
                    try:
                        column = self.table_1[i].sheet_by_index(self.ui.comboBox_list_1.currentIndex() - 1).col_values(
                            self.list_of_keys.index(self.QT_filter_moduls[id_modul]["comboBox_keys"].currentText()))[
                                 int(flag_head):]
                        type = self.table_1[i].sheet_by_index(self.ui.comboBox_list_1.currentIndex() - 1) \
                            .cell(rowx=self.QT_moduls[id_modul]["spinBox_list"].value() + 2,
                                  colx=self.list_of_keys.index(
                                      self.QT_filter_moduls[id_modul]["comboBox_keys"].currentText())).ctype
                        if type == 3:
                            for str_id in range(len(column)):
                                try:
                                    y, m, d, h, i, s = xlrd.xldate_as_tuple(int(column[str_id]),
                                                                            self.table_1[i].datemode)
                                    column[str_id] = str("{0}.{1}.{2}".format(d, m, y))
                                except BaseException:
                                    pass
                        elems |= set(column)
                    except BaseException:
                        pass
            else:
                elems = set()
                for i in range(len(self.table_2)):
                    try:
                        column = self.table_2[i].sheet_by_index(self.ui.comboBox_list_2.currentIndex() - 1).col_values(
                            self.list_of_keys_2.index(self.QT_filter_moduls[id_modul]["comboBox_keys"].currentText()))[
                                 int(flag_head):]
                        type = self.table_2[i].sheet_by_index(self.ui.comboBox_list_2.currentIndex() - 1) \
                            .cell(rowx=self.QT_moduls[id_modul]["spinBox_list"].value() + 2,
                                  colx=self.list_of_keys_2.index(
                                      self.QT_filter_moduls[id_modul]["comboBox_keys"].currentText())).ctype
                        if type == 3:
                            for str_id in range(len(column)):
                                try:
                                    y, m, d, h, i, s = xlrd.xldate_as_tuple(int(column[str_id]),
                                                                            self.table_2[i].datemode)
                                    column[str_id] = str("{0}.{1}.{2}".format(d, m, y))
                                except BaseException:
                                    pass
                        elems |= set(column)
                    except BaseException:
                        pass
            flag = False
            for el in elems:
                el = float_str(el)
                if (bool(re.fullmatch(formula, el)) and el not in self.filter[id_modul][stb]):
                    flag = True
                    self.filter[id_modul][stb].add(el)
                    # Ставит галочку
            if flag:
                list_unique_keys_id_select = self.QT_filter_moduls[id_modul]["comboBox_keys"].currentIndex()
                self.QT_filter_moduls[id_modul]["comboBox_keys"].disconnect()
                self.Apdate_comboBox(id_modul)
                self.QT_filter_moduls[id_modul]["comboBox_keys"].setCurrentIndex(list_unique_keys_id_select)
                self.QT_filter_moduls[id_modul]["comboBox_keys"].currentIndexChanged.connect(
                    lambda: self.choice_case_filter(id_modul))
            self.QT_filter_moduls[id_modul]["list_selected_unique_keys"].clear()
            for el in self.filter[id_modul][stb]:
                self.QT_filter_moduls[id_modul]["list_selected_unique_keys"].addItem(str(el))

    # Замена интерфейса
    def choice_comboBox_function(self):
        size_square = 31
        height = 121
        width = 211

        if self.ui.comboBox_function.currentIndex() in [1, 4]:
            self.ui.pushButton_delete_4.show()
            self.ui.pushButton_add_4.show()
            self.ui.pushButton_up_4.show()
            self.ui.pushButton_down_4.show()
            self.ui.list_selected_keys_4.show()
            self.ui.keys_label_calculations.show()

            if self.ui.comboBox_function.currentIndex() == 4:
                self.ui.pushButton_delete_5.show()
                self.ui.pushButton_add_5.show()
                self.ui.pushButton_up_5.show()
                self.ui.pushButton_down_5.show()
                self.ui.list_selected_keys_5.show()
                self.ui.keys_label_calculations_2.show()


            # LEFT
            self.ui.list_selected_keys.setGeometry(QtCore.QRect(540, 160, width, height))
            self.ui.pushButton_up.setGeometry(QtCore.QRect(600, 290, size_square, size_square))
            self.ui.pushButton_down.setGeometry(QtCore.QRect(650, 290, size_square, size_square))
            self.ui.pushButton_delete.setGeometry(QtCore.QRect(500, 180, size_square, size_square))
            self.ui.pushButton_add.setGeometry(QtCore.QRect(500, 230, size_square, size_square))

            if self.ui.comboBox_function.currentIndex() == 4:
                # RIGHT
                self.ui.list_selected_keys_2.setGeometry(QtCore.QRect(20, 160, width, height))
                self.ui.pushButton_up_2.setGeometry(QtCore.QRect(80, 290, size_square, size_square))
                self.ui.pushButton_down_2.setGeometry(QtCore.QRect(130, 290, size_square, size_square))
                self.ui.pushButton_delete_2.setGeometry(QtCore.QRect(240, 240, size_square, size_square))
                self.ui.pushButton_add_2.setGeometry(QtCore.QRect(240, 190, size_square, size_square))

        else:
            self.ui.pushButton_delete_4.hide()
            self.ui.pushButton_add_4.hide()
            self.ui.pushButton_up_4.hide()
            self.ui.pushButton_down_4.hide()
            self.ui.list_selected_keys_4.hide()
            self.ui.keys_label_calculations.hide()

            self.ui.pushButton_delete_5.hide()
            self.ui.pushButton_add_5.hide()
            self.ui.pushButton_up_5.hide()
            self.ui.pushButton_down_5.hide()
            self.ui.list_selected_keys_5.hide()
            self.ui.keys_label_calculations_2.hide()

            size_square = 31
            # LEFT
            self.ui.list_selected_keys.setGeometry(QtCore.QRect(540, 160, width, 311))
            self.ui.pushButton_up.setGeometry(QtCore.QRect(600, 480, size_square, size_square))
            self.ui.pushButton_down.setGeometry(QtCore.QRect(650, 480, size_square, size_square))
            self.ui.pushButton_delete.setGeometry(QtCore.QRect(500, 340, size_square, size_square))
            self.ui.pushButton_add.setGeometry(QtCore.QRect(500, 290, size_square, size_square))
            # RIGHT
            self.ui.list_selected_keys_2.setGeometry(QtCore.QRect(20, 160, width, 311))
            self.ui.pushButton_up_2.setGeometry(QtCore.QRect(80, 480, size_square, size_square))
            self.ui.pushButton_down_2.setGeometry(QtCore.QRect(130, 480, size_square, size_square))
            self.ui.pushButton_delete_2.setGeometry(QtCore.QRect(240, 340, size_square, size_square))
            self.ui.pushButton_add_2.setGeometry(QtCore.QRect(240, 290, size_square, size_square))

        pass


    def spinBox_list(self, id_modul):
        if (id_modul == 0):
            if len(self.list_of_keys) == 0:
                return
        else:
            if len(self.list_of_keys_2) == 0:
                return
        self.choice_case_list(id_modul)

    # 2.0 Скачивание файла при выборе листа
    def choice_case_list(self, id_modul):
        self.step = self.step + 1
        selected_list = self.QT_moduls[id_modul]["comboBox_list"].currentText()
        if (selected_list in ["Лист не выбран", "Нет"]):
            self.QT_moduls[id_modul]["list_keys"].clear()
            self.QT_moduls[id_modul]["list_selected_keys"].clear()
            # if id_modul == 0:
            self.QT_filter_moduls[id_modul]["comboBox_keys"].currentIndexChanged.disconnect()
            self.QT_filter_moduls[id_modul]["comboBox_keys"].clear()
            self.QT_filter_moduls[id_modul]["comboBox_keys"].addItem("Нет")
            self.QT_filter_moduls[id_modul]["comboBox_keys"].currentIndexChanged.connect(
                lambda: self.choice_case_filter(id_modul))

            return

        if id_modul == 0:
            len_table = self.table_1[0].sheet_by_index(self.sheets_1.index(selected_list)).nrows
        else:
            len_table = self.table_2[0].sheet_by_index(self.sheets_2.index(selected_list)).nrows

        if len_table - 1 < self.QT_moduls[id_modul]["spinBox_list"].value():
            self.QT_moduls[id_modul]["spinBox_list"].setValue(len_table - 1)

        flag_head = self.QT_moduls[id_modul]["comboBox_head"].currentText() == "Да"
        self.QT_filter_moduls[id_modul]["list_unique_keys"].clear()
        self.QT_filter_moduls[id_modul]["list_selected_unique_keys"].clear()
        if id_modul == 0:
            # self.progressBarMsg(0, 'Подгружаем ключи из выбранного листа')
            self.list_of_keys = Analis_Table_keys(self.table_1[0], self.sheets_1.index(selected_list),
                                                  self.QT_moduls[id_modul]["spinBox_list"].value())

            if (not flag_head):
                self.list_of_keys = [str(i + 1) for i in range(0, len(self.list_of_keys))]
            self.list_of_keys_flags = []
            self.list_of_keys_selected = []
            self.list_of_keys_selected_4 = []
            self.filter[0] = {}
            for key in self.list_of_keys:
                self.filter[0].update({key: set()})

        else:

            self.list_of_keys_2 = Analis_Table_keys(self.table_2[0], self.sheets_2.index(selected_list),
                                                    self.QT_moduls[id_modul]["spinBox_list"].value())
            if (not flag_head):
                self.list_of_keys_2 = [str(i + 1) for i in range(0, len(self.list_of_keys_2))]
            self.list_of_keys_flags_2 = []
            self.list_of_keys_selected_2 = []
            self.list_of_keys_selected_3 = []
            self.list_of_keys_selected_5 = []
            self.filter[1] = {}
            for key in self.list_of_keys_2:
                self.filter[1].update({key: set()})

        # заполняем строки данными из Датафрейма
        if id_modul == 0:
            self.ui.list_keys.clear()
            self.ui.list_selected_keys.clear()
            self.ui.list_selected_keys_4.clear()

            for i in range(len(self.list_of_keys)):
                self.ui.list_keys.insertItem(i, self.list_of_keys[i])
                self.list_of_keys_flags.append(False)
                # Заполняем фильтер

            self.QT_filter_moduls[id_modul]["comboBox_keys"].currentIndexChanged.disconnect()
            self.QT_filter_moduls[id_modul]["comboBox_keys"].clear()
            self.QT_filter_moduls[id_modul]["comboBox_keys"].addItem("Нет")

            for i in range(len(self.list_of_keys)):
                self.QT_filter_moduls[id_modul]["comboBox_keys"].addItem(self.list_of_keys[i])
            self.QT_filter_moduls[id_modul]["comboBox_keys"].currentIndexChanged.connect(
                lambda: self.choice_case_filter(id_modul))
        else:
            self.ui.list_keys_2.clear()
            self.ui.list_selected_keys_2.clear()
            self.ui.list_selected_keys_3.clear()
            self.ui.list_selected_keys_5.clear()
            for i in range(len(self.list_of_keys_2)):
                self.ui.list_keys_2.insertItem(i, self.list_of_keys_2[i])
                self.list_of_keys_flags_2.append(False)

            self.QT_filter_moduls[id_modul]["comboBox_keys"].currentIndexChanged.disconnect()
            self.QT_filter_moduls[id_modul]["comboBox_keys"].clear()
            self.QT_filter_moduls[id_modul]["comboBox_keys"].addItem("Нет")

            for i in range(len(self.list_of_keys_2)):
                self.QT_filter_moduls[id_modul]["comboBox_keys"].addItem(self.list_of_keys_2[i])
            self.QT_filter_moduls[id_modul]["comboBox_keys"].currentIndexChanged.connect(
                lambda: self.choice_case_filter(id_modul))
            # self.progressBarMsg(0, '"' + self.name_file_2 + '" загружен')
        # print("Ща будем удалять")
        # print("Удалили : ", gc.collect())

        #Проверка подлиности конфигурации
        if id_modul == 0:
            conf1 = self.create_config(0, "Параметры первого файла")
            # conf2 = self.create_config(1, "Параметры второго файла")
            list = conf1["Выбранный лист"]
            start = int(conf1["Начало таблицы"])
            keys = conf1["Ключи"]
            flag = self.proverka_tables(0, list, start, keys)
            if (not flag):
                self.ui.msg_label.setText("Не все файлы левой колонки\n соответвуют друг другу!")
            else:
                self.ui.msg_label.setText("")
        else:
            conf2 = self.create_config(1, "Параметры второго файла")
            list = conf2["Выбранный лист"]
            start = int(conf2["Начало таблицы"])
            keys = conf2["Ключи"]
            flag = self.proverka_tables(1, list, start, keys)
            if (not flag):
                self.ui.msg_label_2.setText("Не все файлы правой колонки\n соответвуют друг другу!")
            else:
                self.ui.msg_label_2.setText("")

        # self.progressBarMsg(0, 'Готов в анализу')

    #Обновление листа
    def restart_list(self, list, value):
        # print(self.step, "restart_list")
        self.step = self.step + 1

        list.clear()
        for el in value:
            list.addItem(el)

    # 2.0 Выбор кейса фильтра
    def choice_case_filter(self, id_modul):
        self.step = self.step + 1
        self.QT_filter_moduls[id_modul]["list_unique_keys"].clear()
        self.QT_filter_moduls[id_modul]["list_selected_unique_keys"].clear()
        if self.QT_filter_moduls[id_modul]["comboBox_keys"].currentText() not in "Нет":
            flag_head = self.QT_moduls[id_modul]["comboBox_head"].currentText() == "Да"
            if id_modul == 0:
                elems = set()
                for i in range(len(self.table_1)):
                    try:
                        column = self.table_1[i].sheet_by_index(self.ui.comboBox_list_1.currentIndex() - 1).col_values(
                                self.list_of_keys.index(self.QT_filter_moduls[id_modul]["comboBox_keys"].currentText()))[int(flag_head):]
                        type = self.table_1[i].sheet_by_index(self.ui.comboBox_list_1.currentIndex() - 1)\
                            .cell(rowx=self.QT_moduls[id_modul]["spinBox_list"].value() + 2,
                                  colx=self.list_of_keys.index(self.QT_filter_moduls[id_modul]["comboBox_keys"].currentText())).ctype
                        if type == 3:
                            for str_id in range(len(column)):
                                try:
                                    y, m, d, h, i, s = xlrd.xldate_as_tuple(int(column[str_id]), self.table_1[i].datemode)
                                    column[str_id] = str("{0}.{1}.{2}".format(d, m, y))
                                except BaseException:
                                    pass
                        elems |= set(column)
                    except BaseException:
                        pass
     
            else:
                elems = set()
                for i in range(len(self.table_2)):
                    try:
                        column = self.table_2[i].sheet_by_index(self.ui.comboBox_list_2.currentIndex() - 1).col_values(
                            self.list_of_keys_2.index(self.QT_filter_moduls[id_modul]["comboBox_keys"].currentText()))[
                                 int(flag_head):]
                        type = self.table_2[i].sheet_by_index(self.ui.comboBox_list_2.currentIndex() - 1) \
                            .cell(rowx=self.QT_moduls[id_modul]["spinBox_list"].value() + 2,
                                  colx=self.list_of_keys_2.index(
                                      self.QT_filter_moduls[id_modul]["comboBox_keys"].currentText())).ctype
                        if type == 3:
                            for str_id in range(len(column)):
                                try:
                                    y, m, d, h, i, s = xlrd.xldate_as_tuple(int(column[str_id]),
                                                                            self.table_2[i].datemode)
                                    column[str_id] = str("{0}.{1}.{2}".format(d, m, y))
                                except BaseException:
                                    pass
                        elems |= set(column)
                    except BaseException:
                        pass
             
            for el in elems:
                try:
                    el2 = float(el)
                    if el2 % 1 == 0:
                        el = str(el).split(".")[0]
                    else:
                        el = str(int(el)).replace(".", ",")
                except BaseException:
                    pass
                self.QT_filter_moduls[id_modul]["list_unique_keys"].addItem(str(el))
            for el in self.filter[id_modul][self.QT_filter_moduls[id_modul]["comboBox_keys"].currentText()]:
                self.QT_filter_moduls[id_modul]["list_selected_unique_keys"].addItem(str(el))

        return

    def pushButton_down(self, id_modul):
        self.step = self.step + 1
        id = self.QT_moduls[id_modul]["list_selected_keys"].currentRow()
        n = self.QT_moduls[id_modul]["list_selected_keys"].count()

        if id < n - 1:
            # print(id)
            if id_modul == 0:
                self.list_of_keys_selected[id], self.list_of_keys_selected[id + 1] = self.list_of_keys_selected[id + 1], \
                                                                                     self.list_of_keys_selected[id]
                self.restart_list(self.QT_moduls[id_modul]["list_selected_keys"], self.list_of_keys_selected)
            elif (id_modul == 1):
                self.list_of_keys_selected_2[id], self.list_of_keys_selected_2[id + 1] = self.list_of_keys_selected_2[
                                                                                             id + 1], \
                                                                                         self.list_of_keys_selected_2[
                                                                                             id]
                self.restart_list(self.QT_moduls[id_modul]["list_selected_keys"], self.list_of_keys_selected_2)
            elif (id_modul == 2):
                self.list_of_keys_selected_3[id], self.list_of_keys_selected_3[id + 1] = self.list_of_keys_selected_3[
                                                                                             id + 1], \
                                                                                         self.list_of_keys_selected_3[
                                                                                             id]
                self.restart_list(self.QT_moduls[id_modul]["list_selected_keys"], self.list_of_keys_selected_3)
            elif (id_modul == 3):
                self.list_of_keys_selected_4[id], self.list_of_keys_selected_4[id + 1] = self.list_of_keys_selected_4[
                                                                                             id + 1], \
                                                                                         self.list_of_keys_selected_4[
                                                                                             id]
                self.restart_list(self.QT_moduls[id_modul]["list_selected_keys"], self.list_of_keys_selected_4)
            else:
                self.list_of_keys_selected_5[id], self.list_of_keys_selected_5[id + 1] = self.list_of_keys_selected_5[
                                                                                             id + 1], \
                                                                                         self.list_of_keys_selected_5[
                                                                                             id]
                self.restart_list(self.QT_moduls[id_modul]["list_selected_keys"], self.list_of_keys_selected_5)

            self.QT_moduls[id_modul]["list_selected_keys"].setCurrentRow(id + 1)

    def pushButton_up(self, id_modul):
        self.step = self.step + 1
        id = self.QT_moduls[id_modul]["list_selected_keys"].currentRow()

        if id > 0:
            if id_modul == 0:
                self.list_of_keys_selected[id], self.list_of_keys_selected[id - 1] = self.list_of_keys_selected[id - 1], \
                                                                                     self.list_of_keys_selected[id]
                self.restart_list(self.ui.list_selected_keys, self.list_of_keys_selected)
            elif id_modul == 1:
                self.list_of_keys_selected_2[id], self.list_of_keys_selected_2[id - 1] = self.list_of_keys_selected_2[
                                                                                             id - 1], \
                                                                                         self.list_of_keys_selected_2[
                                                                                             id]
                self.restart_list(self.ui.list_selected_keys_2, self.list_of_keys_selected_2)
            elif id_modul == 2:
                self.list_of_keys_selected_3[id], self.list_of_keys_selected_3[id - 1] = self.list_of_keys_selected_3[
                                                                                             id - 1], \
                                                                                         self.list_of_keys_selected_3[
                                                                                             id]
                self.restart_list(self.ui.list_selected_keys_3, self.list_of_keys_selected_3)
            elif id_modul == 3:
                self.list_of_keys_selected_4[id], self.list_of_keys_selected_4[id - 1] = self.list_of_keys_selected_4[
                                                                                             id - 1], \
                                                                                         self.list_of_keys_selected_4[
                                                                                             id]
                self.restart_list(self.ui.list_selected_keys_4, self.list_of_keys_selected_4)
            else:
                self.list_of_keys_selected_5[id], self.list_of_keys_selected_5[id - 1] = self.list_of_keys_selected_5[
                                                                                             id - 1], \
                                                                                         self.list_of_keys_selected_5[
                                                                                             id]
                self.restart_list(self.ui.list_selected_keys_5, self.list_of_keys_selected_5)

            self.QT_moduls[id_modul]["list_selected_keys"].setCurrentRow(id - 1)

    # перезалив листа
    def Auto_select_str(self, id_modul):
        selected_list = self.QT_moduls[id_modul]["comboBox_list"].currentText()

        if (selected_list in ["Лист не выбран", "Нет"]):
            return
        if (id_modul == 0):
            x = Auto_selection_start_str(self.table_1[0], self.sheets_1.index(selected_list),
                                         self.QT_moduls[id_modul]["spinBox_list"].value())
        else:
            x = Auto_selection_start_str(self.table_2[0], self.sheets_2.index(selected_list),
                                         self.QT_moduls[id_modul]["spinBox_list"].value())
 
        self.QT_moduls[id_modul]["spinBox_list"].setValue(x)

    def Add_key(self, id_modul):
   
        self.step = self.step + 1

        # получить индекс (0, n) текущей (выделенной) строки
        value = self.QT_moduls[id_modul]["list_keys"].currentRow()
        # получить количество строк в листе
        # сравниваем по флагу, выбран ли уже элемент или он доступен к выбору

        if value < 0:
            return
        if id_modul == 0:
            if not self.list_of_keys_flags[value]:
                self.list_of_keys_flags[value] = True
                self.list_of_keys_selected.append(self.list_of_keys[value])
                self.restart_list(self.ui.list_selected_keys, self.list_of_keys_selected)
        elif id_modul == 1:
            if not self.list_of_keys_flags_2[value]:
                self.list_of_keys_flags_2[value] = True
                self.list_of_keys_selected_2.append(self.list_of_keys_2[value])
                self.restart_list(self.ui.list_selected_keys_2, self.list_of_keys_selected_2)
        elif id_modul == 2:
            if self.ui.list_keys_2.currentItem().text() not in self.list_of_keys_selected_3:
                self.list_of_keys_selected_3.append(self.ui.list_keys_2.currentItem().text())
                self.restart_list(self.ui.list_selected_keys_3, self.list_of_keys_selected_3)
        elif id_modul == 3:
            if self.ui.list_keys.currentItem().text() not in self.list_of_keys_selected_4:
                self.list_of_keys_selected_4.append(self.ui.list_keys.currentItem().text())
                self.restart_list(self.ui.list_selected_keys_4, self.list_of_keys_selected_4)
        else:
            if self.ui.list_keys_2.currentItem().text() not in self.list_of_keys_selected_5:
                self.list_of_keys_selected_5.append(self.ui.list_keys_2.currentItem().text())
                self.restart_list(self.ui.list_selected_keys_5, self.list_of_keys_selected_5)

    def Delete_key(self, id_modul):
        self.step = self.step + 1

        value = self.QT_moduls[id_modul]["list_selected_keys"].currentRow()
        # value2 = self.QT_moduls[id_modul]["list_selected_keys"].count()
        # если в поле ХВыбранные ключиХ какое-то поле выделено то
        if value < 0:
            return

        if self.QT_moduls[id_modul]["list_selected_keys"].currentItem() is not None:
            # доступ к текстовому значению строки в листе
            text = self.QT_moduls[id_modul]["list_selected_keys"].currentItem().text()
            # выставляем флаги в Фолсе, чтобы повторно можно было выбрать поле
            if id_modul == 0:
                for i in range(len(self.list_of_keys)):
                    if text == self.list_of_keys[i]:
                        self.list_of_keys_flags[i] = False
                        break
                self.list_of_keys_selected.remove(text)
                self.restart_list(self.ui.list_selected_keys, self.list_of_keys_selected)
            elif id_modul == 1:
                for i in range(len(self.list_of_keys_2)):
                    if text == self.list_of_keys_2[i]:
                        self.list_of_keys_flags_2[i] = False
                        break
                self.list_of_keys_selected_2.remove(text)
                self.restart_list(self.ui.list_selected_keys_2, self.list_of_keys_selected_2)
            elif id_modul == 2:
                self.list_of_keys_selected_3.remove(text)
                self.restart_list(self.ui.list_selected_keys_3, self.list_of_keys_selected_3)
            elif id_modul == 3:
                self.list_of_keys_selected_4.remove(text)
                self.restart_list(self.ui.list_selected_keys_4, self.list_of_keys_selected_4)
            else:
                self.list_of_keys_selected_5.remove(text)
                self.restart_list(self.ui.list_selected_keys_5, self.list_of_keys_selected_5)

    # Обновление комбобокса фильтра
    def Apdate_comboBox(self, id_modul):
        self.step = self.step + 1
        self.QT_filter_moduls[id_modul]["comboBox_keys"].clear()
        self.QT_filter_moduls[id_modul]["comboBox_keys"].addItem("Нет")
        # Заполняем фильтер
        if id_modul == 0:
            for i in range(len(self.list_of_keys)):
                if len(self.filter[id_modul][self.list_of_keys[i]]) > 0:
                    self.QT_filter_moduls[id_modul]["comboBox_keys"].addItem(QtGui.QIcon(":/newPrefix/galka.jpg"),
                                                                             self.list_of_keys[i])
                else:
                    self.QT_filter_moduls[id_modul]["comboBox_keys"].addItem(self.list_of_keys[i])
        else:
            for i in range(len(self.list_of_keys_2)):
                if len(self.filter[id_modul][self.list_of_keys_2[i]]) > 0:
                    self.QT_filter_moduls[id_modul]["comboBox_keys"].addItem(QtGui.QIcon(":/newPrefix/galka.jpg"),
                                                                             self.list_of_keys_2[i])
                else:
                    self.QT_filter_moduls[id_modul]["comboBox_keys"].addItem(self.list_of_keys_2[i])

    def Add_key_filter(self, id_modul):
        self.step = self.step + 1

        # получить индекс (0, n) текущей (выделенной) строки
        current_id = self.QT_filter_moduls[id_modul]["list_unique_keys"].currentRow()
        if current_id < 0:
            return
        list_unique_keys_id_select = self.QT_filter_moduls[id_modul]["comboBox_keys"].currentIndex()

        value = self.QT_filter_moduls[id_modul]["list_unique_keys"].currentItem().text()
        value2 = self.QT_filter_moduls[id_modul]["list_unique_keys"].count()
        stb = self.QT_filter_moduls[id_modul]["comboBox_keys"].currentText()

        if value not in self.filter[id_modul][stb]:
            self.filter[id_modul][stb].add(value)
            self.QT_filter_moduls[id_modul]["list_selected_unique_keys"].insertItem(value2, value)

        # Ставит галочку
        if self.QT_filter_moduls[id_modul]["list_selected_unique_keys"].count() == 1:
            self.QT_filter_moduls[id_modul]["comboBox_keys"].disconnect()
            self.Apdate_comboBox(id_modul)
            self.QT_filter_moduls[id_modul]["comboBox_keys"].setCurrentIndex(list_unique_keys_id_select)
            self.QT_filter_moduls[id_modul]["comboBox_keys"].currentIndexChanged.connect(
                lambda: self.choice_case_filter(id_modul))


    def Delete_key_filter(self, id_modul):
        self.step = self.step + 1
        current_id = self.QT_filter_moduls[id_modul]["list_selected_unique_keys"].currentRow()
        if current_id < 0:
            return
        list_unique_keys_id_select = self.QT_filter_moduls[id_modul]["comboBox_keys"].currentIndex()
        text = self.QT_filter_moduls[id_modul]["list_selected_unique_keys"].currentItem().text()
        # если в поле ХВыбранные ключиХ какое-то поле выделено то

        if self.QT_filter_moduls[id_modul]["list_selected_unique_keys"].currentItem() is not None:
            # выставляем флаги в Фолсе, чтобы повторно можно было выбрать поле
            stb = self.QT_filter_moduls[id_modul]["comboBox_keys"].currentText()
            self.filter[id_modul][stb].remove(text)
            self.QT_filter_moduls[id_modul]["list_selected_unique_keys"].takeItem(current_id)
        if current_id == 0:
            self.QT_filter_moduls[id_modul]["comboBox_keys"].disconnect()
            self.Apdate_comboBox(id_modul)
            self.QT_filter_moduls[id_modul]["comboBox_keys"].setCurrentIndex(list_unique_keys_id_select)
            self.QT_filter_moduls[id_modul]["comboBox_keys"].currentIndexChanged.connect(
                lambda: self.choice_case_filter(id_modul))

    def normalization_table(self, table, value = "", ID=0):
        Len = len(table[0])
        for i in range(len(table)):
            dif = Len - len(table[i])
            if dif > 0:
                table[i] = table[i][:ID] + dif * [value] + table[i][ID:]
            elif dif < 0:
                table[i] = table[i][:ID + dif] + table[i][ID:Len - dif]

    def progress(self):  # Нажатие кнопки результаты
        self.logger(level=1)
        self.progressBarMsg(0, 'Начало обработки!')
        mode = self.ui.comboBox_function.currentIndex()

        conf1 = self.create_config(0, "Параметры первого файла")
        conf2 = self.create_config(1, "Параметры второго файла")

        list = conf1["Выбранный лист"]
        start = int(conf1["Начало таблицы"])
        keys = conf1["Ключи"]
        flag = self.proverka_tables(0, list, start, keys)
        if (not flag):
            if self.warning_select_msg("Не все файлы левой колонки соответвуют друг другу!\nИгнорировать?") not in "Yes":
                return
        if self.ui.comboBox_function.currentIndex() in [3, 4]:
            list = conf2["Выбранный лист"]
            start = int(conf2["Начало таблицы"])
            keys = conf2["Ключи"]
            flag = self.proverka_tables(1, list, start, keys)
            if (not flag):
                if self.warning_select_msg(
                        "Не все файлы правой колонки соответвуют друг другу!\nИгнорировать?") not in "Yes":
                    return

        # Проверка на заполненость данных
        self.step = self.step + 1
        if len(self.list_of_keys) == 0:
            self.error_msg("Не выбран обрабатывемый файл.")
            return

        if self.ui.comboBox_function.currentIndex() in [3, 4]:
            if len(self.list_of_keys_selected_2) == 0:
                self.error_msg("Пожалуйста, выберите хотя бы один ключевой столбец во второй таблице!!")
                return
            if len(self.list_of_keys_selected) != len(self.list_of_keys_selected_2) \
                    and not self.ui.checkBox_join_keys.isChecked():
                self.error_msg("Выбрано разное колличество сопоставляемых ключей!")
                return

        if (mode == 4):
            if (len(self.list_of_keys_selected_4) == 0):
                self.error_msg("Не выбран не один вычисляемый столбец.")
                return
            elif len(self.list_of_keys_selected_4) != len(self.list_of_keys_selected_5):
                self.error_msg("Выбрано разное колличество вычисляемых столбцов!")
                return

        flag = False

        i = 0
        j = 0
        keys = []
        keys_ind = []
        for el in self.list_of_keys_flags:
            if el:
                keys.append(self.list_of_keys[i])
                keys_ind.append(j)
            i = i + 1
            j = j + 1
            flag = flag or el

        if not flag:
            self.error_msg("Пожалуйста, выберите хотя бы один ключевой столбец.")
            return

        # Дефолтная обработка

        self.progressBarMsg(5, 'Преобразование таблицы в удобный для анализа вид')
        selected_list_1 = self.QT_moduls[0]["comboBox_list"].currentText()
        flag_head = self.QT_moduls[0]["comboBox_head"].currentText() == "Да"

        df = []
        for i in range(len(self.table_1)):
            df += Analis_Table_List(self.table_1[i], self.sheets_1.index(selected_list_1),
                               self.QT_moduls[0]["spinBox_list"].value() - int(not flag_head))  # add "id", Удалить"
        self.normalization_table(df, ID=len(df[0]) - 2)


        self.progressBarMsg(25, 'Сортировка таблицы по выбранным ключам')
        sort_list(df, keys_ind)

        if self.ui.checkBox_filter.isChecked():
            filter(df, self.filter[0], 0, self.list_of_keys)

        self.progressBarMsg(35, 'Группировка уникальных ключей')

        list_of_keys_selected_ind = []
        for key in self.list_of_keys_selected:
            list_of_keys_selected_ind.append(self.list_of_keys.index(key))

        group_list(df, list_of_keys_selected_ind,
                   self.list_of_keys_selected,
                   self.ui.checkBox_join_keys.isChecked())  # add "Подгруппa" , "Колличество",  add "Уникальный ключ"

        add_stb = []
        list_of_keys_selected_ind_4 = []
        list_of_keys_selected_ind_5 = []

        if mode == 0:
            sort_list(df, [len(df[0]) - 6])
            self.progressBarMsg(65, 'Анализ закончен, приступаем к выгрузке в файл')
        elif mode == 1:
            self.progressBarMsg(45, 'Фильтруем')
            if len(self.list_of_keys_selected_4) > 0:
                list_of_keys_selected_ind_4 = []
                for el in self.list_of_keys_selected_4:
                    list_of_keys_selected_ind_4.append(self.list_of_keys.index(el))

                list_stb_calculations(df, add_stb, list_of_keys_selected_ind_4, self.list_of_keys_selected_4)
                for el in df:
                    print(el)
            else:
                filter(df, keys_ind, 1)
            self.progressBarMsg(65, 'Анализ закончен, приступаем к выгрузке в файл')
        elif mode == 2:
            self.progressBarMsg(45, 'Фильтруем')
            filter(df, keys_ind, 2)
            self.progressBarMsg(65, 'Анализ закончен, приступаем к выгрузке в файл')
        elif mode in [3, 4]:

            self.progressBarMsg(45, 'Преобразование второй таблицы в удобный для анализа вид')
            selected_list_2 = self.QT_moduls[1]["comboBox_list"].currentText()
            flag_head_2 = self.QT_moduls[1]["comboBox_head"].currentText() == "Да"

            for i in range(len(self.table_2)):
                df_2 = Analis_Table_List(self.table_2[i], self.sheets_2.index(selected_list_2),
                                     self.QT_moduls[1]["spinBox_list"].value() - int(
                                         not flag_head_2))  # add "id", Удалить"
            self.normalization_table(df_2, ID=len(df[0]) - 2)

            # Используем фильтер 2
            if self.ui.checkBox_filter_2.isChecked():
                keys_2 = []
                keys_ind_2 = []
                i = 0
                for el in self.list_of_keys_flags_2:
                    if el:
                        keys_2.append(self.list_of_keys_2[i])
                        keys_ind_2.append(i)
                    i = i + 1
                sort_list(df_2, keys_ind_2)
                # print(df_2)
                # print(self.filter[1])
                # print("go")
                filter(df_2, self.filter[1], 0, self.list_of_keys_2)

            list_of_keys_selected_ind_2 = []
            for key in self.list_of_keys_selected_2:
                list_of_keys_selected_ind_2.append(self.list_of_keys_2.index(key))

            self.progressBarMsg(55, 'Сортировка второй таблицы')
            sort_list(df_2, list_of_keys_selected_ind_2)

            self.progressBarMsg(60, 'Группировка второй таблицы по выбраным ключам ')
            group_list(df_2, list_of_keys_selected_ind_2,
                       self.list_of_keys_selected_2,
                       self.ui.checkBox_join_keys.isChecked())  # add "Подгруппa" , "Колличество",  add "Уникальный ключ"

            self.progressBarMsg(67, 'Составление групп уникальных ключей(Полное соотвествие)')

            # 2!
            set_stb_2 = []
            for j in list_of_keys_selected_ind_2:
                set_stb_2.append(set())
                for i in range(len(df_2)):
                    if (df_2[i][-5] != "Да"):
                        set_stb_2[-1].add(df_2[i][j])

            self.progressBarMsg(73, 'Составление групп уникальных ключей(Частичное соотвествие)')
            # 3!
            ind_key_2 = {}
            # Len_stb = len(self.list_of_keys_2)
            for i in range(len(df_2)):
                try:
                    if (df_2[i][-5] != "Да"):
                        if df_2[i][-2] not in ind_key_2:
                            ind_key_2.update({df_2[i][-2]: i})
                except BaseException:
                    pass

            list_of_keys_selected_ind_3 = []
            for el in self.list_of_keys_selected_3:
                list_of_keys_selected_ind_3.append(self.list_of_keys_2.index(el))

            if (mode == 4):
                for el in self.list_of_keys_selected_4:
                    list_of_keys_selected_ind_4.append(self.list_of_keys.index(el))
                for el in self.list_of_keys_selected_5:
                    list_of_keys_selected_ind_5.append(self.list_of_keys_2.index(el))

            self.progressBarMsg(75, 'Сопоставление таблиц')

            for i in range(len(df)):
                if (df[i][-5] == "Да"):
                    df[i] = df[i][:-6] + [" "] + df[i][-6:]
                    if mode == 4:
                        df[i] = df[i][:-7] + [" "] * 3 * len(list_of_keys_selected_ind_4) + df[i][-7:]
                    df[i] = df[i][:-7] + [" "] * len(list_of_keys_selected_ind_3) + df[i][-7:]

                    continue
                if (df[i][-2] in ind_key_2):
                    df[i] = df[i][:-6] + ["Полное"] + df[i][-6:]
                    j = ind_key_2[df[i][-2]]
                    if mode == 4:
                        for k in range(len(list_of_keys_selected_ind_4)):
                            # print(df[i][list_of_keys_selected_ind_4[k]])
                            i2 = i
                            j2 = j
                            x1 = 0
                            x2 = 0
                            while i2 < len(df) and df[i2][-2] == df[i][-2]:
                                try:
                                    if (df[i2][-5] != "Да"):
                                        x1 += float(str(df[i2][list_of_keys_selected_ind_4[k]]).replace(",", "."))
                                except BaseException:
                                    x1 += 0
                                if i2 > i and k + 1 == len(list_of_keys_selected_ind_4):
                                    df[i2][-5] = "Да"
                                i2 += 1
                            while j2 < len(df_2) and df_2[j2][-2] == df_2[j][-2]:
                                try:
                                    if (df_2[j2][-5] != "Да"):
                                        x2 += float(str(df_2[j2][list_of_keys_selected_ind_5[k]]).replace(",", "."))
                                except BaseException:
                                    x2 += 0
                                if i2 > i and k + 1 == len(list_of_keys_selected_ind_5):
                                    df_2[j2][-5] = "Да"
                                j2 += 1
                            df[i] = df[i][:-7] + [str(x1).replace(".", ",")] + df[i][-7:]
                            df[i] = df[i][:-7] + [str(x2).replace(".", ",")] + df[i][-7:]
                            df[i] = df[i][:-7] + [str(x1 - x2).replace(".", ",")] + df[i][-7:]

                    for ind in list_of_keys_selected_ind_3:
                        df[i] = df[i][:-7] + [df_2[j][ind]] + df[i][-7:]

                else:
                    key = ""
                    if not  self.ui.checkBox_join_keys.isChecked():
                        for j in range(len(list_of_keys_selected_ind)):
                            if (df[i][list_of_keys_selected_ind[j]] in set_stb_2[j]):
                                key += self.list_of_keys[list_of_keys_selected_ind[j]] + " - " \
                                       + self.list_of_keys_2[list_of_keys_selected_ind_2[j]] + ", "

                    if (key == ""):
                        df[i] = df[i][:-6] + ["Нет"] + df[i][-6:]
                    else:
                        df[i] = df[i][:-6] + [key[:-2]] + df[i][-6:]
                    for k in range(len(list_of_keys_selected_ind_4)):
                        i2 = i
                        x1 = 0
                        while i2 < len(df) and df[i2][-2] == df[i][-2]:
                            try:
                                if (df[i2][-5] != "Да"):
                                    x1 += float(str(df[i2][list_of_keys_selected_ind_4[k]]).replace(",", "."))
                            except BaseException:
                                x1 += 0
                            if i2 > i:
                                df[i2][-5] = "Да"
                            i2 += 1
                        df[i] = df[i][:-7] + [str(x1).replace(".", ",")] + df[i][-7:]
                        df[i] = df[i][:-7] + [" "] * 2 + df[i][-7:]
                    df[i] = df[i][:-7] + [" "] * len(list_of_keys_selected_ind_3) + df[i][-7:]

            self.progressBarMsg(83, 'Подготовка результатов для вывода')
            if mode == 4:
                for k in range(len(self.list_of_keys_selected_4)):
                    add_stb.append(self.list_of_keys_selected_4[k] + "(Элемент 1)")
                    add_stb.append(self.list_of_keys_selected_5[k] + "(Элемент 2)")
                    add_stb.append(self.list_of_keys_selected_4[k] + "(Разность)")
            for el in self.list_of_keys_selected_3:
                add_stb.append(el + "(Добавленный)")
            add_stb.append("Совпадения")

        for i in range(len(df)):
            if (df[i][-7] == "Полное"):
                df[i][-7] = " Полное"

        sort_list(df, [len(df[0]) - 7, len(df[0]) - 6])
        for i in range(len(df)):
            if (df[i][-7] == " Полное"):
                df[i][-7] = "Полное"

        self.progressBarMsg(87, 'Вывод результатов в файл')
        filesize_ratio = 1.80
        t1 = time.strftime("%M мин %S сек", time.localtime(int(sum(self.files_size) * filesize_ratio)))
        t2 = time.strftime("%H:%M:%S", time.localtime(time.time() + int(sum(self.files_size)) * filesize_ratio))
        self.progressBarMsg(88, f'Примерное время выгрузки {t1}({t2})')
        out_file_name = ".".join(self.name_file_1[0].split(".")[:-1]) + "(Обработанный)"

        keys = self.list_of_keys.copy()
        if self.ui.checkBox_del_stb.isChecked():
            del_stb(df, keys, self.list_of_keys_selected)
            keys_ind = [i for i in range(len(keys))]
        keys += add_stb
        # keys =  self.list_of_keys + add_stb
        if self.ui.checkBox_color.isChecked():
            out_file = write_table(df, keys, out_file_name,
                                   self.ui.comboBox_del.currentIndex(), True,
                                   len(add_stb), keys_ind)
        else:
            out_file = write_table(df, keys, out_file_name,
                                   self.ui.comboBox_del.currentIndex())

        del df
        gc.collect()  # clear_memory
        self.progressBarMsg(97, 'Создан новый файл "' + out_file + '"')

        if self.ui.checkBox_open.isChecked():
            self.progressBarMsg(98, 'Открытие созданного файла')
            try:
                os.startfile(os.getcwd() + "\\Результаты\\" + out_file)
            except BaseException:
                self.progressBarMsg(99, 'Не удалось автоматически открыть файл результата')
        self.progressBarMsg(100, 'Готово!')
        # print("OK")

    def help(self):
        self.step = self.step + 1
        self.switch_window.emit()


class help_window(QtWidgets.QWidget):
    switch_window = QtCore.pyqtSignal()

    def __init__(self):
        super(help_window, self).__init__()
        self.ui = Ui_help_ui()
        self.ui.setupUi(self)
        self.ui.close_btn.clicked.connect(self.close)


# эта штука отвечает за переключение между окнами, запускает их и вообще Зая
class Controller:

    def __init__(self):
        pass

    def show_main_page(self):  # Проказывает главное окно
        self.main_page = mywindow()
        self.main_page.switch_window.connect(self.show_help_page)
        self.main_page.switch_window_list.connect(self.line_config_window)
        self.main_page.show()
        self.main_page.setWindowIcon(QIcon(":/newPrefix/logo.png"))

    def show_help_page(self):  # Проказывает  help
        self.help_page = help_window()
        self.help_page.show()

    def line_config_window(self):
        self.line_config = line_config_window(self.main_page.array_export)
        self.line_config.return_value.connect(lambda: self.main_page.save_config(self.line_config.ret))
        self.line_config.show()

#Обработка ошибок
def excepthook(exc_type, exc_value, exc_tb):
    tb = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    text = "Oбнаружена неизвестная ошибка!\n" \
           "Просим извинения за неудобства, обратитесь к сотрудникам НСИ и в кратчайшие " \
           "сроки ошибка будет устранена."
    print(text)

    controller.main_page.error_msg(text)
    controller.main_page.logger("\n" + tb + "\n")

def logger(path, msg, refresh = False):
    try:
        log_file = open(path, 'a')
        file_stats = os.stat(path)
        file_size = int(file_stats.st_size / (1024 * 1024) * 100) / 100
        if file_size > 25 and refresh:
            log_file.close()
            log_file = open(path, 'w')
            log_file.write("История работы программы")
        log_file.write(msg)
        log_file.close()
    except:
        pass

#    QtWidgets.QApplication.quit()             # !!! если вы хотите, чтобы событие завершилось

# Для отлова ошибок
sys.excepthook = excepthook
#
print("Запуск программы")
local_logger_path = 'log.txt'
msg_start = "\n\n\n" + "_" * 130 + "\n" +\
            str(time.strftime("%d.%m.%Y  %H:%M:%S", time.localtime(time.time()))) +\
            " - Запуск программы\n\n"
logger(local_logger_path, msg_start, True)

try:
    global_logger_name = os.environ.get("USERNAME")
    try:
        global_logger_name += " (" + win32api.GetUserNameEx(3) + ")"
    except BaseException:
        pass
    global_logger_path = "V:\\Обмен МБУ ФК\\23_Отдел НСИ\\Обработчик\\logger\\" +\
                         global_logger_name + ".txt"
    logger(global_logger_path, msg_start, True)
except BaseException:
    pass



app = QtWidgets.QApplication(sys.argv)
controller = Controller()
controller.show_main_page()
sys.exit(app.exec_())
