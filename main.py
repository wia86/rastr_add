# установка модулей:
# Qt Designer для работы с файлами.ui
# pip freeze > requirements11.txt
# pip install -r requirements11.txt
# exe приложение:
# pyinstaller --onefile --noconsole main.py
# pyinstaller -F --noconsole main.py
import win32com.client  # установить pywin32
# Excel.Application https://memotut.com/en/150745ae0cc17cb5c866/
from abc import ABC
from Rastr_Method import RastrMethod
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.styles.numbers import BUILTIN_FORMATS
from typing import Union  # Any
import sys
import shutil
from itertools import combinations
from PyQt5 import QtWidgets
from datetime import datetime
import time
import os
import re
import configparser  # создать ini файл
# import random
import logging
# import webbrowser
from tkinter import messagebox as mb
import numpy as np
import pandas as pd
from tabulate import tabulate
import yaml
from qt_choice import Ui_choice  # pyuic5 qt_choice.ui -o qt_choice.py # Запустить строку в Terminal
from qt_set import Ui_Settings  # pyuic5 qt_set.ui -o qt_set.py
from qt_cor import Ui_cor  # pyuic5 qt_cor.ui -o qt_cor.py
from qt_calc_ur import Ui_calc_ur  # pyuic5 qt_calc_ur.ui -o qt_calc_ur.py
from qt_calc_ur_set import Ui_calc_ur_set  # pyuic5 qt_calc_ur_set.ui -o qt_calc_ur_set.py
from collections import namedtuple, defaultdict
# Если не работает терминал, то в  PowerShell ввести:
# Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned –Force


class Window:
    """ Класс с общими методами для QT. """

    @staticmethod
    def check_status(set_checkbox_element: tuple):
        """
        По состоянию CheckBox изменить состояние видимости соответствующего элемента.
        :param set_checkbox_element: картеж картежей (checkbox, element)
        """
        for CB, element in set_checkbox_element:
            if CB.isChecked():
                element.show()
            else:
                element.hide()

    @staticmethod
    def hide_show(hide_window: tuple, show_window: tuple):
        """ Изменить состояние видимости окон."""
        for element in hide_window:
            element.hide()
        for element in show_window:
            element.show()

    def choice_file(self, directory: str, filter_: str = 'All Files(*)'):
        """
        Выбор файла.
        """
        fileName_choose, _ = QtWidgets.QFileDialog.getOpenFileName(self, directory=directory,
                                                                   filter=filter_)  # "All Files(*);Text Files(*.txt)"
        if fileName_choose:
            log.info(f"GUI. Выбран файл: {fileName_choose}")
            return fileName_choose

    def choice_folder(self, directory: str):
        """
        Выбор папки.
        """
        folder_Name_choose = QtWidgets.QFileDialog.getExistingDirectory(self, directory=directory)
        if folder_Name_choose:
            return folder_Name_choose

    def save_file(self, directory: str, filter_: str = ''):
        """
        Сохранение файла.
        """
        fileName_choose, _ = QtWidgets.QFileDialog.getSaveFileName(self, directory=directory, filter=filter_)
        if fileName_choose:
            log.info(f"GUI. Для сохранения выбран файл: {fileName_choose}, {_}")
            return fileName_choose

    def choice(self, type_choice: str, insert, directory=None):
        """
        Функция выбора папки или файла.
        :param type_choice: 'file', 'folder'
        :param insert: объект QT 'QPlainTextEdit' или 'QLineEdit' для вставки пути выбранного файла.
        :param directory: объект QT 'QPlainTextEdit' c начальной папкой для поиска.
        """
        name = ''
        if type_choice == 'file':
            name = self.choice_file(directory=directory.toPlainText().replace('*', ''))
        elif type_choice == 'folder':
            name = self.choice_folder(directory=directory.toPlainText().replace('*', ''))

        if name:
            name = name.replace('/', '\\')
            if insert.__class__.__name__ == 'QPlainTextEdit':
                insert.setPlainText(name)
            elif insert.__class__.__name__ == 'QLineEdit':
                insert.setText(name)


class MainChoiceWindow(QtWidgets.QMainWindow, Ui_choice, Window):
    """
    Окно главного меню.
    """
    def __init__(self):
        super(MainChoiceWindow, self).__init__()
        self.setupUi(self)
        self.settings.clicked.connect(lambda: gui_set.show())
        self.correction.clicked.connect(lambda: self.hide_show((gui_choice_window,), (gui_edit,)))
        self.calc_ur.clicked.connect(lambda: self.hide_show((gui_choice_window,), (gui_calc_ur,)))


class CalcWindow(QtWidgets.QMainWindow, Ui_calc_ur, Window):
    """
    Окно задания и запуска УР.
    """
    def __init__(self):
        super(CalcWindow, self).__init__()
        self.setupUi(self)
        self.task_calc = {}
        self.b_set.clicked.connect(lambda: gui_calc_ur_set.show())
        self.b_main_choice.clicked.connect(lambda: self.hide_show((gui_calc_ur,), (gui_choice_window,)))

        # Скрыть параметры при старте.
        self.check_status_visibility = (
            (self.cb_filter, self.gb_filter),
            (self.cb_cor_txt, self.te_cor_txt),
            (self.cb_import_model, self.gb_import_model),
            (self.cb_disable_comb, self.gb_disable_comb),
            (self.cb_disable_excel, self.gb_disable_excel),
            (self.cb_control, self.gb_control),
            (self.cb_tab_KO, self.gb_tab_KO),
            (self.cb_results_pic, self.gb_results_pic),
        )
        self.check_status(self.check_status_visibility)

        # CB показать / скрыть параметры.
        for CB, _ in self.check_status_visibility:
            CB.clicked.connect(lambda: self.check_status(self.check_status_visibility))

        # Функциональные кнопки
        # TODO self.b_task_save.clicked.connect(self.task_save_yaml)
        # TODO self.b_task_load.clicked.connect(self.task_load_yaml)

        self.b_choice_path_folder.clicked.connect(lambda: self.choice(type_choice='folder',
                                                                      insert=self.te_path_initial_models,
                                                                      directory=self.te_path_initial_models))
        self.b_choice_path_file.clicked.connect(lambda: self.choice(type_choice='file',
                                                                    insert=self.te_path_initial_models,
                                                                    directory=self.te_path_initial_models))
        self.b_choice_XL.clicked.connect(lambda: self.choice(type_choice='file', insert=self.te_XL_path,
                                                             directory=self.te_path_initial_models))
        self.b_choice_path_import_folder.clicked.connect(lambda: self.choice(type_choice='folder',
                                                                             insert=self.te_path_import_rg2,
                                                                             directory=self.te_path_initial_models))
        self.b_choice_path_import_file.clicked.connect(lambda: self.choice(type_choice='file',
                                                                           insert=self.te_path_import_rg2,
                                                                           directory=self.te_path_initial_models))

        self.run_calc_rg2.clicked.connect(lambda: self.start())
        self.te_path_initial_models.setPlainText(GeneralSettings.read_ini(section='save_form_folder_calc', key="path"))
        # Подсказки
        self.le_control_field.setToolTip("Например, 'sel'. Если '*' - то контролировать все ветви и узлы")
        self.te_path_initial_models.setToolTip("Для расчета файлов во всех вложенных папках нужно в конце поставить *")

    def start(self):
        """
        Запуск расчета моделей
        """
        GeneralSettings.write_ini(section='save_form_folder_calc', key="path",
                                  value=self.te_path_initial_models.toPlainText())
        self.fill_task_calc()
        global cm
        cm = CalcModel(self.task_calc)
        cm.run_calc()

    def fill_task_calc(self):
        self.task_calc = {
            # Окно запуска расчета.
            "calc_folder": self.te_path_initial_models.toPlainText().strip(),
            # Выборка файлов.
            "Filter_file": self.cb_filter.isChecked(),  # QCheckBox
            "file_count_max": self.sb_count_file.value(),  # QSpainBox
            "calc_criterion": {"years": self.le_condition_file_years.text(),  # QLineEdit text()
                               "season": self.le_condition_file_season.currentText(),  # QComboBox
                               "max_min": self.le_condition_file_max_min.currentText(),
                               "add_name": self.le_condition_file_add_name.text()},
            # Корректировка в начале.
            "cor_rm": {'add': self.cb_cor_txt.isChecked(),
                       'txt': self.te_cor_txt.toPlainText()},
            # Импорт ид для расчетов УР из моделей.
            'CB_Import_Rg2': self.cb_import_model.isChecked(),
            "Import_file": self.te_path_import_rg2.toPlainText(),
            'txt_Import_Rg2': self.te_import_rg2.toPlainText(),
            # Расчет всех возможных сочетаний.
            'cb_disable_comb': self.cb_disable_comb.isChecked(),
            "SRS": {'n-1': self.cb_n1.isChecked(),
                    'n-2': self.cb_n2.isChecked(),
                    'n-3': self.cb_n3.isChecked()},
            'cb_comb_field': self.cb_comb_field.isChecked(),
            "comb_field": self.le_comb_field.text(),

            'cb_auto_disable': self.cb_auto_disable.isChecked(),
            "auto_disable_choice": self.le_auto_disable_choice.text(),

            'filter_comb': self.cb_filter_comb.isChecked(),
            "filter_comb_val": self.le_filter_comb_val.text(),
            # Импорт перечня расчетных сочетаний из EXCEL
            'cb_disable_excel': self.cb_disable_excel.isChecked(),
            "srs_XL_path": self.te_XL_path.toPlainText(),
            'srs_XL_sheets': self.le_XL_sheets.text(),
            # Расчет всех возможных сочетаний.
            'cb_control': self.cb_control.isChecked(),
            'cb_control_field': self.cb_control_field.isChecked(),
            "le_control_field": self.le_control_field.text(),

            'cb_auto_control': self.cb_auto_control.isChecked(),
            "le_auto_control_choice": self.le_auto_control_choice.text(),
            'cb_Imax': self.cb_Imax.isChecked(),

            # Результаты в EXCEL: таблицы контролируемые - отключаемые элементы
            'cb_tab_KO': self.cb_tab_KO.isChecked(),
            'le_tab_KO_info': self.le_tab_KO_info.toPlainText(),

            # Результаты в RG2
            'results_RG2': self.cb_results_pic.isChecked(),
            # TODO настройки
        }
        """
        Заполнить task_calc задание взяв данные с формы QT.
        """


class CalcSetWindow(QtWidgets.QMainWindow, Ui_calc_ur_set, Window):
    """
    Окно основных настроек расчета УР.
    """

    def __init__(self):
        super(CalcSetWindow, self).__init__()
        self.setupUi(self)
        self.load_ini_ur()
        self.b_save.clicked.connect(lambda: self.save_ini_ur())

    def load_ini_ur(self):
        """Загрузить, создать или перезаписать файл .ini """
        if os.path.exists(GeneralSettings.ini):
            config = configparser.ConfigParser()
            config.read(GeneralSettings.ini)
            try:
                self.cb_gost.setChecked(eval(config['CalcSetWindow']["gost"]))
                self.cb_skrm.setChecked(eval(config['CalcSetWindow']["skrm"]))
                self.cb_avr.setChecked(eval(config['CalcSetWindow']["avr"]))
                self.cb_add_disabling_repair.setChecked(eval(config['CalcSetWindow']["add_disabling_repair"]))
                self.cb_pa.setChecked(eval(config['CalcSetWindow']["pa"]))
            except LookupError:
                log.error(f'файл {GeneralSettings.ini} не читается, перезаписан')
                self.save_ini_ur()
        else:
            log.info(f'создан файл {GeneralSettings.ini}')
            self.save_ini_ur()

    def save_ini_ur(self):
        config = configparser.ConfigParser()
        config.read(GeneralSettings.ini)
        config['CalcSetWindow'] = {
            "gost": self.cb_gost.isChecked(),
            "skrm": self.cb_skrm.isChecked(),
            "avr": self.cb_avr.isChecked(),
            "add_disabling_repair": self.cb_add_disabling_repair.isChecked(),
            "pa": self.cb_pa.isChecked()}
        with open(GeneralSettings.ini, 'w') as configfile:
            config.write(configfile)


class SetWindow(QtWidgets.QMainWindow, Ui_Settings, Window):
    """
    Окно общих настроек.
    """

    def __init__(self):
        super(SetWindow, self).__init__()
        self.setupUi(self)
        self.load_ini()
        self.set_save.clicked.connect(lambda: self.save_ini())

    def load_ini(self):
        """Загрузить, создать или перезаписать файл .ini """
        if os.path.exists(GeneralSettings.ini):
            config = configparser.ConfigParser()
            config.read(GeneralSettings.ini)
            try:
                self.LE_path.setText(config['DEFAULT']["folder RastrWin3"])
                self.LE_rg2.setText(config['DEFAULT']["шаблон rg2"])
                self.LE_rst.setText(config['DEFAULT']["шаблон rst"])
                self.LE_sch.setText(config['DEFAULT']["шаблон sch"])
                self.LE_amt.setText(config['DEFAULT']["шаблон amt"])
                self.LE_trn.setText(config['DEFAULT']["шаблон trn"])
            except LookupError:
                log.error(f'файл {GeneralSettings.ini} не читается, перезаписан')
                self.save_ini()
        else:
            log.info(f'создан файл {GeneralSettings.ini}')
            self.save_ini()

    def save_ini(self):
        config = configparser.ConfigParser()
        config.read(GeneralSettings.ini)
        config['DEFAULT'] = {
            "folder RastrWin3": self.LE_path.text(),
            "шаблон rg2": self.LE_rg2.text(),
            "шаблон rst": self.LE_rst.text(),
            "шаблон sch": self.LE_sch.text(),
            "шаблон amt": self.LE_amt.text(),
            "шаблон trn": self.LE_trn.text()}
        with open(GeneralSettings.ini, 'w') as configfile:
            config.write(configfile)


class EditWindow(QtWidgets.QMainWindow, Ui_cor, Window):
    """
    Окно корректировки моделей.
    """
    def __init__(self):
        super(EditWindow, self).__init__()  # *args, **kwargs
        self.setupUi(self)
        self.task_ui = {}
        self.check_import = (
            (self.CB_N, 'узлы'),
            (self.CB_V, 'ветви'),
            (self.CB_G, 'генераторы'),
            (self.CB_A, 'районы'),
            (self.CB_A2, 'территории'),
            (self.CB_D, 'объединения'),
            (self.CB_PQ, 'PQ'),
            (self.CB_IT, 'I(T)'),)

        # Скрыть параметры при старте.
        self.check_status_visibility = (
            (self.CB_KFilter_file, self.GB_sel_file),
            (self.CB_cor_b, self.TE_cor_b),
            (self.CB_ImpRg2, self.sel_import),
            (self.CB_import_val_XL, self.GB_import_val_XL),
            (self.CB_cor_e, self.TE_cor_e),
            (self.CB_kontrol_rg2, self.GB_control),
            (self.CB_Filtr_N, self.FFN),
            (self.CB_Filtr_V, self.FFV),
            (self.CB_Filtr_G, self.FFG),
            (self.CB_Filtr_A, self.FFA),
            (self.CB_Filtr_A2, self.FFA2),
            (self.CB_Filtr_D, self.FFD),
            (self.CB_Filtr_PQ, self.FFPQ),
            (self.CB_Filtr_IT, self.FFIT),
            (self.CB_printXL, self.GB_prinr_XL),
            (self.CB_print_tab_log, self.GB_sel_tabl),
            (self.CB_print_parametr, self.TA_parametr_vibor),
            (self.CB_print_balance_Q, self.balance_Q_vibor),)
        self.check_status(self.check_status_visibility)

        # CB показать / скрыть параметры.
        for CB, element in self.check_status_visibility:
            CB.clicked.connect(lambda: self.check_status(self.check_status_visibility))
        # CB показать список импортируемых моделей.
        for CB, _ in self.check_import:
            CB.clicked.connect(lambda: self.import_name_table())

        # Функциональные кнопки
        self.task_save.clicked.connect(self.task_save_yaml)
        self.task_load.clicked.connect(self.task_load_yaml)
        self.choice_from_folder.clicked.connect(lambda: self.choice(type_choice='folder', insert=self.T_IzFolder,
                                                                    directory=self.T_IzFolder))
        self.choice_from_file.clicked.connect(lambda: self.choice(type_choice='file', insert=self.T_IzFolder,
                                                                  directory=self.T_IzFolder))
        self.choice_in_folder.clicked.connect(lambda: self.choice(type_choice='folder', insert=self.T_InFolder,
                                                                  directory=self.T_IzFolder))
        self.choice_XL.clicked.connect(lambda: self.choice(type_choice='file', insert=self.T_PQN_XL_File,
                                                           directory=self.T_IzFolder))
        self.choice_N.clicked.connect(lambda: self.choice(type_choice='file', insert=self.file_N,
                                                          directory=self.T_IzFolder))
        self.choice_V.clicked.connect(lambda: self.choice(type_choice='file', insert=self.file_V,
                                                          directory=self.T_IzFolder))
        self.choice_G.clicked.connect(lambda: self.choice(type_choice='file', insert=self.file_G,
                                                          directory=self.T_IzFolder))
        self.choice_A.clicked.connect(lambda: self.choice(type_choice='file', insert=self.file_A,
                                                          directory=self.T_IzFolder))
        self.choice_A2.clicked.connect(lambda: self.choice(type_choice='file', insert=self.file_A2,
                                                           directory=self.T_IzFolder))
        self.choice_D.clicked.connect(lambda: self.choice(type_choice='file', insert=self.file_D,
                                                          directory=self.T_IzFolder))
        self.choice_PQ.clicked.connect(lambda: self.choice(type_choice='file', insert=self.file_PQ,
                                                           directory=self.T_IzFolder))
        self.choice_IT.clicked.connect(lambda: self.choice(type_choice='file', insert=self.file_IT,
                                                           directory=self.T_IzFolder))

        self.run_krg2.clicked.connect(lambda: self.start())
        self.b_main_choice.clicked.connect(lambda: self.hide_show((gui_edit,), (gui_choice_window,)))
        # Подсказки
        self.T_IzFolder.setToolTip("Для корректировки файлов во всех вложенных папках нужно в конце поставить *")
        # Загрузить из .ini начальный путь для T_IzFolder
        self.T_IzFolder.setPlainText(GeneralSettings.read_ini(section='save_form_folder_edit', key="path"))

    def import_name_table(self):
        """
        Сформировать строку имени CB_ImpRg2 для наглядности выбранных вкладок.
        """
        add_str = 'Импорт из файлa (.rg2)'
        for CB, name in self.check_import:
            if CB.isChecked():
                add_str += f', {name}'
        self.CB_ImpRg2.setText(add_str)

    def task_save_yaml(self):
        name_file_save = self.save_file(directory=self.T_IzFolder.toPlainText(), filter_="YAML Files (*.yaml)")
        if name_file_save:
            self.fill_task_ui()
            with open(name_file_save, 'w') as f:
                yaml.dump(data=self.task_ui, stream=f, default_flow_style=False, sort_keys=False)

    def task_load_yaml(self):
        name_file_load = self.choice_file(directory=self.T_IzFolder.toPlainText().replace('*', ''),
                                          filter_="YAML Files (*.yaml)")
        if not name_file_load:
            return
        with open(name_file_load) as f:
            task_yaml = yaml.safe_load(f)
        if not task_yaml:
            return

        self.T_IzFolder.setPlainText(task_yaml["KIzFolder"])
        self.T_InFolder.setPlainText(task_yaml["KInFolder"])

        self.CB_KFilter_file.setChecked(task_yaml["KFilter_file"])  # QCheckBox
        self.D_count_file.setValue(task_yaml["max_file_count"])  # QSpainBox
        self.condition_file_years.setText(task_yaml["cor_criterion_start"]["years"])  # QLineEdit text()
        self.condition_file_season.setCurrentText(task_yaml["cor_criterion_start"]["season"])  # QComboBox
        self.condition_file_max_min.setCurrentText(task_yaml["cor_criterion_start"]["max_min"])
        self.condition_file_add_name.setText(task_yaml["cor_criterion_start"]["add_name"])

        self.CB_cor_b.setChecked(task_yaml["cor_beginning_qt"]['add'])
        self.TE_cor_b.setPlainText(task_yaml["cor_beginning_qt"]['txt'])

        self.CB_import_val_XL.setChecked(task_yaml["import_val_XL"])
        self.T_PQN_XL_File.setPlainText(task_yaml["excel_cor_file"])
        self.T_PQN_Sheets.setText(task_yaml["excel_cor_sheet"])

        self.CB_cor_e.setChecked(task_yaml["cor_end_qt"]['add'])
        self.TE_cor_e.setPlainText(task_yaml["cor_end_qt"]['txt'])

        self.CB_kontrol_rg2.setChecked(task_yaml["checking_parameters_rg2"])
        self.CB_U.setChecked(task_yaml["control_rg2_task"]['node'])
        self.CB_I.setChecked(task_yaml["control_rg2_task"]['vetv'])
        self.CB_gen.setChecked(task_yaml["control_rg2_task"]['Gen'])
        self.CB_s.setChecked(task_yaml["control_rg2_task"]['section'])
        self.kontrol_rg2_Sel.setText(task_yaml["control_rg2_task"]['sel_node'])

        self.CB_printXL.setChecked(task_yaml["printXL"])
        self.CB_print_sech.setChecked(task_yaml['set_printXL']["sechen"]['add'])
        self.CB_print_area.setChecked(task_yaml['set_printXL']["area"]['add'])
        self.CB_print_area2.setChecked(task_yaml['set_printXL']["area2"]['add'])
        self.CB_print_darea.setChecked(task_yaml['set_printXL']["darea"]['add'])


        self.CB_print_tab_log.setChecked(task_yaml['set_printXL']["tab"]['add'])
        self.print_tab_log_ar_set.setText(task_yaml['set_printXL']["tab"]["sel"])
        self.print_tab_log_ar_tab.setText(task_yaml['set_printXL']["tab"]['tabl'])
        self.print_tab_log_ar_cols.setText(task_yaml['set_printXL']["tab"]['par'])
        self.print_tab_log_rows.setText(task_yaml['set_printXL']["tab"]['rows'])
        self.print_tab_log_cols.setText(task_yaml['set_printXL']["tab"]['columns'])
        self.print_tab_log_vals.setText(task_yaml['set_printXL']["tab"]['values'])

        self.CB_print_parametr.setChecked(task_yaml['print_parameters']['add'])
        self.TA_parametr_vibor.setPlainText(task_yaml['print_parameters']['sel'])

        self.CB_print_balance_Q.setChecked(task_yaml['print_balance_q']['add'])
        self.balance_Q_vibor.setText(task_yaml['print_balance_q']['sel'])

        self.CB_ImpRg2.setChecked(task_yaml['CB_ImpRg2'])
        self.CB_ImpRg2.setText(task_yaml['CB_ImpRg2_name'])

        dict_ = task_yaml['Imp_add']['node']
        self.CB_N.setChecked(dict_['add'])
        self.file_N.setText(dict_['import_file_name'])
        if 'selection' in dict_:
            self.CB_Filtr_N.setChecked(dict_['selection'])
        self.Filtr_god_N.setText(dict_["years"])
        self.Filtr_sez_N.setCurrentText(dict_["season"])
        self.Filtr_max_min_N.setCurrentText(dict_["max_min"])
        self.Filtr_dop_name_N.setText(dict_["add_name"])
        self.tab_N.setText(dict_['tables'])
        self.param_N.setText(dict_['param'])
        self.sel_N.setText(dict_['sel'])
        self.tip_N.setCurrentText(dict_['calc'])

        dict_ = task_yaml['Imp_add']['vetv']
        self.CB_V.setChecked(dict_['add'])
        self.file_V.setText(dict_['import_file_name'])
        if 'selection' in dict_:
            self.CB_Filtr_V.setChecked(dict_['selection'])
        self.Filtr_god_V.setText(dict_["years"])
        self.Filtr_sez_V.setCurrentText(dict_["season"])
        self.Filtr_max_min_V.setCurrentText(dict_["max_min"])
        self.Filtr_dop_name_V.setText(dict_["add_name"])
        self.tab_V.setText(dict_['tables'])
        self.param_V.setText(dict_['param'])
        self.sel_V.setText(dict_['sel'])
        self.tip_V.setCurrentText(dict_['calc'])

        dict_ = task_yaml['Imp_add']['gen']
        self.CB_G.setChecked(dict_['add'])
        self.file_G.setText(dict_['import_file_name'])
        if 'selection' in dict_:
            self.CB_Filtr_G.setChecked(dict_['selection'])
        self.Filtr_god_G.setText(dict_["years"])
        self.Filtr_sez_G.setCurrentText(dict_["season"])
        self.Filtr_max_min_G.setCurrentText(dict_["max_min"])
        self.Filtr_dop_name_G.setText(dict_["add_name"])
        self.tab_G.setText(dict_['tables'])
        self.param_G.setText(dict_['param'])
        self.sel_G.setText(dict_['sel'])
        self.tip_G.setCurrentText(dict_['calc'])

        dict_ = task_yaml['Imp_add']['area']
        self.CB_A.setChecked(dict_['add'])
        self.file_A.setText(dict_['import_file_name'])
        if 'selection' in dict_:
            self.CB_Filtr_A.setChecked(dict_['selection'])
        self.Filtr_god_A.setText(dict_["years"])
        self.Filtr_sez_A.setCurrentText(dict_["season"])
        self.Filtr_max_min_A.setCurrentText(dict_["max_min"])
        self.Filtr_dop_name_A.setText(dict_["add_name"])
        self.tab_A.setText(dict_['tables'])
        self.param_A.setText(dict_['param'])
        self.sel_A.setText(dict_['sel'])
        self.tip_A.setCurrentText(dict_['calc'])

        dict_ = task_yaml['Imp_add']['area2']
        self.CB_A2.setChecked(dict_['add'])
        self.file_A2.setText(dict_['import_file_name'])
        if 'selection' in dict_:
            self.CB_Filtr_A2.setChecked(dict_['selection'])
        self.Filtr_god_A2.setText(dict_["years"])
        self.Filtr_sez_A2.setCurrentText(dict_["season"])
        self.Filtr_max_min_A2.setCurrentText(dict_["max_min"])
        self.Filtr_dop_name_A2.setText(dict_["add_name"])
        self.tab_A2.setText(dict_['tables'])
        self.param_A2.setText(dict_['param'])
        self.sel_A2.setText(dict_['sel'])
        self.tip_A2.setCurrentText(dict_['calc'])

        dict_ = task_yaml['Imp_add']['darea']
        self.CB_D.setChecked(dict_['add'])
        if 'selection' in dict_:
            self.CB_Filtr_D.setChecked(dict_['selection'])
        self.file_D.setText(dict_['import_file_name'])
        self.Filtr_god_D.setText(dict_["years"])
        self.Filtr_sez_D.setCurrentText(dict_["season"])
        self.Filtr_max_min_D.setCurrentText(dict_["max_min"])
        self.Filtr_dop_name_D.setText(dict_["add_name"])
        self.tab_D.setText(dict_['tables'])
        self.param_D.setText(dict_['param'])
        self.sel_D.setText(dict_['sel'])
        self.tip_D.setCurrentText(dict_['calc'])

        dict_ = task_yaml['Imp_add']['PQ']
        self.CB_PQ.setChecked(dict_['add'])
        self.file_PQ.setText(dict_['import_file_name'])
        if 'selection' in dict_:
            self.CB_Filtr_PQ.setChecked(dict_['selection'])
        self.Filtr_god_PQ.setText(dict_["years"])
        self.Filtr_sez_PQ.setCurrentText(dict_["season"])
        self.Filtr_max_min_PQ.setCurrentText(dict_["max_min"])
        self.Filtr_dop_name_PQ.setText(dict_["add_name"])
        self.tab_PQ.setText(dict_['tables'])
        self.param_PQ.setText(dict_['param'])
        self.sel_PQ.setText(dict_['sel'])
        self.tip_PQ.setCurrentText(dict_['calc'])

        dict_ = task_yaml['Imp_add']['IT']
        self.CB_IT.setChecked(dict_['add'])
        self.file_IT.setText(dict_['import_file_name'])
        if 'selection' in dict_:
            self.CB_Filtr_IT.setChecked(dict_['selection'])
        self.Filtr_god_IT.setText(dict_["years"])
        self.Filtr_sez_IT.setCurrentText(dict_["season"])
        self.Filtr_max_min_IT.setCurrentText(dict_["max_min"])
        self.Filtr_dop_name_IT.setText(dict_["add_name"])
        self.tab_IT.setText(dict_['tables'])
        self.param_IT.setText(dict_['param'])
        self.sel_IT.setText(dict_['sel'])
        self.tip_IT.setCurrentText(dict_['calc'])

        self.check_status(self.check_status_visibility)

    def start(self):
        """
        Запуск корректировки моделей.
        """
        GeneralSettings.write_ini(section='save_form_folder_edit', key="path", value=self.T_IzFolder.toPlainText())

        self.fill_task_ui()
        global em
        em = EditModel(self.task_ui)
        self.gui_import()
        em.run_cor()

    def gui_import(self):
        """
        Добавление в ImportFromModel данных с формы.
        """
        if self.CB_ImpRg2.isChecked():
            for tables in self.task_ui['Imp_add']:
                if self.task_ui['Imp_add'][tables]['add']:
                    criterion_start = {}
                    if self.task_ui['Imp_add'][tables]['selection']:
                        criterion_start = {"years": self.task_ui['Imp_add'][tables]['years'],
                                           "season": self.task_ui['Imp_add'][tables]['season'],
                                           "max_min": self.task_ui['Imp_add'][tables]['max_min'],
                                           "add_name": self.task_ui['Imp_add'][tables]['add_name']}

                    ifm = ImportFromModel(import_file_name=self.task_ui['Imp_add'][tables]['import_file_name'],
                                          criterion_start=criterion_start,
                                          tables=self.task_ui['Imp_add'][tables]['tables'],
                                          param=self.task_ui['Imp_add'][tables]['param'],
                                          sel=self.task_ui['Imp_add'][tables]['sel'],
                                          calc=self.task_ui['Imp_add'][tables]['calc'])
                    ImportFromModel.set_import_model.append(ifm)

    def fill_task_ui(self):
        """
        Заполнить task_ui задание взяв данные с формы QT.
        """
        self.task_ui = {
            "KIzFolder": self.T_IzFolder.toPlainText(),  # QPlainTextEdit
            "KInFolder": self.T_InFolder.toPlainText(),
            # Выборка файлов.
            "KFilter_file": self.CB_KFilter_file.isChecked(),  # QCheckBox
            "max_file_count": self.D_count_file.value(),  # QSpainBox
            "cor_criterion_start": {"years": self.condition_file_years.text(),  # QLineEdit text()
                                    "season": self.condition_file_season.currentText(),  # QComboBox
                                    "max_min": self.condition_file_max_min.currentText(),
                                    "add_name": self.condition_file_add_name.text()},
            # Корректировка в начале.
            "cor_beginning_qt": {'add': self.CB_cor_b.isChecked(),
                                 'txt': self.TE_cor_b.toPlainText()},
            # Задание из 'EXCEL'
            "import_val_XL": self.CB_import_val_XL.isChecked(),
            "excel_cor_file": self.T_PQN_XL_File.toPlainText(),
            "excel_cor_sheet": self.T_PQN_Sheets.text(),
            # Корректировка в конце.
            "cor_end_qt": {'add': self.CB_cor_e.isChecked(),
                           'txt': self.TE_cor_e.toPlainText()},
            # Расчет режима и контроль параметров режима
            "checking_parameters_rg2": self.CB_kontrol_rg2.isChecked(),
            "control_rg2_task": {'node': self.CB_U.isChecked(),
                                 'vetv': self.CB_I.isChecked(),
                                 'Gen': self.CB_gen.isChecked(),
                                 'section': self.CB_s.isChecked(),
                                 'sel_node': self.kontrol_rg2_Sel.text()},
            # Выводить данные из моделей в XL
            "printXL": self.CB_printXL.isChecked(),
            "set_printXL": {
                "sechen": {'add': self.CB_print_sech.isChecked(),
                           "sel": 'ns>0',
                           'tabl': "sechen",
                           'par': "ns,name,pmin,pmax,psech",
                           "rows": "ns,name",  # поля строк в сводной
                           "columns": "год,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля столбцов в сводной
                           "values": "psech,pmax"},
                "area": {'add': self.CB_print_area.isChecked(),
                         "sel": 'na>0',
                         'tabl': "area",
                         'par': 'na,name,no,pg,pn,pn_sum,dp,pop,set_pop,qn_sum,pg_max,pg_min,poq,qn,qg,dev_pop',
                         "rows": "na,name,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля строк в сводной
                         "columns": "год",  # поля столбцов в сводной
                         "values": "pop,pg"},
                "area2": {'add': self.CB_print_area2.isChecked(),
                          "sel": 'npa>0',
                          'tabl': "area2",
                          'par': 'npa,name,pg,pn,dp,pop,vnp,qg,qn,dq,poq,vnq,pn_sum,qn_sum,set_pop,dev_pop',
                          "rows": "npa,name,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля строк в сводной
                          "columns": "год",  # поля столбцов в сводной
                          "values": "pop,pg"},
                "darea": {'add': self.CB_print_darea.isChecked(),
                          "sel": 'no>0',
                          'tabl': "darea",
                          'par': 'no,name,pg,pp,pvn,qn_sum,pnr_sum,pn_sum,set_pop,qvn,qp,qg,dev_pop',
                          "rows": "no,name,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля строк в сводной
                          "columns": "год",  # поля столбцов в сводной
                          "values": "pp,pg"},
                "tab": {'add': self.CB_print_tab_log.isChecked(),
                        "sel": self.print_tab_log_ar_set.text(),
                        'tabl': self.print_tab_log_ar_tab.text(),
                        'par': self.print_tab_log_ar_cols.text(),
                        "rows": self.print_tab_log_rows.text(),  # поля строк в сводной
                        "columns": self.print_tab_log_cols.text(),  # поля столбцов в сводной
                        "values": self.print_tab_log_vals.text()}},  # поля значений в сводной
            "print_parameters": {'add': self.CB_print_parametr.isChecked(),
                                 "sel": self.TA_parametr_vibor.toPlainText()},
            "print_balance_q": {'add': self.CB_print_balance_Q.isChecked(), "sel": self.balance_Q_vibor.text()},
            # только для UI
            'CB_ImpRg2_name': self.CB_ImpRg2.text(),
            'CB_ImpRg2': self.CB_ImpRg2.isChecked(),
            'Imp_add': {
                'node': {'add': self.CB_N.isChecked(),
                         'import_file_name': self.file_N.text(),
                         "selection": self.CB_Filtr_N.isChecked(),
                         "years": self.Filtr_god_N.text(),
                         "season": self.Filtr_sez_N.currentText(),
                         "max_min": self.Filtr_max_min_N.currentText(),
                         "add_name": self.Filtr_dop_name_N.text(),
                         'tables': self.tab_N.text(),
                         'param': self.param_N.text(),
                         'sel': self.sel_N.text(),
                         'calc': self.tip_N.currentText(), },
                'vetv': {'add': self.CB_V.isChecked(),
                         'import_file_name': self.file_V.text(),
                         "selection": self.CB_Filtr_V.isChecked(),
                         "years": self.Filtr_god_V.text(),
                         "season": self.Filtr_sez_V.currentText(),
                         "max_min": self.Filtr_max_min_V.currentText(),
                         "add_name": self.Filtr_dop_name_V.text(),
                         'tables': self.tab_V.text(),
                         'param': self.param_V.text(),
                         'sel': self.sel_V.text(),
                         'calc': self.tip_V.currentText(), },
                'gen': {'add': self.CB_G.isChecked(),
                        'import_file_name': self.file_G.text(),
                        "selection": self.CB_Filtr_G.isChecked(),
                        "years": self.Filtr_god_G.text(),
                        "season": self.Filtr_sez_G.currentText(),
                        "max_min": self.Filtr_max_min_G.currentText(),
                        "add_name": self.Filtr_dop_name_G.text(),
                        'tables': self.tab_G.text(),
                        'param': self.param_G.text(),
                        'sel': self.sel_G.text(),
                        'calc': self.tip_G.currentText(), },
                'area': {'add': self.CB_A.isChecked(),
                         'import_file_name': self.file_A.text(),
                         "selection": self.CB_Filtr_A.isChecked(),
                         "years": self.Filtr_god_A.text(),
                         "season": self.Filtr_sez_A.currentText(),
                         "max_min": self.Filtr_max_min_A.currentText(),
                         "add_name": self.Filtr_dop_name_A.text(),
                         'tables': self.tab_A.text(),
                         'param': self.param_A.text(),
                         'sel': self.sel_A.text(),
                         'calc': self.tip_A.currentText(), },
                'area2': {'add': self.CB_A2.isChecked(),
                          'import_file_name': self.file_A2.text(),
                          "selection": self.CB_Filtr_A2.isChecked(),
                          "years": self.Filtr_god_A2.text(),
                          "season": self.Filtr_sez_A2.currentText(),
                          "max_min": self.Filtr_max_min_A2.currentText(),
                          "add_name": self.Filtr_dop_name_A2.text(),
                          'tables': self.tab_A2.text(),
                          'param': self.param_A2.text(),
                          'sel': self.sel_A2.text(),
                          'calc': self.tip_A2.currentText(), },
                'darea': {'add': self.CB_D.isChecked(),
                          'import_file_name': self.file_D.text(),
                          "selection": self.CB_Filtr_D.isChecked(),
                          "years": self.Filtr_god_D.text(),
                          "season": self.Filtr_sez_D.currentText(),
                          "max_min": self.Filtr_max_min_D.currentText(),
                          "add_name": self.Filtr_dop_name_D.text(),
                          'tables': self.tab_D.text(),
                          'param': self.param_D.text(),
                          'sel': self.sel_D.text(),
                          'calc': self.tip_D.currentText(), },
                'PQ': {'add': self.CB_PQ.isChecked(),
                       'import_file_name': self.file_PQ.text(),
                       "selection": self.CB_Filtr_PQ.isChecked(),
                       "years": self.Filtr_god_PQ.text(),
                       "season": self.Filtr_sez_PQ.currentText(),
                       "max_min": self.Filtr_max_min_PQ.currentText(),
                       "add_name": self.Filtr_dop_name_PQ.text(),
                       'tables': self.tab_PQ.text(),
                       'param': self.param_PQ.text(),
                       'sel': self.sel_PQ.text(),
                       'calc': self.tip_PQ.currentText(), },
                'IT': {'add': self.CB_IT.isChecked(),
                       'import_file_name': self.file_IT.text(),
                       "selection": self.CB_Filtr_IT.isChecked(),
                       "years": self.Filtr_god_IT.text(),
                       "season": self.Filtr_sez_IT.currentText(),
                       "max_min": self.Filtr_max_min_IT.currentText(),
                       "add_name": self.Filtr_dop_name_IT.text(),
                       'tables': self.tab_IT.text(),
                       'param': self.param_IT.text(),
                       'sel': self.sel_IT.text(),
                       'calc': self.tip_IT.currentText(), },
            }
        }


class GeneralSettings(ABC):
    """
    Для хранения общих настроек.
    """
    # коллекция настроек, которые хранятся в ini файле
    set_save = {}
    ini = 'settings.ini'
    log_file = 'log_file.log'

    # @abstractmethod
    def __init__(self):
        # коллекция для хранения информации о расчете
        self.set_info = {"calc_val": {1: "ЗАМЕНИТЬ", 2: "ПРИБАВИТЬ", 3: "ВЫЧЕСТЬ", 0: "УМНОЖИТЬ"},
                         'collapse': '', 'end_info': ''}
        # прочитать ini файл
        if os.path.exists(self.ini):
            config = configparser.ConfigParser()
            config.read(self.ini)
            try:
                for key in config['DEFAULT']:
                    GeneralSettings.set_save[key] = config['DEFAULT'][key]
                for key in config['CalcSetWindow']:
                    GeneralSettings.set_save[key] = config['CalcSetWindow'][key]
            except LookupError:
                raise LookupError('файл settings.ini не читается')
        else:
            raise LookupError("Отсутствует файл settings.ini")

        self.file_count = 0  # Счетчик расчетных файлов.
        self.number_comb = 1  # счетчик общего количества расчетных комбинаций
        self.now = datetime.now()
        self.time_start = time.time()
        self.now_start = self.now.strftime("%d-%m-%Y %H:%M:%S")

    def the_end(self):  # по завершению
        execution_time = time.strftime("%H:%M:%S", time.gmtime(time.time() - self.time_start))
        self.set_info['end_info'] = (
            f"РАСЧЕТ ЗАКОНЧЕН! \nНачало расчета {self.now_start}, конец {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}"
            f" \nЗатрачено: {execution_time} (файлов: {self.file_count}).")
        log.info(self.set_info['end_info'])

    @staticmethod
    def read_ini(section: str, key: str):
        """
        Прочитать в файле settings.ini значение в разделе section по ключу key.
        """
        if os.path.exists(GeneralSettings.ini):
            config = configparser.ConfigParser()
            config.read(GeneralSettings.ini)
            try:
                return config[section][key]
            except LookupError:
                log.error(f'В файле {GeneralSettings.ini!r} не найден разделе {section!r} или ключ {key!r}.')

    @staticmethod
    def write_ini(section: str, key: str, value):
        """
        Записать в файл settings.ini значение value в раздел section по ключу key.
        """
        config = configparser.ConfigParser()
        config.read(GeneralSettings.ini)
        config[section] = {key: value}
        with open(GeneralSettings.ini, 'w') as configfile:
            config.write(configfile)


class CalcModel(GeneralSettings):
    """
    Расчет нормативных возмущений.
    """

    def __init__(self, task_calc):
        super(CalcModel, self).__init__()
        self.task_calc = task_calc
        self.all_folder = False  # Не перебирать вложенные папки
        self.set_comb = {}  # {количество отключений: контроль ДТН, 1:"ДДТН",2:"АДТН"}
        self.auto_shunt = {}
        
        self.control_I = None
        self.control_U = None

        self.disable_df_gen = None
        self.disable_df_node = None
        self.disable_df_vetv = None
        # DF для хранения токовых перегрузок и недопустимого снижения U

        self.srs_xl = pd.DataFrame()  # Перечень отключений их excel
        self.overloads_all = pd.DataFrame()  # общий
        self.overloads_srs = pd.DataFrame()  # СРС перегрузки
        self.info_srs = pd.Series(dtype='object')  # СРС
        self.info_action = pd.Series(dtype='object')  # действие ПА

        self.book_path: str = ''  # Путь к файлу excel.

        if self.task_calc['cb_tab_KO']:
            self.name_tab = self.task_calc['le_tab_KO_info'].strip()
            # Нумерация таблиц К-О
            self.num_tab = self.name_tab[self.name_tab.find('[') + 1: self.name_tab.find(']')]
            self.name_tab = self.name_tab.split(f'[{self.num_tab}]')
            self.num_tab = int(self.num_tab) if self.num_tab.isdigit() else 1

    def run_calc(self):
        """
        Запуск расчета нормативных возмущений (НВ) в РМ.
        """
        log.info('Запуск расчета нормативных возмущений (НВ) в расчетной модели (РМ).')
        if "*" in self.task_calc["calc_folder"]:
            self.task_calc["calc_folder"] = self.task_calc["calc_folder"].replace('*', '')
            self.all_folder = True

        if not os.path.exists(self.task_calc["calc_folder"]):
            raise ValueError(f'Не найден путь: {self.task_calc["calc_folder"]}.')

        # папка для сохранения результатов
        self.task_calc['folder_result_calc'] = self.task_calc["calc_folder"] + r"\result"
        if os.path.isfile(self.task_calc["calc_folder"]):
            self.task_calc['folder_result_calc'] = os.path.dirname(self.task_calc["calc_folder"]) + r"\result"
        if not os.path.exists(self.task_calc['folder_result_calc']):
            os.mkdir(self.task_calc['folder_result_calc'])  # создать папку result

        self.task_calc['name_time'] = f"{self.task_calc['folder_result_calc']}" \
                                      f"\\{datetime.now().strftime('%d-%m-%Y %H-%M-%S')}"

        if self.task_calc['cb_disable_excel']:
            self.srs_xl = pd.read_excel(self.task_calc['srs_XL_path'], sheet_name=self.task_calc['srs_XL_sheets'])

            self.srs_xl = self.srs_xl[self.srs_xl['Статус'] != '-']
            self.srs_xl.drop(columns=['Примечание', 'Статус'], inplace=True)
            self.srs_xl.dropna(how='all', axis=0,  inplace=True)
            self.srs_xl.dropna(how='all', axis=1, inplace=True)
            for col in self.srs_xl.columns:
                self.srs_xl[col] = self.srs_xl[col].str.split('#').str[0]
            self.srs_xl.fillna(0, inplace=True)

        # Цикл, если несколько файлов задания.
        if self.task_calc['CB_Import_Rg2'] and os.path.isdir(self.task_calc["Import_file"]):
            task_files = os.listdir(self.task_calc["Import_file"])
            task_files = list(filter(lambda x: x.endswith('.rg2'), task_files))
            for task_file in task_files:  # цикл по файлам '.rg2' в папке
                task_full_name = os.path.join(self.task_calc["Import_file"], task_file)
                log.info(f'Текущий файл задания: {task_full_name}')
                self.run_calc_task(task_full_name)
        else:
            if self.task_calc['CB_Import_Rg2']:
                self.run_calc_task(self.task_calc['Import_file'])
            else:
                self.run_calc_task()

        self.the_end()
        notepad_path = f'{self.task_calc["name_time"]} протокол расчета РМ.log'
        shutil.copyfile(GeneralSettings.log_file, notepad_path)
        with open(self.task_calc['name_time'] + ' задание на расчет РМ.yaml', 'w') as f:
            yaml.dump(data=self.task_calc, stream=f, default_flow_style=False, sort_keys=False)
        # webbrowser.open(notepad_path)  #  Открыть блокнотом лог-файл.
        mb.showinfo("Инфо", self.set_info['end_info'])

    @staticmethod
    def gen_comb_xl(rm, df: pd.DataFrame) -> pd.DataFrame:
        """
        Генератор комбинаций из XL
        :param rm:
        :param df:
        :return:  комбинацию comb_xl
        """
        for _, row in df.iterrows():
            comb_xl = pd.DataFrame(columns=['table',
                                            'index',
                                            'dname',
                                            'status_repair',
                                            'key',
                                            'repair_scheme',
                                            'disable_scheme'])
            if row['Ключ откл.']:
                table, index = rm.index_table_from_key(row['Ключ откл.'])
                if table and index >= 0:
                    scheme = rm.rastr.tables(table).cols.item("disable_scheme").Z(index)
                    if 'Схема при отключении' in row and row['Схема при отключении']:
                        scheme += ',' + row['Схема при отключении']
                    comb_xl.loc[len(comb_xl.index)] = [table,  # 'table'
                                                       index,  # 'index'
                                                       row['Отключение'],  # 'dname'
                                                       False,  # 'status_repair'
                                                       row['Ключ откл.'],  # 'key'
                                                       0,  # 'repair_scheme'
                                                       scheme]  # 'disable_scheme'
                else:
                    log.info(f'Задание комбинаций их XL: {row["Отключение"]!r} не найден ключ {row["Ключ откл."]!r}')
                    continue

            if row['Ключ рем.1']:
                table, index = rm.index_table_from_key(row['Ключ рем.1'])
                if table and index >= 0:
                    scheme = rm.rastr.tables(table).cols.item("repair_scheme").Z(index)
                    if 'Ремонтная схема' in row and row['Ремонтная схема']:
                        scheme += ',' + row['Ремонтная схема']
                    comb_xl.loc[len(comb_xl.index)] = [table,  # 'table'
                                                       index,  # 'index'
                                                       row['Ремонт 1'],  # 'dname'
                                                       True,  # 'status_repair'
                                                       row['Ключ рем.1'],  # 'key'
                                                       0,  # 'repair_scheme'
                                                       scheme]  # 'disable_scheme'
                else:
                    log.info(f'Задание комбинаций их XL: {row["Ремонт 1"]!r} не найден ключ {row["Ключ рем.1"]!r}')
                    continue

            if row['Ключ рем.2']:
                table, index = rm.index_table_from_key(row['Ключ рем.2'])
                if table and index >= 0:
                    comb_xl.loc[len(comb_xl.index)] = [table,  # 'table'
                                                       index,  # 'index'
                                                       row['Ремонт 2'],  # 'dname'
                                                       True,  # 'status_repair'
                                                       row['Ключ рем.2'],  # 'key'
                                                       0,  # 'repair_scheme'
                                                       rm.rastr.tables(table).cols.item("repair_scheme").Z(index)]
                else:
                    log.info(f'Задание комбинаций их XL: {row["Ремонт 2"]!r} не найден ключ {row["Ключ рем.2"]!r}')
                    continue
            yield comb_xl

    def run_calc_task(self, task_full_name: str = ''):
        """
        Запуск расчета с текущим файлом импорта задания или без него.
        :task_full_name: полный путь к текущему файлу задания
        """
        xlApp = None

        # Экспорт из модели исходных данных для расчетов УР.
        if task_full_name:
            ImportFromModel.set_import_model = []
            self.import_id_rg2(path_file=task_full_name, txt_task=self.task_calc['txt_Import_Rg2'])

        if os.path.isdir(self.task_calc["calc_folder"]):
            if self.all_folder:  # с вложенными папками
                for address, dir_, file_ in os.walk(self.task_calc["calc_folder"]):
                    self.for_file(folder_calc=address)
            else:  # без вложенных папок
                self.for_file(folder_calc=self.task_calc["calc_folder"])

        elif os.path.isfile(self.task_calc["calc_folder"]):
            rm = RastrModel(self.task_calc["calc_folder"])
            if not rm.code_name_rg2:
                raise ValueError(f'Имя файла {self.task_calc["calc_folder"]!r} не подходит.')
            self.calc_file(rm=rm)

        # Сохранить в Excel таблицу перегрузки.
        if len(self.overloads_all):
            # https://www.geeksforgeeks.org/how-to-write-pandas-dataframes-to-multiple-excel-sheets/
            mode = 'a' if os.path.exists(self.book_path) else 'w'
            with pd.ExcelWriter(path=self.book_path, mode=mode) as writer:
                if len(self.overloads_all):
                    for col in ["Отключение", "Ремонт 1", "Ремонт 2", "Доп. имя"]:
                        for col_df in self.overloads_all.columns:
                            if col in col_df:
                                self.overloads_all.fillna(value={col_df: 0}, inplace=True)
                                self.overloads_all.loc[self.overloads_all[col_df] == 0, col_df] = '-'
                    self.overloads_all.rename(columns={'dname': 'Контролируемые элементы'}, inplace=True)
                    self.overloads_all.to_excel(excel_writer=writer,
                                                float_format="%.2f",
                                                index_label='Номер сочетания.подсочетания',
                                                freeze_panes=(1, 1),
                                                sheet_name='Перегрузки')

        # Сводная
        if len(self.overloads_all):
            log.info('Формируется сводная таблица.')
            xlApp = win32com.client.Dispatch("Excel.Application")
            xlApp.ScreenUpdating = False  # Обновление экрана
            try:
                book = xlApp.Workbooks.Open(self.book_path)
            except Exception:
                raise Exception(f'Ошибка при открытии файла {self.book_path=}')
            try:
                sheet = book.sheets['Перегрузки']
            except Exception:
                raise Exception(f'Не найден лист Перегрузки')
            # Создать объект таблица из всего диапазона листа.
            tabl_overload = sheet.ListObjects.Add(SourceType=1, Source=sheet.Range(sheet.UsedRange.address))
            tabl_overload.Name = "Таблица_Перегрузки"
            pt_cache = book.PivotCaches().add(1, tabl_overload)  # Создать КЭШ xlDatabase, ListObjects

            task_pivot = []
            task_pivot_i = namedtuple('task_pivot',
                                      ['sheet_name', 'pivot_table_name', 'data_field'])
            # todo сводная если только режимы которые не моделируются
            if "i_max" in self.overloads_all.columns:
                task_pivot.append(task_pivot_i('Сводная_I', "Свод_I",
                                               dict(i_max="Iрасч.,A",
                                                    i_dop_r="Iддтн,A",
                                                    i_zag="Iзагр. ддтн,%",
                                                    i_dop_r_av="Iадтн,A",
                                                    i_zag_av="Iзагр. адтн,%")))
            if "umin" in self.overloads_all.columns:
                task_pivot.append(task_pivot_i('Сводная_Umin', "Свод_Umin",
                                               dict(vras="Uрасч.,кВ",
                                                    umin="Uмдн,кВ",
                                                    otv_min="Uмдн,%",
                                                    umin_av="Uaдн,кВ",
                                                    otv_min_av="Uaдн,%")))
            if "umax" in self.overloads_all.columns:
                task_pivot.append(task_pivot_i('Сводная_Umax', "Свод_Umax",
                                               dict(vras="Uрасч.,кВ",
                                                    umax="Uнаиб.раб.,кВ",
                                                    otv_max="Uнаиб.раб.,%")))

            RowFields = [col for col in ["Контролируемые элементы", "Отключение", "Ремонт 1", "Ремонт 2"]
                         if col in self.overloads_all.columns]

            ColumnFields = ["Год", "Сезон макс/мин"] + [col for col in self.overloads_all.columns if "Доп. имя" in col]

            for task in task_pivot:
                sheet_pivot = book.Sheets.Add()
                sheet_pivot.Name = task.sheet_name

                pt = pt_cache.CreatePivotTable(TableDestination=task.sheet_name+"!R1C1",
                                               TableName=task.pivot_table_name)
                pt.ManualUpdate = True  # True не обновить сводную
                pt.AddFields(RowFields=RowFields,
                             ColumnFields=ColumnFields,
                             PageFields=["Имя файла", 'Кол. откл. эл.', 'Конец', 'Наименование СРС', 'Контроль ДТН',
                                         'Темп.(°C)'],
                             AddToTable=False)
                for field_df, field_pt in task.data_field.items():
                    pt.AddDataField(Field=pt.PivotFields(field_df), Caption=field_pt, Function=-4157)
                    pt.PivotFields(field_pt).NumberFormat = "0"

                pt.PivotFields("Контролируемые элементы").ShowDetail = True  # группировка
                pt.RowAxisLayout(1)  # 1 xlTabularRow показывать в табличной форме!!!!
                pt.DataPivotField.Orientation = 1  # xlRowField = 1 "Значения" в столбцах или строках xlColumnField
                pt.RowGrand = False  # Удалить строку общих итогов
                pt.ColumnGrand = False  # Удалить столбец общих итогов
                pt.MergeLabels = True  # Объединять одинаковые ячейки
                pt.HasAutoFormat = False  # Не обновлять ширину при обновлении
                pt.NullString = "--"  # Заменять пустые ячейки
                pt.PreserveFormatting = False  # Сохранять формат ячеек при обновлении
                pt.ShowDrillIndicators = False  # Показывать кнопки свертывания
                for row in RowFields + ColumnFields:
                    pt.PivotFields(row).Subtotals = [False, False, False, False, False, False, False, False, False,
                                                     False, False, False]  # промежуточные итоги и фильтры
                field = list(task.data_field)[2]
                pt.PivotFields(field).Orientation = 3  # xlPageField = 3
                pt.PivotFields(field).CurrentPage = "(All)"  #
                if len(task_pivot) > 1:
                    pt.PivotFields(field).PivotItems("(blank)").Visible = False
                pt.TableStyle2 = ""  # стиль
                pt.ColumnRange.ColumnWidth = 10  # ширина строк
                pt.RowRange.ColumnWidth = 20
                pt.DataBodyRange.HorizontalAlignment = -4108  # xlCenter = -4108
                pt.TableRange1.WrapText = True  # перенос текста в ячейке
                for i in range(7, 13):
                    pt.TableRange1.Borders(i).LineStyle = 1  # лево
                # Условное форматирование
                for i in range(3, len(task.data_field) + 1, 2):
                    dpz = pt.DataBodyRange.Rows(i).Cells(1)
                    dpz.FormatConditions.AddColorScale(2)  # ColorScaleType:=2
                    dpz.FormatConditions(dpz.FormatConditions.count).SetFirstPriority()
                    dpz.FormatConditions(1).ColorScaleCriteria(1).Type = 0  # xlConditionValueNumber = 0
                    if list(task.data_field)[2] == 'i_zag':
                        dpz.FormatConditions(1).ColorScaleCriteria(1).Value = 100
                    else:
                        dpz.FormatConditions(1).ColorScaleCriteria(1).Value = 0
                    dpz.FormatConditions(1).ColorScaleCriteria(1).FormatColor.ThemeColor = 1  # xlThemeColorDark1 = 1
                    dpz.FormatConditions(1).ColorScaleCriteria(1).FormatColor.TintAndShade = 0
                    dpz.FormatConditions(1).ColorScaleCriteria(2).Type = 2  # xlConditionValueHighestValue = 2
                    dpz.FormatConditions(1).ColorScaleCriteria(2).FormatColor.ThemeColor = 3 + i  # номер темы
                    dpz.FormatConditions(1).ColorScaleCriteria(2).FormatColor.TintAndShade = -0.249977111117893
                    dpz.FormatConditions(1).ScopeType = 2  # xlDataFieldScope = 2 применить ко всем значениям поля
                    pass
                pt.ManualUpdate = False  # обновить сводную
            book.Save()
            book.Close()
        else:
            log.info('Отклонений параметров режима от допустимых значений не выявлено.')

        # Вставить таблицы К-О в word.
        if self.task_calc['cb_tab_KO']:
            log.info('Вставить таблицы К-О в word.')

            xlApp = win32com.client.Dispatch("Excel.Application")
            xlApp.Visible = False
            book = xlApp.Workbooks.Open(self.book_path)

            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.ScreenUpdating = False
            doc = word.Documents.Add()  # doc = word.Documents.Open(r"I:\file.docx")

            doc.PageSetup.PageWidth = 29.7 * 28.35  # CentimetersToPoints( format_list_i (2) ) 1 см = 28,35
            doc.PageSetup.PageHeight = 42.0 * 28.35  # CentimetersToPoints( format_list_i (1) )
            doc.PageSetup.Orientation = 1  # 1 книжная или 0 альбомная

            cursor = word.Selection
            cursor.Font.Size = 12
            cursor.Font.Name = "Times New Roman"
            cursor.EndKey(Unit=6)  # перейти в конец текста

            for i in range(1, book.Worksheets.Count + 1):
                if book.Worksheets(i).name[:1].isnumeric():
                    sheet = book.Worksheets(i)
                    cursor.TypeText(Text=sheet.Cells(1, 1).value)
                    cursor.TypeParagraph()

                    sheet.Range(sheet.UsedRange.address.replace('$', '').replace('A1', 'A2')).Copy()
                    # cursor.PasteExcelTable(LinkedToExcel=False, WordFormatting=False, RTF=False)
                    cursor.PasteAndFormat(Type=13)  # 13 Вставить в виде рисунка.
                    cursor.InsertBreak(Type=0)
                    # разрыв:7 страницы с новой строки, 0-в той же строке,1 и 8 колонки,
                    # 2-5 раздела со след стр,6 и 9-11 перенос на новую стр

            word.ScreenUpdating = True
            doc.SaveAs2(FileName=self.task_calc['name_time'] + ' таблицы К-О.docx')  # FileFormat=16 .docx
            doc.Close()

        if xlApp:
            xlApp.Visible = True
            xlApp.ScreenUpdating = True  # обновление экрана

    def for_file(self, folder_calc: str):
        """
        Цикл по файлам.
        :param folder_calc:
        """
        files_calc = os.listdir(folder_calc)  # список всех файлов в папке
        rm_files = list(filter(lambda x: x.endswith('.rg2'), files_calc))

        for rastr_file in rm_files:  # цикл по файлам '.rg2' в папке
            if self.task_calc["Filter_file"] and self.file_count == self.task_calc["file_count_max"]:
                break  # Если включен фильтр файлов проверяем количество расчетных файлов.
            full_name = os.path.join(folder_calc, rastr_file)
            rm = RastrModel(full_name)
            # если включен фильтр файлов и имя стандартизовано
            if not rm.code_name_rg2:
                log.info(f'Имя файла {full_name} не подходит.')
                continue
            if self.task_calc["Filter_file"]:
                if not rm.test_name(condition=self.task_calc["calc_criterion"],
                                    info=f'Имя файла {full_name} не подходит.'):
                    continue  # пропускаем, если не соответствует фильтру
            self.calc_file(rm)

    def calc_file(self, rm):
        """
        Рассчитать РМ.
        """
        self.file_count += 1
        self.book_path = self.task_calc['name_time'] + ' результаты расчетов.xlsx'
        rm.load()
        log.info(f"Расчетная температура: {rm.temperature}")
        rm.rastr.CalcIdop(rm.temperature, 0.0, "")
        if self.task_calc['cor_rm']['add']:
            log.info("\t*** Внесения изменений в РМ. ***")
            rm.cor_rm_from_txt(self.task_calc['cor_rm']['txt'])
            log.info("\t*** Внесение изменений в РМ выполнено. ***")

        # Импорт моделей
        if ImportFromModel.set_import_model:
            for im in ImportFromModel.set_import_model:
                im.import_data_in_rm(rm)

        # Подготовка.
        rm.voltage_fix_frame()
        if GeneralSettings.set_save['skrm']:
            self.auto_shunt = rm.auto_shunt_rec(selection='')

        # создать поле index
        rm.add_fields_in_table(name_tables='vetv,node,Generator', fields='index', type_fields=0)
        rm.table_index('vetv,node,Generator')

        # Поля для сортировки ветвей и др.
        rm.add_fields_in_table(name_tables='vetv', fields='temp,temp1', type_fields=1)

        # Поля для контроля напряжений
        rm.add_fields_in_table(name_tables='node', fields='umin_av', type_fields=1)
        rm.add_fields_in_table(name_tables='node', fields='otv_min', type_fields=1,
                               prop=((5, 'if(sta=0) (-vras+umin)/umin*100:0'),),
                               replace=True)
        rm.add_fields_in_table(name_tables='node', fields='otv_min_av', type_fields=1,
                               prop=((5, 'if(sta=0) (-vras+umin_av)/umin_av*100:0'),),
                               replace=True)
        rm.add_fields_in_table(name_tables='node', fields='otv_max', type_fields=1,
                               prop=((5, 'if(sta=0) (vras-umax)/umax*100:0'),))
        # Поля для загрузки ветвей
        rm.add_fields_in_table(name_tables='vetv', fields='i_zag_av', type_fields=1,
                               prop=((5, 'if(ktr!=0) zag_it_av:zag_i_av'), (12, 1000),))

        # Поля для автоматики, что бы не было ошибок
        rm.add_fields_in_table(name_tables='vetv,node,Generator',
                               fields='repair_scheme,disable_scheme,automation,dname', type_fields=2)
        # Поля с ключами таблиц
        rm.add_fields_in_table(name_tables='vetv', fields='key', type_fields=2,
                               prop=((5, '"ip="+str(ip)+"&iq="+str(iq)+"&np="+str(np)'),))
        rm.add_fields_in_table(name_tables='node', fields='key', type_fields=2,
                               prop=((5, '"ny="+str(ny)'),))
        rm.add_fields_in_table(name_tables='Generator', fields='key', type_fields=2,
                               prop=((5, '"Num="+str(Num)'),))

        # Сохранить текущее состояние РМ
        rm.add_fields_in_table(name_tables='vetv,node,Generator', fields='staRes', type_fields=3)
        rm.rastr.tables('vetv').cols.item("staRes").calc('sta')
        rm.rastr.tables('node').cols.item("staRes").calc('sta')
        rm.rastr.tables('Generator').cols.item("staRes").calc('sta')
        rm.add_fields_in_table(name_tables='node', fields='pnRes,qnRes,pgRes', type_fields=1)
        rm.rastr.tables('node').cols.item("pnRes").calc('pn')
        rm.rastr.tables('node').cols.item("qnRes").calc('qn')
        rm.rastr.tables('node').cols.item("pgRes").calc('pg')
        rm.add_fields_in_table(name_tables='Generator', fields='PRes', type_fields=1)
        rm.rastr.tables('Generator').cols.item("PRes").calc('P')
        rm.add_fields_in_table(name_tables='vetv', fields='ktrRes', type_fields=1)
        rm.rastr.tables('vetv').cols.item("ktrRes").calc('ktr')

        # Контролируемые элементы сети.
        if self.task_calc['cb_control']:
            # all_control для отметки всех контролируемых узлов и ветвей (авто и field)
            rm.add_fields_in_table(name_tables='vetv,node', fields='all_control', type_fields=3)

            if self.task_calc['cb_auto_control']:
                # todo заполнить all_control в соответствии с self.task_calc['le_auto_control_choice']
                pass

            if self.task_calc['cb_control_field']:

                if '*' in self.task_calc['le_control_field']:
                    rm.rastr.Tables("node").cols.item("all_control").Calc("1")
                    rm.rastr.Tables("vetv").cols.item("all_control").Calc("1")
                else:
                    # Добавит поле отметки отключений если их нет в какой-то таблице.
                    rm.add_fields_in_table(name_tables='vetv,node',
                                           fields=self.task_calc['le_control_field'],
                                           type_fields=3)
                    for table_name in ('vetv', 'node',):
                        rm.group_cor(tabl=table_name,
                                     param="all_control",
                                     selection=self.task_calc['le_control_field'],
                                     formula='1')

                    # all_control_groupid для отметки всех контролируемых ветвей и ветвей с теми же groupid
                    if not self.task_calc['cb_tab_KO']:
                        rm.add_fields_in_table(name_tables='vetv', fields='all_control_groupid', type_fields=3)
                        rm.rastr.tables('vetv').cols.item("all_control_groupid").calc("all_control")
                        tv = rm.rastr.tables('vetv')
                        tv.SetSel('all_control')
                        i = tv.FindNextSel(-1)
                        while i >= 0:
                            if tv.Cols.item("groupid").Z(i):
                                rm.group_cor(tabl='vetv',
                                             param="all_control",
                                             selection="groupid=" + tv.Cols.item("groupid").ZS(i),
                                             formula='1')
                            i = tv.FindNextSel(i)
            # Узлы
            tn = rm.rastr.tables('node')
            tn.SetSel('all_control')
            i = tn.FindNextSel(-1)
            while i >= 0:
                if not tn.Cols.item("dname").ZS(i).strip():
                    if tn.Cols.item("name").ZS(i):
                        tn.Cols.item("dname").SetZ(i, tn.Cols.item("name").ZS(i))
                    else:
                        tn.Cols.item("dname").SetZ(i, f'Узел номер {tn.Cols.item("ny").ZS(i)} ,без имени')
                i = tn.FindNextSel(i)

            # Ветви
            tv = rm.rastr.tables('vetv')
            tv.SetSel('all_control')
            i = tv.FindNextSel(-1)
            while i >= 0:
                if not tv.Cols.item("dname").ZS(i).strip():
                    tv.Cols.item("dname").SetZ(i, tv.Cols.item("name").ZS(i))
                i = tv.FindNextSel(i)

            # Таблицы контроль - отключение.
            if self.task_calc['cb_tab_KO']:
                rm.rastr.tables('vetv').cols.item("temp").calc('ip.uhom')
                rm.rastr.tables('vetv').cols.item("temp1").calc('iq.uhom')
                self.control_I = rm.fd_from_table(table_name='vetv',
                                                  fields='index,dname,temp,temp1,i_dop_r,i_dop_r_av,groupid,key,tip',
                                                  # ip, iq, np, name
                                                  setsel="all_control")
                if len(self.control_I):
                    self.control_I['uhom'] = (self.control_I[['temp', 'temp1']].max(axis=1) * 10000 +
                                              self.control_I[['temp', 'temp1']].min(axis=1))
                    self.control_I.sort_values(by=['tip', 'uhom', 'dname'],  # столбцы сортировки
                                               ascending=(False, False, True),  # обратный порядок
                                               inplace=True)  # изменить df
                    self.control_I.drop(['temp', 'temp1', 'uhom', 'tip'], axis=1, inplace=True)
                    self.control_I['i_dop_r'] = self.control_I['i_dop_r'].round(0).astype(int)
                    self.control_I['i_dop_r_av'] = self.control_I['i_dop_r_av'].round(0).astype(int)
                    self.control_I.rename(columns={'i_dop_r': 'ДДТН, А',
                                                   'i_dop_r_av': 'АДТН, А',
                                                   'dname': 'Контролируемый элемент'}, inplace=True)
                    self.control_I.set_index('index', inplace=True)
                    self.control_I = self.control_I.T
                    self.control_I.index = pd.MultiIndex.from_product([['-'], ['-'], self.control_I.index])

                self.control_U = rm.fd_from_table(table_name='node',
                                                  fields='index,dname,umin,umin_av,uhom',  # ,ny,umax
                                                  setsel="all_control")
                if len(self.control_U):
                    self.control_U.sort_values(by=['uhom', 'dname'],  # столбцы сортировки
                                               ascending=(False, True),  # обратный порядок
                                               inplace=True)  # изменить df
                    self.control_U.drop(['uhom'], axis=1, inplace=True)
                    self.control_U['umin'] = self.control_U['umin'].round(1)
                    self.control_U['umin_av'] = self.control_U['umin_av'].round(1)
                    self.control_U.rename(columns={'umin': 'МДН, кВ',
                                                   'umin_av': 'АДН, кВ',
                                                   'dname': 'Контролируемый элемент'}, inplace=True)
                    self.control_U.set_index('index', inplace=True)
                    self.control_U = self.control_U.T
                    self.control_U.index = pd.MultiIndex.from_product([['-'], ['-'], self.control_U.index])

        # Нормальная схема сети
        self.info_srs = pd.Series(dtype='object')  # СРС
        self.info_srs['Наименование СРС'] = 'Нормальная схема сети.'
        self.info_srs['Номер СРС'] = self.number_comb
        self.info_srs['Кол. откл. эл.'] = 0
        self.info_srs['Контроль ДТН'] = 'ДДТН'
        self.do_action(rm)

        # Отключаемые элементы сети.
        if self.task_calc['cb_disable_comb']:
            # Выбор количества одновременно отключаемых элементов
            if self.task_calc['SRS']['n-1']:
                self.set_comb[1] = 'ДДТН'
            if self.task_calc['SRS']['n-2']:
                self.set_comb[2] = 'ДДТН'
                if 0 < rm.code_name_rg2 < 4 and GeneralSettings.set_save['gost']:
                    self.set_comb[2] = 'AДТН'
            if self.task_calc['SRS']['n-3']:
                if GeneralSettings.set_save['gost']:
                    if rm.code_name_rg2 > 3:
                        self.set_comb[3] = 'АДТН'
                else:
                    self.set_comb[3] = 'ДДТН'
            log.info(f'Расчетные СРС: {self.set_comb}.')

            # В поле all_disable складываем элементы авто отмеченные и отмеченные в поле comb_field
            rm.add_fields_in_table(name_tables='vetv,node,Generator', fields='all_disable', type_fields=3)

            if self.task_calc['cb_auto_disable']:
                # Выбор отключаемых элементов автоматически из выборки в таблице узлы
                # Отмечается в таблицах ветви и узлы поле all_disable
                # todo self.task_calc['auto_disable_choice']
                pass

            # Выбор отключаемых элементов из отмеченных в поле comb_field
            if self.task_calc['cb_comb_field']:
                # Добавит поле отметки отключений если их нет в какой-то таблице
                rm.add_fields_in_table(name_tables='vetv,node,Generator', fields=self.task_calc['comb_field'],
                                       type_fields=3)
                for table_name in 'vetv,node,Generator'.split(','):
                    rm.group_cor(tabl=table_name,
                                 param="all_disable",
                                 selection=self.task_calc['comb_field'],
                                 formula='1')

            # Создать df отключаемых узлов и ветвей и генераторов. Сортировка.
            columns_pa = ',repair_scheme,disable_scheme'  # 'automation,automation_sta'
            # Генераторы
            self.disable_df_gen = rm.fd_from_table(table_name='Generator',
                                                   fields='index,Name,dname,key' + columns_pa,   # ,Num,NodeState,Node
                                                   setsel="all_disable")
            self.disable_df_gen['table'] = 'Generator'
            self.disable_df_gen.rename(columns={'Name': 'name'}, inplace=True)  # , 'Node': 'ny'
            # Узлы
            self.disable_df_node = rm.fd_from_table(table_name='node',
                                                    fields='index,name,uhom,dname,key' + columns_pa,  # ny
                                                    setsel="all_disable")
            # self.disable_df_node.index = self.disable_df_node['index']
            self.disable_df_node['table'] = 'node'
            self.disable_df_node.sort_values(by=['uhom', 'name'],  # столбцы сортировки
                                             ascending=(False, True),  # обратный порядок
                                             inplace=True)  # изменить df
            # Ветви
            self.disable_df_vetv = rm.fd_from_table(table_name='vetv',
                                                    fields='index,name,dname,key,temp,temp1,tip' + columns_pa,
                                                    setsel="all_disable")
            self.disable_df_vetv['table'] = 'vetv'
            self.disable_df_vetv['uhom'] = self.disable_df_vetv[['temp', 'temp1']].max(axis=1) * 10000 + \
                                           self.disable_df_vetv[['temp', 'temp1']].min(axis=1)
            self.disable_df_vetv.sort_values(by=['tip', 'uhom', 'name'],  # столбцы сортировки
                                             ascending=(False, False, True),  # обратный порядок
                                             inplace=True)  # изменить df
            self.disable_df_vetv.drop(['temp', 'temp1', 'tip'], axis=1, inplace=True)

            log.info(f'Количество отключаемых ветвей: {len(self.disable_df_vetv.axes[0])},'
                     f' узлов: {len(self.disable_df_node.axes[0])},'
                     f' генераторов: {len(self.disable_df_gen.axes[0])}.')

            disable_df_all = pd.concat([self.disable_df_vetv, self.disable_df_node, self.disable_df_gen])
            # удалить пробелы и значения после #
            disable_df_all.loc[disable_df_all['dname'] == '', 'dname'] = \
                disable_df_all.loc[disable_df_all['dname'] == '', 'name']
            disable_df_all['dname'] = disable_df_all['dname'].str.replace('  ', ' ').str.split('(').str[0]
            disable_df_all['dname'] = disable_df_all['dname'].str.split(',').str[0].str.strip()
            for col in ['disable_scheme', 'repair_scheme']:
                disable_df_all[col] = disable_df_all[col].str.replace(' ', '').str.split('#').str[0]

            # Цикл по всем возможным сочетаниям отключений
            for n_, self.info_srs['Контроль ДТН'] in self.set_comb.items():  # Цикл н-1 н-2 н-3.
                log.info(f"Количество отключаемых элементов в комбинации: {n_} ({self.info_srs['Контроль ДТН']}).")
                if n_ == 1:
                    disable_all = disable_df_all
                else:
                    disable_all = \
                        disable_df_all[(disable_df_all['uhom'] > 300) | (disable_df_all['table'] != 'node')]
                name_columns = list(disable_all.columns)
                disable_all = tuple(disable_all.itertuples(index=False, name=None))  # df в tuple построчно

                for comb in combinations(disable_all, r=n_):  # Цикл по комбинациям.
                    comb_df = pd.DataFrame(data=comb, columns=name_columns)

                    comb_df['status_repair'] = False  # Истина, если элемент в ремонте. Ложь отключен.
                    comb_df['dif_scheme'] = comb_df['disable_scheme'] != comb_df['repair_scheme']
                    if n_ < 3:
                        # Первое подсочетнание.
                        comb_df.loc[comb_df.index < n_ - 1, 'status_repair'] = True
                        self.calc_comb(rm, comb_df)

                        # Если хотя бы у одного элемента disable_scheme != repair_scheme, то больше 1 подсочетания.
                        if any(comb_df['dif_scheme']):
                            # Второе подсочетнание.
                            comb_df['status_repair'] = False
                            comb_df.loc[comb_df.index == n_ - 1, 'status_repair'] = True
                            self.calc_comb(rm, comb_df)
                            if n_ == 2 and self.info_srs['Контроль ДТН'] == 'ДДТН' and all(comb_df['dif_scheme']):
                                comb_df['status_repair'] = True  # Двойной ремонт.
                                self.calc_comb(rm, comb_df)
                    elif n_ == 3:
                        if len(comb_df[comb_df['dif_scheme'] == True]) == 0:
                            # Одно любое сочетание если все откл = ремонт.
                            comb_df.loc[comb_df.index < n_ - 1, 'status_repair'] = True
                            self.calc_comb(rm, comb_df)
                        elif len(comb_df[comb_df['dif_scheme'] == True]) == 1:
                            # Два сочетания с одним элементом откл != ремонт.
                            for i in range(n_):
                                if comb_df.iloc[i]['dif_scheme']:
                                    comb_df['status_repair'] = False
                                    comb_df.loc[comb_df.index != i, 'status_repair'] = True
                                    self.calc_comb(rm, comb_df)
                                    ii = 0 if i > 0 else 1
                                    comb_df['status_repair'] = False
                                    comb_df.loc[comb_df.index != ii, 'status_repair'] = True
                                    self.calc_comb(rm, comb_df)
                                    break
                        elif len(comb_df[comb_df['dif_scheme'] == True]) > 1:
                            # Два-три сочетания с откл != ремонт.
                            for i in range(n_):
                                if comb_df.iloc[i]['dif_scheme']:
                                    comb_df['status_repair'] = False
                                    comb_df.loc[comb_df.index != i, 'status_repair'] = True
                                    self.calc_comb(rm, comb_df)

        if self.task_calc['cb_disable_excel']:
            if self.srs_xl.empty:
                raise ValueError(f'Таблица отключений из xl отсутствует.')
            # self.srs_xl.fillna(0, inplace=True)
            comb_xl = self.gen_comb_xl(rm, self.srs_xl)
            for comb in comb_xl:
                self.info_srs['Контроль ДТН'] = 'ДДТН'
                if GeneralSettings.set_save['gost']:
                    if comb.shape[0] == 3 or (comb.shape[0] == 2 and rm.code_name_rg2 in [1, 2, 3]):
                        self.info_srs['Контроль ДТН'] = 'АДТН'
                    if rm.code_name_rg2 in [1, 2, 3] and (comb.shape[0] == 3 or
                                                          (comb.shape[0] == 2 and all(comb['status_repair']))):
                        log.info(f'Сочетание отклонено по ГОСТ: ')
                        log.info(tabulate(comb, headers='keys', tablefmt='psql'))
                        continue
                self.calc_comb(rm, comb)

        # Прибавить info_file к overloads_srs.
        if not self.overloads_srs.empty:
            self.overloads_srs.index = self.overloads_srs['Номер СРС'].astype(str) + '.' + \
                                       self.overloads_srs['Номер подсочетания'].astype(str) + '_' + \
                                       self.overloads_srs.index.astype(str)
            self.overloads_all = pd.concat([self.overloads_all,
                                            self.overloads_srs.apply(lambda x: rm.info_file,
                                                                     axis=1).join(self.overloads_srs)])
            self.overloads_srs.drop(self.overloads_srs.index, inplace=True)

        # Вывод таблиц К-О в excel
        if self.task_calc['cb_tab_KO'] and (len(self.control_I) or len(self.control_U)):
            name_sheet = f'{self.file_count}_{rm.info_file["Имя файла"] }'.replace('[', '').replace(']', '')[:28]
            control_df_dict = {}
            if len(self.control_I):
                # todo объединить колонки с одинаковыми контр элементами
                control_df_dict[name_sheet + '{I}'] = self.control_I
                self.control_I = None
            if len(self.control_U):
                control_df_dict[name_sheet + '{U}'] = self.control_U
                self.control_U = None
            # https://www.geeksforgeeks.org/how-to-write-pandas-dataframes-to-multiple-excel-sheets/

            mode = 'a' if os.path.exists(self.book_path) else 'w'
            with pd.ExcelWriter(path=self.book_path, mode=mode, engine="openpyxl") as writer:
                for name_sheet, df_control in control_df_dict.items():
                    # Поиск столбцов с одинаковыми dname; ДДТН, А; АДТН, А; groupid
                    # https/www.geeksforgeeks.org/how-to-find-drop-duplicate-columns-in-a-pandas-dataframe/
                    df_control_head = df_control.iloc[:4].T  # включая groupid
                    duplicated_true = df_control_head.duplicated(keep=False)
                    groupid_true = df_control.loc['-', '-', 'groupid'] > 0
                    selection_columns = duplicated_true & groupid_true  # выборка в столбцах df_control для проверки
                    if selection_columns.any():
                        dict_equals = defaultdict(list)  # {номер:[перечень индексов столбцов с одинаковыми колонками]}
                        df_control_head = df_control_head[selection_columns]
                        duplicated_unique = df_control_head.drop_duplicates()
                        for i in range(len(duplicated_unique)):
                            col_unique = duplicated_unique.iloc[i, :]
                            for ii in range(len(df_control_head)):
                                control_col = df_control_head.iloc[ii, :]
                                if col_unique.equals(control_col):
                                    dict_equals[i].append(int(control_col.name))
                    # Объединить столбцы с одинаковыми dname; ДДТН, А; АДТН, А; groupid
                    if dict_equals:
                        for cols in dict_equals.values():
                            df_control[cols[0]] = df_control[cols].max(axis=1)
                            df_control.drop(columns=cols[1:], inplace=True)

                    df_control.to_excel(excel_writer=writer,
                                        sheet_name=name_sheet,
                                        header=False,
                                        startrow=1,
                                        freeze_panes=(2, 3),
                                        index=True)

            # Форматирование таблиц Отключение - Контроль
            wb = load_workbook(self.book_path)
            for name_sheet in control_df_dict:
                ws = wb[name_sheet]
                ws['A1'] = f'{self.name_tab[0]}{self.num_tab}{self.name_tab[1]} {rm.name_rm}'
                self.num_tab += 1
                ws['A2'] = 'Наименование режима'
                ws['B2'] = 'Номер режима'
                ws['C2'] = 'Наименование параметра'
                # ws.merge_cells('A2:B4')
                thins = Side(border_style="thin", color="000000")
                max_column_lit = get_column_letter(ws.max_column)
                ws.merge_cells(f'A1:{max_column_lit}1')

                # Данные
                for row in range(3, ws.max_row + 1):
                    for col in range(4, ws.max_column + 1):
                        ws.cell(row, col).border = Border(thins, thins, thins, thins)
                        if ws.cell(row, 3).value in ['I, А', 'U, кВ']:
                            ws.cell(row, col).font = Font(bold=True)
                        if 'I, %' in ws.cell(row, 3).value:
                            if ws.cell(row, col).value >= 100:
                                ws.cell(row, col).fill = PatternFill(fill_type='solid', fgColor="00FF9900")
                        if 'U, %' in ws.cell(row, 3).value:
                            if ws.cell(row, col).value > 0:
                                ws.cell(row, col).fill = PatternFill(fill_type='solid', fgColor="00FF9900")
                # Колонки
                for litter, L in {'A': 35, 'B': 6, 'C': 17}.items():
                    ws.column_dimensions[litter].width = L
                for n in range(4, ws.max_column + 1):
                    ws[f'{get_column_letter(n)}2'].alignment = Alignment(textRotation=90, wrap_text=True,
                                                                         horizontal="center", vertical="center")
                    ws.column_dimensions[get_column_letter(n)].width = 9
                    ws[f'{get_column_letter(n)}2'].font = Font(bold=True)
                    ws[f'{get_column_letter(n)}2'].border = Border(thins, thins, thins, thins)
                # Строки
                ws.row_dimensions[5].hidden = True  # Скрыть
                ws.row_dimensions[6].hidden = True
                for n in range(1, ws.max_row + 1):
                    ws[f'A{n}'].alignment = Alignment(wrap_text=True, horizontal="left", vertical="center")
                    ws[f'B{n}'].alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
                    ws[f'C{n}'].alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
                ws.row_dimensions[2].height = 145
            wb.save(self.book_path)

        # rm.save(full_name_new=self.task_calc['folder_result_calc'] + '\\' + rm.Name)  # todo удалить? сохранение rg2

    def calc_comb(self, rm, comb: pd.DataFrame):
        """
        Смоделировать отключение элементов в комбинации.
        :param rm:
        :param comb:  'table', 'index', "dname", 'status_repair', "key"
        :return:
        """
        comb.sort_values(by='status_repair', inplace=True)
        comb['scheme_info'] = ''  # Для добавления в 'Наименование СРС' данных о disable_scheme и repair_scheme
        self.restore_rm(rm=rm)
        # Отключаем
        for i in range(len(comb)):
            if rm.sta(table=comb['table'].iloc[i], index=comb['index'].iloc[i]):  # отключаем элемент
                log.info(f'Комбинация отклонена, тк элемент {comb["dname"].iloc[i]!r} уже был отключен.')
                return False
            if comb['repair_scheme'].iloc[i]:
                pass  # todo доделать comb['scheme_info'].iloc[i] = ' ()'
            if comb['disable_scheme'].iloc[i]:
                pass  # todo доделать comb['scheme_info'].iloc[i] = ' ()'

        # Имя сочетания
        self.info_srs.drop(labels=['Отключение', 'Ключ откл.', 'Ремонт 1', 'Ключ рем.1', 'Ремонт 2', 'Ключ рем.2'],
                           inplace=True, errors='ignore')
        if comb.iloc[0]["status_repair"]:
            self.info_srs['Наименование СРС'] = 'Ремонт '
            self.info_srs['Ремонт 1'] = comb["dname"].iloc[0] + comb['scheme_info'].iloc[0]
            self.info_srs['Ключ рем.1'] = comb["key"].iloc[0]
        else:
            self.info_srs['Наименование СРС'] = 'Отключение '
            self.info_srs['Отключение'] = comb["dname"].iloc[0] + comb['scheme_info'].iloc[0]
            self.info_srs['Ключ откл.'] = comb["key"].iloc[0]

        self.info_srs['Наименование СРС'] += comb["dname"].iloc[0] + comb['scheme_info'].iloc[0]

        if len(comb) > 1:
            self.info_srs['Наименование СРС'] += ' при ремонте' if 'Откл' in self.info_srs['Наименование СРС'] else ' и'
            self.info_srs['Наименование СРС'] += f' {comb["dname"].iloc[1] + comb["scheme_info"].iloc[1]}'
            self.info_srs['Ремонт 1'] = comb["dname"].iloc[1] + comb["scheme_info"].iloc[1]
            self.info_srs['Ключ рем.1'] = comb["key"].iloc[1]
        if len(comb) == 3:
            self.info_srs['Наименование СРС'] += f', {comb["dname"].iloc[2] + comb["scheme_info"].iloc[2]}'
            self.info_srs['Ремонт 2'] = comb["dname"].iloc[2] + comb["scheme_info"].iloc[2]
            self.info_srs['Ключ рем.2'] = comb["key"].iloc[2]
        self.info_srs['Наименование СРС'] += '.'

        log.info(f"Сочетание {self.number_comb}: {self.info_srs['Наименование СРС']}")
        # log.info(f'Комбинация {self.number_comb}:\n{comb[["table","name", "status_repair"]]}')

        self.info_srs['Номер СРС'] = self.number_comb
        self.info_srs['Кол. откл. эл.'] = comb.shape[0]

        self.do_action(rm)

    def do_action(self, rm):
        """
        Цикл по действиям для ввода режима в область допустимых значений.
        :param rm:
        :return:
        """
        self.info_action = pd.Series(dtype='object')
        self.info_action['Номер подсочетания'] = 0

        # Если False - значит есть ПА, True - конец расчета сочетания (перегрузку нечем ликвидировать или отсутствует).

        # Цикл по действиям (ПА или ОП)
        while True:
            self.info_action['Конец'] = True

            self.do_control(rm)
            # прибавить info_srs и info_action к overloads_srs
            self.info_action['Номер подсочетания'] += 1
            self.info_action.drop(labels=['АРВ', 'СКРМ', 'Действие'], inplace=True, errors='ignore')

            if self.info_action['Конец']:
                break
        self.number_comb += 1  # код комбинации
        # rm.save(full_name_new=self.task_calc['folder_result_calc'] + '\\' + self.info_srs['Наименование СРС']+'.rg2')
        # todo сохранение rg2

    def do_control(self, rm):
        """
        Проверка параметров режима.
        :return:  Наполняет overloads_srs
        """
        test = rm.rgm()
        if GeneralSettings.set_save['avr']:
            self.info_action['АРВ'] = self.node_include(rm)
            if self.info_action['АРВ']:
                test = rm.rgm()
        if GeneralSettings.set_save['skrm']:
            self.info_action['СКРМ'] = rm.auto_shunt_cor(all_auto_shunt=self.auto_shunt)
            if self.info_action['СКРМ']:
                test = rm.rgm()

        if not test:
            overloads = pd.DataFrame({'dname': ['Режим не моделируется'], 'i_zag': [-1], 'otv_min': [-1]})
        else:
            overloads = pd.DataFrame()
            # проверка на наличие перегрузок ветвей (ЛЭП, трансформаторов, выключателей)
            if self.info_srs['Контроль ДТН'] == 'АДТН':
                selection_v = 'all_control & i_zag_av > 0.1'
                selection_n = 'all_control & vras<umin_av & !sta'
            else:
                selection_v = 'all_control & i_zag > 0.1'
                selection_n = 'all_control & vras<umin & !sta'

            tv = rm.rastr.tables('vetv')
            tv.SetSel(selection_v)
            if tv.count:
                overloads = rm.fd_from_table(table_name='vetv',
                                             fields='dname,'  # 'Контролируемые элементы,'
                                                    'key,'  # 'Ключ контроль,'
                                                    'txt_zag,'  # 'txt_zag,' 
                                                    'i_max,'  # 'Iрасч.(A),'
                                                    'i_dop_r,'  # 'Iддтн(A),'
                                                    'i_zag,'  # 'Iзагр.ддтн(%),'
                                                    'i_dop_r_av,'  # 'Iадтн(A),'
                                                    'i_zag_av',  # 'Iзагр.адтн(%),'
                                             setsel=selection_v)
            # проверка на наличие недопустимого снижение напряжения
            # todo округлить i_max
            tn = rm.rastr.tables('node')
            tn.SetSel(selection_n)
            if tn.count:
                overloads = pd.concat([overloads, rm.fd_from_table(table_name='node',
                                                                   fields='dname,'  # 'Контролируемые элементы,'
                                                                          'key,'  # 'Ключ контроль,'
                                                                   # 'txt_zag,'  # 'txt_zag,'
                                                                   # todo сделать что бы в txt_zag были значения узлов?
                                                                          'vras,'  # 'Uрасч.(кВ),'
                                                                          'umin,'  # 'Uмин.доп.(кВ),'
                                                                          'umin_av,'  # 'U ав.доп.(кВ),'
                                                                          'otv_min,'  
                                                                          # отклонение vras от 'Uмин.доп.' (%)
                                                                          'otv_min_av',
                                                                          # отклонение vras от 'U ав.доп.' (%)
                                                                   setsel=selection_n)])

            # проверка на наличие недопустимого повышения напряжения
            tn.SetSel('all_control & umax<vras & umax>0 & !sta')
            if tn.count:
                overloads = pd.concat([overloads, rm.fd_from_table(table_name='node',
                                                                   fields='dname,'  # 'Контролируемые элементы,'
                                                                          'key,'  # 'Ключ контроль,'
                                                                          'vras,'  # 'Uрасч.(кВ),'
                                                                          'umax,'  # 'Uнаиб.раб.(кВ)'
                                                                          'otv_max',  # 'Uнаиб.раб.(кВ)'
                                                                   setsel='all_control & umax<vras & umax>0 & !sta')])
            if self.task_calc['cb_tab_KO']:
                if len(self.control_I):
                    ci = rm.fd_from_table(table_name='vetv',
                                          fields='index,i_max,i_zag,i_zag_av',
                                          setsel="all_control")
                    ci.set_index('index', inplace=True)
                    ci['i_max'] = ci['i_max'].round(0)
                    ci['i_zag'] = ci['i_zag'].round(0)
                    ci['i_zag_av'] = ci['i_zag_av'].round(0)
                    ci = ci.T
                    ci.index = pd.MultiIndex.from_product([[self.info_srs['Наименование СРС']],
                                                           [f'{self.number_comb}.'
                                                            f'{self.info_action["Номер подсочетания"]}'],
                                                           ['I, А', 'I, % от ДДТН', 'I, % от АДТН']])
                    self.control_I = pd.concat([self.control_I, ci], axis=0)

                if len(self.control_U):
                    cu = rm.fd_from_table(table_name='node',
                                          fields='index,vras,otv_min,otv_min_av',
                                          setsel="all_control")
                    cu.set_index('index', inplace=True)
                    cu['vras'] = cu['vras'].round(1)
                    cu['otv_min'] = cu['otv_min'].round(2)
                    cu['otv_min_av'] = cu['otv_min_av'].round(2)
                    cu = cu.T
                    cu.index = pd.MultiIndex.from_product([[self.info_srs['Наименование СРС']],
                                                           [f'{self.number_comb}.'
                                                            f'{self.info_action["Номер подсочетания"]}'],
                                                           ['U, кВ', 'U, % от МДН', 'I, % от АДН']])
                    self.control_U = pd.concat([self.control_U, cu], axis=0)

        if not overloads.empty:
            overloads.index = range(len(overloads))
            self.overloads_srs = pd.concat([self.overloads_srs,
                                            overloads.apply(lambda x: pd.concat([self.info_srs, self.info_action]),
                                                            axis=1).join(other=overloads)])

    @staticmethod
    def import_id_rg2(path_file: str, txt_task):
        """
        Преобразует txt формат в ImportFromModel.
        :param path_file:
        :param txt_task:
        # таблица: параметры (обновить)
        node: otkl1, otkl2
        vetv: sta, otkl2
        """
        for row in txt_task.split('\n'):
            row = row.replace(' ', '').split('#')[0]  # удалить текст после '#'
            if ':' in row:
                name_table, name_fields = row.split(':')
                if name_table and name_fields:
                    ifm = ImportFromModel(
                        import_file_name=path_file,
                        # criterion_start={"years": "",
                        #                  "season": "",
                        #                  "max_min": "",
                        #                  "add_name": ""},
                        tables=name_table,
                        param=name_fields,
                        # sel="",
                        calc=2)  # обновить
                    ImportFromModel.set_import_model.append(ifm)

    @staticmethod
    def restore_rm(rm):
        """
        Восстановить поля РМ.
        """
        rm.rastr.tables('vetv').cols.item("sta").calc('staRes')
        rm.rastr.tables('node').cols.item("sta").calc('staRes')
        rm.rastr.tables('Generator').cols.item("sta").calc('staRes')
        # меняется после действия изменения схемы или ПА
        rm.rastr.tables('node').cols.item("pn").calc('pnRes')
        rm.rastr.tables('node').cols.item("qn").calc('qnRes')
        rm.rastr.tables('node').cols.item("pg").calc('pgRes')
        rm.rastr.tables('Generator').cols.item("P").calc('PRes')
        rm.rastr.tables('vetv').cols.item("ktr").calc('ktrRes')

    @staticmethod
    def node_include(rm) -> str:
        """
        Восстановление питания узлов путем включения выключателей (r<0.011 & x<0.011).
        :return: информация
        """
        node_info = ''
        node_all = set()
        node_include = set()
        tv = rm.rastr.tables('vetv')
        tn = rm.rastr.tables('node')
        tn.SetSel("sta&(!staRes)&(pn!=0|qn!=0|pg!=0)")
        if tn.count:
            tv.cols.item("temp").calc('ip.sta')
            tv.cols.item("temp1").calc('iq.sta')
            i = tn.FindNextSel(-1)
            while i >= 0:
                ny = tn.Cols("ny").ZS(i)
                node_all.add((ny, tn.Cols("name").ZS(i)))
                tv.SetSel(f"(ip={ny}|iq={ny}) & r<0.011 & x<0.011")
                if tv.count:
                    iv = tv.FindNextSel(-1)
                    while iv >= 0:
                        if tv.Cols("temp").Z(iv) + tv.Cols("temp1").Z(iv) < 2:  # ip.sta+iq.sta
                            tv.Cols.item("sta").SetZ(iv, False)
                            tn.Cols.item("sta").SetZ(i, False)
                            node_include.add((ny, tn.Cols("name").ZS(i)))
                        iv = tv.FindNextSel(iv)

                i = tn.FindNextSel(i)

        if node_include:
            node_info = "Восстановлено питание узлов:"
            for ny, name in node_include:
                node_info += f' {name} ({ny}),'
            node_info = node_info.strip(',') + ". "

        node_not_include = node_all - node_include
        if node_not_include:
            node_info += "Не восстановлено питание узлов:"
            for ny, name in node_not_include:
                node_info += f' {name} ({ny}),'
            node_info = node_info.strip(',') + ". "

        if node_info:
            log.info('\tnode_include: ' + node_info)
        return node_info.strip()


class EditModel(GeneralSettings):
    """
    Коррекция файлов.
    """

    def __init__(self, task):
        super(EditModel, self).__init__()
        self.print_xl = None
        self.cor_xl = None
        self.task = task
        self.rastr_files = None
        self.all_folder = False  # Не перебирать вложенные папки
        self.load_additional = []

    def run_cor(self):
        """
        Запуск корректировки моделей.
        """
        log.info('\n!!! Запуск корректировки РМ !!!\n')
        self.task["KIzFolder"] = self.task["KIzFolder"].strip()
        if "*" in self.task["KIzFolder"]:
            self.task["KIzFolder"] = self.task["KIzFolder"].replace('*', '')
            self.all_folder = True

        if not os.path.exists(self.task["KIzFolder"]):
            raise ValueError(f'Не найден путь: {self.task["KIzFolder"]}.')

        self.task['folder_result'] = self.task["KIzFolder"] + r"\result"
        if os.path.isfile(self.task["KIzFolder"]):
            self.task['folder_result'] = os.path.dirname(self.task["KIzFolder"]) + r"\result"

        self.task["KInFolder"] = self.task["KInFolder"].strip()
        # папка для сохранения result и KInFolder
        if self.task["KInFolder"] and not os.path.exists(self.task["KInFolder"]):
            if os.path.isdir(self.task["KIzFolder"]):
                log.info("Создана папка: " + self.task["KInFolder"])
                os.makedirs(self.task["KInFolder"])  # создать папку
                self.task['folder_result'] = self.task["KInFolder"] + r"\result"
            else:
                self.task['folder_result'] = os.path.dirname(self.task["KIzFolder"]) + r"\result"

        if not os.path.exists(self.task['folder_result']):
            os.mkdir(self.task['folder_result'])  # создать папку result

        self.task['name_time'] = f"{self.task['folder_result']}\\{datetime.now().strftime('%d-%m-%Y %H-%M-%S')}"

        if "import_val_XL" in self.task:
            if self.task["import_val_XL"]:  # Задать параметры узла по значениям в таблице excel (имя книги, имя листа)
                self.cor_xl = CorXL(excel_file_name=self.task["excel_cor_file"],
                                    sheets=self.task["excel_cor_sheet"])
                self.cor_xl.init_export_model()

        # Загрузить файл сечения.
        if "printXL" in self.task:
            if ((self.task["printXL"] and self.task["set_printXL"]["sechen"]['add']) or
                    (self.task["checking_parameters_rg2"] and self.task["control_rg2_task"]["section"])):
                self.load_additional.append('sch')

        if os.path.isdir(self.task["KIzFolder"]):  # корр файлы в папке

            if self.all_folder:  # с вложенными папками
                for address, dirs, files in os.walk(self.task["KIzFolder"]):
                    in_dir = address.replace(self.task["KIzFolder"], self.task["KInFolder"])
                    if not os.path.exists(in_dir):
                        os.makedirs(in_dir)
                    self.for_file_in_dir(from_dir=address, in_dir=in_dir)

            else:  # без вложенных папок
                self.for_file_in_dir(from_dir=self.task["KIzFolder"], in_dir=self.task["KInFolder"])

        elif os.path.isfile(self.task["KIzFolder"]):  # корр файл
            rm = RastrModel(full_name=self.task["KIzFolder"])
            rm.load()
            if self.load_additional:
                rm.downloading_additional_files(self.load_additional)

            self.cor_file(rm)
            if self.task["KInFolder"]:
                if os.path.isdir(self.task["KInFolder"]):
                    rm.save(self.task["KInFolder"] + '\\' + rm.Name)
                else:  # if os.path.isfile(self.task["KInFolder"]):
                    rm.save(self.task["KInFolder"])

        # для нескольких запусков через GUI
        if ImportFromModel.set_import_model:
            ImportFromModel.set_import_model = []

        if self.print_xl:
            self.print_xl.finish()

        self.the_end()
        if self.set_info['collapse']:
            self.set_info['end_info'] += f"\nВНИМАНИЕ! Развалились модели:\n[{self.set_info['collapse']}].\n"

        notepad_path = self.task['name_time'] + ' протокол коррекции файлов.log'
        shutil.copyfile(GeneralSettings.log_file, notepad_path)
        with open(self.task['name_time'] + ' задание на корректировку.yaml', 'w') as f:
            yaml.dump(data=self.task, stream=f, default_flow_style=False, sort_keys=False)
        # webbrowser.open(notepad_path)  #  Открыть блокнотом лог-файл.
        mb.showinfo("Инфо", self.set_info['end_info'])

    def for_file_in_dir(self, from_dir: str, in_dir: str):
        files = os.listdir(from_dir)  # список всех файлов в папке
        self.rastr_files = list(filter(lambda x: x.endswith('.rg2') | x.endswith('.rst'), files))

        for rastr_file in self.rastr_files:  # цикл по файлам .rg2 .rst в папке KIzFolder
            if self.task["KFilter_file"] and self.file_count == self.task["max_file_count"]:
                break  # Если включен фильтр файлов проверяем количество расчетных файлов.
            full_name = os.path.join(from_dir, rastr_file)
            full_name_new = os.path.join(in_dir, rastr_file)
            rm = RastrModel(full_name)
            # если включен фильтр файлов и имя стандартизовано
            if self.task["KFilter_file"] and rm.code_name_rg2:
                if not rm.test_name(condition=self.task["cor_criterion_start"], info='Цикл по файлам.'):
                    continue  # пропускаем если не соответствует фильтру
            rm.load()
            if self.load_additional:
                rm.downloading_additional_files(self.load_additional)

            self.cor_file(rm)
            if self.task["KInFolder"]:
                rm.save(full_name_new)

    def cor_file(self, rm):
        """Корректировать файл rm"""
        self.file_count += 1
        try:
            if self.task['cor_beginning_qt']['add']:
                log.info("\t*** Начало корректировку модели 'до импорта' ***")
                rm.cor_rm_from_txt(self.task['cor_beginning_qt']['txt'])
                log.info("\t*** Конец выполнения корректировки моделей 'до импорта' ***\n")
        except KeyError:
            pass

        if 'block_beginning' in self.task:
            if self.task['block_beginning']:
                log.info("\t***Блок начала ***")
                block_b(rm)
                log.info("\t*** Конец блока начала ***")

        # Импорт моделей
        if ImportFromModel.set_import_model:
            for im in ImportFromModel.set_import_model:
                im.import_data_in_rm(rm)

        # Задать параметры по значениям в таблице excel
        if "import_val_XL" in self.task:
            if self.task["import_val_XL"]:
                self.cor_xl.run_xl(rm)

        try:
            if self.task['cor_end_qt']['add']:
                log.info("\t*** Начало корректировку модели 'после импорта' ***")
                rm.cor_rm_from_txt(self.task['cor_end_qt']['txt'])
                log.info("\t*** Конец выполнения корректировки моделей 'после импорта' ***\n")
        except KeyError:
            pass

        if 'block_end' in self.task:
            if self.task['block_end']:
                log.info("\t*** Блок конца ***")
                block_e(rm)
                log.info("\t*** Конец блока конца ***")
        # Исправить пробелы, заменить английские буквы на русские.
        if "cor_name" in self.task:
            if self.task["cor_name"]:
                rm.txt_field_right(table_field=self.task["cor_name_task"])

        if 'checking_parameters_rg2' in self.task:
            if self.task['checking_parameters_rg2']:
                if not rm.checking_parameters_rg2(self.task['control_rg2_task']):  # расчет и контроль параметров режима
                    self.set_info['collapse'] += rm.name_base + ', '

        if 'printXL' in self.task:
            if self.task['printXL']:
                if not type(self.print_xl) == PrintXL:
                    self.print_xl = PrintXL(self.task)
                self.print_xl.add_val(rm)


class RastrModel(RastrMethod):
    """
    Для хранения параметров текущего расчетного файла.
    """

    def __init__(self, full_name: str):
        super(RastrModel, self).__init__()
        self.full_name = full_name
        self.dir = os.path.dirname(full_name)
        self.Name = os.path.basename(full_name)  # вернуть имя с расширением "2020 зим макс.rg2"
        self.name_base, self.type_file = self.Name.split('.')  # имя без расширения "2020 зим макс" # без rst или rg2
        self.pattern = GeneralSettings.set_save["шаблон " + self.type_file]
        self.code_name_rg2 = 0  # 0 не распознан, 1 зим макс 2 зим мин 3 ПЭВТ 4 лет макс 5 лет мин 6 паводок
        self.all_auto_shunt = {}
        self.temperature: float = 0
        self.rastr = None
        self.name_list = ["-", "-", "-"]
        self.additional_name_list = None
        self.season_name: str = ''
        self.god: str = ''
        self.name_rm: str = self.Name
        self.info_file = pd.Series(dtype='object')  # имя файла

        # "^(20[1-9][0-9])\s(лет\w?|зим\w?|паводок)\s?(макс|мин)?"
        match = re.search(re.compile(r"^(20[1-9][0-9])\s(лет\w*|зим\w*|паводок)\s?(макс\w*|мин\w*)?"), self.name_base)
        if match:
            if match.re.groups == 3:
                self.name_list = [match[1], match[2], match[3]]
                if not self.name_list[2]:
                    self.name_list = "-"
                if self.name_list[1] == "паводок":
                    self.code_name_rg2 = 6
                    self.season_name = "Паводок"
                if "зим" in self.name_list[1] and "макс" in self.name_list[2]:
                    self.code_name_rg2 = 1
                    self.season_name = "Зимний максимум нагрузки"
                if "зим" in self.name_list[1] and "мин" in self.name_list[2]:
                    self.code_name_rg2 = 2
                    self.season_name = "Зимний минимум нагрузки"
                if "лет" in self.name_list[1] and "макс" in self.name_list[2]:
                    self.code_name_rg2 = 4
                    self.season_name = "Летний максимум нагрузки"
                if "лет" in self.name_list[1] and "мин" in self.name_list[2]:
                    self.code_name_rg2 = 5
                    self.season_name = "Летний минимум нагрузки"

        self.god = self.name_list[0]
        if self.code_name_rg2:
            if self.code_name_rg2 in [4, 5] and ("ПЭВТ" in self.name_base):
                self.code_name_rg2 = 3
        self.name_rm = f'{self.season_name} {self.god} г.'

        # поиск в строке значения в ()
        match = re.search(re.compile(r"\((.+)\)"), self.name_base)
        if match:
            self.additional_name_list = match[1].split(";")

        if "°C" in self.name_base:
            match = re.search(re.compile(r"(-?\d+((,|\.)\d*)?)\s?°C"), self.name_base)  # -45.,14 °C
            if match:
                self.temperature = float(match[1].replace(',', '.'))  # число
                self.name_rm += f' Расчетная температура {self.temperature} °C.'

        self.info_file['Имя файла'] = self.name_base
        self.info_file['Год'] = self.god
        self.info_file['Сезон макс/мин'] = self.season_name
        self.info_file['Темп.(°C)'] = self.temperature
        if self.additional_name_list:
            for i, additional_name in enumerate(self.additional_name_list, 1):
                self.info_file['Доп. имя' + str(i)] = additional_name

    def test_name(self, condition: dict, info: str = "") -> bool:
        """
         Проверка имени файла на соответствие условию condition.
        :param condition:
        {"years":"2020,2023...2025","season": "лет,зим,паводок","max_min":"макс","add_name":"-41С;МДП:ТЭ-У"}
        :param info: для вывода в протокол
        :return: True если удовлетворяет
        """
        if not condition:
            return True
        if not (any(condition.values())):  # условие пустое
            return True
        # Проверка года
        if 'years' in condition:
            if condition['years']:
                if not int(self.god) in str_yeas_in_list(str(condition['years'])):
                    log.info(f"{info} {self.Name!r}. Год не проходит по условию: {condition['years']!r}")
                    return False
        # Проверка "зим" "лет" "паводок"
        if 'season' in condition:
            if condition['season']:
                if not self.name_list[1] in condition['season'].replace(' ', '').split(","):
                    log.info(f'{info} {self.Name!r}. Сезон не проходит по условию: {condition["season"]!r}')
                    return False
        # Проверка "макс" "мин"
        if 'max_min' in condition:
            if condition['max_min']:
                if self.name_list[2] not in condition['max_min'].replace(' ', '').split(","):
                    log.info(f'{info} {self.Name!r}. Не проходит по условию: {condition["max_min"]!r}')
                    return False
        # Проверка доп имени, например (-41С;МДП:ТЭ-У)
        if 'add_name' in condition:
            if condition['add_name']:
                if condition['add_name'].strip():
                    for us in condition['add_name'].split(";"):
                        if us not in self.additional_name_list:
                            log.debug(f'{info} {self.Name}. Не проходит по условию: {us!r}')
                            return False
        return True

    def load(self):
        """
        Загрузить модель в Rastr
        """
        if not self.rastr:
            try:
                self.rastr = win32com.client.Dispatch("Astra.Rastr")
            except Exception:
                raise Exception('Com объект Astra.Rastr не найден')

        self.rastr.Load(1, self.full_name, self.pattern)  # загрузить или перезагрузить
        log.info(f"\n\nЗагружен файл: {self.full_name}")

    def downloading_additional_files(self, load_additional: list = None):
        """
        Загрузка в Rastr дополнительных файлов из папки с РМ.
        :param load_additional: ['amt','sch','trn']
        """
        for extension in load_additional:
            files = os.listdir(self.dir)
            names = list(filter(lambda x: x.endswith('.' + extension), files))
            if len(names) > 0:
                self.rastr.Load(1, f'{self.dir}\\{names[0]}', GeneralSettings.set_save[f"шаблон {extension}"])
                log.info(f"Загружен файл: {names[0]}")
            else:
                raise ValueError(f'Файл с расширением {extension!r} не найден в папке {self.dir}')

    def save(self, full_name_new):
        self.rastr.Save(full_name_new, self.pattern)
        log.info("Файл сохранен: " + full_name_new)

    def checking_parameters_rg2(self, dict_task: dict):
        """  контроль  dict_task = {'node': True, 'vetv': True, 'Gen': True, 'section': True,
             'area': True, 'area2': True, 'darea': True, 'sel_node': "na>0"}  """
        if not self.rgm("checking_parameters_rg2"):
            return False

        node = self.rastr.tables("node")
        branch = self.rastr.tables("vetv")
        # Также проверяется наличие узлов без ветвей, ветвей без узлов начала или конца, генераторов без узлов.
        all_ny = set([x[0] for x in node.writesafearray("ny", "000")])
        all_ip = set([x[0] for x in branch.writesafearray("ip", "000")])
        all_iq = set([x[0] for x in branch.writesafearray("iq", "000")])
        all_iq_ip = all_ip.union(all_iq)

        # Узлы без ветвей.
        all_ny_not_branches = all_ny - all_iq_ip
        if all_ny_not_branches:
            log.error(f'В таблице node узлы без ветвей: {all_ny_not_branches}')
        # Ветви без узлов.
        all_ip_iq_not_node = all_iq_ip - all_ny
        if all_ip_iq_not_node:
            log.error(f'В таблице vetv есть ссылка на узлы которых нет в таблице node: {all_ip_iq_not_node}')
        # Генераторы без узлов.
        generator = self.rastr.tables("Generator")
        if generator.size:
            all_gen_ny = set([x[0] for x in generator.writesafearray("Node", "000")])
            all_gen_not_node = all_gen_ny - all_ny
            if all_gen_not_node:
                log.error(f'В таблице Generator есть ссылка на узлы которых нет в таблице node: {all_gen_not_node}')

        if dict_task["sel_node"]:
            self.add_fields_in_table(name_tables='node', fields='sel1', type_fields=3)
            node.cols.item("sel1").calc(0)
            node.setsel(dict_task["sel_node"])
            node.cols.item("sel1").calc(1)

        # Напряжения
        if dict_task["node"]:
            log.info("\tКонтроль напряжений.")
            self.voltage_nominal(choice=(dict_task["sel_node"] + '&uhom>30'))
            self.voltage_deviation(choice=dict_task["sel_node"])
            self.voltage_fine(choice=dict_task["sel_node"])
            self.voltage_error(choice=dict_task["sel_node"])

        # Токи
        if dict_task['vetv']:
            # Контроль токовой загрузки
            log.info(f"Расчет загрузки ветвей для температуры {self.temperature}.")
            self.rastr.CalcIdop(self.temperature, 0.0, "")
            if dict_task["sel_node"]:
                sel_vetv = "i_zag>=0.1&(ip.sel1|iq.sel1)"
                presence_n_it = {'n_it': "(ip.sel1|iq.sel1)&n_it>0",
                                 'n_it_av': "(ip.sel1|iq.sel1)&n_it_av>0"}
            else:
                sel_vetv = "i_zag>=0.1"
                presence_n_it = {'n_it': "n_it>0",
                                 'n_it_av': "n_it_av>0"}

            log.debug('Контроль токовой загрузки.')
            branch.setsel(sel_vetv)
            if branch.count:  # есть превышения
                j = branch.FindNextSel(-1)
                while j > -1:
                    log.info(f"\t\tВНИМАНИЕ ТОКИ! vetv:{branch.SelString(j)}, "
                             f"{branch.cols.item('name').ZS(j)} - {branch.cols.item('i_zag').ZS(j)} %")
                    j = branch.FindNextSel(j)

            log.debug('Проверка наличия n_it,n_it_av в таблице График_Iдоп_от_Т(graphikIT).')
            graph_it = self.rastr.tables("graphikIT")
            if graph_it.size:
                all_graph_it = set([x[0] for x in graph_it.writesafearray("Num", "000")])
                for field, sel_vetv_n_it in presence_n_it.items():
                    branch.setsel(sel_vetv_n_it)
                    for i in branch.writesafearray(field + ",name,ip,iq,np", "000"):
                        if i[0] > 0 and i[0] not in all_graph_it:
                            log.error(f"\t\tВНИМАНИЕ graphikIT! vetv: {i[1]} [{i[2]},{i[3]},{i[4]}] "
                                      f"{field}={i[0]} не найден в таблице График_Iдоп_от_Т")

        #  ГЕНЕРАТОРЫ
        if dict_task['Gen']:
            log.info("\tКонтроль генераторов")
            chart_pq = set([x[0] for x in self.rastr.tables("graphik2").writesafearray("Num", "000")])
            sel_gen = "!sta&Node.sel1" if dict_task["sel_node"] else "!sta"
            generator.setsel(sel_gen)
            col = {'Num': 0, 'Node': 1, 'Name': 2, 'Pmin': 3, 'Pmax': 4, 'P': 5, 'NumPQ': 6}
            if generator.count:
                for i in generator.writesafearray(','.join(col), "000"):
                    Pmin = i[col['Pmin']]
                    Pmax = i[col['Pmax']]
                    P = i[col['P']]
                    Name = i[col['Name']]
                    Num = i[col['Num']]
                    Node = i[col['Node']]
                    NumPQ = i[col['NumPQ']]
                    if P < Pmin and Pmin:
                        log.info(f"\t\tВНИМАНИЕ! ГЕНЕРАТОР: {Name}, {Num=},{Node=}, {P=} < {Pmin=}")
                    if P > Pmax and Pmax:
                        log.info(f"\t\tВНИМАНИЕ! ГЕНЕРАТОР: {Name}, {Num=},{Node=}, {P=} > {Pmax=}")
                    if NumPQ and self.rastr.tables("graphik2").size:
                        chart_pq = set([x[0] for x in self.rastr.tables("graphik2").writesafearray("Num", "000")])
                        if NumPQ not in chart_pq:
                            log.info(f"\t\tВНИМАНИЕ! ГЕНЕРАТОР: {Name}, {Num=},{Node=}, "
                                     f"{NumPQ=} не найден в таблице PQ-диаграммы (graphik2)")
        # сечения
        if dict_task['section']:
            if self.rastr.tables.Find("sechen") >= 0:
                section = self.rastr.tables("sechen")
                if section.size == 0:
                    log.error("\tCечения отсутствуют")
                else:
                    log.info("\tКонтроль сечений")
                    section.setsel("")
                    j = section.FindNextSel(-1)
                    while j != -1:
                        name = section.cols.item("name").ZS(j)
                        ns = section.cols.item("ns").ZS(j)
                        pmax = section.cols.item("pmax").Z(j)
                        psech = section.cols.item("psech").Z(j)
                        if psech > pmax + 0.01:
                            log.info(f"\t\tВНИМАНИЕ! сечение: {name} {ns!r}, P = {round(psech)}, "
                                     f"pmax = {pmax}, отклонение: {round(pmax - psech)}")
                        j = section.FindNextSel(j)
            else:
                raise ValueError("Файл сечений не загружен")

        return True

    def cor_rm_from_txt(self, task_txt: str):
        """
        Корректировать модели по заданию в текстовом формате
        :param task_txt:
        """
        task_rows = task_txt.split('\n')
        for task_row in task_rows:
            task_row = task_row.split('#')[0]  # удалить текст после '#'
            name_fun = task_row.split('[', 1)[0]  # Имя функции стоит перед "[".
            name_fun = name_fun.replace(' ', '')
            if not name_fun:
                continue  # К следующей строке.

            # Условие выполнения в фигурных скобках
            match = re.search(re.compile(r"\{(.+?)}"), task_row)
            if match:
                conditions = match[1].strip()
                if not self.conditions_test(conditions):
                    continue  # К следующей строке.

            # Параметры функции в квадратных скобках
            function_parameters = []
            match = re.search(re.compile(r"\[(.+?)]"), task_row)
            if match:
                function_parameters = match[1].split(':', maxsplit=1)
            function_parameters += ['', '']
            self.txt_task_cor(name=name_fun,
                              sel=function_parameters[0],
                              value=function_parameters[1])

    def conditions_test(self, conditions: str) -> bool:
        """
        В строке типа "years : 2026...2029& ny=1: vras>125|(not ny=1: na==2)" проверяет выполнение условий.
        :param conditions:
        :return:
        """
        conditions_s = conditions
        conditions = self.replace_links(conditions)
        conditions_list = re.split('\*|/|\^|\+|-|\(|\)==|!=|&|\||not|>|<|<=|=<|>=|=>', conditions)
        for condition in conditions_list:
            if ':' in condition:
                for key_txt in ['years', 'season', 'max_min', 'add_name']:
                    if not self.code_name_rg2:  # Если имя не стандартное, то True.
                        conditions = conditions.replace(condition, 'True')
                        continue
                    else:
                        if key_txt in condition:
                            par, value = condition.split(':')

                            if self.test_name(condition={par.replace(' ', ''): value.strip()}, info=condition):
                                conditions = conditions.replace(condition, 'True')
                            else:
                                conditions = conditions.replace(condition, 'False')
        if ':' in conditions:
            raise ValueError("Ошибка в условии: " + conditions)
        try:
            return bool(eval(conditions))
        except Exception:
            raise ValueError(f'Ошибка у условии: {conditions_s!r}.')

    def txt_task_cor(self, name: str, sel: str = '', value: str = ''):
        """
        Функция для выполнения задания в текстовом формате
        :param name: Имя функции.
        :param sel: Выборка, нр, 15145; 12,13.
        :param value: Значение, нр, name=Промплощадка: изм name; pg=qn*2+10.
        """
        name = name.lower()
        if 'уд' in name:
            self.cor(keys=sel, values='del', del_all=('*' in name), print_log=True)
        elif 'изм' in name:
            self.cor(keys=sel, values=value, print_log=True)
        elif 'импорт' in name:
            self.txt_import_rm(type_import=sel, description=value)
        elif 'снять' in name:
            self.cor(keys='(node); (vetv); (Generator)', values='sel=0', print_log=True)
        elif 'расчет' in name:
            self.rgm(txt='txt_task_cor')
        elif 'добавить' in name:
            self.table_add_row(table=sel, tasks=value)
        elif 'текст' in name:
            self.txt_field_right(table_field=sel)
        elif 'схн' in name:
            self.shn(choice=sel)
        elif 'сечение' in name:
            sel = sel.replace(' ', '')
            for i in ['ns:', 'psech:', 'выбор:', 'тип:']:
                if i not in sel:
                    raise ValueError(f'В задании "сечение": {sel!r} отсутствует ключ {i!r}')
            sel = sel.split(';')
            sd = {}
            for _ in sel:
                key, val = _.split(':')
                sd[key] = val
            self.loading_section(ns=sd['ns'], p_new=sd['psech'], type_correction=sd['тип'])
        elif 'напряжения' in name:
            self.voltage_nominal(choice=sel, edit=True)
            self.voltage_error(choice=sel, edit=True)
        elif 'скрм' in name:
            if 'скрм*' in name:
                self.all_auto_shunt = self.auto_shunt_rec(selection=sel)
            else:
                self.all_auto_shunt = self.auto_shunt_rec(selection=sel, only_auto_bsh=True)
            self.auto_shunt_cor(all_auto_shunt=self.all_auto_shunt)
        else:
            raise ValueError(f'Задание {name=} не распознано ({sel=}, {value=})')

    def txt_import_rm(self, type_import: str, description: str):
        """
        Импорт данных из РМ.
        :param type_import: Если 'папка', то переносить данные из одноименных файлов из папки,
         'файл' -из указанного файла
        :param description: "(I:\ОЭС Урала\Тюм_ЭС\!КПР ХМАО ЯНАО ТО\Модели4 - 2023\v29\без МДП pop);
                             таблица:node; тип:2; поле: pn,qn; выборка:"
        :return:
        """
        description_dict = {}
        path = description[description.find('(') + 1:description.find(')')]
        description_list = description.replace(path, '').replace(' ', '').split(';')
        dict_name = {'таблица': 'tables', 'тип': 'calc', 'поле': 'param', 'выборка': 'sel'}

        for i in description_list:
            if ':' in i:
                key, val = i.split(':')
                for x in dict_name:
                    if x in key:
                        description_dict[dict_name[x]] = val

        if type_import == 'папка':
            file_name = path + '\\' + self.Name
            if os.path.isfile(file_name):
                path = file_name
        if os.path.isfile(path):
            ifm = ImportFromModel(import_file_name=path, **description_dict)
            ifm.import_data_in_rm(rm=self)
        else:
            log.error(f'Файл для импорта не найден {path}')

    def loading_section(self, ns: int, p_new: Union[float, str], type_correction: str = 'pg'):
        """
        Изменить переток мощности в сечении номер ns до величины p_new за счет изменения нагрузки('pn') или
        генерации ('qn') в отмеченных узлах и генераторах
        :param ns: номер сечения
        :param p_new:
        :param type_correction:  'pn' изменения нагрузки или 'pg' генерации
        """
        # --------------настройки----------
        choice = 'sel&!sta'
        max_cycle = 30  # максимальное количество циклов
        accuracy = 0.05  # процент, точность задания мощности сечения, но не превышает заданную
        dr_p_zad = 0.01  # величина реакции начальная

        log.info(f'\tИзменить переток мощности в сечении {ns=}: P={p_new}, выборка: {choice}, тип: {type_correction}.')
        if self.rastr.tables.Find("sechen") == -1:
            self.downloading_additional_files(['sch'])

        index_ns = self.index_in_table('sechen', f'ns={ns}')
        if index_ns == -1:
            raise ValueError(f'сечение {ns=} отсутствует в файле сечений')
        grline = self.rastr.Tables("grline")
        sechen = self.rastr.tables('sechen')
        name_ns = sechen.cols('name').ZS(index_ns)
        if p_new in ['pmax', 'pmin']:
            p_new = sechen.cols(p_new).Z(index_ns)
        try:
            p_new = float(p_new)
        except ValueError:
            raise ValueError(f'Заданная величина перетока мощности не распознано {p_new!r}')
        if not p_new:
            p_new = 0.01
        p_current = round(self.rastr.Calc("sum", "sechen", "psech", f"ns={ns}"), 2)
        log.info(f'\tТекущий переток мощности в сечении {name_ns!r}: {p_current}.')

        self.rastr.sensiv_start("")
        grline.SetSel(f'ns={ns}')
        index_grline = grline.FindNextSel(-1)
        while not index_grline == -1:
            self.rastr.sensiv_back(4, 1., grline.Cols("ip").Z(index_grline), grline.Cols("iq").Z(index_grline), 0)
            index_grline = grline.FindNextSel(index_grline)

        self.rastr.sensiv_write("")
        self.rastr.sensiv_end()

        node = self.rastr.tables("node")
        change_p = round(p_new - p_current, 2)
        db = 0
        # 'pn'
        p_sum = 0
        dr_p_sum = 0
        # 'pg'
        node_all = {}

        if type_correction == 'pn':  # изменение нагрузки

            choice_dr_p = f"!sta & abs(dr_p) > {dr_p_zad}"  # !sta вкл
            db = abs(self.rastr.Calc("sum", "node", "dr_p", choice_dr_p + "&dr_p>0"))
            db += abs(self.rastr.Calc("sum", "node", "dr_p", choice_dr_p + "&dr_p<0"))

        elif type_correction == 'pg':

            if self.rastr.Tables("Generator").cols.Find("sel") < 0:
                log.info('В таблицу Generator добавляется отсутствующее поле sel')
                self.rastr.Tables("Generator").Cols.Add('sel', 3)

            # Доотметить узлы и генераторы которые нужно корректировать
            # отметить генераторы у отмеченных узлов
            node.SetSel("sel")
            i = node.FindNextSel(-1)
            while i >= 0:
                self.group_cor("Generator", "sel", f"Node={node.cols('ny').ZS(i)}", "1")
                i = node.FindNextSel(i)
            # отметить узлы у отмеченных генераторов
            generators = self.rastr.tables("Generator")
            generators.SetSel("sel")
            i = generators.FindNextSel(-1)
            while i >= 0:
                self.group_cor("node", "sel", f"ny={generators.cols('Node').ZS(i)}", "1")
                i = generators.FindNextSel(i)
            choice_dr_p = f"tip>1 &!sta & abs(dr_p) > {dr_p_zad}"  # tip>1 ген   !sta вкл
            db = abs(self.rastr.Calc("sum", "node", "dr_p", choice_dr_p + "&dr_p>0"))
            db += abs(self.rastr.Calc("sum", "node", "dr_p", choice_dr_p + "&dr_p<0"))

            node.SetSel(choice)
            i = node.FindNextSel(-1)

            while i >= 0:
                nd = NodeGeneration(rastr=self.rastr, i=i)
                node_all[node.cols("ny").Z(i)] = nd
                i = node.FindNextSel(i)

        if db < dr_p_zad:
            log.error("Невозможно изменить мощность по сечению (с учетом отмеченных узлов и/или генераторов)")
            return False

        for cycle in range(max_cycle):

            p_current = round(self.rastr.Calc("sum", "sechen", "psech", f"ns={ns}"), 2)
            change_p = round(p_new - p_current, 2)
            log.debug(f'\t{cycle=}, {p_current=}, {p_new=}, {change_p=} МВт ({round(abs(change_p / p_new) * 100)} %)')

            if abs(change_p / p_new) * 100 < accuracy:
                if (p_current < p_new and p_new > 0) or (p_current > p_new and p_new < 0):
                    log.info(f'\tЗаданная точность достигнута P={p_current},'
                             f' отклонение {change_p}. {cycle + 1} итераций')
                    break

            # изменение нагрузки
            if type_correction == 'pn':
                node.SetSel(choice)
                i = node.FindNextSel(-1)
                while not i == -1:
                    p_sum += node.cols("pn").Z(i)
                    dr_p_sum += node.cols("dr_p").Z(i)
                    i = node.FindNextSel(i)
                if not p_sum:
                    log.error('Изменение мощности сечения: сумма нагрузки узлов равна 0')
                    break
                if dr_p_sum < 0:
                    coefficient = 1 + (1 - (p_sum - change_p) / p_sum)
                else:
                    coefficient = (p_sum - change_p) / p_sum
                node.cols("pn").Calc(f"pn*({coefficient})")
                node.cols("qn").Calc(f"qn*({coefficient})")

            # изменение генерации
            elif type_correction == 'pg':
                NodeGeneration.change_p = change_p
                section_up_sum = 0
                section_down_sum = 0
                for nd in node_all:
                    if nd.use:
                        nd.reserve_p()
                        if change_p * nd.dr_p > 0:
                            if nd.reserve_p_up:
                                section_up_sum += nd.reserve_p_up
                                nd.up_pgen = True
                        elif change_p * nd.dr_p < 0:
                            if nd.reserve_p_down:
                                section_down_sum += nd.reserve_p_down
                                nd.up_pgen = False
                log.debug(f'')
                if not (section_up_sum and section_down_sum):
                    log.error(f'Не удалось добиться заданной точности в сечении')

                # на сколько МВт нужно снизить Р
                reduce_p = abs(section_down_sum / (section_down_sum + section_up_sum) * change_p)
                if section_down_sum < reduce_p:
                    reduce_p = section_down_sum
                # на сколько МВт нужно увеличить Р
                increase_p = abs(abs(change_p) - reduce_p)
                if section_up_sum < increase_p:
                    increase_p = section_up_sum

                if (section_down_sum + section_up_sum) < change_p:
                    log.info("Генерации не хватает")
                # Коэффициент на сколько нужно умножить резерв Рген и прибавить к резерву Рген, для снижения генерации
                koef_p_down = 0
                # Коэффициент: на сколько нужно умножить резерв Рген и прибавить его к резерву Рген,
                # для увеличения генерации
                koef_p_up = 0
                if section_down_sum:
                    koef_p_down = 1 - (section_down_sum - reduce_p) / section_down_sum
                if section_up_sum:
                    koef_p_up = 1 - (section_up_sum - increase_p) / section_up_sum

                for nd in node_all:
                    if nd.use:
                        nd.change(koef_p_down=koef_p_down, koef_p_up=koef_p_up)

            self.rgm()
        else:
            log.info(f'Заданная точность не достигнута P={p_current}, отклонение {change_p}.')


class NodeGeneration:
    """Класс для хранения информации об узле для изменения мощности в сечении."""
    dr_p_koeff = 0  # если 1, то умножаем дополнительно на dr_p в этом случае больше загружаются
    # генераторы которые меньше влияют на изменение мощности в сечении

    no_pmin = True  # ' не учитывать Pmin
    abs_change_p = None  # todo что это?
    change_p = 0
    unbalance_p = 0

    def __init__(self, i: int, rastr):
        """
        :param i: Индекс в таблице узлы
        :param rastr: 
        """
        self.gen_available = False  # Узел с генераторами
        self.use = True
        self.up_pgen = True
        self.reserve_p_up = 0
        self.reserve_p_down = 0
        self.rastr = rastr
        self.i = i
        self.node_t = self.rastr.tables("node")
        gen_t = self.rastr.tables("Generator")
        self.ny = self.node_t.Cols("ny").Z(self.i)
        self.dr_p = self.node_t.Cols("dr_p").Z(self.i)
        self.gen_all = {}
        dr_p = self.node_t.Cols("dr_p").Z(self.i)
        self.name = self.node_t.Cols("name").ZS(self.i)
        txt = f'\t\tУзел {self.ny}: {self.name}'
        gen_t.SetSel(f"Node={self.ny}")
        if gen_t.count:
            gen_t.SetSel(f"Node={self.ny}&sel")  # все генераторы дб отмечены, если не отмечен то не используем
            i = gen_t.FindNextSel(-1)
            while i >= 0:  # ЦИКЛ ген
                self.gen_available = True  # узел с генераторами
                gen = Gen(rastr=self.rastr, i=i)
                self.gen_all[gen.Num] = gen
                i = gen_t.FindNextSel(i)
        else:
            self.pg_max = self.node_t.Cols("pg_max").Z(self.i)
            self.pg_min = self.node_t.Cols("pg_min").Z(self.i)

    def reserve_p(self):
        self.reserve_p_up = 0
        self.reserve_p_down = 0
        if self.gen_available:
            for gen in self.gen_all:
                if gen.use:
                    gen.reserve_p()
                    self.reserve_p_up += gen.reserve_p_up
                    self.reserve_p_down += gen.reserve_p_down
        else:
            if self.pg_max:
                self.reserve_p_up = self.pg_max - self.node_t.Cols("pg").Z(self.i)
            else:
                log.info(f"в узле {self.ny} {self.name} не задано поле pg_max")
            self.reserve_p_down = self.node_t.Cols("pg").Z(self.i)

    def change(self, koef_p_down: float = 0, koef_p_up: float = 0):
        # --------------настройки----------
        change_p = abs(NodeGeneration.abs_change_p)
        unbalance_p = NodeGeneration.unbalance_p
        pg_node = self.node_t.Cols("pg").Z(self.i)

        if self.up_pgen:
            deviation_pg = koef_p_up * self.reserve_p_up  # На сколько нужно изменить генерацию в узле
        else:
            deviation_pg = pg_node * koef_p_down

        if not deviation_pg:
            return False

        if unbalance_p > 0:
            if unbalance_p > deviation_pg:
                unbalance_p = unbalance_p - deviation_pg
                deviation_pg = 0
            if unbalance_p < deviation_pg:
                deviation_pg = deviation_pg - unbalance_p
                unbalance_p = 0

        if not self.gen_available:  # нет генераторов
            if self.up_pgen:  # увеличиваем генерацию узла, koef_p_up
                if self.pg_max and self.pg_max > pg_node:
                    if self.pg_min > pg_node + deviation_pg:  # (от 0 до pg_min)
                        if self.pg_min and not self.no_pmin:  # если есть Рмин и учитываем Рмин то
                            if change_p > self.pg_min:
                                self.node_t.cols.Item("pg").SetZ(self.i, self.pg_min)
                                # unbalance_p = unbalance_p + (self.pg_min - deviation_pg)
                                change_p = change_p - self.pg_min
                        else:  # нет Рмин или не учитываем Рмин
                            self.node_t.cols.Item("pg").SetZ(self.i, pg_node + deviation_pg)
                            change_p = change_p - deviation_pg
                    elif self.pg_max > pg_node + deviation_pg and (
                            self.pg_min < pg_node + deviation_pg or self.pg_min == pg_node + deviation_pg):
                        # (от pg_min (включительно) до pg_max)v
                        self.node_t.cols.Item("pg").SetZ(self.i, pg_node + deviation_pg)
                        change_p = change_p - deviation_pg
                    elif self.pg_max < pg_node + deviation_pg or self.pg_max == pg_node + deviation_pg:
                        # (больше или равно pg_max)
                        self.node_t.cols.Item("pg").SetZ(self.i, self.pg_max)
                        change_p = change_p - (self.pg_max - pg_node)

            else:  # снижаем генерацию узла,KefPG_Down
                if self.pg_min < pg_node - deviation_pg or self.pg_min == pg_node - deviation_pg:
                    # (от pg_min (включительно) до pg_node)
                    self.node_t.cols.Item("pg").SetZ(self.i, pg_node - deviation_pg)
                    change_p = change_p - deviation_pg

                elif self.pg_min > pg_node - deviation_pg and (pg_node - deviation_pg) > 0:  # (от 0 до pg_min)
                    if self.pg_min > 0 and not self.no_pmin:  # если есть Рмин и учитываем Рмин то
                        self.node_t.cols.Item("pg").SetZ(self.i, self.pg_min)
                        change_p = change_p - (pg_node - self.pg_min)
                        deviation_pg = deviation_pg - (pg_node - self.pg_min)
                        if change_p > self.pg_min:
                            self.node_t.cols.Item("sta").SetZ(self.i, True)
                            # unbalance_p = unbalance_p + (self.pg_min - deviation_pg)
                            change_p = change_p - self.pg_min

                    else:  # если Рмин не учитываем
                        self.node_t.cols.Item("pg").SetZ(self.i, pg_node - deviation_pg)
                        change_p = change_p - deviation_pg

                elif pg_node - deviation_pg < 0 or pg_node == deviation_pg:  # (меньше или равно 0)
                    self.node_t.cols.Item("pg").SetZ(self.i, 0)
                    change_p = change_p - pg_node


class Gen:
    """Класс для хранения информации о генераторах в узле для изменения мощности в сечении."""

    def __init__(self, i: int, rastr):
        self.reserve_p_up = 0
        self.reserve_p_down = 0
        self.use = True
        self.rastr = rastr
        self.i = i
        self.gen_t = self.rastr.tables("Generator")
        self.Num = self.gen_t.Cols("Num").Z(self.i)
        self.gen_name = self.gen_t.Cols("Name").Z(self.i)
        self.Pmax = self.gen_t.Cols("Pmax").Z(self.i)
        if not self.Pmax:
            log.debug(f"У генератора {self.Num!r}: {self.gen_name!r}  не задано Pmax")
        self.Pmin = self.gen_t.Cols("Pmin").Z(self.i)

    def reserve_p(self):
        if self.gen_t.Cols("sta").Z(self.i):
            self.reserve_p_up = self.Pmax
        else:  # если генератор включен
            self.reserve_p_down = self.gen_t.Cols("P").Z(self.i)
            if self.Pmax:
                self.reserve_p_up = self.Pmax - self.gen_t.Cols("P").Z(self.i)


class CorSheet:
    """
    Клас лист для хранения листов книги excel и работы с ними.
    """
    SHAPE = {"Параметры импорта из файлов RastrWin": 'import_model',  # Импорт из моделей(ИМ)
             'Выполнить изменение модели по строкам': 'list_cor',  # Строковая форма(СФ)
             'Имя таблицы:': 'table_import'}  # Импорт таблиц(ИТ)

    def __init__(self, name: str, obj):
        """
        :param name: Имя листа
        :param obj: Объект лист
        """
        #  Сводная таблица корректировок 'tab_cor', нр корр потребления или нагрузка узлов / имя файла.
        #  Таблица корректировок по списку 'list_cor', нр изм, удалить, снять отметку.
        # 'import_model' - Импорт моделей.
        # 'table_import' - Импорт таблиц(ИТ).
        self.name = name
        self.xls = obj  # xls.cell(1,1).value [строки][столбцы] xls.max_row xls.max_column
        self.import_model_all = []  # Для хранения объектов ImportFromModel
        self.calc_val = self.xls.cell(1, 1).value
        if isinstance(self.calc_val, int):
            self.type = 'tab_cor'
        else:
            try:
                self.type = self.SHAPE[self.calc_val]
            except KeyError:
                raise ValueError(f'Тип задания листа {name!r} не распознан.')

    def run_sheets(self, rm: RastrModel):
        log.info(f'\tВыполнение задания листа {self.name!r}')
        if self.type == 'import_model':
            self.import_model(rm)
        elif self.type == 'list_cor':
            self.list_cor(rm)
        elif self.type == 'tab_cor':
            self.tab_cor(rm)
        elif self.type == 'table_import':
            self.tab_import(rm)
        log.info(f'\tКонец выполнения задания листа {self.name!r}\n')

    def export_model(self):
        """"Экспорт из моделей"""
        for row in range(3, self.xls.max_row + 1):
            if self.xls.cell(row, 1).value and '#' not in self.xls.cell(row, 1).value:
                """ ИД для импорта из модели(выполняется после блока начала)"""
                ifm = ImportFromModel(import_file_name=self.xls.cell(row, 1).value,
                                      criterion_start={"years": self.xls.cell(row, 6).value,
                                                       "season": self.xls.cell(row, 7).value,
                                                       "max_min": self.xls.cell(row, 8).value,
                                                       "add_name": self.xls.cell(row, 9).value},
                                      tables=self.xls.cell(row, 2).value,
                                      param=self.xls.cell(row, 4).value,
                                      sel=self.xls.cell(row, 3).value,
                                      calc=self.xls.cell(row, 5).value)
                self.import_model_all.append(ifm)

    def tab_import(self, rm: RastrModel) -> None:
        """Импорт таблицы из XL в Rastr"""
        tables_name = self.xls.cell(2, 1).value
        type_import = self.xls.cell(4, 1).value
        field_import = []
        field_column = []
        for column in range(2, self.xls.max_column+1):
            if self.xls.cell(1, column).value:
                field_column.append(column)
                field_import.append(self.xls.cell(1, column).value)
        field_import = ','.join(field_import)
        data = [[self.xls.cell(row, col).value for col in field_column] for row in range(2, self.xls.max_row+1)]
        table = rm.rastr.Tables(tables_name)
        table.ReadSafeArray(type_import, field_import, data)

    def import_model(self, rm: RastrModel) -> None:
        """Импорт в модели"""
        if self.import_model_all:
            for im in self.import_model_all:
                im.import_data_in_rm(rm)

    def list_cor(self, rm: RastrModel) -> None:
        """
        Таблица корректировок по списку, нр изм, удалить, снять отметку.
        """
        # номера столбцов
        C_SELECTION = 2
        C_VALUE = 3

        for row in range(3, self.xls.max_row + 1):
            name_fun = self.xls.cell(row, 1).value
            if name_fun:
                if '#' not in name_fun:
                    sel = str(self.xls.cell(row, C_SELECTION).value)
                    value = self.xls.cell(row, C_VALUE).value
                    year = self.xls.cell(row, 4).value
                    season = self.xls.cell(row, 5).value
                    max_min = self.xls.cell(row, 6).value
                    add_name = self.xls.cell(row, 7).value
                    statement = self.xls.cell(row, 8).value

                    if any([year, season, max_min, add_name]) and rm.code_name_rg2:  # any если хотя бы один истина
                        if not rm.test_name(condition={"years": year, "season": season,
                                                       "max_min": max_min, "add_name": add_name},
                                            info=f'\t\tcor_x:{sel=}, {value=}'):
                            continue
                    if statement:
                        if not rm.conditions_test(statement):
                            continue
                    rm.txt_task_cor(name=name_fun, sel=sel, value=value)

    def tab_cor(self, rm: RastrModel) -> None:
        """
        Корректировка моделей по заданию в табличном виде
        """
        name_files = ""
        dict_param_column = {}  # {10: "pn"}-столбец: параметр
        # Шаг по колонкам и запись в словарь всех столбцов для коррекции
        for column_name_file in range(2, self.xls.max_column + 1):
            if self.xls.cell(1, column_name_file).value not in ["", None]:
                name_files = self.xls.cell(1, column_name_file).value.split("|")  # list [name_file, name_file]
            if self.xls.cell(2, column_name_file).value:
                duct_add = False
                for name_file in name_files:
                    if name_file in [rm.name_base, "*"]:
                        duct_add = True
                    if "*" in name_file and len(name_file) > 7:
                        pattern_name = re.compile(r"\[(.*)\].*\[(.*)\].*\[(.*)\].*\[(.*)\]")
                        match = re.search(pattern_name, name_file)
                        if match.re.groups == 4 and rm.code_name_rg2:
                            if rm.test_name(condition={"years": match[1], "season": match[2],
                                                       "max_min": match[3], "add_name": match[4]},
                                            info=f"\tcor_xl, условие: {name_file}, ") or not rm.code_name_rg2:
                                duct_add = True
                if duct_add:
                    _ = self.xls.cell(2, column_name_file).value
                    dict_param_column[column_name_file] = _.replace(' ', '')

        if not dict_param_column:
            log.info(f"\t {rm.name_base} НЕ НАЙДЕН на листе {self.name} книги excel")
        else:
            log.info(f'\t\tРасчетной модели соответствуют столбцы: параметры {dict_param_column}')
            calc_vals = {1: "ЗАМЕНИТЬ", 2: "+", 3: "-", 0: "*"}
            # 1: "ЗАМЕНИТЬ", 2: "ПРИБАВИТЬ", 3: "ВЫЧЕСТЬ", 0: "УМНОЖИТЬ"
            for row in range(3, self.xls.max_row + 1):
                for column, param in dict_param_column.items():
                    short_key = self.xls.cell(row, 1).value
                    if short_key not in [None, ""]:
                        new_val = self.xls.cell(row, column).value
                        if new_val is not None:
                            if param not in ["pop", "pp"]:
                                if self.calc_val == 1:
                                    rm.cor(keys=str(short_key),
                                           values=f"{param}={new_val}",
                                           print_log=True)
                                else:
                                    rm.cor(keys=str(short_key),
                                           values=f"{param}={param}{calc_vals[self.calc_val]}{new_val}",
                                           print_log=True)
                            else:
                                rm.cor_pop(zone=short_key, new_pop=new_val)  # изменить потребление


class CorXL:
    """
    Изменить параметры модели по заданию в таблице excel.
    """
    def __init__(self, excel_file_name: str, sheets: str) -> None:
        """
        Проверить наличие книги и листов, создать классы CorSheet для листов.
        :param excel_file_name: Полное имя файла excel;
        :param sheets: имя листов, нр [импорт из моделей][XL->RastrWin], если '*', то все листы по порядку
        """
        self.sheets_list = []  # для хранения объектов CorSheet
        log.info(f"Изменить модели по заданию из книги: {excel_file_name}, листы: {sheets}")
        if not os.path.exists(excel_file_name):
            raise ValueError("Ошибка в задании, не найден файл: " + excel_file_name)
        else:
            self.excel_file_name = excel_file_name
            # data_only - Загружать расчетные значения ячеек, иначе будут формулы.
            self.wb = load_workbook(excel_file_name, data_only=True)

            if sheets == '*':  # все листы
                self.sheets = self.wb.sheetnames
                for sheet in self.sheets:
                    if '#' not in sheet:  # все листы
                        self.sheets_list.append(CorSheet(name=sheet, obj=self.wb[sheet]))
            else:
                self.sheets = re.findall(r"\[(.+?)\]", sheets)
                for sheet in self.sheets:
                    if sheet not in self.wb.sheetnames:
                        raise ValueError(f"Ошибка в задании, не найден лист: {sheet} в файле {excel_file_name}")
                    else:
                        self.sheets_list.append(CorSheet(name=sheet, obj=self.wb[sheet]))

    def init_export_model(self) -> None:
        """Экспорт данных из растр в csv"""
        for sheet in self.sheets_list:
            if sheet.type == 'import_model':
                sheet.export_model()

    def run_xl(self, rm: RastrModel) -> None:
        """Запуск всех корректировок"""
        for sheet in self.sheets_list:
            sheet.run_sheets(rm)


class ImportFromModel:
    # __slots__ = 'set_import_model', 'calc_str'
    set_import_model = []  # хранение объектов класса ImportFromModel созданных в GUI и коде
    calc_str = {"обновить": 2, "загрузить": 1, "присоединить": 0, "присоединить-обновить": 3, "объединить": 3}
    number = 0  # для создания уникального имени csv файла

    def __init__(self,
                 import_file_name: str,
                 criterion_start: Union[dict, None] = None,
                 tables: str = '',
                 param='',
                 sel: Union[str, None] = '',
                 calc: Union[int, str] = '2',
                 way='array'):
        """
        Импорт данных из файлов '.rg2', '.rst' и др.
        :param import_file_name: полное имя файла
        :param criterion_start: {"years": "","season": "","max_min": "", "add_name": ""} условие выполнения
        :param tables: таблица для импорта, нр "node, vetv"
        :param param: параметры для импорта: "" все параметры или перечисление, нр 'sel, sta'(ключи необязательно)
        :param sel: выборка нр "sel" или "" - все
        :param calc: число типа int, строка или ключевое слово:
        {"обновить": 2 , "загрузить": 1, "присоединить": 0, "присоединить-обновить": 3}
        :param way: 'csv' или 'array'
        'csv' - Создает папку temp в папке с файлом и сохраняет в ней .csv файлы
        """
        self.way = way
        folder_temp = ''
        if not os.path.exists(import_file_name):
            raise ValueError("Ошибка в задании, не найден файл: " + import_file_name)
        else:
            log.info(f'Экспорт данных из файла "{import_file_name}".')
            self.import_file_name = import_file_name

            if way == 'csv':
                ImportFromModel.number += 1
                folder_temp = os.path.dirname(import_file_name) + '\\temp'
                if not os.path.exists(folder_temp):
                    log.debug(f'Создана папка {folder_temp}.')
                    os.mkdir(folder_temp)

            self.criterion_start = criterion_start
            self.sel = sel if sel else ''

            if type(calc) == int:
                self.calc = calc
            elif calc.isdigit():
                self.calc = int(calc)
            else:
                if calc in self.calc_str:
                    self.calc = self.calc_str[calc]
                else:
                    raise ValueError(f"ImportFromModel. Ошибка в задании, не распознано задание '{calc=}'.")

            self.param = []
            self.import_data = []  # if way == 'array': tuple(данных)
            self.import_csv_file = []  # if way == 'csv': полный путь к CSV
            self.tables = tables.replace(' ', '').split(",")  # разделить на ["таблицы"]

            import_rm = RastrModel(full_name=self.import_file_name)
            import_rm.load()

            for i, tabl in enumerate(self.tables):
                # Параметры
                if param:  # Добавить к строке параметров ключи текущей таблицы
                    self.param.append(param + ',' + import_rm.rastr.Tables(tabl).Key)
                else:  # если все параметры
                    self.param.append(import_rm.all_cols(tabl))

                log.info(f"\tТаблица: {tabl}, выборка: {self.sel}, параметры: {self.param[i]!r}.")
                tab = import_rm.rastr.Tables(tabl)
                tab.setsel(self.sel)
                if tab.count:
                    # Данные
                    if way == 'csv':
                        self.import_csv_file.append(f"{folder_temp}\\{os.path.basename(import_file_name)}_{tabl}_"
                                                    f"{ImportFromModel.number}.csv")
                        # Экспорт данных из файла в .csv файлы в папку temp
                        log.info(f"\tФайл CSV: {self.import_csv_file[i]!r}.")
                        tab.WriteCSV(1, self.import_csv_file[i], self.param[i], ";")  # 0 дописать, 1 заменить
                    elif way == 'array':
                        self.import_data.append(tab.writesafearray(self.param[i], "000"))
                
    def import_data_in_rm(self, rm: RastrModel) -> None:
        """
        Импорт данных в файлы
        """
        log.info(f"\tИмпорт из файла {self.import_file_name} в РМ.")
        if not rm.code_name_rg2 or rm.test_name(condition=self.criterion_start,
                                                info='\tImportFromModel '):
            for i, tab in enumerate(self.tables):
                log.info(f"\tТаблица: {self.tables[i]}, выборка: {self.sel}, тип: {self.calc}, "
                         f"параметры: {self.param[i]}.")
                rm_tab = rm.rastr.Tables(self.tables[i])

                if self.way == 'csv':
                    log.info(f"\tФайл CSV: {self.import_csv_file[i]}")
                    rm_tab.ReadCSV(self.calc, self.import_csv_file[i], self.param[i], ";", '')
                elif self.way == 'array':
                    rm_tab.ReadSafeArray(self.calc, self.param[i], self.import_data[i])
        ImportFromModel.number = 0


class PrintXL:
    """Класс печать данных в excel"""
    list_name_z = []
    short_name_tables = {'n': 'node',
                         'v': 'vetv',
                         'g': 'Generator',
                         'na': 'area',
                         'npa': 'area2',
                         'no': 'darea',
                         'nga': 'ngroup',
                         'ns': 'sechen'}

    #  ...._log  лист протокол для сводной

    def __init__(self, task):
        """
        Добавить листы и первая строка с названиями
        """
        self.sheet_couple = {}  # имя листа_log: имя листа_сводная
        self.name_xl_file = ''  # Имя файла EXCEL для сохранения
        self.wbook = None
        self.sheets = {}  # Для хранения ссылок на листы excel {'имя листа': ссылка}
        self.task = task
        self.list_name = ["name_rg2", "год", "лет/зим", "макс/мин", "доп_имя1", "доп_имя2", "доп_имя3"]
        self.book = Workbook()
        #  Создать лист xl и присвоить ссылку на него
        for name_table in self.task['set_printXL']:
            if self.task['set_printXL'][name_table]['add']:
                self.sheets[name_table] = self.book.create_sheet(name_table + "_log")
                # Записать первую строку параметров.
                header_list = self.list_name + self.task['set_printXL'][name_table]['par'].split(',')
                self.sheets[name_table].append(header_list)

        if self.task['print_parameters']['add']:
            self.sheets['parameters'] = self.book.create_sheet('Значения')

        if self.task['print_balance_q']['add']:
            self.sheet_q = self.book.create_sheet("balance_Q")
            self.row_q = {}
            # (имя ключа, название в ячейке XL, комментарий ячейки)
            name_row = (
                ('row_name',
                 'Наименование', ''),
                ('row_qn',
                 'Реактивная мощность нагрузки', 'Calc(sum,area,qn,vibor)'),
                ('row_dq_sum',
                 'Нагрузочные потери', ''),
                ('row_dq_line',
                 'в т.ч. потери в ЛЭП', 'потери в ЛЭП: \nCalc(sum,area,dq_line,vibor)'),
                ('row_dq_tran',
                 'потери в трансформаторах', 'Calc(sum,area,dq_tran,vibor)'),
                ('row_shq_tran',
                 'потери Х.Х. в трансформаторах', 'Calc(sum,area,shq_tran,vibor)'),
                ('row_skrm_potr',
                 'Потребление реактивной мощности СКРМ (ШР, УШР, СК, СТК)',
                 'Calc(sum,node,qsh,qsh>0 & vibor) - Calc(sum,node,qg,qg<0&pg<0.1&pg>-0.1 & vibor)'),
                ('row_sum_port_Q',
                 'Суммарное потребление реактивной мощности', ''),
                ('row_qg',
                 'Генерация реактивной мощности электростанциями', 'Calc(sum,node,qg,(pg>0.1|pg<-0.1) & vibor)'),
                ('row_skrm_gen',
                 'Генерация реактивной мощности СКРМ (БСК, СК, СТК)', ''),
                ('row_qg_min',
                 'Минимальная генерация реактивной мощности электростанциями', 'Calc(sum,node,qmin,pg>0.1& vibor)'),
                ('row_qg_max',
                 'Максимальная генерация реактивной мощности электростанциями', 'Calc(sum,node,qmax,pg>0.1& vibor)'),
                ('row_shq_line',
                 'Зарядная мощность ЛЭП', 'Calc(sum,area,shq_line, vibor)'),
                ('row_sum_QG',
                 'Суммарная генерация реактивной мощности', ''),
                ('row_Q_itog',
                 'Внешний переток реактивной мощности (избыток/дефицит +/-)', ''),
                ('row_Q_itog_gmin',
                 'Внешний переток реактивной мощности при минимальной генерации '
                 'реактивной мощности электростанциями и КУ(избыток/дефицит +/-)', ''),
                ('row_Q_itog_gmax',
                 'Внешний переток реактивной мощности при максимальной генерации '
                 'реактивной мощности электростанциями и КУ(избыток/дефицит +/-)',
                 ''),
            )
            self.sheet_q.cell(1, 1, 'Таблица 1 - Баланс реактивной мощности, Мвар')
            for n, row_info in enumerate(name_row, 2):
                self.row_q[row_info[0]] = n
                self.sheet_q.cell(n, 1, row_info[1])
                if row_info[2]:
                    self.sheet_q.cell(n, 1).comment = Comment(row_info[2], '')

    def add_val(self, rm: RastrModel):

        log.info("\tВывод данных из моделей в XL")
        if not rm.code_name_rg2 or not rm.additional_name_list:
            dop_name_list = ['-'] * 3
        else:
            dop_name_list = rm.additional_name_list[:3]
            if len(dop_name_list) < 3:
                dop_name_list += ['-'] * (3 - len(dop_name_list))

        self.list_name_z = [rm.name_base, rm.god, rm.name_list[1], rm.name_list[2]] + dop_name_list

        self.add_val_table(rm)

        if self.task['print_parameters']['add']:
            self.add_val_parameters(rm.rastr, sel=self.task['print_parameters']['sel'])

        if self.task['print_balance_q']['add']:
            self.add_val_balance_q(rm)

    def add_val_table(self, rm):

        for key in self.task['set_printXL']:
            if not self.task['set_printXL'][key]['add']:
                continue
            # проверка наличия таблицы
            tabl_name = self.task['set_printXL'][key]['tabl']
            if rm.rastr.Tables.Find(tabl_name) < 0:
                raise ValueError("В RastrWin не загружена таблица: " + self.task['set_printXL'][key]['tabl'])

            # принт данных из растр в таблицу для СВОДНОЙ
            r_table = rm.rastr.tables(self.task['set_printXL'][key]['tabl'])
            param = self.task['set_printXL'][key]['par'].replace(' ', '')
            if not param:
                param = rm.all_cols(tabl_name)

            param_list = param.split(',')
            param_list = [param_list[i] if r_table.cols.Find(param_list[i]) > -1 else '-' for i in
                          range(len(param_list))]

            setsel = self.task['set_printXL'][key]['sel'] if self.task['set_printXL'][key]['sel'] else ""
            r_table.setsel(setsel)
            index = r_table.FindNextSel(-1)
            while index >= 0:
                self.sheets[key].append(
                    self.list_name_z + [r_table.cols.item(val).ZN(index) if val != '-' else '-' for val in param_list])
                index = r_table.FindNextSel(index)

    def add_val_parameters(self, rastr, sel):
        """
        Вывод заданных параметров в формате: "v=15105,15113,0;15038,15037,4|r;x;b / n=15198|pg;qg".
        Таблица: n-node,v-vetv,g-Generator,na-area,npa-area2,no-darea,nga-ngroup,ns-sechen.
        """
        sheet = self.sheets['parameters']
        one_row_list = None
        if sheet.max_row == 1:
            one_row_list = self.list_name[:]
        val_list = self.list_name_z[:]
        for task_i in sel.replace(' ', '').split('/'):
            key_row, key_column = task_i.split("|")  # нр"ny=8;9", "pn;qn"
            key_column = key_column.split(';')  # ['pn','qn']
            key_row = key_row.split('=')  # ['n','8|9']
            set_key_row = key_row[1].split(';')  # ['8','9']
            try:
                tabl_key = self.short_name_tables[key_row[0]]
            except KeyError:
                tabl_key = key_row[0]

            if rastr.Tables.Find(tabl_key) < 0:
                raise ValueError("print_parameters, в Rastrwin не найдена таблица: " + key_row[0])

            t_print = rastr.Tables(tabl_key)

            for key_i in set_key_row:
                choice = ''
                if key_i.count(","):
                    if t_print.Key.count(',') == key_i.count(","):
                        fields = t_print.Key.split(",")
                        values = key_i.split(",")
                        for n, field in enumerate(fields, 1):
                            choice += field + '=' + values[n - 1] + '&'
                        choice = choice.rstrip('&')
                else:
                    choice = t_print.Key + '=' + key_i

                if not choice:
                    raise ValueError(f"print_parameters: Ошибка формата задания: {key_i}")

                name_tek = ""
                if t_print.cols.Find("name") > 0:
                    name_tek = "name"
                elif t_print.cols.Find("Name") > 0:
                    name_tek = "name"

                t_print.setsel(choice)
                ndx = t_print.FindNextSel(-1)

                for key_column_i in key_column:
                    if ndx > -1:
                        if sheet.max_row == 1:
                            name_add = f'[{t_print.cols.item(name_tek).ZS(ndx)}]' if name_tek else ''
                            one_row_list.append(f'{choice}[{key_column_i}]{name_add}')
                        try:
                            val_list.append(t_print.cols.item(key_column_i).ZN(ndx))
                        except Exception:
                            raise ValueError(f'В таблице {tabl_key!r} отсутствует поле {key_column_i!r}')
                    else:
                        if sheet.max_row == 1:
                            one_row_list.append(f"не найден:  {key_i}, {key_column_i}")
                        val_list.append("не найдено")

        if sheet.max_row == 1:
            sheet.append(one_row_list)
        sheet.append(val_list)

    def add_val_balance_q(self, rm):
        column = self.sheet_q.max_column + 1
        choice = self.task["print_balance_q"]["sel"]
        self.sheet_q.cell(2, column, rm.name_base)
        area = rm.rastr.Tables("area")
        area.SetSel(self.task["print_balance_q"]["sel"])
        # ndx = area.FindNextSel(-1)

        # Нагрузка Q
        address_qn = self.sheet_q.cell(self.row_q['row_qn'], column,
                                       rm.rastr.Calc("sum", "area", "qn", choice)).coordinate
        # Потери Q в ЛЭП
        address_dq_line = self.sheet_q.cell(self.row_q['row_dq_line'], column,
                                            rm.rastr.Calc("sum", "area", "dq_line", choice)).coordinate
        # Потери Q в Трансформаторах
        address_dq_tran = self.sheet_q.cell(self.row_q['row_dq_tran'], column,
                                            rm.rastr.Calc("sum", "area", "dq_tran", choice)).coordinate
        # Потери Q_ХХ в Трансформаторах
        address_shq_tran = self.sheet_q.cell(self.row_q['row_shq_tran'], column,
                                             rm.rastr.Calc("sum", "area", "shq_tran", choice)).coordinate
        # ШР УШР без бСК
        address_SHR = self.sheet_q.cell(self.row_q['row_skrm_potr'], column,
                                        rm.rastr.Calc("sum", "node", "qsh", f"qsh>0&{choice}") - rm.rastr.Calc(
                                            "sum", "node", "qg", f"qg<0&pg<0.1&pg>-0.1&{choice}")).coordinate
        # Генерация Q генераторов
        address_qg = self.sheet_q.cell(self.row_q['row_qg'], column,
                                       rm.rastr.Calc("sum", "node", "qg", f"(pg>0.1|pg<-0.1)&{choice}")).coordinate
        # Генерация БСК шунтом и СТК СК
        address_skrm_gen = self.sheet_q.cell(self.row_q['row_skrm_gen'], column,
                                             -rm.rastr.Calc("sum", "node", "qsh", f"qsh<0&{choice}") + rm.rastr.Calc(
                                                 "sum", "node", "qg", f"qg>0&pg<0.1&pg>-0.1&{choice}")).coordinate
        # Минимальная генерация реактивной мощности в узлах выборки
        address_qg_min = self.sheet_q.cell(self.row_q['row_qg_min'], column,
                                           rm.rastr.Calc("sum", "node", "qmin", f"pg>0.1&{choice}")).coordinate
        # Максимальная генерация реактивной мощности в узлах выборки
        address_qg_max = self.sheet_q.cell(self.row_q['row_qg_max'], column,
                                           rm.rastr.Calc("sum", "node", "qmax", f"pg>0.1&{choice}")).coordinate
        # Генерация Q в ЛЭП
        address_shq_line = self.sheet_q.cell(self.row_q['row_shq_line'], column,
                                             - rm.rastr.Calc("sum", "area", "shq_line", choice)).coordinate
        address_losses = self.sheet_q.cell(self.row_q['row_dq_sum'], column,
                                           f"={address_dq_line}+{address_dq_tran}+{address_shq_tran}").coordinate
        address_load = self.sheet_q.cell(self.row_q['row_sum_port_Q'], column,
                                         f"={address_qn}+{address_losses}+{address_SHR}").coordinate
        address_sum_gen = self.sheet_q.cell(self.row_q['row_sum_QG'], column,
                                            f"={address_qg}+{address_shq_line}+{address_skrm_gen}").coordinate
        self.sheet_q.cell(self.row_q['row_Q_itog'], column,
                          f"=-{address_load}+{address_sum_gen}")
        self.sheet_q.cell(self.row_q['row_Q_itog_gmin'], column,
                          f"=-{address_load}+{address_qg_min}+{address_shq_line}")
        self.sheet_q.cell(self.row_q['row_Q_itog_gmax'], column,
                          f"=-{address_load}+{address_qg_max}+{address_shq_line}")

    def finish(self):
        """
        Преобразовать в объект таблицу и удалить листы с одной строкой.
        """
        for sheet_name in self.book.sheetnames:
            sheet = self.book[sheet_name]
            if sheet.max_row == 1:
                del self.book[sheet_name]  # удалить пустой лист
            else:
                if 'log' in sheet_name or 'Значения' == sheet_name:
                    PrintXL.create_table(sheet, sheet_name)  # Создать объект таблица.
                    if 'log' in sheet_name:
                        self.book.create_sheet(sheet_name.replace('log', 'сводная'))
                        self.sheet_couple[sheet_name] = sheet_name.replace('log', 'сводная')

        self.name_xl_file = self.task['name_time'] + ' вывод данных.xlsx'

        if self.task['print_balance_q']['add']:
            self.configure_balance_q()

        self.book.save(self.name_xl_file)
        self.book = None
        self.create_pivot()

    @staticmethod
    def create_table(sheet, sheet_name):
        """
        Создать объект таблица из всего диапазона листа.
        :param sheet: Объект лист excel
        :param sheet_name: Имя таблицы.
        """
        tab = Table(displayName=sheet_name, ref='A1:' + get_column_letter(sheet.max_column) + str(sheet.max_row))
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=True)
        tab.tableStyleInfo = style
        sheet.add_table(tab)

    def configure_balance_q(self):
        self.sheet_q.row_dimensions[2].height = 140
        self.sheet_q.column_dimensions['A'].width = 40
        thins = Side(border_style="thin", color="000000")
        for row in range(2, self.sheet_q.max_row + 1):
            for col in range(1, self.sheet_q.max_column + 1):
                if row > 2 and col > 1:
                    self.sheet_q.cell(row, col).number_format = BUILTIN_FORMATS[1]
                self.sheet_q.cell(row, col).border = Border(thins, thins, thins, thins)
                self.sheet_q.cell(row, col).font = Font(name='Times New Roman', size=11)
                if row == 2:
                    self.sheet_q.cell(row, col).alignment = Alignment(text_rotation=90,
                                                                      wrap_text=True, horizontal="center")
                if col == 1:
                    self.sheet_q.cell(row, col).alignment = Alignment(wrap_text=True)
                if row in [12, 13, 17, 18]:
                    self.sheet_q.cell(row, col).fill = PatternFill('solid', fgColor="00FF0000")
                if row in [9, 15, 16]:
                    self.sheet_q.cell(row, col).font = Font(bold=True)

    def create_pivot(self):
        """
        Открыть excel через win32com.client и создать сводные.
        :return:
        """
        excel = win32com.client.Dispatch("Excel.Application")
        excel.ScreenUpdating = False  # обновление экрана
        # excel.Calculation = -4135  # xlCalculationManual
        excel.EnableEvents = False  # отслеживание событий
        excel.StatusBar = False  # отображение информации в строке статуса excel
        try:
            self.wbook = excel.Workbooks.Open(self.name_xl_file)
        except Exception:
            raise Exception(f'Ошибка при открытии файла {self.name_xl_file=}')

        for n in range(self.wbook.sheets.count):
            if self.wbook.sheets[n].Name in self.sheet_couple:
                self.pivot_tables(self.wbook.sheets[n].Name, self.sheet_couple[self.wbook.sheets[n].Name])
        if self.task['folder_result']:
            self.wbook.Save()
        excel.Visible = True
        excel.ScreenUpdating = True  # обновление экрана
        excel.Calculation = -4105  # xlCalculationAutomatic
        excel.EnableEvents = True  # отслеживание событий
        excel.StatusBar = True  # отображение информации в строке статуса excel

    def pivot_tables(self, s_log: str, s_pivot: str) -> None:
        """Создать сводную таблицу
        s_log: имя листа с исходной таблицей
        s_pivot: имя листа для вставки сводной"""
        tab_log = None
        for n in range(self.wbook.sheets.count):
            if s_log == self.wbook.sheets[n].Name:
                tab_log = self.wbook.sheets[n].ListObjects[0]
        rows = self.task['set_printXL'][s_log[:-4]]['rows'].split(",")
        columns = self.task['set_printXL'][s_log[:-4]]['columns'].split(",")
        values = self.task['set_printXL'][s_log[:-4]]['values'].split(",")

        pt_cache = self.wbook.PivotCaches().add(1, tab_log)  # создать КЭШ xlDatabase, ListObjects
        pt = pt_cache.CreatePivotTable(s_pivot + "!R1C1", "Сводная " + s_log[:-4])  # создать сводную таблицу
        pt.ManualUpdate = True  # не обновить сводную
        # print(s_log, s_pivot)
        pt.AddFields(RowFields=rows, ColumnFields=columns, PageFields=["name_rg2"], AddToTable=False)
        # добавить поля фильтра которых нет в столбцах и строках
        # pt.AddFields RowFields:=poleRow_arr, ColumnFields:=poleCol_arr,
        # PageFields:=Array("name_rg", "лет/зим", "макс/мин", "доп_имя1", "доп_имя2") #  добавить поля

        for val in values:
            pt.AddDataField(pt.PivotFields(val), val + " ", -4157)  # xlSum добавить поле в область значений
            # Field                      Caption             def формула расчета
            pt.PivotFields(val + " ").NumberFormat = "0"

        # .PivotFields("na").ShowDetail = True #  группировка
        pt.RowAxisLayout(1)  # xlTabularRow показывать в табличной форме!!!!
        if len(values) > 1:
            pt.DataPivotField.Orientation = 1  # xlRowField"Значения в столбцах или строках xlColumnField

        # .DataPivotField.Position = 1 #  позиция в строках
        pt.RowGrand = False  # удалить строку общих итогов
        pt.ColumnGrand = False  # удалить столбец общих итогов
        pt.MergeLabels = True  # объединять одинаковые ячейки
        pt.HasAutoFormat = False  # не обновлять ширину при обновлении
        pt.NullString = "--"  # заменять пустые ячейки
        pt.PreserveFormatting = False  # сохранять формат ячеек при обновлении
        pt.ShowDrillIndicators = False  # показывать кнопки свертывания
        # pt.PivotCache.MissingItemsLimit = 0 # xlMissingItemsNone
        # xlMissingItemsNone для норм отображения уникальных значений автофильтра
        for row in rows:
            pt.PivotFields(row).Subtotals = [False, False, False, False, False, False, False, False, False, False,
                                             False, False]  # промежуточные итоги и фильтры
        for column in columns:
            pt.PivotFields(column).Subtotals = [False, False, False, False, False, False, False, False, False, False,
                                                False, False]  # промежуточные итоги и фильтры
        pt.ManualUpdate = False  # обновить сводную
        pt.TableStyle2 = ""  # стиль
        if s_log[:-4] in ["area", "area2", "darea"]:
            pt.ColumnRange.ColumnWidth = 10  # ширина строк
            pt.RowRange.ColumnWidth = 9
            pt.RowRange.Columns(1).ColumnWidth = 7
            pt.RowRange.Columns(2).ColumnWidth = 20
            pt.RowRange.Columns(3).ColumnWidth = 10
            pt.RowRange.Columns(6).ColumnWidth = 20
        pt.DataBodyRange.HorizontalAlignment = -4108  # xlCenter
        # .DataBodyRange.NumberFormat = "#,##0"
        # формат
        pt.TableRange1.WrapText = True  # перенос текста в ячейке
        pt.TableRange1.Borders(7).LineStyle = 1  # лево
        pt.TableRange1.Borders(8).LineStyle = 1  # верх
        pt.TableRange1.Borders(9).LineStyle = 1  # низ
        pt.TableRange1.Borders(10).LineStyle = 1  # право
        pt.TableRange1.Borders(11).LineStyle = 1  # внутри вертикаль
        pt.TableRange1.Borders(12).LineStyle = 1  #


def str_yeas_in_list(id_str: str):
    """
    Преобразует перечень годов.
    :param id_str: "2021,2023...2025"
    :return: [2021,2023,2024,2025] или []
    """
    years_list = id_str.replace(" ", "").split(',')
    if years_list:
        years_list_new = np.array([], int)
        for it in years_list:
            if "..." in it:
                i_years = it.split('...')
                years_list_new = np.hstack(
                    [years_list_new, np.array(np.arange(int(i_years[0]), int(i_years[1]) + 1), int)])
            else:
                years_list_new = np.hstack([years_list_new, int(it)])
        return np.sort(years_list_new)
    else:
        return []


def block_b(rm):
    rm.sel0('block_b')
    rm.rgm("block_b")


def block_e(rm):
    rm.sel0('block_e')
    rm.rgm("block_e")


def my_except_hook(func):
    """
    Переназначить функцию для добавления информации об ошибке в диалоговое окно.
    :param func:
    :return:
    """

    def new_func(*args, **kwargs):
        log.error(f"Критическая ошибка: {args[0]}, {args[1]}", exc_info=True)
        mb.showerror("Ошибка", f"Критическая ошибка: {args[0]}, {args[1]}")
        # https://python-scripts.com/python-traceback
        func(*args, **kwargs)

    return new_func


if __name__ == '__main__':
    em = None  # глобальный объект класса EditModel
    cm = None  # глобальный объект класса CalcModel
    sys.excepthook = my_except_hook(sys.excepthook)

    # DEBUG, INFO, WARNING, ERROR и CRITICAL
    # logging.basicConfig(filename="log_file.log", level=logging.DEBUG, filemode='w',
    #                     format='%(asctime)s %(name)s  %(levelname)s:%(message)s')

    log = logging.getLogger(__name__)
    log.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s %(name)s %(levelname)s:%(message)s')

    file_handler = logging.FileHandler(filename=GeneralSettings.log_file, mode='w')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    # file_handler.close()
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG)
    console_handler.setFormatter(formatter)

    log.addHandler(file_handler)
    log.addHandler(console_handler)

    app = QtWidgets.QApplication([])  # Новый экземпляр QApplication
    # app.setApplicationName("Правка моделей RastrWin")

    gui_choice_window = MainChoiceWindow()
    gui_choice_window.show()
    gui_calc_ur = CalcWindow()
    # gui_calc_ur.show()
    gui_calc_ur_set = CalcSetWindow()
    # gui_calc_ur_set.show()
    gui_edit = EditWindow()
    # gui_edit.show()
    gui_set = SetWindow()
    # gui_set.show()
    sys.exit(app.exec_())  # Запуск

# TODO дописать: перенос параметров из одноименных файлов
# TODO дописать: сравнение файлов
# TODO спросить про перезапись файлов
