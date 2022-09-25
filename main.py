# установка модулей:
# pip freeze > requirements.txt
# pip install -r requirements.txt
# exe приложение:
# pyinstaller --onefile --noconsole main.py
# pyinstaller -F --noconsole main.py
import win32com.client
from abc import ABC
from Rastr_Method import RastrMethod
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.styles.numbers import BUILTIN_FORMATS
# import pandas as pd
from typing import Union  # Any
import sys
import shutil
from PyQt5 import QtWidgets
from datetime import datetime, timedelta
from time import time
import os
import re
import configparser  # создать ini файл
# import random
import logging
import webbrowser
from tkinter import messagebox as mb
import numpy as np
import yaml
from qt_choice import Ui_choice  # pyuic5 qt_choice.ui -o qt_choice.py
from qt_set import Ui_Settings  # pyuic5 qt_set.ui -o qt_set.py
from qt_cor import Ui_cor  # pyuic5 qt_cor.ui -o qt_cor.py
from qt_calc_ur import Ui_calc_ur  # pyuic5 qt_calc_ur.ui -o qt_calc_ur.py
from qt_calc_ur_set import Ui_calc_ur_set  # pyuic5 qt_calc_ur_set.ui -o qt_calc_ur_set.py


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
            log.info(f"Выбран файл: {fileName_choose}")
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
            log.info(f"Для сохранения выбран файл: {fileName_choose}, {_}")
            return fileName_choose


class MainChoiceWindow(QtWidgets.QMainWindow, Ui_choice, Window):
    def __init__(self):
        super(MainChoiceWindow, self).__init__()
        self.setupUi(self)
        self.settings.clicked.connect(lambda: gui_set.show())
        self.correction.clicked.connect(lambda: self.hide_show((gui_choice_window,), (gui_edit,)))
        self.calc_ur.clicked.connect(lambda: self.hide_show((gui_choice_window,), (gui_calc_ur,)))


class CalcURWindow(QtWidgets.QMainWindow, Ui_calc_ur, Window):
    def __init__(self):
        super(CalcURWindow, self).__init__()
        self.setupUi(self)
        self.b_set.clicked.connect(lambda: gui_calc_ur_set.show())
        self.b_main_choice.clicked.connect(lambda: self.hide_show((gui_calc_ur,), (gui_choice_window,)))

        # Скрыть параметры при старте.
        self.check_status_visibility = (
            (self.cb_filter, self.gb_filter),
            (self.cb_cor_txt, self.te_cor_txt),
            (self.cb_import_model, self.gb_import_model),
            (self.cb_combinations, self.gb_combinations),
            (self.cb_excel, self.gb_excel),
            (self.cb_results_tab, self.gb_results_tab),
            (self.cb_tab_KO, self.gb_tab_KO),
            (self.cb_results_pic, self.gb_results_pic),
        )
        self.check_status(self.check_status_visibility)

        # CB показать / скрыть параметры.
        for CB, _ in self.check_status_visibility:
            CB.clicked.connect(lambda: self.check_status(self.check_status_visibility))


class CalcURSetWindow(QtWidgets.QMainWindow, Ui_calc_ur_set, Window):
    def __init__(self):
        super(CalcURSetWindow, self).__init__()
        self.setupUi(self)


class SetWindow(QtWidgets.QMainWindow, Ui_Settings, Window):
    def __init__(self):
        super(SetWindow, self).__init__()
        self.setupUi(self)
        self.load_ini()
        self.set_save.clicked.connect(lambda: self.save_ini())

    def load_ini(self):
        """Загрузить, создать или перезаписать файл .ini """
        if os.path.exists('settings.ini'):
            config = configparser.ConfigParser()
            config.read('settings.ini')
            try:
                self.LE_path.setText(config['DEFAULT']["folder RastrWin3"])
                self.LE_rg2.setText(config['DEFAULT']["шаблон rg2"])
                self.LE_rst.setText(config['DEFAULT']["шаблон rst"])
                self.LE_sch.setText(config['DEFAULT']["шаблон sch"])
                self.LE_amt.setText(config['DEFAULT']["шаблон amt"])
                self.LE_trn.setText(config['DEFAULT']["шаблон trn"])
            except LookupError:
                log.error('файл settings.ini не читается, перезаписан')
                self.save_ini()
        else:
            log.info('создан файл settings.ini')
            self.save_ini()

    def save_ini(self):
        config = configparser.ConfigParser()
        config['DEFAULT'] = {
            "folder RastrWin3": self.LE_path.text(),
            "шаблон rg2": self.LE_rg2.text(),
            "шаблон rst": self.LE_rst.text(),
            "шаблон sch": self.LE_sch.text(),
            "шаблон amt": self.LE_amt.text(),
            "шаблон trn": self.LE_trn.text()}
        with open('settings.ini', 'w') as configfile:
            config.write(configfile)


class EditWindow(QtWidgets.QMainWindow, Ui_cor, Window):
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
        self.choice_from_folder.clicked.connect(lambda: self.choice(type_choice='folder', insert=self.T_IzFolder))
        self.choice_from_file.clicked.connect(lambda: self.choice(type_choice='file', insert=self.T_IzFolder))
        self.choice_in_folder.clicked.connect(lambda: self.choice(type_choice='folder', insert=self.T_InFolder))
        self.choice_XL.clicked.connect(lambda: self.choice(type_choice='file', insert=self.T_PQN_XL_File))
        self.choice_N.clicked.connect(lambda: self.choice(type_choice='file', insert=self.file_N))
        self.choice_V.clicked.connect(lambda: self.choice(type_choice='file', insert=self.file_V))
        self.choice_G.clicked.connect(lambda: self.choice(type_choice='file', insert=self.file_G))
        self.choice_A.clicked.connect(lambda: self.choice(type_choice='file', insert=self.file_A))
        self.choice_A2.clicked.connect(lambda: self.choice(type_choice='file', insert=self.file_A2))
        self.choice_D.clicked.connect(lambda: self.choice(type_choice='file', insert=self.file_D))
        self.choice_PQ.clicked.connect(lambda: self.choice(type_choice='file', insert=self.file_PQ))
        self.choice_IT.clicked.connect(lambda: self.choice(type_choice='file', insert=self.file_IT))

        self.run_krg2.clicked.connect(lambda: self.start())
        self.b_main_choice.clicked.connect(lambda: self.hide_show((gui_edit,), (gui_choice_window,)))
        # Подсказки
        # self.CB_KFilter_file.setToolTip("Всплывающее окно")
        # Загрузить из .ini начальный путь для T_IzFolder
        if os.path.exists('settings.ini'):
            config = configparser.ConfigParser()
            config.read('settings.ini')
            try:
                self.T_IzFolder.setPlainText(config['save_form_folder']["path"])
            except LookupError:
                log.error('файл settings.ini не читается')

    def save_ini_form_folder(self):
        """
        Сохранить в .ini путь к папке с исходными моделями.
        """
        if os.path.exists('settings.ini'):
            config = configparser.ConfigParser()
            config.read('settings.ini')
            config['save_form_folder'] = {"path": self.T_IzFolder.toPlainText()}
            with open('settings.ini', 'w') as configfile:
                config.write(configfile)

    def choice(self, type_choice: str, insert):
        """
        Функция выбора папки или файла.
        :param type_choice: 'file', 'folder'
        :param insert: объект QT  для вставки пути выбранного файла.
        """
        name = ''
        if type_choice == 'file':
            name = self.choice_file(directory=self.T_IzFolder.toPlainText().replace('*', ''))
        elif type_choice == 'folder':
            name = self.choice_folder(directory=self.T_IzFolder.toPlainText().replace('*', ''))
        if name:
            name = name.replace('/', '\\')
            if insert.__class__.__name__ == 'QPlainTextEdit':
                insert.setPlainText(name)
            elif insert.__class__.__name__ == 'QLineEdit':
                insert.setText(name)

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
        name_file_load = self.choice_file(directory=self.T_IzFolder.toPlainText(), filter_="YAML Files (*.yaml)")
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

        self.CB_kontrol_rg2.setChecked(task_yaml["control_rg2"])
        self.CB_U.setChecked(task_yaml["control_rg2_task"]['node'])
        self.CB_I.setChecked(task_yaml["control_rg2_task"]['vetv'])
        self.CB_gen.setChecked(task_yaml["control_rg2_task"]['Gen'])
        self.CB_s.setChecked(task_yaml["control_rg2_task"]['section'])
        self.CB_na.setChecked(task_yaml["control_rg2_task"]['area'])
        self.CB_npa.setChecked(task_yaml["control_rg2_task"]['area2'])
        self.CB_no.setChecked(task_yaml["control_rg2_task"]['darea'])
        self.kontrol_rg2_Sel.setText(task_yaml["control_rg2_task"]['sel_node'])

        self.CB_printXL.setChecked(task_yaml["printXL"])
        self.CB_print_sech.setChecked(task_yaml['set_printXL']["sechen"]['add'])
        self.setsel_sech.setText(task_yaml['set_printXL']["sechen"]["sel"])
        self.CB_print_area.setChecked(task_yaml['set_printXL']["area"]['add'])
        self.setsel_area.setText(task_yaml['set_printXL']["area"]["sel"])
        self.CB_print_area2.setChecked(task_yaml['set_printXL']["area2"]['add'])
        self.setsel_area2.setText(task_yaml['set_printXL']["area2"]["sel"])
        self.CB_print_darea.setChecked(task_yaml['set_printXL']["darea"]['add'])
        self.setsel_darea.setText(task_yaml['set_printXL']["darea"]["sel"])

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
        Добавить ImportFromModel и запуск
        """
        file_handler.close()
        self.save_ini_form_folder()
        self.fill_task_ui()
        # Убрать 'file:///'
        for str_name in ["KIzFolder", "KInFolder", "excel_cor_file"]:
            self.task_ui[str_name].lstrip('file:///')
        """Запуск корректировки моделей"""
        global cm
        cm = CorModel(self.task_ui)
        self.gui_import()
        cm.run_cor()

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
                    ImportFromModel.ui_import_model.append(ifm)

    def fill_task_ui(self):
        """
        Заполнить task_ui задание взяв данные с формы QT (task_ui).
        """
        self.task_ui = {
            "KIzFolder": self.T_IzFolder.toPlainText(),  # QPlainTextEdit
            "KInFolder": self.T_InFolder.toPlainText(),
            "KFilter_file": self.CB_KFilter_file.isChecked(),  # QCheckBox
            "max_file_count": self.D_count_file.value(),  # QSpainBox
            "cor_criterion_start": {"years": self.condition_file_years.text(),  # QLineEdit text()
                                    "season": self.condition_file_season.currentText(),  # QComboBox
                                    "max_min": self.condition_file_max_min.currentText(),
                                    "add_name": self.condition_file_add_name.text()},
            # Корректировка в начале
            "cor_beginning_qt": {'add': self.CB_cor_b.isChecked(),
                                 'txt': self.TE_cor_b.toPlainText()},
            # Задание из 'EXCEL'
            "import_val_XL": self.CB_import_val_XL.isChecked(),
            "excel_cor_file": self.T_PQN_XL_File.toPlainText(),
            "excel_cor_sheet": self.T_PQN_Sheets.text(),
            # Корректировка в конце
            "cor_end_qt": {'add': self.CB_cor_e.isChecked(),
                           'txt': self.TE_cor_e.toPlainText()},
            # Расчет режима и контроль параметров режима
            "control_rg2": self.CB_kontrol_rg2.isChecked(),
            "control_rg2_task": {'node': self.CB_U.isChecked(),
                                 'vetv': self.CB_I.isChecked(),
                                 'Gen': self.CB_gen.isChecked(),
                                 'section': self.CB_s.isChecked(),
                                 'area': self.CB_na.isChecked(),
                                 'area2': self.CB_npa.isChecked(),
                                 'darea': self.CB_no.isChecked(),
                                 'sel_node': self.kontrol_rg2_Sel.text()},
            # Выводить данные из моделей в XL
            "printXL": self.CB_printXL.isChecked(),
            "set_printXL": {
                "sechen": {'add': self.CB_print_sech.isChecked(), "sel": self.setsel_sech.text(),
                           'tabl': "sechen", 'par': "ns,name,pmin,pmax,psech",
                           "rows": "ns,name",  # поля строк в сводной
                           "columns": "год,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля столбцов в сводной
                           "values": "psech,pmax"},
                "area": {'add': self.CB_print_area.isChecked(), "sel": self.setsel_area.text(), 'tabl': "area",
                         'par': 'na,name,no,pg,pn,pn_sum,dp,pop,pop_zad,qn_sum,pg_max,pg_min,pn_max,pn_min,poq,qn,qg',
                         "rows": "na,name,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля строк в сводной
                         "columns": "год",  # поля столбцов в сводной
                         "values": "pop,pg"},
                "area2": {'add': self.CB_print_area2.isChecked(),
                          "sel": self.setsel_area2.text(), 'tabl': "area2",
                          'par': 'npa,name,pg,pn,dp,pop,vnp,qg,qn,dq,poq,vnq,pn_sum,qn_sum,pop_zad',
                          "rows": "npa,name,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля строк в сводной
                          "columns": "год",  # поля столбцов в сводной
                          "values": "pop,pg"},
                "darea": {'add': self.CB_print_darea.isChecked(),
                          "sel": self.setsel_darea.text(), 'tabl': "darea",
                          'par': 'no,name,pg,pp,pvn,qn_sum,pnr_sum,pn_sum,pop_zad,qvn,qp,qg',
                          "rows": "no,name,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля строк в сводной
                          "columns": "год",  # поля столбцов в сводной
                          "values": "pp,pg"},
                "tab": {'add': self.CB_print_tab_log.isChecked(), "sel": self.print_tab_log_ar_set.text(),
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

    # @abstractmethod
    def __init__(self):
        # коллекция для хранения информации о расчете
        self.set_info = {"calc_val": {1: "ЗАМЕНИТЬ", 2: "ПРИБАВИТЬ", 3: "ВЫЧЕСТЬ", 0: "УМНОЖИТЬ"},
                         'collapse': '', 'end_info': ''}
        # прочитать ini файл
        if os.path.exists('settings.ini'):
            config = configparser.ConfigParser()
            config.read('settings.ini')
            try:
                for key in config['DEFAULT']:
                    GeneralSettings.set_save[key] = config['DEFAULT'][key]
            except LookupError:
                raise LookupError('файл settings.ini не читается')
        else:
            raise LookupError("Отсутствует файл settings.ini")

        self.file_count = 0  # счетчик расчетных файлов
        self.now = datetime.now()
        self.time_start = time()
        self.now_start = self.now.strftime("%d-%m-%Y %H:%M")

    # @abstractmethod
    def the_end(self):  # по завершению
        self.set_info['end_info'] = (
            f"РАСЧЕТ ЗАКОНЧЕН! \nНачало расчета {self.now_start}, конец {self.now.strftime('%d-%m-%Y %H:%M')}"
            f" \nЗатрачено: {timedelta(seconds=time() - self.time_start)} c.")
        log.info(self.set_info['end_info'])


class CorModel(GeneralSettings):
    """
    Коррекция файлов.
    """

    def __init__(self, task):
        super(CorModel, self).__init__()
        self.print_xl = None
        self.cor_xl = None
        self.task = task
        self.rastr_files = None
        self.all_folder = False  # Не перебирать вложенные папки
        self.load_additional = []

    def run_cor(self):
        """Запуск корректировки моделей"""
        # определяем корректировать файл или файлы в папке по анализу "KIzFolder"
        if 'KIzFolder' not in self.task:
            raise ValueError('В задании отсутствует папка для корректировки: KIzFolder')

        if "*" in self.task["KIzFolder"]:
            self.task["KIzFolder"] = self.task["KIzFolder"].replace('*', '')
            self.all_folder = True

        if os.path.isdir(self.task["KIzFolder"]):
            self.task["folder_file"] = 'folder'  # если корр папка
        elif os.path.isfile(self.task["KIzFolder"]):
            self.task["folder_file"] = 'file'  # если корр файл
        else:
            mb.showerror("Ошибка в задании", "Не найден: " + self.task["KIzFolder"] + ", выход")
            return False
        # создать папку KInFolder
        if "KInFolder" in self.task:
            if self.task["KInFolder"]:
                if not os.path.exists(self.task["KInFolder"]):
                    log.info("Создана папка: " + self.task["KInFolder"])
                    os.makedirs(self.task["KInFolder"])
            folder_save = self.task["KInFolder"] if self.task["KInFolder"] else self.task["KIzFolder"]
        else:
            self.task["KInFolder"] = ''
            folder_save = self.task["KIzFolder"]

        self.task['folder_result'] = folder_save + r"\result"  # папка для сохранения результатов
        now = datetime.now()
        self.task['name_time'] = f"{self.task['folder_result']}\\{now.strftime('%d-%m-%Y %H-%M')}"
        if not os.path.exists(self.task['folder_result']):
            os.mkdir(self.task['folder_result'])  # создать папку result
        # self.task['folder_temp'] = self.task['folder_result'] + r"\temp"  # папка для сохранения рабочих файлов
        # if not os.path.exists(self.task['folder_temp']):
        #     os.mkdir(self.task['folder_temp'])  # создать папку temp

        # ЭКСПОРТ ИЗ МОДЕЛЕЙ
        if 'block_import' in self.task:
            if self.task['block_import']:
                import_model()  # ИД для импорта

        if "import_val_XL" in self.task:
            if self.task["import_val_XL"]:  # Задать параметры узла по значениям в таблице excel (имя книги, имя листа)
                self.cor_xl = CorXL(excel_file_name=self.task["excel_cor_file"], sheets=self.task["excel_cor_sheet"])
                self.cor_xl.init_export_model()

        # Загрузить файл сечения.
        if "printXL" in self.task:
            if ((self.task["printXL"] and self.task["set_printXL"]["sechen"]['add']) or
                    (self.task["control_rg2"] and self.task["control_rg2_task"]["section"])):
                self.load_additional.append('sch')

        if self.task["folder_file"] == 'folder':  # корр файлы в папке

            if self.all_folder:  # с вложенными папками
                for address, dirs, files in os.walk(self.task["KIzFolder"]):
                    in_dir = address.replace(self.task["KIzFolder"], self.task["KInFolder"])
                    if not os.path.exists(in_dir):
                        os.makedirs(in_dir)
                    self.for_file_in_dir(from_dir=address, in_dir=in_dir)

            else:  # без вложенных папок
                self.for_file_in_dir(from_dir=self.task["KIzFolder"], in_dir=self.task["KInFolder"])

        elif self.task["folder_file"] == 'file':  # корр файл
            rm = RastrModel(full_name=self.task["KIzFolder"])
            rm.load(load_additional=self.load_additional)
            self.cor_file(rm)
            if self.task["KInFolder"]:
                rm.save(self.task["KInFolder"] + '\\' + rm.Name)
        # для нескольких запусков через GUI
        if ImportFromModel.ui_import_model:
            ImportFromModel.ui_import_model = []

        if self.print_xl:
            self.print_xl.finish()

        self.the_end()

        if 'collapse' in self.set_info:
            if self.set_info['collapse']:
                self.set_info['end_info'] += f"\nВНИМАНИЕ! развалились модели:\n[{self.set_info['collapse']}].\n"

        notepad_path = self.task['name_time'] + ' протокол коррекции файлов.log'
        shutil.copyfile('log_file.log', notepad_path)
        with open(self.task['name_time'] + ' задание на корректировку.yaml', 'w') as f:
            yaml.dump(data=self.task, stream=f, default_flow_style=False, sort_keys=False)
        # webbrowser.open(notepad_path)  #  Открыть блокнотом лог-файл.
        mb.showinfo("Инфо", self.set_info['end_info'])

    def for_file_in_dir(self, from_dir: str, in_dir: str):
        files = os.listdir(from_dir)  # список всех файлов в папке
        self.rastr_files = list(filter(lambda x: x.endswith('.rg2') | x.endswith('.rst'), files))

        for rastr_file in self.rastr_files:  # цикл по файлам .rg2 .rst в папке KIzFolder
            full_name = os.path.join(from_dir, rastr_file)
            full_name_new = os.path.join(in_dir, rastr_file)
            rm = RastrModel(full_name)
            # если включен фильтр файлов и имя стандартизовано
            if self.task["KFilter_file"] and rm.kod_name_rg2:
                if not rm.test_name(condition=self.task["cor_criterion_start"], info='Цикл по файлам KIzFolder: '):
                    continue  # пропускаем если не соответствует фильтру

            self.file_count += 1
            #  если включен фильтр файлов проверяем количество расчетных файлов
            if self.task["KFilter_file"] and self.file_count == self.task["max_file_count"] + 1:
                break
            rm.load(load_additional=self.load_additional)
            self.cor_file(rm)
            if self.task["KInFolder"]:
                rm.save(full_name_new)

    def cor_file(self, rm):
        """Корректировать файл rm"""
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
        if ImportFromModel.ui_import_model:
            for im in ImportFromModel.ui_import_model:
                im.import_csv(rm)

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
                rm.cor_txt_field(table_field=self.task["cor_name_task"])

        if 'control_rg2' in self.task:
            if self.task['control_rg2']:
                if not rm.control_rg2(self.task['control_rg2_task']):  # расчет и контроль параметров режима
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
        self.name_base = self.Name[:-4]  # вернуть имя без расширения "2020 зим макс"
        self.tip_file = self.Name[-3:]  # rst или rg2
        self.pattern = GeneralSettings.set_save["шаблон " + self.tip_file]
        self.kod_name_rg2 = 0  # 0 не распознан, 1 зим макс 2 зим мин 3 лет макс 4 лет мин 5 паводок макс
        self.temp_a_v_gost = False  # True температуры:  а-в - зима + лето ПЭВТ
        self.TabRgmCount = 1  # счетчик режимов в каждой таблице
        self.all_auto_shunt = {}
        self.txt_dop = ""
        self.degree_int = 0
        self.degree_str = ""
        self.loadRGM = False
        self.DopNameStr = ""
        self.DopName = ""
        self.rastr = None
        self.name_list = ["-", "-", "-"]
        pattern_name = re.compile("^(20[1-9][0-9])\s(лет\w?|зим\w?|паводок)\s?(макс|мин)?")
        match = re.search(pattern_name, self.name_base)
        if match:
            if match.re.groups == 3:
                self.name_list = [match[1], match[2], match[3]]
                if not self.name_list[2]:
                    self.name_list = "-"
                if self.name_list[1] == "паводок":
                    self.kod_name_rg2 = 5
                    self.season_name = "Паводок"
                if self.name_list[1] == "зим" and self.name_list[2] == "макс":
                    self.kod_name_rg2 = 1
                    self.season_name = "Зимний максимум нагрузки"
                if self.name_list[1] == "зим" and self.name_list[2] == "мин":
                    self.kod_name_rg2 = 2
                    self.season_name = "Зимний минимум нагрузки"
                if self.name_list[1] == "лет" and self.name_list[2] == "макс":
                    self.kod_name_rg2 = 3
                    self.season_name = "Летний максимум нагрузки"
                if self.name_list[1] == "лет" and self.name_list[2] == "мин":
                    self.kod_name_rg2 = 4
                    self.season_name = "Летний минимум нагрузки"

        self.god = self.name_list[0]
        if self.kod_name_rg2 > 0:
            self.name_standard = self.god + " " + self.name_list[1]
            if self.kod_name_rg2 < 5:
                self.name_standard += " " + self.name_list[2]
            if (self.kod_name_rg2 in [1, 2]) or ("ПЭВТ" in self.name_base):
                self.temp_a_v_gost = True  # зима + период экстремально высоких температур -ПЭВТ
        else:
            self.name_standard = "не стандарт"  # отсеиваем файлы задание и прочее
        # поиск в строке значения в ()
        pattern_name = re.compile(r"\((.+)\)")
        match = re.search(pattern_name, self.name_base)
        if match:
            self.DopNameStr = match[1]

        if self.DopNameStr.replace(" ", "") != "":
            if "," in self.DopNameStr:
                self.DopName = self.DopNameStr.split(",")
            elif ";" in self.DopNameStr:
                self.DopName = self.DopNameStr.split(";")
            else:
                self.DopName = [self.DopNameStr]
        if "°C" in self.name_base:
            pattern_name = re.compile("(-?\d+((,|\.)\d*)?)\s?°C")  # -45.,14 °C
            match = re.search(pattern_name, self.name_base)
            if match:
                self.degree_str = match[1].replace(',', '.')
                self.degree_int = float(self.degree_str)  # число
                self.txt_dop = "Расчетная температура " + self.degree_str + " °C. "

        # if calc_set == 2:  # расчет режимов
        #     if self.kod_name_rg2 > 0:
        # if GLR.zad_temperature == 1:
        #     if self.name_list[1] == "зим":
        #         self.degree_int = GLR.temperature_zima
        #     else:
        #         self.degree_int = GLR.temperature_leto
        #
        #     self.degree_str = str(self.degree_int)
        #     self.txt_dop = "Расчетная температура " + self.degree_str + " °C. "

        # for DopName_tek in self.DopName:
        #     for each ii in GLR.rg2_name_metka
        #         if trim (DopName_tek) = trim (ii (0)):
        #             txt_dop = txt_dop + ii (1)

        # self.NAME_RG2_plus = self.season_name + " " + self.god + " г"
        # if self.txt_dop != "":
        #     self.NAME_RG2_plus += ". " + self.txt_dop
        # self.NAME_RG2_plus2 = self.season_name + "(" + self.degree_str + " °C)"
        # self.TEXT_NAME_TAB = GLR.tabl_name_OK1 + str(
        #     GLR.Ntabl_OK) + GLR.tabl_name_OK2 + self.season_name + " " + self.god + " г. " + self.txt_dop

    def test_name(self, condition: dict, info: str = "") -> bool:
        """
        Проверка имени файла на соответствие условию condition.
        Возвращает True, если имя режима соответствует условию condition:
        нр, год("2020,2023-2025"), зим/лет/паводок("лет,зим"), макс/мин("макс"), доп имя("-41С;МДП:ТЭ-У")
        condition = {"years":"","season":"","max_min":"","add_name":""}-всегда истина
        str = для вывода в протокол
        """
        if not self.kod_name_rg2:
            return True
        if not condition:
            return True
        if not (any(condition.values())):  # условие пустое
            return True

        if 'years' in condition:
            if condition['years']:
                fff = False
                for us in str_yeas_in_list(str(condition['years'])):
                    if int(self.god) == us:
                        fff = True
                if not fff:
                    log.debug(info + self.Name + f" Год '{self.god}' не проходит по условию: "
                              + str(condition['years']))
                    return False

        if 'season' in condition:
            if condition['season']:
                if condition['season'].strip():  # ПРОВЕРКА "зим" "лет" "паводок"
                    fff = False
                    _ = condition['season'].replace(' ', '')
                    for us in _.split(","):
                        if self.name_list[1] == us:
                            fff = True
                    if not fff:
                        log.debug(info + self.Name + f" Сезон '{self.name_list[1]}' не проходит по условию: "
                                  + condition['season'])
                        return False

        if 'max_min' in condition:
            if condition['max_min']:
                if condition['max_min'].strip():  # ПРОВЕРКА "макс" "мин"
                    if self.name_list[2] != condition['max_min'].replace(' ', ''):
                        log.debug(info + self.Name + f" '{self.name_list[2]}' не проходит по условию: "
                                  + condition['max_min'])
                        return False

        if 'add_name' in condition:
            if condition['add_name']:  # ПРОВЕРКА (-41С;МДП:ТЭ-У)
                if condition['add_name'].strip():  # ПРОВЕРКА (-41С;МДП:ТЭ-У)
                    if ";" in condition['add_name']:
                        _ = condition['add_name'].split(";")
                    else:
                        _ = condition['add_name'].split(",")
                    fff = False
                    for us in _:
                        for DopName_i in self.DopName:
                            if DopName_i == us:
                                fff = True
                    if not fff:
                        log.debug(
                            info + self.Name + f" Доп. имя {self.DopNameStr} не проходит по условию: " + condition[
                                'add_name'])
                        return False
        return True

    def load(self, load_additional: list = None):
        """загрузить модель в Rastr
        load_additional=['amt','sch','trn'] расширения файлов которые нужно загрузить
        загружается первый попавшийся файл в папке IzFolder"""
        if not self.rastr:
            try:
                self.rastr = win32com.client.Dispatch("Astra.Rastr")
            except Exception:
                raise Exception('Com объект Astra.Rastr не найден')

        self.rastr.Load(1, self.full_name, self.pattern)  # загрузить или перезагрузить
        log.info(f"\n\nЗагружен файл: {self.full_name}\n")
        # Загрузить файлы load_additional
        if load_additional:
            self.downloading_additional_files(load_additional)

    def downloading_additional_files(self, load_additional: list = None):
        """
        Загрузка в Rastr дополнительных файлов.
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

    def control_rg2(self, dict_task: dict):
        """  контроль  dict_task = {'node': True, 'vetv': True, 'Gen': True, 'section': True,
             'area': True, 'area2': True, 'darea': True, 'sel_node': "na>0"}  """
        if not self.rgm("control_rg2"):
            return False

        node = self.rastr.tables("node")
        branch = self.rastr.tables("vetv")
        generator = self.rastr.tables("Generator")
        chart_pq = self.rastr.tables("graphik2")
        graph_it = self.rastr.tables("graphikIT")
        # Также проверяется наличие узлов без ветвей, ветвей без узлов начала или конца, генераторов без узлов.
        all_ny = set([x[0] for x in node.writesafearray("ny", "000")])
        all_ip = set([x[0] for x in branch.writesafearray("ip", "000")])
        all_iq = set([x[0] for x in branch.writesafearray("iq", "000")])
        all_iq_ip = all_ip.union(all_iq)
        all_gen_ny = set([x[0] for x in generator.writesafearray("Node", "000")])
        # Узлы без ветвей.
        all_ny_not_branches = all_ny - all_iq_ip
        if all_ny_not_branches:
            log.error(f'В таблице node узлы без ветвей: {all_ny_not_branches}')
        # Ветви без узлов.
        all_ipiq_not_node = all_iq_ip - all_ny
        if all_ipiq_not_node:
            log.error(f'В таблице vetv есть ссылка на узлы которых нет в таблице node: {all_ipiq_not_node}')
        # Генераторы без узлов.
        all_gen_not_node = all_gen_ny - all_ny
        if all_gen_not_node:
            log.error(f'В таблице Generator есть ссылка на узлы которых нет в таблице node: {all_gen_not_node}')

        # Напряжения
        if dict_task["node"]:
            log.info("\tКонтроль напряжений.")
            self.voltage_nominal(choice=(dict_task["sel_node"] + '&uhom>30'))
            self.voltage_normal(choice=dict_task["sel_node"])
            self.voltage_deviation(choice=dict_task["sel_node"])
            self.voltage_error(choice=dict_task["sel_node"])

        # Токи
        if dict_task['vetv']:
            # Контроль токовой загрузки
            log.info("\tКонтроль токовой загрузки, расчетная температура: " + self.degree_str)
            self.rastr.CalcIdop(self.degree_int, 0.0, "")
            if dict_task["sel_node"] != "":
                if node.cols.Find("sel1") < 0:
                    node.Cols.Add("sel1", 3)  # добавить столбцы sel1
                node.cols.item("sel1").calc(0)
                node.setsel(dict_task["sel_node"])
                node.cols.item("sel1").calc(1)
                sel_vetv = "i_zag>=0.1&(ip.sel1|iq.sel1)"
                presence_n_it = {'n_it': "(ip.sel1|iq.sel1)&n_it>0",
                                 'n_it_av': "(ip.sel1|iq.sel1)&n_it_av>0"}
            else:
                sel_vetv = "i_zag>=0.1"
                presence_n_it = {'n_it': "n_it>0",
                                 'n_it_av': "n_it_av>0"}

            branch.setsel(sel_vetv)
            if branch.count:  # есть превышений
                j = branch.FindNextSel(-1)
                while j > -1:
                    log.info(f"\t\tВНИМАНИЕ ТОКИ! vetv:{branch.SelString(j)}, "
                             f"{branch.cols.item('name').ZS(j)} - {branch.cols.item('i_zag').ZS(j)} %")
                    j = branch.FindNextSel(j)

            # проверка наличия n_it,n_it_av в таблице График_Iдоп_от_Т(graphikIT)
            all_graph_it = set(graph_it.writesafearray("Num", "000"))
            for field, sel_vetv_n_it in presence_n_it.items():
                branch.setsel(sel_vetv_n_it)
                if branch.count:
                    j = branch.FindNextSel(-1)
                    while j > -1:
                        n_it = branch.cols.item(field).Z(j)
                        if (n_it,) not in all_graph_it and n_it > 0:
                            log.error(f"\t\tВНИМАНИЕ graphikIT! vetv: {branch.SelString(j)!r}, "
                                      f"{branch.cols.item('name').ZS(j)!r}, "
                                      f"{field}={n_it} не найден в таблице График_Iдоп_от_Т")
                        j = branch.FindNextSel(j)
        #  ГЕНЕРАТОРЫ
        if dict_task['Gen']:
            log.info("\tКонтроль генераторов")
            if dict_task["sel_node"] != "":
                if node.cols.Find("sel1") < 0:
                    node.Cols.Add("sel1", 3)  # добавить столбцы
                node.cols.item("sel1").calc(0)
                node.setsel(dict_task["sel_node"])
                node.cols.item("sel1").calc(1)
                sel_gen = "!sta&Node.sel1"
            else:
                sel_gen = "!sta"

            generator.setsel(sel_gen)
            j = generator.FindNextSel(-1)
            while j != -1:
                Pmin = generator.cols.item("Pmin").Z(j)
                Pmax = generator.cols.item("Pmax").Z(j)
                P = generator.cols.item("P").Z(j)
                Name = generator.cols.item("Name").ZS(j)
                Num = generator.cols.item("Num").ZS(j)
                Node = generator.cols.item("Node").ZS(j)
                NumPQ = generator.cols.item("NumPQ").Z(j)
                if P < Pmin and Pmin:
                    log.info(f"\t\tВНИМАНИЕ! {Name}, Num={Num},ny={Node}, P={str(round(P))} < Pmin={str(Pmin)}")
                if P > Pmax and Pmax:
                    log.info(f"\t\tВНИМАНИЕ! {Name}, Num={Num},ny={Node}, P={str(round(P))} > Pmax={str(Pmax)}")
                if NumPQ:
                    chart_pq.setsel("Num=" + str(NumPQ))
                    if chart_pq.count == 0:
                        log.info(f"\t\tВНИМАНИЕ! ГЕНЕРАТОР: {Name}, {Num=},ny={Node}, "
                                 f"NumPQ={NumPQ} не найден в таблице PQ-диаграммы (graphik2)")
                j = generator.FindNextSel(j)
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

        if dict_task['area']:
            self.control_pop('area')
        if dict_task['area2']:
            self.control_pop('area2')
        if dict_task['darea']:
            self.control_pop('darea')
        return True

    def control_pop(self, zone: str):
        """zone =  'darea', 'area', 'area2'"""
        key_zone = {'darea': 'no', 'area': 'na', 'area2': 'npa',
                    'darea_pop': 'pp', 'area_pop': 'pop', 'area2_pop': 'pop',
                    'darea_name': 'объединений', 'area_name': 'районов', 'area2_name': 'территорий'}

        log.info("\tКонтроль pop_zad " + key_zone[zone + '_name'])
        tabl = self.rastr.tables(zone)
        if tabl.cols.Find("pop_zad") < 0:
            log.error("Поле pop_zad отсутствует в таблице " + key_zone[zone + '_name'])
        else:
            tabl.setsel("pop_zad>0")
            j = tabl.FindNextSel(-1)
            while j != -1:
                pop_zad = round(tabl.cols.item("pop_zad").Z(j))
                pp = round(tabl.cols.item(key_zone[zone + '_pop']).Z(j))
                deviation = round(abs(pop_zad - pp) / pop_zad, 2)
                if deviation > 0.01:
                    name = tabl.cols.item("name").ZS(j)
                    no = tabl.cols.item(key_zone[zone]).ZS(j)
                    log.info(f"\t\tВНИМАНИЕ: {name} ({no}), pop: {str(pp)}, pop_zad: {str(pop_zad)}, "
                             + f"отклонение: {str(round(pop_zad - pp))} или {str(round(deviation * 100))} %")
                j = tabl.FindNextSel(j)

    def cor_rm_from_txt(self, task_txt: str):
        """
        Корректировать модели по заданию в текстовом формате
        :param task_txt:
        """
        task_rows = task_txt.split('\n')
        for task_row in task_rows:
            task_row = task_row.split('#', 1)[0]  # удалить текст после '#'
            # Имя функции стоит перед "(" и "["
            name_fun = task_row.split('(', 1)[0]
            name_fun = name_fun.split('[', 1)[0]
            name_fun = name_fun.replace(' ', '')
            if not name_fun:
                continue  # К следующей строке.

            # Условие выполнения в фигурных скобках
            condition_dict = {}
            statements = ''
            match = re.search(re.compile(r"\{(.+?)}"), task_row)
            if match:
                conditions = match[1].split('|')
                for condition in conditions:
                    condition = condition.strip()
                    if 'add_name' not in condition:
                        condition = condition.replace(' ', '')
                    if ":" not in condition:
                        raise ValueError(f'Ошибка в задании {condition!r}')
                    parameter, value = condition.split(':')
                    if parameter in ['years', 'season', 'max_min', 'add_name']:
                        condition_dict[parameter] = value
                    else:
                        statements += condition + '|'
            if condition_dict and self.kod_name_rg2:
                if not self.test_name(condition=condition_dict):
                    continue  # К следующей строке.
            if statements:
                if not self.test_parameter_rm_all(statements):
                    continue
            # Параметры функции в квадратных скобках
            function_parameters = []
            match = re.search(re.compile(r"\[(.+?)]"), task_row)
            if match:
                function_parameters = match[1].split('|')
            function_parameters += ['', '']
            self.txt_task_cor(name=name_fun, sel=function_parameters[0], value=function_parameters[1])

    def txt_task_cor(self, name: str, sel: str = '', value: str = ''):
        """
        Функция для выполнения задания в текстовом формате
        :param name: имя функции
        :param sel: выборка
        :param value: значение
        """
        name = name.lower()
        if 'уд' in name:
            self.cor(keys=sel, tasks='del', del_all=('*' in name))
        elif 'изм' in name:
            self.cor(keys=sel, tasks=value)
        elif 'снять' in name:
            self.cor(keys='(node); (vetv); (Generator)', tasks='sel=0')
        elif 'расчет' in name:
            self.rgm(txt='txt_task_cor')
        elif 'добавить' in name:
            self.table_add_row(table=sel, tasks=value)
        elif 'текст' in name:
            self.cor_txt_field(table_field=sel)
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
        elif 'ном' in name:  # номинальные напряжения
            self.voltage_nominal(choice=sel, edit=True)
        elif 'скрм' in name:
            if 'скрм*' in name:
                self.all_auto_shunt = self.auto_shunt_rec(selection=sel, only_AutoBsh=False)
            else:
                self.all_auto_shunt = self.auto_shunt_rec(selection=sel, only_AutoBsh=True)
            self.auto_shunt_cor(all_auto_shunt=self.all_auto_shunt)
        else:
            raise ValueError(f'Задание {name=} не распознано ({sel=}, {value=})')

    def loading_section(self, ns: int, p_new: Union[float, str], type_correction: str = 'pg'):
        """
        Изменить переток мощности в сечении номер ns до величины p_new за счет изменения нагрузки('pn') или
        генерации ('qn') в отмеченых узлах и генераторах
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
            log.debug(f'\t{cycle=}, {p_current=}, {p_new=}, {change_p=} МВт ({round (abs(change_p/p_new)*100)} %)')

            if abs(change_p / p_new) * 100 < accuracy:
                if (p_current < p_new and p_new > 0) or (p_current > p_new and p_new < 0):
                    log.info(f'\tЗаданная точность достигнута P={p_current},'
                             f' отклонение {change_p}. {cycle+1} итераций')
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
                NodeGeneration.change_p =change_p
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
                # Коэффициент на сколько нужно умножить резерв Рген и прибавивить его к резерву Рген, для увеличения генерации
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
    """Класс для хранения информации о узле для изменения мощности в сечении."""
    dr_p_koeff = 0  # если  1 то умножаем дополнительно на dr_p в этом случае больше загружаются
    # генераторы которые меньше влияют на изменение мощности в сечении

    no_pmin = True  # ' не учитывать Pmin 
    change_p = 0
    unbalance_p = 0

    def __init__(self, i: int, rastr):
        """
        :param i: Индекс в таблице узлы
        :param rastr: 
        """
        self.gen_available = False  # узел c генераторами
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
                self.gen_available = True  # узел c генераторами
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
            deviation_pg = koef_p_up * self.reserve_p_up   # На сколько нужно изменить генерацию в узле
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

        if not self.gen_available :  # нет генераторов
            if self.up_pgen:  # увеличиваем генерацию узла, koef_p_up
                if self.pg_max and self.pg_max > pg_node:
                    if self.pg_min > pg_node + deviation_pg:  # (от 0 до pg_min)
                        if self.pg_min and not self.no_pmin:  # если есть Рмин и учитываем Рмин то
                            if change_p > self.pg_min:
                                self.node_t.cols.Item("pg").SetZ(self.i,self.pg_min)
                                # unbalance_p = unbalance_p + (self.pg_min - deviation_pg)
                                change_p = change_p - self.pg_min
                        else:  # нет Рмин или не учитываем Рмин
                            self.node_t.cols.Item("pg").SetZ(self.i, pg_node + deviation_pg)
                            change_p = change_p - deviation_pg
                    elif self.pg_max > pg_node + deviation_pg and (self.pg_min < pg_node + deviation_pg or self.pg_min == pg_node + deviation_pg):  # (от pg_min (включительно) до pg_max)v
                        self.node_t.cols.Item("pg").SetZ(self.i,pg_node + deviation_pg)
                        change_p = change_p - deviation_pg
                    elif self.pg_max < pg_node + deviation_pg or self.pg_max == pg_node + deviation_pg:  # (больше или равно pg_max)
                        self.node_t.cols.Item("pg").SetZ(self.i, self.pg_max)
                        change_p = change_p - (self.pg_max - pg_node)

            else:  # снижаем генерацию узла,KefPG_Down
                if self.pg_min < pg_node - deviation_pg or self.pg_min == pg_node - deviation_pg:  # (от pg_min (включительно) до pg_node)
                    self.node_t.cols.Item("pg").SetZ(self.i, pg_node - deviation_pg)
                    change_p = change_p - deviation_pg

                elif self.pg_min > pg_node - deviation_pg and pg_node - deviation_pg:  # (от 0 до pg_min)
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

            # kod = self.rastr.rgm('')
            # if kod != 0 or fKontrSech (ns)  == False:  # fKontrSech возвращает истина если мощность в сечениях отмеченных контроль (sta) не превышена   или ложь (исключение)
            #     self.node_t.cols.Item("pg").SetZ(self.i, pg_node)
            #     log.info( f"Разваливается узел {self.ny=} генерацию вернули назад")
            #     kod = self.rastr.rgm('')
            #     if not kod:
            #         log.error(f"Аварийное завершение расчета режима, подвел узел {self.ny=} " )
            #         KSC.inini = KSC.Ncikl
            #         return False

            # if KSC.print_gen = 1: log.info (txt & ": увеличение на " & cstr(pg_max -  pg_node))
        # elif self.gen_available:  # если есть ГЕНЕРАТОРЫ в узле
        #     FOR             UGMC            IN            DGen.Items
        #     pgGen = tG.Cols("P").Z(UGMC.ndxGen)  # запомнить начальное состояние
        #     StaGen = tG.Cols("sta").Z(UGMC.ndxGen)  # запомнить начальное состояние
        #     if StaGen: pgGen = 0
        #
        #     if self.up_pgen and deviation_pg <> 0:  # если УВЕЛИЧИТЬ ГЕНЕРАЦИЮ
        #         if UGMC.Pmax > 0 and pgGen < UGMC.Pmax:  # если задан  UGMC.Pmax
        #             if pgGen + deviation_pg < UGMC.Pmin:  # (от 0 до Pmin ) сравнивеем резерв снижения Р в узле и величину Р на которую нада изменить мощнсть узла
        #                 if UGMC.Pmin > 0 and not self.no_pmin:  # если есть Pmin и учитываем ее
        #                     if change_p > UGMC.Pmin:
        #                         if StaGen: tG.Cols("sta").Z(UGMC.ndxGen) = false
        #                         tG.Cols("P").Z(UGMC.ndxGen) = UGMC.Pmin
        #                         # unbalance_p = unbalance_p + (UGMC.Pmin - deviation_pg)
        #                         change_p = change_p - UGMC.Pmin
        #                         deviation_pg = 0
        #
        #                 else:  # если нет Pmin или не  учитываем ее
        #                     if StaGen:
        #                         tG.Cols("sta").Z(UGMC.ndxGen) = false
        #                     tG.Cols("P").Z(UGMC.ndxGen) = pgGen + deviation_pg  #
        #                     change_p = change_p - deviation_pg
        #                     deviation_pg = 0  # если этого ген достаточно для снижения и другие ген трогать не нада
        #
        #             elif pgGen + deviation_pg < UGMC.Pmax and (pgGen + deviation_pg > UGMC.Pmin or pgGen + deviation_pg
        #             = UGMC.Pmin):  # (от Pmin (включительно) до  Pmax)
        #                 if StaGen:
        #                     tG.Cols("sta").Z(UGMC.ndxGen) = false
        #                 tG.Cols("P").Z(UGMC.ndxGen) = pgGen + deviation_pg
        #                 change_p = change_p - deviation_pg
        #                 deviation_pg = 0  # если этого ген достаточно для снижения и другие ген трогать не нада
        #             elif pgGen + deviation_pg > UGMC.Pmax or pgGen + deviation_pg =  UGMC.Pmax:  # (равно или больше Pmax)
        #                 if StaGen:
        #                     tG.Cols("sta").Z(UGMC.ndxGen) = false
        #                 tG.Cols("P").Z(UGMC.ndxGen) = UGMC.Pmax
        #                 change_p = change_p - (UGMC.Pmax - pgGen)
        #                 deviation_pg = deviation_pg - (UGMC.Pmax - pgGen)
        #
        #         else:
        #             if UGMC.Pmax =  0: log.info(                        vbtab & "Не задан Pmax у генератора " & cstr(UGMC.Num) & "в узле " & cstr(
        #                     ny) & "исключен из рассматрения")
        #
        #     elif not self.up_pgen and deviation_pg <> 0:  # если СНИЖЕНИЕ ГЕНЕРАЦИИ
        #         if StaGen = false:  # если ген включен
        #             if pgGen - UGMC.Pmin > deviation_pg or pgGen - UGMC.Pmin = deviation_pg:  # (от Pmin(включительно) до pgGen) сравнивеем резерв снижения Р в узле и величину Р на которую нада изменить мощнсть узла
        #                 tG.Cols("P").Z(UGMC.ndxGen) = pgGen - deviation_pg  #
        #                 change_p = change_p - deviation_pg
        #                 deviation_pg = 0  # если этого ген достаточно для снижения и другие ген трогать не нада
        #             elif pgGen < deviation_pg or pgGen = deviation_pg:  # (меньше или равно 0)
        #                 tG.Cols("sta").Z(UGMC.ndxGen) = True
        #                 change_p = change_p - (pgGen)
        #                 deviation_pg = deviation_pg - pgGen
        #             elif pgGen - deviation_pg > 0 and UGMC.Pmin > pgGen - deviation_pg:  # (от 0 до Pmin)
        #                 if UGMC.Pmin > 0 and not self.no_pmin:  # если есть Pmin и учитываем ее
        #                     tG.Cols("P").Z(UGMC.ndxGen) = UGMC.Pmin
        #                     change_p = change_p - (pgGen - UGMC.Pmin)
        #                     deviation_pg = deviation_pg - (pgGen - UGMC.Pmin)
        #                     if change_p > UGMC.Pmin:
        #                         tG.Cols("sta").Z(UGMC.ndxGen) = true
        #                         # unbalance_p = unbalance_p + (UGMC.Pmin - deviation_pg)
        #                         change_p = change_p - UGMC.Pmin
        #
        #                 else:  # если нет Pmin или не  учитываем ее
        #                     tG.Cols("P").Z(UGMC.ndxGen) = pgGen - deviation_pg  #
        #                     change_p = change_p - deviation_pg
        #                     deviation_pg = 0  # если этого ген достаточно для снижения и другие ген трогать не нада
        #     NodeGeneration.abs_change_p = change_p
        #     NodeGeneration.unbalance_p = unbalance_p
        #     kod = Rastr.rgm("")
        #     if kod <> 0 OR fKontrSech (ns)  = False:  # fKontrSech возвращает истина если мощность в сечениях отмеченных контроль (sta) не превышена   или ложь (исключение)
        #         log.info("Разваоивается узел ny = " & cstr(ny) & " генерацию вернули назад")
        #         tG.Cols("P").Z(UGMC.ndxGen) = pgGen
        #         tG.Cols("sta").Z(UGMC.ndxGen) = StaGen
        #         kod = Rastr.rgm("")
        #         if kod <> 0:
        #             log.info(                        vbtab & "Аварийное завершение расчета режима, подвел узел ny = " & cstr(ny) & "ген N= " & cstr(
        #                     Num)): KSC.inini = KSC.Ncikl
        #         if kod <> 0:
        #             exit                    sub




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
    SHAPE = {"Параметры импорта из файлов RastrWin": 'import_model',
             'Выполнить изменение модели по строкам': 'list_cor'}

    def __init__(self, name: str, obj):
        """
        :param name: имя листа
        :param obj: объект лист
        """
        # type:
        # 'tab_cor' - сводная таблица корректировок, нр корр потребления или нагрузка узлов / имя файла;
        # 'list_cor' - таблица корректировок по списку, нр изм, удалить, снять отметку;
        # 'import_model' - импорт моделей;
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

    def import_model(self, rm: RastrModel) -> None:
        """Импорт в модели"""
        if self.import_model_all:
            for im in self.import_model_all:
                im.import_csv(rm)

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

                    if any([year, season, max_min, add_name]) and rm.kod_name_rg2:  # any если хотя бы один истина
                        if not rm.test_name(condition={"years": year, "season": season,
                                                       "max_min": max_min, "add_name": add_name},
                                            info=f'\t\tcor_x:{sel=}, {value=}'):
                            continue
                    if statement:
                        if not rm.test_parameter_rm_all(statement):
                            continue
                    rm.txt_task_cor(name=name_fun, sel=sel, value=value)

    def tab_cor(self, rm: RastrModel) -> None:
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
                        pattern_name = re.compile("\[(.*)\].*\[(.*)\].*\[(.*)\].*\[(.*)\]")
                        match = re.search(pattern_name, name_file)
                        if match.re.groups == 4:
                            if rm.test_name(condition={"years": match[1], "season": match[2],
                                                       "max_min": match[3], "add_name": match[4]},
                                            info=f"\tcor_xl, условие: {name_file}, ") or not rm.kod_name_rg2:
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
                                    rm.cor(keys=str(short_key), tasks=f"{param}={new_val}", print_log=True)
                                else:
                                    rm.cor(keys=str(short_key),
                                           tasks=f"{param}={param}{calc_vals[self.calc_val]}{new_val}",
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
        :param excel_file_name: полное имя файла excel, нр I:\примеры.xlsx;
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
                self.sheets = re.findall("\[(.+?)\]", sheets)
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
    # __slots__ = 'ui_import_model', 'calc_str'
    ui_import_model = []  # хранение объектов класса ImportFromModel созданных в GUI и коде
    calc_str = {"обновить": 2, "загрузить": 1, "присоединить": 0, "присоединить-обновить": 3, "объединить": 3}
    number = 0  # для создания уникального имени csv файла

    def __init__(self, import_file_name: str = '', criterion_start: Union[dict, None] = None, tables: str = '',
                 param='', sel: Union[str, None] = '', calc: Union[int, str] = '2'):
        """
        Импорт данных из файлов .rg2, .rst и др.
        Создает папку temp в папке с файлом и сохраняет в ней .csv файлы
        :param import_file_name: полное имя файла
        :param criterion_start: {"years": "","season": "","max_min": "", "add_name": ""} условие выполнения
        :param tables: таблица для импорта, нр "node;vetv"
        :param param: параметры для импорта: "" все параметры или перечисление, нр 'sel,sta'(ключи не обязательно)
        :param sel: выборка нр "sel" или "" - все
        :param calc: число типа int, строка или ключевое слово:
        {"обновить": 2 , "загрузить": 1, "присоединить": 0, "присоединить-обновить": 3}
        """
        self.import_rm = None
        if not os.path.exists(import_file_name):
            raise ValueError("Ошибка в задании, не найден файл: " + import_file_name)
        else:
            self.folder_temp = os.path.dirname(import_file_name) + '\\temp'
            if not os.path.exists(self.folder_temp):
                os.mkdir(self.folder_temp)

            self.import_file_name = import_file_name
            self.basename = os.path.basename(import_file_name)
            self.criterion_start = criterion_start
            self.tables = tables.replace(' ', '').split(";")  # разделить на ["таблицы"]
            self.param = []
            if sel:
                self.sel = sel
            else:
                self.sel = ''
                pass
            if type(calc) == int:
                self.calc = calc
            elif calc.isdigit():
                self.calc = int(calc)
            else:
                if calc in self.calc_str:
                    self.calc = self.calc_str[calc]
                else:
                    raise ValueError("Ошибка в задании, не распознано задание calc ImportFromModel: " + calc)
            self.file_csv = []
            ImportFromModel.number += 1
            for tabl in self.tables:
                self.file_csv.append(f"{self.folder_temp}\\{self.basename}_{tabl}_{ImportFromModel.number}.csv")
                self.param.append(param)

            # Экспорт данных из файла в .csv файлы в папку temp
            if self.import_file_name:
                log.info(f'Экспорт из файла <{self.import_file_name}> в CSV')
                self.import_rm = RastrModel(full_name=self.import_file_name)
                self.import_rm.load()
                for index in range(len(self.tables)):
                    if not self.param[index]:  # если все параметры
                        self.param[index] = self.import_rm.all_cols(self.tables[index])
                    else:  # добавить к строке параметров ключи текущей таблицы
                        if self.import_rm.rastr.Tables(self.tables[index]).Key not in self.param[index]:
                            self.param[index] += ',' + self.import_rm.rastr.Tables(self.tables[index]).Key

                    log.info(f"\n\tТаблица: {self.tables[index]}. Выборка: {self.sel}\n"
                             + f"\tПараметры: {self.param[index]}\n\tФайл CSV: {self.file_csv[index]}")

                    tab = self.import_rm.rastr.Tables(self.tables[index])
                    tab.setsel(self.sel)
                    tab.WriteCSV(1, self.file_csv[index], self.param[index], ";")  # 0 дописать, 1 заменить

    def import_csv(self, rm: RastrModel) -> None:
        """Импорт данных из csv в файла"""
        if self.import_file_name:
            log.info(f"\tИмпорт из CSV <{self.import_file_name}> в модель:")
            if rm.test_name(condition=self.criterion_start, info='\tImportFromModel ') or not rm.kod_name_rg2:
                for index in range(len(self.tables)):
                    log.info(f"\n\tТаблица: {self.tables[index]}. Выборка: {self.sel}. тип: {self.calc}" +
                             f"\n\tФайл CSV: {self.file_csv[index]}" +
                             f"\n\tПараметры: {self.param[index]}")
                    """{"обновить": 2 , "загрузить": 1, "присоединить": 0, "присоединить-обновить": 3}"""
                    tab = rm.rastr.Tables(self.tables[index])
                    tab.ReadCSV(self.calc, self.file_csv[index], self.param[index], ";", '')
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

    def __init__(self, task):  # добавить листы и первая строка с названиями
        self.sheet_couple = {}  # имя листа_log: имя листа_сводная
        self.name_xl_file = ''  # Имя файла EXCEL для сохранения
        self.excel = None
        self.wbook = None
        self.sheets = {}  # Для хранения ссылок на листы excel {'имя листа': ссылка}
        self.task = task
        self.list_name = ["name_rg2", "год", "лет/зим", "макс/мин", "доп_имя1", "доп_имя2", "доп_имя3"]
        self.book = Workbook()
        #  Создать лист xl и присвоить ссылку на него
        for key in self.task['set_printXL']:
            if self.task['set_printXL'][key]['add']:
                self.sheets[key] = self.book.create_sheet(key + "_log")
                # Записать первую строку параметров.
                header_list = self.list_name + self.task['set_printXL'][key]['par'].split(',')
                self.sheets[key].append(header_list)

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
        if rm.name_standard == "не стандарт" or not rm.DopNameStr:
            dop_name_list = ['-'] * 3
        else:
            dop_name_list = rm.DopName[:3]
            if len(dop_name_list) < 3:
                dop_name_list += ['-'] * (3 - len(dop_name_list))
        self.list_name_z = [rm.name_base, rm.god, rm.name_list[1], rm.name_list[2]] + dop_name_list

        self.add_val_table(rm)

        if self.task['print_parameters']['add']:
            self.add_val_parameters(rm.rastr, sel=self.task['print_parameters']['sel'])

        if self.task['print_balance_q']['add']:
            self.add_val_balance_q(rm)

    def add_val_table(self, rm):
        rastr = rm.rastr
        for key in self.task['set_printXL']:
            if not self.task['set_printXL'][key]['add']:
                continue
            # проверка наличия таблицы
            if rastr.Tables.Find(self.task['set_printXL'][key]['tabl']) < 0:
                raise ValueError("В RastrWin не загружена таблица: " + self.task['set_printXL'][key]['tabl'])

            # принт данных из растр в таблицу для СВОДНОЙ
            r_table = rastr.tables(self.task['set_printXL'][key]['tabl'])
            param_list = self.task['set_printXL'][key]['par'].split(',')
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
        ndx = area.FindNextSel(-1)

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
        address_nagruz = self.sheet_q.cell(self.row_q['row_sum_port_Q'], column,
                                           f"={address_qn}+{address_losses}+{address_SHR}").coordinate
        address_sum_gen = self.sheet_q.cell(self.row_q['row_sum_QG'], column,
                                            f"={address_qg}+{address_shq_line}+{address_skrm_gen}").coordinate
        self.sheet_q.cell(self.row_q['row_Q_itog'], column,
                          f"=-{address_nagruz}+{address_sum_gen}")
        self.sheet_q.cell(self.row_q['row_Q_itog_gmin'], column,
                          f"=-{address_nagruz}+{address_qg_min}+{address_shq_line}")
        self.sheet_q.cell(self.row_q['row_Q_itog_gmax'], column,
                          f"=-{address_nagruz}+{address_qg_max}+{address_shq_line}")

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
        :param sheet: объект лист excel
        :param sheet_name: Имя таблицы.
        """
        tab = Table(displayName=sheet_name, ref='A1:' + get_column_letter(sheet.max_column) + str(sheet.max_row))
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=True)
        tab.tableStyleInfo = style
        sheet.add_table(tab)

    def create_pivot(self):
        """
        Открыть excel через win32com.client и создать сводные.
        :return:
        """
        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.ScreenUpdating = False  # обновление экрана
        # self.excel.Calculation = -4135  # xlCalculationManual
        self.excel.EnableEvents = False  # отслеживание событий
        self.excel.StatusBar = False  # отображение информации в строке статуса excel
        try:
            self.wbook = self.excel.Workbooks.Open(self.name_xl_file)
        except Exception:
            raise Exception(f'Ошибка при открытии файла {self.name_xl_file=}')

        for n in range(self.wbook.sheets.count):
            if self.wbook.sheets[n].Name in self.sheet_couple:
                self.pivot_tables(self.wbook.sheets[n].Name, self.sheet_couple[self.wbook.sheets[n].Name])
        if self.task['folder_result']:
            self.wbook.Save()
        self.excel.Visible = True
        self.excel.ScreenUpdating = True  # обновление экрана
        self.excel.Calculation = -4105  # xlCalculationAutomatic
        self.excel.EnableEvents = True  # отслеживание событий
        self.excel.StatusBar = True  # отображение информации в строке статуса excel

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
    """Функция из строки "2021,2023-2025" делает [2021,2023,2024,2025]"""
    years_list = id_str.replace(" ", "").split(',')
    if years_list != "":
        years_list_new = np.array([], int)
        for it in years_list:
            if "-" in it:
                i_years = it.split('-')
                years_list_new = np.hstack(
                    [years_list_new, np.array(np.arange(int(i_years[0]), int(i_years[1]) + 1), int)])
            else:
                years_list_new = np.hstack([years_list_new, int(it)])
        return np.sort(years_list_new)
    else:
        return []


def start_calc():
    """Запуск расчета моделей"""
    pass


def block_b(rm):
    rm.sel0('block_b')
    rm.rgm("block_b")


def import_model():
    """ ИД для импорта из модели(выполняется после блока начала)"""
    ifm = ImportFromModel(import_file_name=r"H:\ОЭС Урала без ТЭ\Пермская ЭС\КПР ПЭ 2021\ТКЗ\импорт перспективы.rg2",
                          criterion_start={"years": "",
                                           "season": "",
                                           "max_min": "",
                                           "add_name": ""},
                          tables="vetv;node",
                          param="",
                          sel="sel",
                          calc=3)
    ImportFromModel.ui_import_model.append(ifm)


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
    VISUAL_CHOICE = 1  # 1 задание через QT, 0 - в коде
    calc_set = 1  # 1 -Изменить модели, 2-Расчет установившихся режимов, 3-Расчет токов КЗ
    cm = None  # глобальный объект класса CorModel
    sys.excepthook = my_except_hook(sys.excepthook)

    # DEBUG, INFO, WARNING, ERROR и CRITICAL
    # logging.basicConfig(filename="log_file.log", level=logging.DEBUG, filemode='w',
    #                     format='%(asctime)s %(name)s  %(levelname)s:%(message)s')

    log = logging.getLogger(__name__)
    log.setLevel(logging.DEBUG)
    formatter = logging.Formatter('%(asctime)s %(name)s %(levelname)s:%(message)s')

    file_handler = logging.FileHandler(filename='log_file.log', mode='w')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    # file_handler.close()
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG)
    console_handler.setFormatter(formatter)

    log.addHandler(file_handler)
    log.addHandler(console_handler)

    if VISUAL_CHOICE:  # в коде
        app = QtWidgets.QApplication([])  # Новый экземпляр QApplication
        # app.setApplicationName("Правка моделей RastrWin")

        gui_choice_window = MainChoiceWindow()
        gui_choice_window.show()
        gui_calc_ur = CalcURWindow()
        # gui_calc_ur.show()
        gui_calc_ur_set = CalcURSetWindow()
        # gui_calc_ur_set.show()
        gui_edit = EditWindow()
        # gui_edit.show()
        gui_set = SetWindow()
        # gui_set.show()
        sys.exit(app.exec_())  # Запуск

    else:
        if calc_set == 1:
            cor_task = {
                # в KIzFolder абсолютный путь к папке с файлами или файлу
                "KIzFolder": r"I:\rastr_add\test",
                # KInFolder папка в которую сохранять измененные файлы(или файл), "" не сохранять
                # результаты работы программы (.xlsx) сохраняются в папку KInFolder, если ее нет то в KIzFolder
                "KInFolder": r"I:\rastr_add\test\test_result",
                # ФИЛЬТР ФАЙЛОВ: False все файлы, True в соответствии с фильтром---------------------------------------
                "KFilter_file": False,
                "max_file_count": 999,  # максимальное количество расчетных файлов
                # нр("2019,2021-2027","зим","мин","1°C;МДП") (год, зим, макс, доп имя разделитель , или ;)
                "cor_criterion_start": {"years": "",
                                        "season": "",
                                        "max_min": "",
                                        "add_name": "", },
                # Корректировка в начале
                "cor_beginning_qt": {'add': False,
                                     'txt': ''},
                # импорт по excel------------------------------------------------------
                "import_val_XL": False,
                "excel_cor_file": r"I:\rastr_add\test\пример задания.xlsx",
                "excel_cor_sheet": "*",  # листы [импорт из моделей][XL->RastrWin], если'*', то все листы по порядку
                # Корректировка в конце
                "cor_end_qt": {'add': False,
                               'txt': ''},
                # Исправить пробелы, заменить английские буквы на русские.
                "cor_name": False,
                "cor_name_task": 'node:name,dname vetv:dname Generator:Name',
                # ----------------------------------------------------------------------------------------------------
                # TODO import или export из xl в растр
                # export_xl [table:Generator; parametrs:... ]
                # import_xl [table:Generator; xl:C:\Users\User\Desktop\1.xlsx list:Generator tip:1]
                # "XL_table": [r"C:\Users\User\Desktop\1.xlsx", "Generator"],  # полный адрес и имя листа
                # "tip_export_xl": 1,  # 1 загрузить, 0 присоединить 2 обновить
                # Проверка параметров режима---------------------------------------------------------------------------
                "control_rg2": True,
                "control_rg2_task": {'node': False, 'vetv': True, 'Gen': False, 'section': False, 'area': False,
                                     'area2': False, 'darea': False, 'sel_node': "na>0"},
                # выводить данные из моделей в XL---------------------------------------------------------------------
                "printXL": False,
                "set_printXL": {
                    "sechen": {'add': False, "sel": "", 'tabl': "sechen", 'par': "ns,name,pmin,pmax,psech",
                               "rows": "ns,name",  # поля строк в сводной
                               "columns": "год,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля столбцов в сводной
                               "values": "psech,pmax"},
                    "area": {'add': False, "sel": "", 'tabl': "area",
                             'par': 'na,name,no,pg,pn,pn_sum,dp,pop,pop_zad,qn_sum,pn_max,pn_min,vnq,vnp,poq,qn,qg',
                             "rows": "na,name,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля строк в сводной
                             "columns": "год",  # поля столбцов в сводной
                             "values": "pop,pg"},
                    "area2": {'add': False, "sel": "", 'tabl': "area2",
                              'par': 'npa,name,pg,pn,dp,pop,vnp,qg,qn,dq,poq,vnq,pn_sum,qn_sum,pop_zad',
                              "rows": "npa,name,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля строк в сводной
                              "columns": "год",  # поля столбцов в сводной
                              "values": "pop,pg"},
                    "darea": {'add': False, "sel": "", 'tabl': "darea",
                              'par': 'no,name,pg,pp,pvn,qn_sum,pnr_sum,pn_sum,pop_zad,qvn,qp,qg',
                              "rows": "no,name,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля строк в сводной
                              "columns": "год",  # поля столбцов в сводной
                              "values": "pp,pg"},
                    # Из любой таблицы растр, нр "Generator" ,"P,Pmax" или "" все параметры, "Num>0" выборка)
                    "tab": {'add': False, "sel": "Num>0", 'tabl': "Generator",
                            'par': "Num,Name,sta,Node,P,Pmax,Pmin,value",
                            "rows": "Num,Name",  # поля строк в сводной
                            "columns": "год,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля столбцов в сводной
                            "values": "P,Pmax"}},  # поля значений в сводной

                # вывод заданных параметров в следующем формате "v=15105,15113,0|15038,15037,4/r|x|b; n=15198/pg|qg"
                # таблица: n-node,v-vetv,g-Generator,na-area,npa-area2,no-darea,nga-ngroup,ns-sechen
                "print_parameters": {'add': False,
                                     "sel": "v=15105,15113,0|15038,15037,4/r|x|b; n=151980/pg|qg"},
                # БАЛАНС Q ! 0 тоже район, даже если в районах не задан "na>13&na<201"
                "print_balance_q": {'add': True, "sel": "na=11"},
                # ----------------------------------------------------------------------------------------------------
                "block_import": False,  # начало
            }
            """Запуск корректировки моделей"""
            cm = CorModel(cor_task)
            cm.run_cor()
        if calc_set == 2:
            start_calc()  # calc

# TODO дописать: перенос параметров из одноименных файлов
# TODO дописать: сравнение файлов
# TODO : в изм разделитель , а в добавить . !!!??????
