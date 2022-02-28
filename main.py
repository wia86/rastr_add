# pip freeze > requirements.txt
# pip install -r requirements.txt
# I:\rastr_add> pyinstaller --onefile --noconsole main.py
# I:\rastr_add> pyinstaller -F --noconsole main.py
import win32com.client
from abc import ABC, abstractmethod
from Rastr_Method import RastrMethod
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
# import pandas as pd
from typing import Union  # Any
import sys
import shutil
from PyQt5 import QtWidgets
from datetime import datetime
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
from qt_cor import Ui_MainCor  # импорт ui: pyuic5 qt_cor.ui -o qt_cor.py
from qt_set import Ui_Settings  # импорт ui: pyuic5 qt_set.ui -o qt_set.py


class Window:
    @staticmethod
    def check_status(set_checkbox_element):
        """
        По состоянию CheckBox изменить состояние видимости соответствующего элемента.
        :param set_checkbox_element: картеж картежей (checkbox, element)
        """
        for CB, element in set_checkbox_element:
            if CB.isChecked():
                element.show()
            else:
                element.hide()

    def choose_file(self, directory: str, filter_: str = ''):
        """
        Выбор файла
        """
        fileName_choose, _ = QtWidgets.QFileDialog.getOpenFileName(self, caption="Выбрать файл", directory=directory,
                                                                   filter=filter_, )  # "All Files(*);Text Files(*.txt)"
        if fileName_choose:
            logging.info(f"Выбран файл: {fileName_choose}, {_}")
            return fileName_choose

    def save_file(self, directory: str, filter_: str = ''):
        """
        Сохранение файла
        """
        fileName_choose, _ = QtWidgets.QFileDialog.getSaveFileName(self, caption="Сохранение файла",
                                                                   directory=directory, filter=filter_)
        if fileName_choose:
            logging.info(f"Для сохранения выбран файл: {fileName_choose}, {_}")
            return fileName_choose


class SetWindow(QtWidgets.QMainWindow, Ui_Settings):
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
                logging.error('файл settings.ini не читается, перезаписан')
                self.save_ini()
        else:
            logging.info('создан файл settings.ini')
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


class EditWindow(QtWidgets.QMainWindow, Ui_MainCor, Window):
    def __init__(self):
        super(EditWindow, self).__init__()  # *args, **kwargs
        self.task_ui = {}
        self.setupUi(self)
        self.check_import = (
            (self.CB_N, 'узлы'),
            (self.CB_V, 'ветви'),
            (self.CB_G, 'генераторы'),
            (self.CB_A, 'районы'),
            (self.CB_A2, 'территории'),
            (self.CB_D, 'объединения'),
            (self.CB_PQ, 'PQ'),
            (self.CB_IT, 'I(T)'),
        )

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
            (self.CB_print_balance_Q, self.balance_Q_vibor),
        )
        # Скрыть параметры при старте.
        self.check_status(self.check_status_visibility)
        # CB показать / скрыть параметры.
        for CB, element in self.check_status_visibility:
            CB.clicked.connect(lambda: self.check_status(self.check_status_visibility))
        # CB показать список импортируемых моделей.
        for CB, _ in self.check_import:
            CB.clicked.connect(lambda: self.import_name_table())
        # Функциональные кнопки
        self.task_save.clicked.connect(lambda: self.task_save_yaml())
        self.task_load.clicked.connect(lambda: self.task_load_yaml())
        self.run_krg2.clicked.connect(lambda: self.gui_start())
        self.SetBut.clicked.connect(lambda: ui_set.show())
        # Подсказки
        # self.CB_KFilter_file.setToolTip("Всплывающее окно")

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
        name_file_save = self.save_file(directory=self.path_beginnings(), filter_="YAML Files (*.yaml)")
        if name_file_save:
            self.fill_task_ui()
            with open(name_file_save, 'w') as f:
                yaml.dump(data=self.task_ui, stream=f, default_flow_style=False, sort_keys=False)

    def path_beginnings(self):
        if self.T_InFolder.toPlainText():
            return self.T_InFolder.toPlainText()
        else:
            return self.T_IzFolder.toPlainText()

    def task_load_yaml(self):
        name_file_load = self.choose_file(directory=self.path_beginnings(), filter_="YAML Files (*.yaml)")
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
        self.Filtr_god_N.setText(dict_["years"])
        self.Filtr_sez_N.setEditText(dict_["season"])
        self.Filtr_max_min_N.setEditText(dict_["max_min"])
        self.Filtr_dop_name_N.setText(dict_["add_name"])
        self.tab_N.setText(dict_['tables'])
        self.param_N.setText(dict_['param'])
        self.sel_N.setText(dict_['sel'])
        self.tip_N.setEditText(dict_['calc'])

        dict_ = task_yaml['Imp_add']['vetv']
        self.CB_V.setChecked(dict_['add'])
        self.file_V.setText(dict_['import_file_name'])
        self.Filtr_god_V.setText(dict_["years"])
        self.Filtr_sez_V.setEditText(dict_["season"])
        self.Filtr_max_min_V.setEditText(dict_["max_min"])
        self.Filtr_dop_name_V.setText(dict_["add_name"])
        self.tab_V.setText(dict_['tables'])
        self.param_V.setText(dict_['param'])
        self.sel_V.setText(dict_['sel'])
        self.tip_V.setEditText(dict_['calc'])

        dict_ = task_yaml['Imp_add']['gen']
        self.CB_G.setChecked(dict_['add'])
        self.file_G.setText(dict_['import_file_name'])
        self.Filtr_god_G.setText(dict_["years"])
        self.Filtr_sez_G.setEditText(dict_["season"])
        self.Filtr_max_min_G.setEditText(dict_["max_min"])
        self.Filtr_dop_name_G.setText(dict_["add_name"])
        self.tab_G.setText(dict_['tables'])
        self.param_G.setText(dict_['param'])
        self.sel_G.setText(dict_['sel'])
        self.tip_G.setEditText(dict_['calc'])

        dict_ = task_yaml['Imp_add']['area']
        self.CB_A.setChecked(dict_['add'])
        self.file_A.setText(dict_['import_file_name'])
        self.Filtr_god_A.setText(dict_["years"])
        self.Filtr_sez_A.setEditText(dict_["season"])
        self.Filtr_max_min_A.setEditText(dict_["max_min"])
        self.Filtr_dop_name_A.setText(dict_["add_name"])
        self.tab_A.setText(dict_['tables'])
        self.param_A.setText(dict_['param'])
        self.sel_A.setText(dict_['sel'])
        self.tip_A.setEditText(dict_['calc'])

        dict_ = task_yaml['Imp_add']['area2']
        self.CB_A2.setChecked(dict_['add'])
        self.file_A2.setText(dict_['import_file_name'])
        self.Filtr_god_A2.setText(dict_["years"])
        self.Filtr_sez_A2.setEditText(dict_["season"])
        self.Filtr_max_min_A2.setEditText(dict_["max_min"])
        self.Filtr_dop_name_A2.setText(dict_["add_name"])
        self.tab_A2.setText(dict_['tables'])
        self.param_A2.setText(dict_['param'])
        self.sel_A2.setText(dict_['sel'])
        self.tip_A2.setEditText(dict_['calc'])

        dict_ = task_yaml['Imp_add']['darea']
        self.CB_D.setChecked(dict_['add'])
        self.file_D.setText(dict_['import_file_name'])
        self.Filtr_god_D.setText(dict_["years"])
        self.Filtr_sez_D.setEditText(dict_["season"])
        self.Filtr_max_min_D.setEditText(dict_["max_min"])
        self.Filtr_dop_name_D.setText(dict_["add_name"])
        self.tab_D.setText(dict_['tables'])
        self.param_D.setText(dict_['param'])
        self.sel_D.setText(dict_['sel'])
        self.tip_D.setEditText(dict_['calc'])

        dict_ = task_yaml['Imp_add']['PQ']
        self.CB_PQ.setChecked(dict_['add'])
        self.file_PQ.setText(dict_['import_file_name'])
        self.Filtr_god_PQ.setText(dict_["years"])
        self.Filtr_sez_PQ.setEditText(dict_["season"])
        self.Filtr_max_min_PQ.setEditText(dict_["max_min"])
        self.Filtr_dop_name_PQ.setText(dict_["add_name"])
        self.tab_PQ.setText(dict_['tables'])
        self.param_PQ.setText(dict_['param'])
        self.sel_PQ.setText(dict_['sel'])
        self.tip_PQ.setEditText(dict_['calc'])

        dict_ = task_yaml['Imp_add']['IT']
        self.CB_IT.setChecked(dict_['add'])
        self.file_IT.setText(dict_['import_file_name'])
        self.Filtr_god_IT.setText(dict_["years"])
        self.Filtr_sez_IT.setEditText(dict_["season"])
        self.Filtr_max_min_IT.setEditText(dict_["max_min"])
        self.Filtr_dop_name_IT.setText(dict_["add_name"])
        self.tab_IT.setText(dict_['tables'])
        self.param_IT.setText(dict_['param'])
        self.sel_IT.setText(dict_['sel'])
        self.tip_IT.setEditText(dict_['calc'])

        self.check_status(self.check_status_visibility)

    def gui_start(self):
        """
        Добавить ImportFromModel и запуск start_cor
        """
        self.record_task_ui()
        # Импорт параметров режима
        if self.CB_ImpRg2.isChecked():
            if self.task_ui['CB_ImpRg2']:
                for tables in self.task_ui['Imp_add']:
                    if self.task_ui['Imp_add'][tables]['add']:
                        ifm = ImportFromModel(import_file_name=self.task_ui['Imp_add'][tables]['import_file_name'],
                                              criterion_start={"years": self.task_ui['Imp_add'][tables]['years'],
                                                               "season": self.task_ui['Imp_add'][tables]['season'],
                                                               "max_min": self.task_ui['Imp_add'][tables]['max_min'],
                                                               "add_name": self.task_ui['Imp_add'][tables]['add_name']},
                                              tables=self.task_ui['Imp_add'][tables]['tables'],
                                              param=self.task_ui['Imp_add'][tables]['param'],
                                              sel=self.task_ui['Imp_add'][tables]['sel'],
                                              calc=self.task_ui['Imp_add'][tables]['calc'])
                        ImportFromModel.ui_import_model.append(ifm)
        # Убрать 'file:///'
        for str_name in ["KIzFolder", "KInFolder", "excel_cor_file"]:
            self.task_ui[str_name].lstrip('file:///')
        start_cor(self.task_ui)

    def fill_task_ui(self):
        """
        Заполнить task_ui задание (task_ui).
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
    Для хранения общих настроек
    """
    # коллекция настроек, которые хранятся в ini файле
    set_save = {}
    # коллекция для хранения информации о расчете
    set_info = {"calc_val": {1: "ЗАМЕНИТЬ", 2: "ПРИБАВИТЬ", 3: "ВЫЧЕСТЬ", 0: "УМНОЖИТЬ"},
                'collapse': '',
                'end_info': ''}

    # @abstractmethod
    def __init__(self):
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
        time_spent = round(time() - self.time_start, 2)
        time_spent_minutes = round(time_spent / 60, 1)
        self.set_info['end_info'] = (
                f"РАСЧЕТ ЗАКОНЧЕН! \nНачало расчета {self.now_start}, конец {self.now.strftime('%d-%m-%Y %H-%M')}" +
                f" \nЗатрачено: {str(time_spent)} секунд или {str(time_spent_minutes)} минут")
        logging.info(self.set_info['end_info'])


class CorModel(GeneralSettings):
    """
    Коррекция файлов
    """

    def __init__(self, task):
        super(CorModel, self).__init__()
        self.pxl = None
        self.cor_xl = None
        self.task = task
        self.rastr_files = None

    def run_cor(self):
        """Запуск корректировки моделей"""
        # определяем корректировать файл или файлы в папке по анализу "KIzFolder"
        if 'KIzFolder' not in self.task:
            logging.error('В задании отсутствует папка для корректировки')
            return False

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
                    logging.info("Создана папка: " + self.task["KInFolder"])
                    os.mkdir(self.task["KInFolder"])
            folder_save = self.task["KInFolder"] if self.task["KInFolder"] else self.task["KIzFolder"]
        else:
            self.task["KInFolder"] = ''
            folder_save = self.task["KIzFolder"]

        self.task['folder_result'] = folder_save + r"\result"  # папка для сохранения результатов
        now = datetime.now()
        self.task['name_time'] = f"{self.task['folder_result']}\\{now.strftime('%d-%m-%Y %H-%M')}"
        if not os.path.exists(self.task['folder_result']):
            os.mkdir(self.task['folder_result'])  # создать папку result
        self.task['folder_temp'] = self.task['folder_result'] + r"\temp"  # папка для сохранения рабочих файлов
        if not os.path.exists(self.task['folder_temp']):
            os.mkdir(self.task['folder_temp'])  # создать папку temp

        # ЭКСПОРТ ИЗ МОДЕЛЕЙ
        if 'block_import' in self.task:
            if self.task['block_import']:
                import_model()  # ИД для импорта

        if "import_val_XL" in self.task:
            if self.task["import_val_XL"]:  # задать параметры узла по значениям в таблице excel (имя книги, имя листа)
                self.cor_xl = CorXL(excel_file_name=self.task["excel_cor_file"], sheets=self.task["excel_cor_sheet"])
                self.cor_xl.init_export_model()

        load_add = []
        if "printXL" in self.task:
            if ((self.task["printXL"] and self.task["set_printXL"]["sechen"]) or
                    (self.task["control_rg2"] and self.task["control_rg2_task"]["section"])):
                load_add.append('sch')

        if self.task["folder_file"] == 'folder':  # корр файлы в папке
            files = os.listdir(self.task["KIzFolder"])  # список всех файлов в папке
            self.rastr_files = list(filter(lambda x: x.endswith('.rg2') | x.endswith('.rst'), files))

            for rastr_file in self.rastr_files:  # цикл по файлам .rg2 .rst в папке KIzFolder
                full_name = self.task["KIzFolder"] + '\\' + rastr_file
                full_name_new = self.task["KInFolder"] + '\\' + rastr_file
                rm = RastrModel(full_name)
                # если включен фильтр файлов и имя стандартизовано
                if self.task["KFilter_file"] and rm.kod_name_rg2:
                    if not rm.test_name(condition=self.task["cor_criterion_start"], info='Цикл по файлам KIzFolder'):
                        continue  # пропускаем если не соответствует фильтру

                self.file_count += 1
                #  если включен фильтр файлов проверяем количество расчетных файлов
                if self.task["KFilter_file"] and self.file_count == self.task["max_file_count"] + 1:
                    break
                rm.load(load_add=load_add)
                self.cor_file(rm)
                if self.task["KInFolder"]:
                    rm.save(full_name_new)

        elif self.task["folder_file"] == 'file':  # корр файл
            rm = RastrModel(full_name=self.task["KIzFolder"])
            rm.load(load_add=load_add)
            self.cor_file(rm)
            if self.task["KInFolder"]:
                rm.save(self.task["KInFolder"] + '\\' + rm.Name)
        # для нескольких запусков через GUI
        if ImportFromModel.ui_import_model:
            ImportFromModel.ui_import_model = []

        if self.pxl:
            self.pxl.finish()

        if 'collapse' in self.set_info:
            if self.set_info['collapse']:
                self.set_info['end_info'] += f"\nВНИМАНИЕ! развалились модели:\n[{self.set_info['collapse']}]. "

        self.the_end()
        notepad_path = self.task['name_time'] + ' протокол коррекции файлов.log'
        shutil.copyfile('log_file.log', notepad_path)
        with open(self.task['name_time'] + ' задание на корректировку.yaml', 'w') as f:
            yaml.dump(data=self.task, stream=f, default_flow_style=False, sort_keys=False)
        webbrowser.open(notepad_path)
        mb.showinfo("Инфо", self.set_info['end_info'])

    def cor_file(self, rm):
        """Корректировать файл rm"""
        try:
            if self.task['cor_beginning_qt']['add']:
                logging.info("\t*** Начало корректировку модели 'до импорта' ***")
                rm.cor_rm_from_txt(self.task['cor_beginning_qt']['txt'])
                logging.info("\t*** Конец выполнения корректировки моделей 'до импорта' ***")
        except KeyError:
            pass

        if 'block_beginning' in self.task:
            if self.task['block_beginning']:
                logging.info("\t***Блок начала ***")
                block_b(rm)
                logging.info("\t*** Конец блока начала ***")
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
                logging.info("\t*** Начало корректировку модели 'после импорта' ***")
                rm.cor_rm_from_txt(self.task['cor_end_qt']['txt'])
                logging.info("\t*** Конец выполнения корректировки моделей 'после импорта' ***")
        except KeyError:
            pass

        if 'block_end' in self.task:
            if self.task['block_end']:
                logging.info("\t*** Блок конца ***")
                block_e(rm)
                logging.info("\t*** Конец блока конца ***")
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
                if not type(self.pxl) == PrintXL:
                    self.pxl = PrintXL(self.task)
                self.pxl.add_val(rm)


"""<<<<<<<<<<<<<<<<<<<<СПРАВКА>>>>>>>>>>>>>>>>>>>>>>>>>
# <<<УДАЛИТЬ>>>
#  Del_sel ()  #  удалить отмеченные узлы (c ветвями) ветви и генераторы
#  Del(tabl,viborka)  # viborka = "net" - удалить узлы или ветви без связей или без узла начала конца
# <<<ИЗМЕНИТЬ СЕТЬ>>>
#  uhom_korr_sub (set_sel) #  исправить номинальные напряжения в узлах
# sta_node ("str_ny", on_off)#  узлы с ветвями (СТРОКА номера узлов через пробел) включить False; отключить True

# rastr.RenumWP=True     # включить ссылки, отключить
# vzd0 ()           #  поиск узлов где напряжение vzd задано а диапазона реактивки нет и удаляет vzd
# name0 ()           #  поиск узлов и генераторов без имени
# nyNum0 ()           #  поиск узлов и генераторов с номером 0
# <<<прочее>>>
# = otklonenie_seshen (nomer_sesh)   #   возвращает величину отклонения psech от  pmax   + превышение; - недобор
# = rastr.Calc("sum,max,min,val","area","qn","vibor") - функция (vibor не может быть "")
#  ГЕНЕРАТОРЫ  PGen_cor ("sel")  # если мощность P больше Pmax то изменить мощность генератора  на Pmax, если P меньше
#  Pmin но больше 0 - то на Pmin #  если P ген = 0 то отключить генератор, чтоб реактивка не выдавалась
СЕЧЕНИЕ # KorSech  (ns,newp,vibor , tip, net_Pmin_zad) #  номер сеч, новая мощность в сеч (значение или "max" "min"),
#  выбор корр узлов  (нр "sel"или "" - авто) ,  tip - "pn" или "pg", net_Pmin_zad #  1 не учитывать Pmin
#  Qgen_node_in_gen_sub ()  #  посчитать Q ГЕН по  Q в узле
# <<<настройки rastr>>>
#  rastr.tables("com_regim").cols.item("gen_p").Z(0) = 0
#  #0- "да"; 1- "да"; 2- только Р; 3- только Q ///it_max  количество расчетов///neb_p точность расчектов////
"""


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
        self.txt_dop = ""
        self.degree_int = 0
        self.degree_str = ""
        self.loadRGM = False
        self.DopNameStr = ""
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

        # if CALC_SET == 2:  # расчет режимов
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
                    logging.debug(info + self.Name + f" Год '{self.god}' не проходит по условию: "
                                  + str(condition['years']))
                    return False

        if 'season' in condition:
            if condition['season']:
                if condition['season'].strip():  # ПРОВЕРКА "зим" "лет" "паводок"
                    fff = False
                    temp = condition['season'].replace(' ', '')
                    for us in temp.split(","):
                        if self.name_list[1] == us:
                            fff = True
                    if not fff:
                        logging.debug(info + self.Name + f" Сезон '{self.name_list[1]}' не проходит по условию: "
                                      + condition['season'])
                        return False

        if 'max_min' in condition:
            if condition['max_min']:
                if condition['max_min'].strip():  # ПРОВЕРКА "макс" "мин"
                    if self.name_list[2] != condition['max_min'].replace(' ', ''):
                        logging.debug(info + self.Name + f" '{self.name_list[2]}' не проходит по условию: "
                                      + condition['max_min'])
                        return False

        if 'add_name' in condition:
            if condition['add_name']:  # ПРОВЕРКА (-41С;МДП:ТЭ-У)
                if condition['add_name'].strip():  # ПРОВЕРКА (-41С;МДП:ТЭ-У)
                    if ";" in condition['add_name']:
                        temp = condition['add_name'].split(";")
                    else:
                        temp = condition['add_name'].split(",")
                    fff = False
                    for us in temp:
                        for DopName_i in self.DopName:
                            if DopName_i == us:
                                fff = True
                    if not fff:
                        logging.debug(
                            info + self.Name + f" Доп. имя {self.DopNameStr} не проходит по условию: " + condition[
                                'add_name'])
                        return False
        return True

    def load(self, rastr=None, load_add: list = None):
        """загрузить модель в Rastr
        load_add=['amt','sch','trn'] расширения файлов которые нужно загрузить
        загружается первый попавшийся файл в папке IzFolder"""
        if rastr:
            self.rastr = rastr
        else:
            # try:
            self.rastr = win32com.client.Dispatch("Astra.Rastr")
            # except NameError:
            #     logging.critical('Com объект Astra.Rastr не найден')

        self.rastr.Load(1, self.full_name, self.pattern)  # загрузить
        logging.info(f"\n\nЗагружен файл: {self.full_name}\n")
        # Загрузить файлы load_add
        if load_add:
            for extension in load_add:
                files = os.listdir(self.dir)
                names = list(filter(lambda x: x.endswith('.' + extension), files))
                if len(names) > 0:
                    self.rastr.Load(1, self.dir + '\\' + names[0], GeneralSettings.set_save["шаблон " + extension])
                    logging.info("Загружен файл: " + names[0])

    def save(self, full_name_new):
        self.rastr.Save(full_name_new, self.pattern)
        logging.info("Файл сохранен: " + full_name_new)

    def control_rg2(self, dict_task):
        """  контроль  dict_task = {'node': True, 'vetv': True, 'Gen': True, 'section': True,
             'area': True, 'area2': True, 'darea': True, 'sel_node': "na>0"}  """
        if not self.rgm("control_rg2"):
            return False

        node = self.rastr.tables("node")
        branch = self.rastr.tables("vetv")
        generator = self.rastr.tables("Generator")
        chart_pq = self.rastr.tables("graphik2")
        graph_it = self.rastr.tables("graphikIT")

        # Напряжения
        if dict_task["node"]:
            logging.info("\tКонтроль напряжений.")
            self.rastr.voltage_nominal(choice=dict_task["sel_node"])
            self.rastr.voltage_normal(choice=dict_task["sel_node"])
            self.rastr.voltage_deviation(choice=dict_task["sel_node"])

        # Токи
        if dict_task['vetv']:
            self.rastr.CalcIdop(self.degree_int, 0.0, "")
            logging.info("\tКонтроль токовой загрузки, расчетная температура: " + self.degree_str)
            if dict_task["sel_node"] != "":
                if node.cols.Find("sel1") < 0:
                    node.Cols.Add("sel1", 3)  # добавить столбцы
                node.cols.item("sel1").calc(0)
                node.setsel(dict_task["sel_node"])
                node.cols.item("sel1").calc(1)
                sel_vetv = "i_zag>=0.1&(ip.sel1|iq.sel1)"
                sel_vetv2 = "(ip.sel1|iq.sel1)&(n_it_av>0|n_it>0)"
            else:
                sel_vetv = "i_zag>=0.1"
                sel_vetv2 = "(n_it_av>0|n_it>0)"

            branch.setsel(sel_vetv)
            if branch.count > 0:  # есть превышений
                j = branch.FindNextSel(-1)
                while j > -1:
                    name = branch.cols.item('name').ZS(j)
                    i_zag = branch.cols.item('i_zag').ZS(j)
                    logging.info(f"\t\tВНИМАНИЕ ТОКИ! vetv:{branch.SelString(j)}, {name} - {i_zag}%")
                    j = branch.FindNextSel(j)

            branch.setsel(sel_vetv2)  # проверка наличия n_it,n_it_av в таблице График_Iдоп_от_Т(graphikIT)
            if branch.count > 0:
                j = branch.FindNextSel(-1)
                while j > -1:
                    if branch.cols.item("n_it").Z(j) > 0:
                        graph_it.setsel("Num=" + branch.cols.item("n_it").ZS(j))
                        if graph_it.count == 0:
                            name = branch.cols.item('name').ZS(j)
                            n_it = branch.cols.item('n_it').ZS(j)
                            logging.info(f"\t\tВНИМАНИЕ graphikIT! vetv: {branch.SelString(j)}, {name}, "
                                         + f"n_it={n_it} не найден в таблице График_Iдоп_от_Т")

                    if branch.cols.item("n_it_av").Z(j) > 0:
                        graph_it.setsel("Num=" + branch.cols.item("n_it_av").ZS(j))
                        if graph_it.count == 0:
                            name = branch.cols.item('name').ZS(j)
                            n_it_av = branch.cols.item('n_it_av').ZS(j)
                            logging.info(f"\t\tВНИМАНИЕ graphikIT! vetv: {branch.SelString(j)}, {name},"
                                         + f" n_it_av={n_it_av} не найден в таблице График_Iдоп_от_Т")
                    j = branch.FindNextSel(j)
        #  ГЕНЕРАТОРЫ
        if dict_task['Gen']:
            logging.info("\tКонтроль генераторов")
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
                if P < Pmin > 0:
                    logging.info(f"\t\tВНИМАНИЕ! {Name}, Num={Num},ny={Node}, P={str(round(P))} < Pmin={str(Pmin)}")
                if P > Pmax > 0:
                    logging.info(f"\t\tВНИМАНИЕ! {Name}, Num={Num},ny={Node}, P={str(round(P))} > Pmax={str(Pmax)}")
                if NumPQ > 0:
                    chart_pq.setsel("Num=" + str(NumPQ))
                    if chart_pq.count == 0:
                        logging.info(f"\t\tВНИМАНИЕ! ГЕНЕРАТОР: {Name}, Num={Num},ny={Node}, "
                                     + f"NumPQ={str(NumPQ)} не найден в таблице PQ-диаграммы (graphik2)")
                j = generator.FindNextSel(j)
        # сечения
        if self.rastr.tables.Find("sechen") > 0:
            section = self.rastr.tables("sechen")
            if dict_task['section']:
                if section.size == 0:
                    logging.error("\tCечения отсутствуют")
                else:
                    logging.info("\tКонтроль сечений")
                    section.setsel("")
                    j = section.FindNextSel(-1)
                    while j != -1:
                        name = section.cols.item("name").ZS(j)
                        ns = section.cols.item("ns").ZS(j)
                        pmax = section.cols.item("pmax").Z(j)
                        psech = section.cols.item("psech").Z(j)
                        if psech > pmax + 0.01:
                            logging.info(f"\t\tВНИМАНИЕ! сечение: {name}({ns}), P: {str(round(psech))}, "
                                         + f"pmax: {str(pmax)}, отклонение:{str(round(pmax - psech))}")
                        j = section.FindNextSel(j)
        else:
            logging.error("\tФайл сечений не загружен")

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

        logging.info("\tКонтроль pop_zad " + key_zone[zone + '_name'])
        tabl = self.rastr.tables(zone)
        if tabl.cols.Find("pop_zad") < 0:
            logging.error("Поле pop_zad отсутствует в таблице " + key_zone[zone + '_name'])
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
                    logging.info(f"\t\tВНИМАНИЕ: {name} ({no}), pop: {str(pp)}, pop_zad: {str(pop_zad)}, "
                                 + f"отклонение: {str(round(pop_zad - pp))} или {str(round(deviation * 100))} %")
                j = tabl.FindNextSel(j)

    def cor_rm_from_txt(self, task_txt: str):
        """
        Корректировать модели по заданию в текстовом формате
        :param task_txt:
        :return:
        """
        task_rows = task_txt.split('\n')
        for task_row in task_rows:
            task_row = task_row.split('#', 1)[0]  # удалить текст после '#'
            # Имя функции стоит перед "(" и "["
            name_fun = task_row.split('(', 1)[0]
            name_fun = name_fun.split('[', 1)[0]
            name_fun = name_fun.replace(' ', '')
            if not name_fun:
                continue

            # Условие выполнения в фигурных скобках
            condition_dict = {}
            match = re.search(re.compile(r"\{(.+?)}"), task_row)
            if match:
                conditions = match[1].split('|')
                for condition in conditions:
                    parameter, value = condition.split('=')
                    condition_dict[parameter.strip()] = value.strip()
            if condition_dict and self.kod_name_rg2:
                if not self.test_name(condition=condition_dict):
                    continue
            # Параметры функции в круглых скобках
            function_parameters = []
            match = re.search(re.compile(r"\[(.+?)]"), task_row)
            if match:
                function_parameters = match[1].split('|')
            function_parameters += ['', '']
            self.txt_task_cor(name=name_fun, sel=function_parameters[0], value=function_parameters[1])


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


class CorSheet:
    """
    Клас лист для хранения листов книги excel и работы с ними
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
        logging.info(f'\tВыполнение задания листа {self.name!r}')
        if self.type == 'import_model':
            self.import_model(rm)
        if self.type == 'list_cor':
            self.list_cor(rm)
        if self.type == 'tab_cor':
            self.tab_cor(rm)
        logging.info(f'\tКонец выполнения задания листа {self.name!r}')

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

    def list_cor(self, rm: RastrModel) -> None:  # cor таблица корректировок по списку, нр изм, удалить, снять отметку
        # номера столбцов
        C_VALUE = 3
        C_SELECTION = 2

        for row in range(3, self.xls.max_row + 1):
            name_fun = self.xls.cell(row, 1).value
            if name_fun:
                if '#' not in name_fun:
                    test_condition = True
                    year = self.xls.cell(row, 4).value
                    season = self.xls.cell(row, 5).value
                    max_min = self.xls.cell(row, 6).value
                    add_name = self.xls.cell(row, 7).value
                    sel = str(self.xls.cell(row, C_SELECTION).value)
                    value = self.xls.cell(row, C_VALUE).value
                    if any([year, season, max_min, add_name]) and rm.kod_name_rg2:  # any если хотя бы один истина
                        if not rm.test_name(condition={"years": year, "season": season,
                                                       "max_min": max_min, "add_name": add_name},
                                            info=f'\t\tcor_x:{sel=}, {value=}'):
                            continue

                    rm.txt_task_cor(name=name_fun, sel=sel,
                                    value=value)

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
                        pattern_name = re.compile("\[(.*)\]\[(.*)\]\[(.*)\]\[(.*)\]")
                        match = re.search(pattern_name, name_file)
                        if match.re.groups == 4:
                            if rm.test_name(condition={"years": match[1], "season": match[2],
                                                       "max_min": match[3], "add_name": match[4]},
                                            info=f"\tcor_xl, условие: {name_file}, ") or not rm.kod_name_rg2:
                                duct_add = True
                if duct_add:
                    dict_param_column[column_name_file] = self.xls.cell(2, column_name_file).value

        if len(dict_param_column) == 0:
            logging.info(f"\t {rm.name_base} НЕ НАЙДЕН на листе {self.name} книги excel")
        else:
            logging.info(f'\t\tРасчетной модели соответствуют столбцы: параметры {dict_param_column}')
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

    def __init__(self, excel_file_name: str, sheets: str) -> None:  # , rm: RastrModel
        """
        Проверить наличие книги и листов, создать классы CorSheet для листов
        :param excel_file_name: полное имя файла excel, нр I:\примеры.xlsx;
        :param sheets: имя листов, нр [импорт из моделей][XL->RastrWin], если '*', то все листы по порядку
        """

        self.sheets_list = []  # для хранения объектов CorSheet
        logging.info(f"Изменить модели по заданию из книги: {excel_file_name}, листы: {sheets}")
        if not os.path.exists(excel_file_name):
            raise ValueError("Ошибка в задании, не найден файл: " + excel_file_name)
        else:
            self.excel_file_name = excel_file_name
            self.wb = load_workbook(excel_file_name, data_only=True)  # data_only - загружать расчетные значения ячеек

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
    calc_str = {"обновить": 2, "загрузить": 1, "присоединить": 0, "присоединить-обновить": 3}
    number = 0  # для создания уникального имени csv файла

    def __init__(self, import_file_name: str = '', criterion_start: Union[dict, None] = None, tables: str = '',
                 param='', sel: Union[str, None] = '', calc: Union[int, str] = '2'):
        """
        Импорт данных из файлов .rg2, .rst и др.
        Создает папку temp в папке с файлом и сохраняет в ней .csv файлы
        import_file_name = полное имя файла
        criterion_start={"years": "","season": "","max_min": "", "add_name": ""} условие выполнения
        tables = таблица для импорта, нр "node;vetv"
        param= параметры для импорта: "" все параметры или перечисление, нр 'sel,sta'(ключи не обязательно)
        sel= выборка нр "sel" или "" - все
        calc= число типа int, строка или ключевое слово:
        {"обновить": 2 , "загрузить": 1, "присоединить": 0, "присоединить-обновить": 3}
        """
        self.import_rm = None
        if not os.path.exists(import_file_name):
            logging.error("Ошибка в задании, не найден файл: " + import_file_name)
            self.import_file_name = ''
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
                    logging.error("Ошибка в задании, не распознано задание calc ImportFromModel: " + str(calc))
                    self.import_file_name = ''
            self.file_csv = []
            ImportFromModel.number += 1
            number = str(ImportFromModel.number)
            for tabl in self.tables:
                self.file_csv.append(f"{self.folder_temp}\\{self.basename}_{tabl}_{number}.csv")
                self.param.append(param)

            # Экспорт данных из файла в .csv файлы в папку temp
            if self.import_file_name:
                logging.info(f'Экспорт из файла <{self.import_file_name}> в CSV')
                self.import_rm = RastrModel(full_name=self.import_file_name)
                self.import_rm.load()
                for index in range(len(self.tables)):
                    if not self.param[index]:  # если все параметры
                        self.param[index] = self.import_rm.all_cols(self.tables[index])
                    else:  # добавить к строке параметров ключи текущей таблицы
                        if self.import_rm.rastr.Tables(self.tables[index]).Key not in self.param[index]:
                            self.param[index] += ',' + self.import_rm.rastr.Tables(self.tables[index]).Key

                    logging.info(f"\n\tТаблица: {self.tables[index]}. Выборка: {self.sel}\n"
                                 + f"\tПараметры: {self.param[index]}\n\tФайл CSV: {self.file_csv[index]}")

                    tab = self.import_rm.rastr.Tables(self.tables[index])
                    tab.setsel(self.sel)
                    tab.WriteCSV(1, self.file_csv[index], self.param[index], ";")  # 0 дописать, 1 заменить

    def import_csv(self, rm: RastrModel) -> None:
        """Импорт данных из csv в файла"""
        if self.import_file_name:
            logging.info(f"\tИмпорт из CSV <{self.import_file_name}> в модель:")
            if rm.test_name(condition=self.criterion_start, info='\tImportFromModel ') or not rm.kod_name_rg2:
                for index in range(len(self.tables)):
                    logging.info(f"\n\tТаблица: {self.tables[index]}. Выборка: {self.sel}. тип: {str(self.calc)}" +
                                 f"\n\tФайл CSV: {self.file_csv[index]}" +
                                 f"\n\tПараметры: {self.param[index]}")
                    """{"обновить": 2 , "загрузить": 1, "присоединить": 0, "присоединить-обновить": 3}"""
                    tab = rm.rastr.Tables(self.tables[index])
                    tab.ReadCSV(self.calc, self.file_csv[index], self.param[index], ";", '')


class PrintXL:
    """Класс печать данных в excel"""
    list_name_z = []
    short_name_tables = {'n': 'node', 'v': 'vetv', 'g': 'Generator', 'na': 'area', 'npa': 'area2', 'no': 'darea',
                         'nga': 'ngroup', 'ns': 'sechen'}

    #  ...._log  лист протокол для сводной

    def __init__(self, task):  # добавить листы и первая строка с названиями
        self.sheet_couple = {}  # имя листа_log: имя листа_сводная
        self.name_xl_file = ''  # Имя файла EXCEL для сохранения
        self.excel = None
        self.wbook = None
        self.task = task
        self.list_name = ["name_rg2", "год", "лет/зим", "макс/мин", "доп_имя1", "доп_имя2", "доп_имя3"]
        self.book = Workbook()
        #  создать лист xl и присвоить ссылку на него
        for key in self.task['set_printXL']:
            if self.task['set_printXL'][key]['add']:
                self.task['set_printXL'][key]["sheet"] = self.book.create_sheet(key + "_log")
                # записать первую строку параметров
                header_list = self.list_name + self.task['set_printXL'][key]['par'].split(',')
                self.task['set_printXL'][key]["sheet"].append(header_list)

        if self.task['print_parameters']['add']:
            self.task['print_parameters']["sheet"] = self.book.create_sheet('parameters')

        if self.task['print_balance_q']['add']:
            self.task['print_balance_q']["sheet"] = self.book.create_sheet("balance_Q")
            self.balance_q_x0 = 5

    def add_val(self, rm: RastrModel):

        logging.info("\tВывод данных из моделей в XL")
        if rm.name_standard == "не стандарт":
            dop_name_list = ['-'] * 3
        else:
            dop_name_list = rm.DopName[:3]
            if len(dop_name_list) < 3:
                dop_name_list += ['-'] * (3 - len(dop_name_list))
        self.list_name_z = [rm.name_base, rm.god, rm.name_list[1], rm.name_list[2]] + dop_name_list

        self.add_val_table(rm)

        if self.task['print_parameters']['add']:
            self.add_val_parameters(rm.rastr)

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
            sheet = self.task['set_printXL'][key]["sheet"]
            param_list = self.task['set_printXL'][key]['par'].split(',')
            param_list = [param_list[i] if r_table.cols.Find(param_list[i]) > -1 else '-' for i in
                          range(len(param_list))]

            setsel = self.task['set_printXL'][key]['sel'] if self.task['set_printXL'][key]['sel'] else ""
            r_table.setsel(setsel)
            index = r_table.FindNextSel(-1)
            while index >= 0:
                sheet.append(
                    self.list_name_z + [r_table.cols.item(val).ZN(index) if val != '-' else '-' for val in param_list])
                index = r_table.FindNextSel(index)

    def add_val_parameters(self, rastr):
        """
        Вывод заданных параметров в формате: "v=42,48,0|43,49,0|27,11,3/r|x|b; n=8|6/pg|qg|pn|qn".
        Таблица: n-node,v-vetv,g-Generator,na-area,npa-area2,no-darea,nga-ngroup,ns-sechen.
        :param rastr:
        """
        sheet = self.task['print_parameters']["sheet"]
        one_row_list = None
        if sheet.max_row == 1:
            one_row_list = self.list_name[:]
        val_list = self.list_name_z[:]

        for task_i in self.task['print_parameters']['sel'].replace(' ', '').split(';'):
            key_row, key_column = task_i.split("/")  # нр"ny=8|9", "pn|qn"
            key_column = key_column.split('|')  # ['pn','qn']
            key_row = key_row.split('=')  # ['n','8|9']
            set_key_row = key_row[1].split('|')  # ['8','9']
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
        pass

    def finish(self):
        """
        Преобразовать в объект таблицу и удалить листы с одной строкой.
        """
        for sheet_name in self.book.sheetnames:
            sheet = self.book[sheet_name]
            if sheet.max_row == 1:
                del self.book[sheet_name]  # удалить пустой лист
            else:
                # Создать объект таблица.
                tab = Table(displayName=sheet_name,
                            ref='A1:' + get_column_letter(sheet.max_column) + str(sheet.max_row))
                style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                       showLastColumn=False, showRowStripes=True, showColumnStripes=True)
                tab.tableStyleInfo = style
                sheet.add_table(tab)
                if 'log' in sheet_name:
                    self.book.create_sheet(sheet_name.replace('log', 'сводная'))
                    self.sheet_couple[sheet_name] = sheet_name.replace('log', 'сводная')
        self.name_xl_file = self.task['name_time'] + ' вывод данных.xlsx'
        self.book.save(self.name_xl_file)
        self.book = None
        for key in self.task['set_printXL']:
            if self.task['set_printXL'][key]['add']:
                self.create_pivot()
                break

        if self.task['print_balance_q']['add']:
            self.configure_balance_q()

    def create_pivot(self):
        # Открыть win32com.client для создания сводных.
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
        pass
        # if self.print_balance_Q:
        #     XL_print_balance_Q.Columns(4).ColumnWidth = 33
        #     diapozon = XL_print_balance_Q.UsedRange.address
        #     With
        #     XL_print_balance_Q.Range(diapozon)
        #     .HorizontalAlignment = -4108  # выравнивание по центру
        #     .VerticalAlignment = -4108
        #     .NumberFormat = "0"
        # diapozon = XL_print_balance_Q.UsedRange.address
        # XL_print_balance_Q.Range(diapozon)
        # .Borders(7).LineStyle = 1  # лево
        # .Borders(8).LineStyle = 1  # верх
        # .Borders(9).LineStyle = 1  # низ
        # .Borders(10).LineStyle = 1  # право
        # .Borders(11).LineStyle = 1  # внутри вертикаль
        # .Borders(12).LineStyle = 1  #
        # .WrapText = True  # перенос текста в ячейке
        # .Font.Name = "Times  Roman"

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


def start_cor(cor_task_current: dict):
    """Запуск корректировки моделей"""
    global cm
    cm = CorModel(cor_task_current)
    cm.run_cor()


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


if __name__ == '__main__':
    VISUAL_CHOICE = 1  # 1 задание через QT, 0 - в коде
    CALC_SET = 1  # 1 -корректировать модели CorModel, 2-рассчитать модели
    cm = None  # глобальный объект класса CorModel
    # https://docs.python.org/3/library/logging.html
    # 'w' - перезаписать лог, иначе будет добавляться
    logging.basicConfig(filename="log_file.log", level=logging.DEBUG, filemode='w',
                        format='%(asctime)s %(levelname)s:%(message)s')  # DEBUG, INFO, WARNING, ERROR и CRITICAL

    if not VISUAL_CHOICE:  # в коде
        if CALC_SET == 1:
            cor_task = {
                # в KIzFolder абсолютный путь к папке с файлами или файлу
                "KIzFolder": r"I:\rastr_add\test",
                # KInFolder папка в которую сохранять измененные файлы(или файл), "" не сохранять
                # результаты работы программы (.xlsx) сохраняются в папку KInFolder, если ее нет то в KIzFolder
                "KInFolder": r"I:\rastr_add\test\test_result",
                # ФИЛЬТР ФАЙЛОВ: False все файлы, True в соответствии с фильтром---------------------------------------
                "KFilter_file": False,
                "max_file_count": 1,  # максимальное количество расчетных файлов
                # нр("2019,2021-2027","зим","мин","1°C;МДП") (год, зим, макс, доп имя разделитель , или ;)
                "cor_criterion_start": {"years": "",
                                        "season": "",
                                        "max_min": "",
                                        "add_name": "", },
                # Корректировка в начале
                "cor_beginning_qt": {'add': False,
                                     'txt': ''},
                # импорт по excel------------------------------------------------------
                "import_val_XL": True,
                "excel_cor_file": r"I:\rastr_add\test\пример задания.xlsx",
                "excel_cor_sheet": "*",  # листы [импорт из моделей][XL->RastrWin], если'*', то все листы по порядку
                # Корректировка в конце
                "cor_end_qt": {'add': False,
                               'txt': ''},
                # Исправить пробелы, заменить английские буквы на русские.
                "cor_name": False,
                "cor_name_task": 'node:name,dname vetv:dname Generator:Name',
                # ----------------------------------------------------------------------------------------------------
                # "import_export_xl": False,  # False нет, True  import или export из xl в растр
                # "table": "Generator",  # нр "oborudovanie"
                # "export_xl": True,  # False нет, True - export из xl в растр
                # "XL_table": [r"C:\Users\User\Desktop\1.xlsx", "Generator"],  # полный адрес и имя листа
                # "tip_export_xl": 1,  # 1 загрузить, 0 присоединить 2 обновить
                # ----------------------------------------------------------------------------------------------------
                # что бы узел с скрм  вкл и отк этот  сопротивление единственной ветви r+x<0.2 и pn:qn:0
                # "AutoShuntForm": False,  # False нет, True сущ bsh записать в автошунт
                # "AutoShuntFormSel": "(na>0|na<13)",  # строка выборка узлов
                # "AutoShuntIzm": False,  # False нет, True вкл откл шунтов  autobsh
                # "AutoShuntIzmSel": "(na>0|na<13)",  # строка выборка узлов
                # Проверка параметров режима---------------------------------------------------------------------------
                # напряжений в узлах; дтн  в линиях(rastr.CalcIdop по degree_int);
                # pmax pmin относительно P у генераторов и pop_zad у территорий, объединений и районов; СЕЧЕНИЯ
                # выборка в таблице узлы "na=1|na=8)"
                "control_rg2": False,
                "control_rg2_task": {'node': False, 'vetv': True, 'Gen': False, 'section': False, 'area': False,
                                     'area2': False, 'darea': False, 'sel_node': "na>0"},
                # выводить данные из моделей в XL---------------------------------------------------------------------
                "printXL": True,
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
                    # из любой таблицы растр, нр "Generator" ,"P,Pmax" или "" все параметры, "Num>0" выборка)
                    "tab": {'add': False, "sel": "Num>0", 'tabl': "Generator",
                            'par': "Num,Name,sta,Node,P,Pmax,Pmin,value",
                            "rows": "Num,Name",  # поля строк в сводной
                            "columns": "год,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля столбцов в сводной
                            "values": "P,Pmax"}},  # поля значений в сводной

                # вывод заданных параметров в следующем формате "v=15105,15113,0|15038,15037,4/r|x|b; n=15198/pg|qg"
                # таблица: n-node,v-vetv,g-Generator,na-area,npa-area2,no-darea,nga-ngroup,ns-sechen
                "print_parameters": {'add': True,
                                     "sel": "v=15105,15113,0|15038,15037,4/r|x|b; n=151980/pg|qg"},
                # TODO БАЛАНС PQ_kor !!! 0 тоже район, даже если в районах не задан "na>13&na<201"
                "print_balance_q": {'add': False, "sel": "na=3012"},
                # ----------------------------------------------------------------------------------------------------
                "block_import": False,  # начало
            }
            start_cor(cor_task)  # cor
        if CALC_SET == 2:
            start_calc()  # calc
    else:
        if CALC_SET == 1:
            app = QtWidgets.QApplication([])  # Новый экземпляр QApplication
            app.setApplicationName("Правка моделей RastrWin")
            ui_edit = EditWindow()
            ui_edit.show()
            ui_set = SetWindow()
            sys.exit(app.exec_())  # Запуск

# TODO дописать: перенос параметров из одноименных файлов
# TODO дописать: сравнение файлов
