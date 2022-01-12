import win32com.client
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
# import pandas as pd
from typing import Union, Any
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
from qt_cor import Ui_MainCor  # импорт ui: pyuic5 qt_cor.ui -o qt_cor.py
from qt_set import Ui_Settings  # импорт ui: pyuic5 qt_set.ui -o qt_set.py


class SetWindow(QtWidgets.QMainWindow, Ui_Settings):
    def __init__(self):
        super(SetWindow, self).__init__()
        self.setupUi(self)
        self.load_ini()
        self.set_save.clicked.connect(lambda: self.save_ini())

    def load_ini(self):
        """Загрузить, создать или перезапичать ini файл"""
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


class EditWindow(QtWidgets.QMainWindow, Ui_MainCor):
    def __init__(self):
        super(EditWindow, self).__init__()  # *args, **kwargs
        self.setupUi(self)
        # скрыть параметры при старте
        self.GB_sel_file.hide()
        self.bloki_file.hide()
        self.sel_import.hide()
        self.FFV.hide()
        self.GB_import_val_XL.hide()
        self.GB_control.hide()
        self.GB_prinr_XL.hide()
        self.GB_sel_tabl.hide()
        self.TA_parametr_vibor.hide()
        self.balans_Q_vibor.hide()
        # CB показать скрыть параметры
        self.CB_KFilter_file.clicked.connect(lambda: self.show_hide(self.CB_KFilter_file, self.GB_sel_file))
        self.CB_bloki.clicked.connect(lambda: self.show_hide(self.CB_bloki, self.bloki_file))
        self.CB_ImpRg2.clicked.connect(lambda: self.show_hide(self.CB_ImpRg2, self.sel_import))
        self.CB_import_val_XL.clicked.connect(lambda: self.show_hide(self.CB_import_val_XL, self.GB_import_val_XL))
        self.CB_Filtr_V.clicked.connect(lambda: self.show_hide(self.CB_Filtr_V, self.FFV))
        self.CB_kontrol_rg2.clicked.connect(lambda: self.show_hide(self.CB_kontrol_rg2, self.GB_control))
        self.CB_printXL.clicked.connect(lambda: self.show_hide(self.CB_printXL, self.GB_prinr_XL))
        self.CB_print_tab_log.clicked.connect(lambda: self.show_hide(self.CB_print_tab_log, self.GB_sel_tabl))
        self.CB_print_parametr.clicked.connect(lambda: self.show_hide(self.CB_print_parametr, self.TA_parametr_vibor))
        self.CB_print_balans_Q.clicked.connect(lambda: self.show_hide(self.CB_print_balans_Q, self.balans_Q_vibor))
        self.run_krg2.clicked.connect(lambda: start_cor())
        self.SetBut.clicked.connect(lambda: ui_set.show())

    def show_hide(self, source, receiver):
        if source.isChecked():
            receiver.show()
        else:
            receiver.hide()


class GeneralSettings:
    """
    Для хранения общих настроек
    """
    # коллекция настроек, которые хранятся в ini файле
    set_save = {}
    # коллекция для хранения информации о расчете
    set_info = {"calc_val": {1: "ЗАМЕНИТЬ", 2: "ПРИБАВИТЬ", 3: "ВЫЧЕСТЬ", 0: "УМНОЖИТЬ"},
                'collapse': '',
                'end_info': ''}

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

        self.file_num = 0  # счетчик расчетных файлов
        self.now = datetime.now()
        self.time_start = time()
        self.now_start = self.now.strftime("%d-%m-%Y %H:%M")

    def the_end(self):  # по завершению
        time_spent = round(time() - self.time_start, 2)
        time_spent_minut = round(time_spent / 60, 1)
        self.set_info['end_info'] = (
                f"РАСЧЕТ ЗАКОНЧЕН! \nНачало расчета {self.now_start}, конец {self.now.strftime('%d-%m-%Y %H-%M')}" +
                f" \nЗатрачено: {str(time_spent)} секунд или {str(time_spent_minut)} минут")
        logging.info(self.set_info['end_info'])


class CorModel(GeneralSettings):
    """Коррекция файлов"""

    def __init__(self):
        super(CorModel, self).__init__()
        self.set = {
            "KIzFolder": r"I:\rastr_add\test",  # в KIzFolder абсолютный путь к папке с файлами или файлу
            # KInFolder папка в которую сохранять измененные файлы(или файл), "" не сохранять
            # результаты работы программы (.xlsx) сохраняются в папку KInFolder, если ее нет то в KIzFolder
            "KInFolder": r"I:\rastr_add\test_result",
            # ФИЛЬТР ФАЙЛОВ: False все файлы, True в соответствии с фильтром--------------------------------------------
            "KFilter_file": True,
            "max_file_count": 2,  # максимальное количество расчетных файлов
            # нр("2019,2021-2027","зим","мин","1°C;МДП") (год, зим, макс, доп имя разделитель , или ;)
            "cor_criterion_start": {"years": "2026",
                                    "season": "",
                                    "max_min": "",
                                    "add_name": ""},
            # импорт значений из excel, коррекция потребления-----------------------------------------------------------
            "import_val_XL": True,
            "excel_cor_file": r"I:\rastr_add\test\примеры.xlsx",
            "excel_cor_sheet": "[импорт из моделей][XL->RastrWin]",
            # ----------------------------------------------------------------------------------------------------------
            # "import_export_xl": False,  # False нет, True  import или export из xl в растр
            # "table": "Generator",  # нр "oborudovanie"
            # "export_xl": True,  # False нет, True - export из xl в растр
            # "XL_table": [r"C:\Users\User\Desktop\1.xlsx", "Generator"],  # полный адрес и имя листа
            # "tip_export_xl": 1,  # 1 загрузить, 0 присоединить 2 обновить
            # ----------------------------------------------------------------------------------------------------------
            # что бы узел с скрм  вкл и отк этот  сопротивление единственной ветви r+x<0.2 и pn:qn:0
            # "AutoShuntForm": False,  # False нет, True сущ bsh записать в автошунт
            # "AutoShuntFormSel": "(na>0|na<13)",  # строка выборка узлов
            # "AutoShuntIzm": False,  # False нет, True вкл откл шунтов  autobsh
            # "AutoShuntIzmSel": "(na>0|na<13)",  # строка выборка узлов
            # проверка параметров режима--------------------------------------------------------------------------------
            # напряжений в узлах; дтн  в линиях(rastr.CalcIdop по degree_int);
            # pmax pmin относительно P у генераторов и pop_zad у территорий, объединений и районов; СЕЧЕНИЯ
            # выборка в таблице узлы "na=1|na=8)"
            "control_rg2": True,
            "control_rg2_task": {'node': True, 'vetv': True, 'Gen': True, 'section': True, 'area': True,
                                 'area2': True, 'darea': True, 'sel_node': "na>0"},
            # выводить данные из моделей в XL---------------------------------------------------------------------------
            "printXL": True,
            "set_printXL": {
                "sechen": {'add': True, "sel": "", 'tabl': "sechen", 'par': "ns,name,pmin,pmax,psech",
                           "rows": "ns,name",  # поля строк в сводной
                           "columns": "год,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля столбцов в сводной
                           "values": "psech,pmax"},
                "area": {'add': True, "sel": "", 'tabl': "area",
                         'par': 'na,name,no,pg,pn,pn_sum,dp,pop,pop_zad,qn_sum,pg_max,pg_min,pn_max,pn_min,vnq,vnp,poq,qn,qg',
                         "rows": "na,name,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля строк в сводной
                         "columns": "год",  # поля столбцов в сводной
                         "values": "pop,pg"},
                "area2": {'add': True, "sel": "", 'tabl': "area2",
                          'par': 'npa,name,pg,pn,dp,pop,vnp,qg,qn,dq,poq,vnq,pn_sum,qn_sum,pop_zad',
                          "rows": "npa,name,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля строк в сводной
                          "columns": "год",  # поля столбцов в сводной
                          "values": "pop,pg"},
                "darea": {'add': True, "sel": "", 'tabl': "darea",
                          'par': 'no,name,pg,pp,pvn,qn_sum,pnr_sum,pn_sum,pop_zad,qvn,qp,qg',
                          "rows": "no,name,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля строк в сводной
                          "columns": "год",  # поля столбцов в сводной
                          "values": "pp,pg"},
                # из любой таблицы растр, нр "Generator" ,"P,Pmax" или "" все параметры, "Num>0" выборка)
                "tab": {'add': True, "sel": "Num>0", 'tabl': "Generator",
                        'par': "Num,Name,sta,Node,P,Pmax,Pmin,value",
                        "rows": "Num,Name",  # поля строк в сводной
                        "columns": "год,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля столбцов в сводной
                        "values": "P,Pmax"}},  # поля значений в сводной
            # вывод заданных параметров в следующем формате "v=42,48,0|43,49,0|27,11,3/r|x|b; n=8|6/pg|qg|pn|qn"
            # таблица: n-node,v-vetv,g-Generator,na-area,npa-area2,no-darea,nga-ngroup,ns-sechen
            "print_parameters": {'add': True, "sel": "v=15113,15112,1|15038,15037,4/r|x|b; n=15013|15021/pn|qn"},
            # БАЛАНС PQ_kor !!!0 тоже район,даже если в районах не задан "na>13&na<201"
            "print_balance_q": {'add': True, "sel": "na=3012"},
            # ---------------------------------------------------------------------------------------------------------
            "block_beginning": False,  # начало
            "block_import": False,  # начало
            "block_end": False,  # конец
            # ПРОЧИЕ НАСТРОЙКИ
        }
        self.pxl = None

        if VISUAL_SET == 1:
            self.set["KIzFolder"] = ui_edit.T_IzFolder.toPlainText()  # QPlainTextEdit
            self.set["KInFolder"] = ui_edit.T_InFolder.toPlainText()
            # фильтр
            self.set["KFilter_file"] = ui_edit.CB_KFilter_file.isChecked()  # QCheckBox
            self.set["file_count"] = ui_edit.D_count_file.value()  # QSpainBox
            self.set["cor_criterion_start"]["years"] = ui_edit.condition_file_years.text()  # QLineEdit text()
            self.set["cor_criterion_start"]["season"] = ui_edit.condition_file_season.currentText()  # QComboBox
            self.set["cor_criterion_start"]["max_min"] = ui_edit.condition_file_max_min.currentText()  #
            self.set["cor_criterion_start"]["add_name"] = ui_edit.condition_file_add_name.text()  #
            # задание из XL
            self.set["import_val_XL"] = ui_edit.CB_import_val_XL.isChecked()
            self.set["excel_cor_file"] = ui_edit.T_PQN_XL_File.toPlainText()
            self.set["excel_cor_sheet"] = ui_edit.T_PQN_Sheets.text()
            # расчет режима и контроль параметров режима
            self.set["control_rg2"] = ui_edit.CB_kontrol_rg2.isChecked()
            self.set["control_rg2_task"]['node'] = ui_edit.CB_U.isChecked()
            self.set["control_rg2_task"]['vetv'] = ui_edit.CB_I.isChecked()
            self.set["control_rg2_task"]['Gen'] = ui_edit.CB_gen.isChecked()
            self.set["control_rg2_task"]['section'] = ui_edit.CB_s.isChecked()
            self.set["control_rg2_task"]['area'] = ui_edit.CB_na.isChecked()
            self.set["control_rg2_task"]['area2'] = ui_edit.CB_npa.isChecked()
            self.set["control_rg2_task"]['darea'] = ui_edit.CB_no.isChecked()
            self.set["control_rg2_task"]['sel_node'] = ui_edit.kontrol_rg2_Sel.text()
            # импорт параметров режима
            if ui_edit.CB_ImpRg2.isChecked():
                # узлы -----------------------------------------------------------------------------
                if ui_edit.CB_N.isChecked():
                    ifm = ImportFromModel(import_file_name=ui_edit.file_N.text(),
                                          criterion_start={"years": ui_edit.Filtr_god_N.text(),
                                                           "season": ui_edit.Filtr_sez_N.currentText(),
                                                           "max_min": ui_edit.Filtr_max_min_N.currentText(),
                                                           "add_name": ui_edit.Filtr_dop_name_N.text()},
                                          tables=ui_edit.tab_N.text(),
                                          param=ui_edit.param_N.text(),
                                          sel=ui_edit.sel_N.text(),
                                          calc=ui_edit.tip_N.currentText())
                    ImportFromModel.all_import_model.append(ifm)
                # ветви -----------------------------------------------------------------------------
                if ui_edit.CB_V.isChecked():
                    ifm = ImportFromModel(import_file_name=ui_edit.file_V.text(),
                                          criterion_start={"years": ui_edit.Filtr_god_V.text(),
                                                           "season": ui_edit.Filtr_sez_V.currentText(),
                                                           "max_min": ui_edit.Filtr_max_min_V.currentText(),
                                                           "add_name": ui_edit.Filtr_dop_name_V.text()},
                                          tables=ui_edit.tab_V.text(),
                                          param=ui_edit.param_V.text(),
                                          sel=ui_edit.sel_V.text(),
                                          calc=ui_edit.tip_V.currentText())
                    ImportFromModel.all_import_model.append(ifm)
                # генераторы -----------------------------------------------------------------------------
                if ui_edit.CB_G.isChecked():
                    ifm = ImportFromModel(import_file_name=ui_edit.file_G.text(),
                                          criterion_start={"years": ui_edit.Filtr_god_G.text(),
                                                           "season": ui_edit.Filtr_sez_G.currentText(),
                                                           "max_min": ui_edit.Filtr_max_min_G.currentText(),
                                                           "add_name": ui_edit.Filtr_dop_name_G.text()},
                                          tables=ui_edit.tab_G.text(),
                                          param=ui_edit.param_G.text(),
                                          sel=ui_edit.sel_G.text(),
                                          calc=ui_edit.tip_G.currentText())
                    ImportFromModel.all_import_model.append(ifm)
                # районы -----------------------------------------------------------------------------
                if ui_edit.CB_A.isChecked():
                    ifm = ImportFromModel(import_file_name=ui_edit.file_A.text(),
                                          criterion_start={"years": ui_edit.Filtr_god_A.text(),
                                                           "season": ui_edit.Filtr_sez_A.currentText(),
                                                           "max_min": ui_edit.Filtr_max_min_A.currentText(),
                                                           "add_name": ui_edit.Filtr_dop_name_A.text()},
                                          tables=ui_edit.tab_A.text(),
                                          param=ui_edit.param_A.text(),
                                          sel=ui_edit.sel_A.text(),
                                          calc=ui_edit.tip_A.currentText())
                    ImportFromModel.all_import_model.append(ifm)
                # территории -----------------------------------------------------------------------------
                if ui_edit.CB_A2.isChecked():
                    ifm = ImportFromModel(import_file_name=ui_edit.file_A2.text(),
                                          criterion_start={"years": ui_edit.Filtr_god_A2.text(),
                                                           "season": ui_edit.Filtr_sez_A2.currentText(),
                                                           "max_min": ui_edit.Filtr_max_min_A2.currentText(),
                                                           "add_name": ui_edit.Filtr_dop_name_A2.text()},
                                          tables=ui_edit.tab_A2.text(),
                                          param=ui_edit.param_A2.text(),
                                          sel=ui_edit.sel_A2.text(),
                                          calc=ui_edit.tip_A2.currentText())
                    ImportFromModel.all_import_model.append(ifm)
                # объединения -----------------------------------------------------------------------------
                if ui_edit.CB_D.isChecked():
                    ifm = ImportFromModel(import_file_name=ui_edit.file_D.text(),
                                          criterion_start={"years": ui_edit.Filtr_god_D.text(),
                                                           "season": ui_edit.Filtr_sez_D.currentText(),
                                                           "max_min": ui_edit.Filtr_max_min_D.currentText(),
                                                           "add_name": ui_edit.Filtr_dop_name_D.text()},
                                          tables=ui_edit.tab_D.text(),
                                          param=ui_edit.param_D.text(),
                                          sel=ui_edit.sel_D.text(),
                                          calc=ui_edit.tip_D.currentText())
                    ImportFromModel.all_import_model.append(ifm)
                # объединения -----------------------------------------------------------------------------
                if ui_edit.CB_PQ.isChecked():
                    ifm = ImportFromModel(import_file_name=ui_edit.file_PQ.text(),
                                          criterion_start={"years": ui_edit.Filtr_god_PQ.text(),
                                                           "season": ui_edit.Filtr_sez_PQ.currentText(),
                                                           "max_min": ui_edit.Filtr_max_min_PQ.currentText(),
                                                           "add_name": ui_edit.Filtr_dop_name_PQ.text()},
                                          tables=ui_edit.tab_PQ.text(),
                                          param=ui_edit.param_PQ.text(),
                                          sel=ui_edit.sel_PQ.text(),
                                          calc=ui_edit.tip_PQ.currentText())
                    ImportFromModel.all_import_model.append(ifm)
                # IT -----------------------------------------------------------------------------
                if ui_edit.CB_IT.isChecked():
                    ifm = ImportFromModel(import_file_name=ui_edit.file_IT.text(),
                                          criterion_start={"years": ui_edit.Filtr_god_IT.text(),
                                                           "season": ui_edit.Filtr_sez_IT.currentText(),
                                                           "max_min": ui_edit.Filtr_max_min_IT.currentText(),
                                                           "add_name": ui_edit.Filtr_dop_name_IT.text()},
                                          tables=ui_edit.tab_IT.text(),
                                          param=ui_edit.param_IT.text(),
                                          sel=ui_edit.sel_IT.text(),
                                          calc=ui_edit.tip_IT.currentText())
                    ImportFromModel.all_import_model.append(ifm)

            for str_name in ["KIzFolder", "KInFolder", "excel_cor_file"]:
                if 'file:///' in self.set[str_name]:
                    self.set[str_name] = self.set[str_name][8:]

    def run_cor(self):
        """Запуск корректировки моделей"""
        # определяем корректировать файл или файлы в папке по анализу "KIzFolder"
        if os.path.isdir(self.set["KIzFolder"]):
            self.set["folder_file"] = 'folder'  # если корр папка
        elif os.path.isfile(self.set["KIzFolder"]):
            self.set["folder_file"] = 'file'  # если корр файл
            rm = RastrModel(full_name=self.set["KIzFolder"])
        else:
            mb.showerror("Ошибка в задании", "Не найден: " + self.set["KIzFolder"] + ", выход")
            return False
        # создать папку KInFolder
        if self.set["KInFolder"]:
            if not os.path.exists(self.set["KInFolder"]):
                logging.info("Создана папка: " + self.set["KInFolder"])
                os.mkdir(self.set["KInFolder"])

        folder_save = self.set["KInFolder"] if self.set["KInFolder"] else self.set["KIzFolder"]

        self.set['folder_result'] = folder_save + r"\result"  # папка для сохранения результатов
        now = datetime.now()
        self.set['name_time'] = self.set['folder_result'] + f"\\коррекция файлов ({now.strftime('%d-%m-%Y %H-%M')})"
        if not os.path.exists(self.set['folder_result']):
            os.mkdir(self.set['folder_result'])  # создать папку result
        self.set['folder_temp'] = self.set['folder_result'] + r"\temp"  # папка для сохранения рабочих файлов
        if not os.path.exists(self.set['folder_temp']):
            os.mkdir(self.set['folder_temp'])  # создать папку temp

        # if VISUAL_SET == 1 :
        #     if IE_kform.CB_bloki.checked :
        #         if len (IE_kform.bloki_file.value) > 0 :
        #             logging.info( "загружен файл: " + IE_kform.bloki_file.value)
        #             executeGlobal (CreateObject("Scripting.FileSystemObject").openTextFile(IE_kform.bloki_file.value).readAll())
        #         else:
        #             logging.info( "!!!НЕ УКАЗАН АДРЕС ФАЙЛА ЗАДАНИЯ!!!" )

        # ЭКСПОРТ ИЗ МОДЕЛЕЙ
        if self.set['block_import'] and VISUAL_SET == 0:
            import_model()  # ИД для импорта
        if self.set["import_val_XL"]:  # задать параметры узла по значениям в таблице excel (имя книги, имя листа)
            sheets = re.findall("\[(.+?)\]", self.set["excel_cor_sheet"])
            for sheet in sheets:
                cor_xl(self.set["excel_cor_file"], sheet, tip='export')

        load_add = []
        if ((self.set["printXL"] and self.set["set_printXL"]["sechen"]) or
                (self.set["control_rg2"] and self.set["control_rg2_task"]["section"])):
            load_add.append('sch')

        if self.set["folder_file"] == 'folder':  # корр файлы в папке
            files = os.listdir(self.set["KIzFolder"])  # список всех файлов в папке
            self.rastr_files = list(filter(lambda x: x.endswith('.rg2') | x.endswith('.rst'), files))

            for rastr_file in self.rastr_files:  # цикл по файлам .rg2 .rst в папке KIzFolder
                full_name = self.set["KIzFolder"] + '\\' + rastr_file
                full_name_new = self.set["KInFolder"] + '\\' + rastr_file
                rm = RastrModel(full_name, self.set["KFilter_file"], self.set["cor_criterion_start"])
                # отключен фильтр или соответствует ему
                if not self.set["KFilter_file"] or rm.Name_st != "не подходит":
                    self.file_num += 1
                    if self.set["KFilter_file"]:
                        if self.set["max_file_count"] > 0:
                            self.set["max_file_count"] -= 1
                        else:
                            break
                    rm.load(load_add=load_add)
                    self.cor_file(rm)
                    if self.set["KInFolder"]:
                        rm.save(full_name_new)
                else:
                    logging.debug("Файл отклонен, не соответствует фильтру: " + rastr_file)

        elif self.set["folder_file"] == 'file':  # корр файл
            rm.load(load_add=load_add)
            self.cor_file(rm)
            if self.set["KInFolder"]:
                rm.save(self.set["KInFolder"] + '\\' + rm.Name)

        ImportFromModel.all_import_model=[]

        if self.set['printXL']:
            self.pxl.finish()

        if self.set_info['collapse']:
            self.set_info['end_info'] += f"\nВНИМАНИЕ! развалились модели:\n[{self.set_info['collapse']}]. "

        self.the_end()
        shutil.copyfile('log_file.log', self.set['name_time'] + '.log')
        webbrowser.open(self.set['name_time'] + '.log')
        mb.showinfo("Инфо", self.set_info['end_info'])

    def cor_file(self, rm):
        """Обработать файл rm"""

        if 'block_beginning' in self.set:
            if self.set['block_beginning']:
                logging.info("\t***Блок начала ***")
                block_b(rm)
                logging.info("\t*** Конец блока начала ***")
        # if VISUAL_SET == 1:
        #    if self.IE_bloki:
        #        logging.info( "\t" & "*** блок начала (bloki.rbs)***")
        #        blok_n2 ()
        #        logging.info( "\t" & "*** конец блока начала (bloki.rbs)***")

        if ImportFromModel.all_import_model:
            for im in ImportFromModel.all_import_model:
                im.import_csv(rm)

        if "import_val_XL" in self.set:  # задать параметры узла по значениям в таблице excel (имя книги, имя листа)
            if self.set["import_val_XL"]:
                sheets = re.findall("\[(.+?)\]", self.set["excel_cor_sheet"])
                for sheet in sheets:
                    cor_xl(self.set["excel_cor_file"], sheet, rm=rm, tip='XL->RastrWin')
        # if self.import_export_xl:
        #     rastr_xl_tab (self.table , self.export_xl  , self.XL_table (0) , self.XL_table (1), self.tip_export_xl  )
        # if self.AutoShuntForm:
        #     add_AutoBsh (self.AutoShuntFormSel) #  процедура записывает из поля bsh в поле AutoBsh (выборка)
        # if self.AutoShuntIzm:
        #     AutoShunt_class_rec (self.AutoShuntIzmSel)#  процедура формирует Umin , Umax, AutoBsh , nBsh
        #     AutoShunt_class_kor ()  #  процедура меняет Bsh  и записывает AutoShunt_list
        #     AutoShunt_list = ""
        #
        # if VISUAL_SET = 1:
        #     if self.IE_CB_np_zad_sub:
        #         np_zad_sub ()   #  задать номер паралельности у ветвей с одинаковым ip i iq
        #     if self.IE_CB_name_txt_korr:
        #         name_txt_korr ()#   name_probel (r_table , r_tabl_pole), izm_bukvi(r_table , r_tabl_pole)#  удалить пробелы в начале и конце, заменить два пробела на один, английские менять на русские буквы
        #     if self.IE_CB_uhom_korr_sub:
        #         uhom_korr_sub ("")      #  исправить номинальные напряжения в узлах для ряда 6,10,35,110,150,220,330,500,750
        #     if self.IE_CB_SHN_ADD:
        #         SHN_ADD () #  добавить зависимость СХН
        #     if self.IE_bloki:
        #         logging.info( "\t" & "блок конца (bloki.rbs)" )
        #         blok_k2 ()
        #         logging.info( "\t" & "*** конец блока конца *** " )

        if 'block_end' in self.set:
            if self.set['block_end']:
                logging.info("\t*** Блок конца ***")
                block_e(rm)
                logging.info("\t*** Конец блока конца ***")

        if 'control_rg2' in self.set:
            if self.set['control_rg2']:
                if not control_rg2(rm, self.set['control_rg2_task']):  # расчет и контроль параметров режима
                    self.set_info['collapse'] += rm.name_base + ', '

        if 'printXL' in self.set:
            if self.set['printXL']:
                if not type(self.pxl) == PrintXL:
                    self.pxl = PrintXL(self.set)
                self.pxl.add_val(rm)


def block_b(rm):
    sel0(rm.rastr, 'block_b')
    rm.rgm("block_b")


def import_model():
    """ ИД для импорта из модели(выполняется после блока начала)"""
    ifm = ImportFromModel(import_file_name=r"I:\rastr_add\test\импорт.rg2",
                          criterion_start={"years": "2026",
                                           "season": "зим",
                                           "max_min": "макс",
                                           "add_name": ""},
                          tables="node;vetv;Generator",
                          param="sel",
                          sel="sel",
                          calc='2')
    ImportFromModel.all_import_model.append(ifm)


def block_e(rm):
    sel0(rm.rastr, 'block_e')
    rm.rgm("block_e")


# <<<<<<<<<<<<<<<<<<<<СПРАВКА>>>>>>>>>>>>>>>>>>>>>>>>>
# <<<ДОБАВИТЬ>>>
#  =fTAB_str_add ( "ngroup" , "nga=15" ) #  добавить запись в таблицу и вернуть indx ( "vetv" , "ip=1 iq=2 np=10 i_dop=100" )
#  =fVetv_add_ndx (dname , ip , iq , np , r , x , b) #  добавить ветвь и вернуть indx
#  =fNode_add (name , na , npa , uhom,ny)            #  добавить узел и вернуть номер (ny или 0)
#  vetv_vikl_add (viborka) #  для ветвей добавить выкл в начале и в конце
#  node_ku_add (viborka) #  к узлам присоединить новый узел и перенести ШР БСК УШР
#  sel_ssh2_add ()       #  к отмеченным узлам присоединить новый узел через выключатель и перенести верви с np=2,4,6 - мб не последняя версия
#  groupid_sel_sub ()#  0 нет 1 задать groupid отмеченных узлов ()
# <<<УДАЛИТЬ>>>
#  Del_sel ()            #  удалить отмеченные узлы (c ветвями) ветви и генераторы
#  Del(tabl,viborka)  # viborka = "net" - удалить узлы или ветви без связей или без узла начала конца
# <<<ИЗМЕНИТЬ СЕТЬ>>>
#  uhom_korr_sub (set_sel) #  исправить номинальные напряжения в узлах
# sel0 ()                    #  снять выделение узлов и ветвей  и генераторов
# SEL ("zadanie" , no_off) #  отметить, например "123 123,312,1 g,12",  no_off = 0 снять отметку 1 отметить
# kor  ("kkluch" , "zadanie")#  коррекция , например  kor "125 25" , "pn=10.2 qn=5.4" для узла, "g,125 g,125" , "Pmax=10 " для ген , "1,2,0 12,125,1" , "r=10.2 x=1" для ветви , также есть no npa na nga (принцип grup_cor)
# kor1  (k_kluch , param_kor , value_param)#  коррекция одного уникальнгого занчения(краткийй ключ, параметр корр, значение) например("7","name","Юж")
# grup_cor ( "tabl","param","viborka","formula")#  групповая коррекция "node","bsh","ny=87",-3036/1000000
# sta_node ("str_ny", on_off)#  узлы с ветвями (СТРОКА номера узлов через пробел) включить False; отключить True
# tN.cols.item("qn").calc ("0")
# rastr.RenumWP=True     # включить ссылки, отключить
# vzd0 ()           #  поиск узлов где напряжение vzd задано а диапозона реактивки нет и удаляет vzd
# name0 ()           #  поиск узлов и генераторов без имени
# nyNum0 ()           #  поиск узлов и генераторов с номером 0
# <<<прочее>>>
# if rm.test_name (array ("2020","","","")) :
# = otklonenie_seshen (nomer_sesh)   #   возвращает величину отклонения psech от  pmax   + превышение; - недобор
# = rastr.Calc("sum,max,min,val","area","qn","vibor") - функция (vibor не может быть "")
#  ПОТРЕБЛЕНИЕ cor_pop(zone,new_pop, task_save) npa_cor_pop(zone,new_pop, task_save)#  територия CorOb(ob,new_pop, task_save)#  обединение
#  ГЕНЕРАТОРЫ  PGen_cor ("sel")  # если мощность P больше Pmax то изменить мощность генератора  на Pmax, если P меньше Pmin но больше 0 - то на Pmin #  если P ген = 0 то отключить генератор, чтоб реактивка не выдавалась
#  СЕЧЕНИЕ # KorSech  (ns,newp,vibor , tip, net_Pmin_zad) #  номер сеч, новая мощность в сеч (значение или "max" "min"), выбор корр узлов  (нр "sel"или "" - авто) ,  tip - "pn" или "pg", net_Pmin_zad #  1 не учитывать Pmin
#  Qgen_node_in_gen_sub ()  #  посчитать Q ГЕН по  Q в узле
# <<<настройки rastr>>>
#  rastr.tables("com_regim").cols.item("gen_p").Z(0) = 0 #    0- "да"; 1- "да"; 2- только Р; 3- только Q ///it_max  количество расчетов///neb_p точность расчектов////
# <<<ТКЗ>>>
#    Delet_node_VL_sub () #  удалить промежкточные точки на ЛЭП при отсутствии магнитной связи


class RastrModel(GeneralSettings):
    """
    для хранения параметров текущего расчетного файла
    """

    def __init__(self, full_name: str, filter_file: bool = False, condition_file: dict = None):
        self.full_name = full_name
        self.dir = os.path.dirname(full_name)
        self.Name = os.path.basename(full_name)  # вернуть имя с расширением "2020 зим макс.rg2"
        self.name_base = self.Name[:-4]  # вернуть имя без расширения "2020 зим макс"
        self.tip_file = self.Name[-3:]  # rst или rg2
        self.pattern = self.set_save["шаблон " + self.tip_file]
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
                if self.name_list[2] == None:
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
            self.Name_st = self.god + " " + self.name_list[1]
            if self.kod_name_rg2 < 5:
                self.Name_st += " " + self.name_list[2]
            if (self.kod_name_rg2 in [1, 2]) or ("ПЭВТ" in self.name_base):
                self.temp_a_v_gost = True  # зима + период экстремально высоких температур -ПЭВТ
        else:
            self.Name_st = "не подходит"  # отсеиваем файлы задание и прочее

        pattern_name = re.compile("\((.+)\)")
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

        if filter_file and self.kod_name_rg2 > 0:
            if not (''.join(condition_file.values())).replace(' ', ''):  # условие не пустое
                if not self.test_name(condition_file):
                    self.Name_st = "не подходит"

    def test_name(self, condition: dict, info: str = "") -> bool:
        """
        Проверка имени файла на соответствие условию condition
        Возвращает True, если имя режима соответствует условию condition:
        нр, год("2020,2023-2025"), зим/лет/паводок("лет,зим"), макс/мин("макс"), доп имя("-41С;МДП:ТЭ-У")
        condition = {"years":"","season":"","max_min":"","add_name":""}-всегда истина
        str = для вывода в протокол
        """
        if not condition:
            return True
        if self.Name_st == "не подходит":
            return False

        if 'years' in condition:
            if condition['years']:
                fff = False
                for us in str_in_list(str(condition['years'])):
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

    def load(self, rastr='', load_add: list = None):
        """загрузить модель в Rastr
        load_add=['amt','sch','trn'] расширения файлов которые нужно загрузить
        загружается первый попавшийся файл в папке IzFolder"""
        if rastr:
            self.rastr = rastr
        else:
            self.rastr = win32com.client.Dispatch("Astra.Rastr")
        self.rastr.Load(1, self.full_name, self.pattern)  # загрузить
        logging.info("Загружен файл: " + self.full_name)
        # Загрузить файлы load_add
        if load_add:
            for extension in load_add:
                files = os.listdir(self.dir)
                names = list(filter(lambda x: x.endswith('.' + extension), files))
                if len(names) > 0:
                    self.rastr.Load(1, self.dir + '\\' + names[0], self.set_save["шаблон " + extension])
                    logging.info("Загружен файл: " + names[0])

    def save(self, full_name_new):
        self.rastr.Save(full_name_new, self.pattern)
        logging.info("Файл сохранен: " + full_name_new)

    def rgm(self, txt: str = "") -> bool:
        """Расчет режима"""
        for i in ['', '', '', 'p', 'p', 'p']:
            kod_rgm = self.rastr.rgm(i)  # 0 сошелся, 1 развалился
            if not kod_rgm:  # 0 сошелся
                if txt:
                    logging.debug(f"\tрасчет режима: {txt}")
                return True
        # развалился
        logging.error(f"расчет режима: {txt} !!!РАЗВАЛИЛСЯ!!!")
        return False


def str_in_list(id_str: str):
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


def cor_xl(excel_file_name: str, sheet: str, rm=None, tip: str = ''):
    """задать параметры по значениям в таблице excel (имя книги, имя листа,
    тип tip='export' или tip='XL->RastrWin'  )"""
    logging.info(f"\t Задать значения из excel, книга: {excel_file_name}, лист: {sheet}")
    if not os.path.exists(excel_file_name):
        logging.error("Ошибка в задании, не найден файл: " + excel_file_name)
        return False

    wb = load_workbook(excel_file_name)
    if sheet not in wb.sheetnames:
        logging.error(f"Ошибка в задании, не найден лист: {sheet} в файле {excel_file_name}")
        return False

    xl = wb[sheet]  # xl.cell(1,1).value [строки][столбцы] xl.max_row xl.max_column
    calc_val = xl.cell(1, 1).value
    if tip == 'export' and calc_val == "Параметры импорта из файлов RastrWin":  # импорт из rg2 rst
        # шаг по строкам
        for row in range(3, xl.max_row + 1):
            if xl.cell(row, 1).value and '#' not in xl.cell(row, 1).value:
                """ ИД для импорта из модели(выполняется после блока начала)"""
                ifm = ImportFromModel(import_file_name=xl.cell(row, 1).value,
                                      criterion_start={"years": xl.cell(row, 6).value,
                                                       "season": xl.cell(row, 7).value,
                                                       "max_min": xl.cell(row, 8).value,
                                                       "add_name": xl.cell(row, 9).value},
                                      tables=xl.cell(row, 2).value,
                                      param=xl.cell(row, 4).value,
                                      sel=xl.cell(row, 3).value,
                                      calc=xl.cell(row, 5).value)
                ImportFromModel.all_import_model.append(ifm)

    elif tip == 'XL->RastrWin' and calc_val != "Параметры импорта из файлов RastrWin":
        name_files = ""
        dict_param_column = {}  # {"pn":10-столбец}
        # Шаг по колонкам и запись в словарь всех столбцов для коррекции
        for column_name_file in range(2, xl.max_column + 1):
            if xl.cell(1, column_name_file).value not in ["", None]:
                name_files = xl.cell(1, column_name_file).value.split("|")  # list [name_file, name_file]
            if xl.cell(2, column_name_file).value:
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
                                            info=f"\tcor_xl, условие: {name_file}, "):
                                duct_add = True
                if duct_add:
                    dict_param_column[xl.cell(2, column_name_file).value] = column_name_file
        logging.debug("\t" + str(dict_param_column))

        if len(dict_param_column) == 0:
            logging.info(f"\t {rm.name_base} НЕ НАЙДЕН на листе {sheet} книги {excel_file_name}")
        else:
            calc_vals = {1: "ЗАМЕНИТЬ", 2: "+", 3: "-", 0: "*"}
            # 1: "ЗАМЕНИТЬ", 2: "ПРИБАВИТЬ", 3: "ВЫЧЕСТЬ", 0: "УМНОЖИТЬ"
            for row in range(3, xl.max_row + 1):
                for param, column in dict_param_column.items():
                    kkey = xl.cell(row, 1).value
                    if kkey not in [None, ""]:
                        new_val = xl.cell(row, column).value
                        if new_val != None:
                            if param not in ["pop", "pp"]:
                                if calc_val == 1:
                                    cor(rm.rastr, str(kkey), f"{param}={new_val}", True)
                                else:
                                    cor(rm.rastr, str(kkey), f"{param}={param}{calc_vals[calc_val]}{new_val}", True)
                            else:
                                cor_pop(rm.rastr, kkey, new_val)  # изменить потребление


def cor_pop(rastr, zone: str, new_pop: Union[int, float], task_save: str = None):
    """Изменить потребление
    (rastr, zone:"na=3", "npa=2" или "no=1", new_pop - значение потребления,
    задание на сохранение нагрузки узлов)"""
    eps = 0.003 if new_pop < 50 else 0.0003  # точность расчета, *100=%
    react_cor = True  # менять реактивное потребление пропорционально
    if '=' not in str(zone):
        logging.error(f"Ошибка в задании, cor_pop /{zone}/{str(new_pop)}/{str(task_save)}")
        return False
    zone_id = zone.partition('=')[0]
    name_zone = {"na": "area", "npa": "area2", "no": "darea",
                 "name_na": "район", "name_npa": "территория", "name_no": "объединение",
                 "p_na": "pop", "p_npa": "pop", "p_no": "pp"}
    # if task_save:
    #     nod = rastr.tables("node")
    #     if task_save != "sel": sel0 (): SEL (task_save , 1)
    #     nod.setsel ("sel")
    #     if rastr.tables("node").cols.Find("value1") < 1: rastr.tables("node").Cols.Add "value1",1
    #     if rastr.tables("node").cols.Find("value2") < 1: rastr.tables("node").Cols.Add "value2",1
    #     nod.cols.item("value1").calc ("pn")
    #     nod.cols.item("value2").calc ("qn")

    t_node = rastr.tables("node")
    t_zone = rastr.tables(name_zone[zone_id])
    t_zone.setsel(zone)
    ndx_z = t_zone.FindNextSel(-1)
    if zone_id == "no":
        t_area = rastr.tables("area")
        t_area.setsel(zone)
    if t_zone.cols.Find("pop_zad") > 0:
        t_zone.cols.Item("pop_zad").SetZ(ndx_z, new_pop)
    name_z = t_zone.cols.item('name').ZS(ndx_z)
    pop = t_zone.cols.item(name_zone['p_' + zone_id]).ZS(ndx_z)
    logging.info(f"\tизменить потребление: {name_z}({zone} текущее потребление {pop})")
    for i in range(10):  # максимальное число итераций
        pop = rastr.Calc("val", name_zone[zone_id], name_zone['p_' + zone_id], zone)
        # нр("val", "darea", "pp", "no=1")
        koef = new_pop / pop  # 50/55=0.9
        if abs(koef - 1) > eps:
            if zone_id == "no":
                ndx_na = t_area.FindNextSel(-1)
                while ndx_na != -1:
                    i_na = t_area.cols.item("na").Z(ndx_na)
                    t_node.setsel("na=" + str(i_na))
                    t_node.cols.item("pn").Calc("pn*" + str(koef))
                    if (react_cor): t_node.cols.item("qn").Calc("qn*" + str(koef))
                    ndx_na = t_area.FindNextSel(ndx_na)

            elif zone_id in ["npa", "na"]:
                t_node.setsel(zone)
                t_node.cols.item("pn").Calc("pn*" + str(koef))
                if react_cor:
                    t_node.cols.item("qn").Calc("qn*" + str(koef))
            # if task_save != "":
            #     if task_save != "sel": sel0 (): SEL (task_save , 1)
            #     nod.setsel ("sel")
            #     nod.cols.item("pn").calc ("value1")
            #     nod.cols.item("qn").calc ("value2")

            kod = rastr.rgm("")
            if kod != 0:
                logging.error(f"Аварийное завершение расчета, cor_pop /{zone}/{str(new_pop)}/{str(task_save)}")
                return False
        else:
            logging.info(
                f"\tпотребление {name_z}({zone}) изменено на {round(pop)} (должно быть {str(new_pop)}, {str(i + 1)} ит.)")
            return True


def cor(rastr, keys: str, tasks: str, cor_print: bool = True):
    """  коррекция  в таблицах rastr, например:
    ("125 ny=25", "pn=10.2 qn=5.4") для узла,
    ("ny", "Tc=0") для всех узлов таблицы,
    ("1,2,0 ip,iq,np=12,125,1", "r=10.2 x=1") для ветви,
    ("Num=25 g=12", "Pmax=10 ") для генераторов,
     аналогично: no npa na nga"""
    for key in keys.strip().split(" "):  # например:['125', 'g=125']
        # разделение ключей
        key_comma = key.split(",")  # нр для ветви [,,], для узла [], прочее [,]
        key_comma = [x.partition('.')[0] for x in key_comma]  # x.partition('.') разделить на 3 части, для округления
        key_equally = key.split("=")  # есть = [,], нет равно []
        # разделение задания
        for task in tasks.strip().split(" "):  # например:['pn=10.2', 'qn=5.4']
            task_equally = task.split("=")

            if key.isdigit() or "ny" in key:  # Узел
                if key == "ny":
                    set_row = ""
                else:
                    set_row = "ny=" + key_equally[-1]
                grup_cor(rastr, "node", task_equally[0], set_row, task_equally[1])
                #  (таблица, корр параметр, выборка, формула для расчета параметра)

            elif len(key_comma) > 2:  # Ветвь
                if key == "ip,iq,np":
                    set_row = ""
                else:
                    if len(key_comma) == 3:  # 1,2,0
                        set_row = f"ip={key_comma[0]}&iq={key_comma[1]}&np={key_comma[2]}"
                    else:
                        key_comma2 = key_equally[1].split(",")
                        set_row = f"ip={key_comma2[0]}&iq={key_comma2[1]}&np={key_comma2[2]}"
                grup_cor(rastr, "vetv", task_equally[0], set_row, task_equally[1])

            elif key_equally[0] in ["g", "Num"]:  # генератор
                set_row = "" if (key == "g" or key == "Num") else "Num=" + key_equally[1]
                grup_cor(rastr, "Generator", task_equally[0], set_row, task_equally[1])

            elif key_equally[0] == "no":  # объединене
                set_row = '' if key == "no" else "no=" + key_equally[1]
                grup_cor(rastr, "darea", task_equally[0], set_row, task_equally[1])

            elif key_equally[0] == "na":  # район
                set_row = '' if key == "na" else "na=" + key_equally[1]
                grup_cor(rastr, "area", task_equally[0], set_row, task_equally[1])

            elif key_equally[0] == "npa":  # территория
                set_row = '' if key == "npa" else "npa=" + key_equally[1]
                grup_cor(rastr, "area2", task_equally[0], set_row, task_equally[1])

            elif key_equally[0] == "nga":  # нагрузочные группы
                set_row = '' if key == "nga" else "nga=" + key_equally[1]
                grup_cor(rastr, "ngroup", task_equally[0], set_row, task_equally[1])

    if cor_print:
        logging.info(f"\t cor {keys},  {tasks}")


def grup_cor(rastr, tabl: str, param: str, selection: str, formula: str):
    """Групповая коррекция (rastr, таблица, параметр корр, выборка, формула для расчета параметра)"""
    if rastr.tables.Find(tabl) < 0:
        logging.error(f"\tВНИМАНИЕ! в rastrwin не загружена таблица {tabl}")
        return False
    table = rastr.tables(tabl)
    if table.cols.Find(param) < 0:
        logging.error(f"ВНИМАНИЕ! в таблице {tabl} нет параметра {param}")
        return False
    pparam = table.cols.item(param)
    table.setsel(selection)
    pparam.Calc(formula)
    return True


def control_rg2(rm, dict_task):
    """  контроль  dict_task = {'node': True, 'vetv': True, 'Gen': True, 'section': True, 'area': True, 'area2': True,
        'darea': True, 'sel_node': "na>0"}  """
    if not rm.rgm("control_rg2"):
        return False

    rastr = rm.rastr
    node = rastr.tables("node")
    branch = rastr.tables("vetv")
    generator = rastr.tables("Generator")
    chart_pq = rastr.tables("graphik2")
    graph_it = rastr.tables("graphikIT")

    # Напряжения
    if dict_task["node"]:
        logging.info("\tКонтроль напряжений.")

        uh = [6, 10, 35, 110, 220, 330, 500, 750]  # номинальные напряжения
        umin_n = [5.8, 9.7, 32, 100, 205, 315, 490, 730]  # минимальные нормальное напряжения для контроля
        unr = [7.2, 12, 42, 126, 252, 363, 525, 787]  # наибольшее рабочее напряжения

        node.setsel(dict_task["sel_node"])
        j = node.FindNextSel(-1)
        while j != -1:
            uhom = node.cols.item("uhom").Z(j)
            if uhom not in uh and uhom > 30:
                ny = node.cols.item('ny').ZS(j)
                name = node.cols.item('name').ZS(j)
                uhom = node.cols.item('uhom').ZS(j)
                logging.info(f"\t\tВНИМАНИЕ НАПРЯЖЕНИЕ! ny={ny}, имя: {name}, uhom={uhom} != Uном")
            j = node.FindNextSel(j)

        for i in range(len(uh)):  # напряжение меньше наибольшего рабочего и
            sel_node = "!sta&uhom=" + str(uh[i])
            if dict_task["sel_node"] != "":
                sel_node += "&" + dict_task["sel_node"]
            node.setsel(sel_node)
            j = node.FindNextSel(-1)
            while j != -1:
                if umin_n[i] > node.cols.item("vras").Z(j) > unr[i]:
                    ny = node.cols.item('ny').ZS(j)
                    name = node.cols.item('name').ZS(j)
                    vras = node.cols.item('vras').ZS(j)
                    logging.info(f"\t\tВНИМАНИЕ НАПРЯЖЕНИЕ! ny={ny}, имя: {name}, vras={vras},uhom={uh[i]}")
                j = node.FindNextSel(j)

        sel_node = "otv_min<0"  # Отклонение напряжения от umin минимально допустимого
        if dict_task["sel_node"] != "":
            sel_node += "&" + dict_task["sel_node"]
        node.setsel(sel_node)
        if node.count > 0:
            j = node.FindNextSel(-1)
            while j != -1:
                ny = node.cols.item('ny').ZS(j)
                name = node.cols.item('name').ZS(j)
                vras = node.cols.item('vras').ZS(j)
                umin = node.cols.item('umin').ZS(j)
                logging.info(f"\t\tВНИМАНИЕ НАПРЯЖЕНИЕ! ny={ny}, имя: {name}, vras={vras},umin={umin}")
                j = node.FindNextSel(j)
    # ТОКИ
    if dict_task['vetv']:
        rastr.CalcIdop(rm.degree_int, 0.0, "")
        logging.info("\tКонтроль токовой загрузки, расчетная температура: " + rm.degree_str)
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
                logging.info(f"\t\tВНИМАНИЕ ТОКИ! vetv:{branch.SelString(j)}, {name} - {round(i_zag)}%")
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
    if rastr.tables.Find("sechen") > 0:
        section = rastr.tables("sechen")
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
        logging.error("\tФайл сечений не загружеин")

    if dict_task['area']:
        control_pop(rastr, 'area')
    if dict_task['area2']:
        control_pop(rastr, 'area2')
    if dict_task['darea']:
        control_pop(rastr, 'darea')

    return True


def control_pop(rastr, zone: str):
    """zone =  'darea', 'area', 'area2'"""
    key_sone = {'darea': 'no', 'area': 'na', 'area2': 'npa'
        , 'darea_pop': 'pp', 'area_pop': 'pop', 'area2_pop': 'pop'
        , 'darea_name': 'объединений', 'area_name': 'районов', 'area2_name': 'территорий'}

    logging.info("\tКонтроль pop_zad " + key_sone[zone + '_name'])
    tabl = rastr.tables(zone)
    if tabl.cols.Find("pop_zad") < 0:
        logging.error("Поле pop_zad отсутствует в таблице " + key_sone[zone + '_name'])
    else:
        tabl.setsel("pop_zad>0")
        j = tabl.FindNextSel(-1)
        while j != -1:
            pop_zad = round(tabl.cols.item("pop_zad").Z(j))
            pp = round(tabl.cols.item(key_sone[zone + '_pop']).Z(j))
            deviation = round(abs(pop_zad - pp) / pop_zad, 2)
            if deviation > 0.01:
                name = tabl.cols.item("name").ZS(j)
                no = tabl.cols.item(key_sone[zone]).ZS(j)
                logging.info(f"\t\tВНИМАНИЕ: {name} ({no}), pop: {str(pp)}, pop_zad: {str(pop_zad)}, "
                             + f"отклонение: {str(round(pop_zad - pp))} или {str(round(deviation * 100))} %")
            j = tabl.FindNextSel(j)


def sheet_exists(book, sh_name: str):  # проверка существования лист в книге
    for sheet_i in book.Sheets:
        if sheet_i.name == sh_name:
            return True
    return False


class ImportFromModel:
    all_import_model = []  # хранение объектов класса ImportFromModel
    calc_str = {"обновить": 2, "загрузить": 1, "присоединить": 0, "присоединить-обновить": 3}

    def __init__(self, import_file_name: str = '', criterion_start: dict = {}, tables: str = '', param='',
                 sel: str = '', calc: Union[int, str] = '2'):
        """
        Импорт данных из файлов .rg2, .rst и др.
        Создает папку temp в папке с файлом и сохраняет в ней .csv файлы
        import_file_name = полное имя файла
        criterion_start={"years": "","season": "","max_min": "", "add_name": ""} условие выполнения
        tables = таблица для импорта, нр "node;vetv"
        param= параметры для импорта: "" все параметры или перечисление, нр 'sel,sta'(ключи не обязательно)
        sel= выборка нр "sel" или "" - все
        calc= число типа int или слово {"обновить": 2 , "загрузить": 1, "присоединить": 0, "присоединить-обновить": 3}
        """
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
            self.tables = tables.split(";")  # разделить на ["таблицы"]
            self.param = []
            self.sel = sel
            if type(calc) == int:
                self.calc = calc
            else:
                if calc in self.calc_str:
                    self.calc = self.calc_str[calc]
                else:
                    logging.error("Ошибка в задании, не распознано задание calc ImportFromModel: " + str(calc))
                    self.import_file_name = ''
            self.file_csv = []
            number = str(len(self.all_import_model))
            for tabl in self.tables:
                self.file_csv.append(f"{self.folder_temp}\\{self.basename}_{tabl}_{number}.csv")
                self.param.append(param)

            # Экспорт данных из файла в .csv файлы в папку temp
            if self.import_file_name:
                rastr = win32com.client.Dispatch("Astra.Rastr")
                rastr.Load(1, self.import_file_name, GeneralSettings.set_save['шаблон ' + self.import_file_name[-3:]])
                logging.info("\tЭкспорт из файла:" + self.import_file_name + ' в CSV')
                for index in range(len(self.tables)):
                    if not self.param[index]:  # если все параметры
                        self.param[index] = all_cols(rastr, self.tables[index])
                    else:
                        if rastr.Tables(self.tables[index]).Key not in self.param[index]:
                            self.param[index] += ',' + rastr.Tables(self.tables[index]).Key

                    logging.info(f"\t\tТаблица: {self.tables[index]}. Выборка: {self.sel}"
                                 + f"\n\t\tПараметры: {self.param[index]}"
                                 + f"\n\t\tФайл CSV: {self.file_csv[index]}")

                    tab = rastr.Tables(self.tables[index])
                    tab.setsel(self.sel)
                    tab.WriteCSV(1, self.file_csv[index], self.param[index], ";")  # 0 дописать, 1 заменить

    def import_csv(self, rm):
        """Импорт данных из csv в файла"""
        if self.import_file_name:
            logging.info("\tИмпорт из CSV в модель:" + self.import_file_name)
            if rm.test_name(condition=self.criterion_start, info='ImportFromModel'):
                for index in range(len(self.tables)):
                    logging.info(f"\t\tТаблица: {self.tables[index]}. Выборка: {self.sel}. тип: {str(self.calc)}" +
                                 f"\n\t\tФайл CSV: {self.file_csv[index]}" +
                                 f"\n\t\tПараметры: {self.param[index]}")
                    """{"обновить": 2 , "загрузить": 1, "присоединить": 0, "присоединить-обновить": 3}"""
                    tab = rm.rastr.Tables(self.tables[index])
                    tab.ReadCSV(self.calc, self.file_csv[index], self.param[index], ";", '')


def all_cols(rastr, tab):
    """Возвращает все колонки таблицы: ny,uhom...."""
    cls = rastr.Tables(tab).Cols
    cols_list = []
    for col in range(cls.Count):
        if cls(col).Name not in ["kkluch", "txt_zag", "txt_adtn_zag", "txt_ddtn", "txt_adtn", "txt_ddtn_zag"]:
            # print(str(cls(col).Name))
            cols_list.append(str(cls(col).Name))
    return ','.join(cols_list)


class PrintXL:
    """Класс печать данных в excel"""

    #  ...._log  лист протокол для сводной
    def __init__(self, set):  # добавить листы и первая строка с названиями
        self.excel = None
        self.wbook = None
        self.set = set
        self.list_name = ["name_rg2", "год", "лет/зим", "макс/мин", "доп_имя1", "доп_имя2", "доп_имя3"]
        self.book = Workbook()
        #  создать лист xl и присвоить ссылку на него
        for key in self.set['set_printXL']:
            if self.set['set_printXL'][key]['add']:
                self.set['set_printXL'][key]["sheet"] = self.book.create_sheet(key + "_log")
                # записать первую строку параметров
                header_list = self.list_name + self.set['set_printXL'][key]['par'].split(',')
                self.set['set_printXL'][key]["sheet"].append(header_list)

        if self.set['print_parameters']['add']:
            self.set['print_parameters']["sheet"] = self.book.create_sheet('parameters')

        if self.set['print_balance_q']['add']:
            self.set['print_balance_q']["sheet"] = self.book.create_sheet("balance_Q")
            self.balans_q_x0 = 5

    def add_val(self, rm):
        rastr = rm.rastr
        logging.info("\tВывод данных из моделей в XL")
        if rm.Name_st == "не подходит":
            dop_name_list = ['-'] * 3
        else:
            dop_name_list = rm.DopName[:3]
            if len(dop_name_list) < 3:
                dop_name_list += ['-'] * (3 - len(dop_name_list))
        list_name_z = [rm.name_base, rm.god, rm.name_list[1], rm.name_list[2]] + dop_name_list

        for key in self.set['set_printXL']:
            if not self.set['set_printXL'][key]['add']:
                continue
            # проверка наличия таблицы
            if rastr.Tables.Find(self.set['set_printXL'][key]['tabl']) < 0:
                logging.error("В RastrWin не загружена таблица " + self.set['set_printXL'][key]['tabl'])
                self.set['set_printXL'][key]['add'] = False
                continue

            # принт данных из растр в таблицу для СВОДНОЙ
            r_table = rastr.tables(self.set['set_printXL'][key]['tabl'])
            sheet = self.set['set_printXL'][key]["sheet"]
            param_list = self.set['set_printXL'][key]['par'].split(',')
            param_list = [param_list[i] if r_table.cols.Find(param_list[i]) > -1 else '-' for i in
                          range(len(param_list))]

            setsel = self.set['set_printXL'][key]['sel'] if self.set['set_printXL'][key]['sel'] else ""
            r_table.setsel(setsel)
            index = r_table.FindNextSel(-1)
            while index >= 0:
                sheet.append(
                    list_name_z + [r_table.cols.item(val).ZN(index) if val != '-' else '-' for val in param_list])
                index = r_table.FindNextSel(index)

        if self.set['print_parameters']['add']:
            dict_tables = {'n': 'node', 'v': 'vetv', 'g': 'Generator', 'na': 'area', 'npa': 'area2', 'no': 'darea',
                           'nga': 'ngroup', 'ns': 'sechen'}
            sheet = self.set['print_parameters']["sheet"]
            if sheet.max_row == 1:
                one_row_list = self.list_name[:]
            val_list = list_name_z[:]

            for task_i in self.set['print_parameters']['sel'].replace(' ', '').split(';'):
                key_row, key_column = task_i.split("/")  # нр key_row = "ny=8|9"   key_column = "pn|qn"
                key_column = key_column.split('|')  # ['pn','qn']
                key_row = key_row.split('=')  # ['n','8|9']
                set_key = key_row[1].split('|')  # ['8','9']
                tabl_key = dict_tables[key_row[0]]
                if rastr.Tables.Find(tabl_key) < 0:
                    logging.error("print_parameters, не найден: " + key_row[0])
                    continue
                t_print = rastr.Tables(tabl_key)

                for key_i in set_key:
                    if ',' in key_i:
                        ipiqnp = key_i.split(",")  # ветвь
                        if len(ipiqnp) != 3:
                            logging.error("print_parameters, ошибка: " + key_i)

                    for key_column_i in key_column:
                        choice = key_row[0] + '=' + key_i
                        if tabl_key == "vetv":
                            choice = f"ip={ipiqnp[0]}&iq={ipiqnp[1]}&np={ipiqnp[2]}"
                        if tabl_key == "Generator":
                            choice = 'Num=' + key_i
                        if tabl_key == "node":
                            choice = 'ny=' + key_i

                        name_tek = "name" if t_print.cols.Find("name") > 0 else "Name"
                        t_print.setsel(choice)
                        ndx = t_print.FindNextSel(-1)
                        if ndx > -1:
                            if sheet.max_row == 1:
                                one_row_list.append(f'{choice}/{key_column_i}({t_print.cols.item(name_tek).Z(ndx)})')
                            val_list.append(t_print.cols.item(key_column_i).ZN(ndx))
                        else:
                            if sheet.max_row == 1:
                                one_row_list.append("не найдено" + key_i)
                            val_list.append("не найдено")
            if sheet.max_row == 1:
                sheet.append(one_row_list)
            sheet.append(val_list)

        if self.set['print_balance_q']['add']:
            pass

    def finish(self):
        # преобразовать в объект таблицу и удалить листы с одной строкой
        sheet_couple = {}

        for sheet_name in self.book.sheetnames:
            sheet = self.book[sheet_name]
            if sheet.max_row == 1:
                del self.book[sheet_name]
            else:
                tab = Table(displayName=sheet_name,
                            ref='A1:' + get_column_letter(sheet.max_column) + str(sheet.max_row))
                style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                       showLastColumn=False, showRowStripes=True, showColumnStripes=True)
                tab.tableStyleInfo = style
                sheet.add_table(tab)
                if 'log' in sheet_name:
                    self.book.create_sheet(sheet_name.replace('log', 'сводная'))
                    sheet_couple[sheet_name] = sheet_name.replace('log', 'сводная')
        name_xl_file = self.set['name_time'] + '.xlsx'
        self.book.save(name_xl_file)
        self.book = None

        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.ScreenUpdating = False  # обновление экрана
        # self.excel.Calculation = -4135  # xlCalculationManual
        self.excel.EnableEvents = False  # отслеживание событий
        self.excel.StatusBar = False  # отображение информации в строке статуса excel

        self.wbook = self.excel.Workbooks.Open(name_xl_file)
        for n in range(self.wbook.sheets.count):
            if self.wbook.sheets[n].Name in sheet_couple:
                self.pivot_tables(self.wbook.sheets[n].Name, sheet_couple[self.wbook.sheets[n].Name])
        if self.set['folder_result']:
            self.wbook.Save()
        self.excel.Visible = True
        self.excel.ScreenUpdating = True  # обновление экрана
        self.excel.Calculation = -4105  # xlCalculationAutomatic
        self.excel.EnableEvents = True  # отслеживание событий
        self.excel.StatusBar = True  # отображение информации в строке статуса excel

        # if self.print_balans_Q:
        #     XL_print_balans_Q.Columns(4).ColumnWidth = 33
        #     diapozon = XL_print_balans_Q.UsedRange.address
        #     With
        #     XL_print_balans_Q.Range(diapozon)
        #     .HorizontalAlignment = -4108  # выравнивание по центру
        #     .VerticalAlignment = -4108
        #     .NumberFormat = "0"
        # diapozon = XL_print_balans_Q.UsedRange.address
        # XL_print_balans_Q.Range(diapozon)
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
        for n in range(self.wbook.sheets.count):
            if s_log == self.wbook.sheets[n].Name:
                tab_log = self.wbook.sheets[n].ListObjects[0]
        rows = self.set['set_printXL'][s_log[:-4]]['rows'].split(",")
        columns = self.set['set_printXL'][s_log[:-4]]['columns'].split(",")
        values = self.set['set_printXL'][s_log[:-4]]['values'].split(",")

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
        if len(values) > 0:
            pt.DataPivotField.Orientation = 1  # xlRowField"Значения в столбцах или строках xlColumnField

        # .DataPivotField.Position = 1 #  позиция в строках
        pt.RowGrand = False  # удалить строку общих итогов
        pt.ColumnGrand = False  # удалить столбец общих итогов
        pt.MergeLabels = True  # объединять одинаковые ячейки
        pt.HasAutoFormat = False  # не обновлять ширину при обновлении
        pt.NullString = "--"  # заменять пустые ячейки
        pt.PreserveFormatting = False  # сохранять формат ячеек при обнавлении
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


def sel0(rastr, txt=''):
    """ Снять отметку узлов, ветвей и генераторов"""
    rastr.Tables("node").cols.item("sel").Calc("0")
    rastr.Tables("vetv").cols.item("sel").Calc("0")
    rastr.Tables("Generator").cols.item("sel").Calc("0")
    if txt != '':
        logging.info("\tСнять отметку узлов, ветвей и генераторов")


def start_cor():
    """Запуск корректировки моделей"""
    global CM
    CM = CorModel()
    CM.run_cor()


def start_calc():
    """Запуск расчета моделей"""
    pass


if __name__ == '__main__':
    VISUAL_SET = 1  # 1 задание через QT, 0 - в коде
    CALC_SET = 1  # 1 -корректировать модели CorModel, 2-рассчитать модели Global_raschot_class
    CM = None  # глобальный объект класса CorModel
    # https://docs.python.org/3/library/logging.html
    logging.basicConfig(filename="log_file.log", level=logging.DEBUG, filemode='w',
                        format='%(asctime)s %(levelname)s:%(message)s')  # debug, INFO, WARNING, ERROR и CRITICAL

    if not VISUAL_SET:  # в коде
        if CALC_SET == 1:
            start_cor()  # corr
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
