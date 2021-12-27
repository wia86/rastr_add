import win32com.client
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
# import pandas as pd
import sys
from PyQt5 import QtWidgets
from datetime import datetime
from time import time
import os
import re
import configparser
# import random
import logging
import webbrowser
from tkinter import messagebox as mb
import numpy as np
from qt_cor import Ui_MainCor  # импорт ui: pyuic5 qt_cor.ui -o qt_cor.py
from qt_set import Ui_Settings  # импорт ui: pyuic5 qt_set.ui -o qt_set.py


class SetWindow(QtWidgets.QMainWindow, Ui_Settings):
    def __init__(self, *args, **kwargs):
        super(SetWindow, self).__init__()  # *args, **kwargs
        self.setupUi(self)
        self.set_save.clicked.connect(lambda: self.set_save_ini())

    def set_save_ini(self):
        config = configparser.ConfigParser()
        config['DEFAULT'] = {
            "folder RastrWin3": self.LE_path.text(),
            "шаблон rg2":       self.LE_rg2.text(),
            "шаблон rst":       self.LE_rst.text(),
            "шаблон sch":       self.LE_sch.text(),
            "шаблон amt":       self.LE_amt.text(),
            "шаблон trn":       self.LE_trn.text()}
        with open('settings.ini', 'w') as configfile:
            config.write(configfile)


class EditWindow(QtWidgets.QMainWindow, Ui_MainCor):
    def __init__(self, *args, **kwargs):
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

        self.run_krg2.clicked.connect(lambda: start())
        self.SetBut.clicked.connect(lambda: ui_set.show())
    # @staticmethod
    def show_hide(self, source, receiver):
        if source.isChecked():
            receiver.show()
        else:
            receiver.hide()


class GeneralSettings:  # GS. для хранения общих параметров
    def __init__(self):
        self.calc_set = 1  # 1 -корректировать модели CorSettings, 2-расчитать модели Global_raschot_class
        self.N_rg2_File = 0  # счетчик расчетных файлов
        self.kod_rgm = 0  # 0 сошелся, 1 развалился
        self.now = datetime.now()

        self.time_start = time()
        self.now_start = self.now.strftime("%d-%m-%Y %H:%M")
        self.set_save = {
            "folder RastrWin3": r"C:\Users\User\Documents\RastrWin3",
            "шаблон rg2": r"C:\Users\User\Documents\RastrWin3\SHABLON\режим.rg2",
            "шаблон rst": r"C:\Users\User\Documents\RastrWin3\SHABLON\динамика.rst",
            "шаблон sch": r"C:\Users\User\Documents\RastrWin3\SHABLON\сечения.sch",
            "шаблон amt": r"C:\Users\User\Documents\RastrWin3\SHABLON\автоматика.amt",
            "шаблон trn": r"C:\Users\User\Documents\RastrWin3\SHABLON\трансформаторы.trn"}


    # self.set = {
    #     "calc_val": {1: "ЗАМЕНИТЬ", 2: "ПРИБАВИТЬ", 3: "ВЫЧЕСТЬ", 0: "УМНОЖИТЬ"}
    # }

    def end_gl(self):  # по завершению макроса
        if self.calc_set == 2:
            if (GLR.kol_test_da + GLR.kol_test_net) > 0:
                percentages = str(round(GLR.kol_test_net / (GLR.kol_test_da + GLR.kol_test_net) * 100))
            else:
                percentages = "0"

        result_info = (f"""РАСЧЕТ ЗАКОНЧЕН!
          Начало расчета {self.now_start} конец {self.now.strftime('%d-%m-%Y %H-%M')}
          Затрачено: {str(round(time() - self.time_start))} с. ({str(round((time() - self.time_start) / 60, 1))} мин)""")
        if self.calc_set == 2:
            result_info += f"""\n Сочетаний отфильтровано: {str(GLR.kol_test_net)} 
                из {str(GLR.kol_test_da + GLR.kol_test_net)} ({percentages} %)
                Скорость расчета: {str(round(GLR.kol_test_da / (time() - self.time_start), 1))} сочетаний/сек."""
        logging.info(result_info)
        mb.showinfo("Инфо", result_info)
        webbrowser.open("log_file.log")


class CorSettings():  # CS. для хранения общих параметров - КОРРЕКЦИЯ ФАЙЛОВ
    dict_import_model = {}  # хранение объектов класса ImportFromModel

    def __init__(self):
        self.set = {
            # в KIzFolder абсолютный путь к папке с файлами или файлу
            "KIzFolder": r"I:\rastr_add\test",  # расчетный файл или папка с файлами
            # KInFolder папка в которую сохранять измененные файлы, "" не сохранять
            # результаты работы программы (.xlsx) сохраняются в папку KInFolder, если ее нет то в KIzFolder
            "KInFolder": r"I:\rastr_add\test_result",
            # ФИЛЬТР ФАЙЛОВ: False все файлы, True в соответствии с фильтром
            "KFilter_file": True,
            "max_file_count": 2,  # максимальное количество расчетных файлов
            # нр("2019,2021-2027","зим","мин","1°C;МДП") (год, зим, макс, доп имя разделитель , или ;)
            "cor_criterion_start": {"years": "2026-2027",
                                    "season": "",
                                    "max_min": "",
                                    "add_name": ""},
            # импорт значений из excel, коррекция потребления-----------------------------------------------------------
            "import_val_XL": False,
            "excel_cor_file": r"I:\rastr_add\test\примеры.xlsx",
            "excel_cor_sheet": "[импорт из моделей][XL->RastrWin][pop]",
            # ----------------------------------------------------------------------------------------------------------
            "import_export_xl": False,  # False нет, True  import или export из xl в растр
            "table": "Generator",  # нр "oborudovanie"
            "export_xl": True,  # False нет, True - export из xl в растр
            "XL_table": [r"C:\Users\User\Desktop\1.xlsx", "Generator"],  # полный адрес и имя листа
            "tip_export_xl": 1,  # 1 загрузить, 0 присоединить 2 обновить
            # ----------------------------------------------------------------------------------------------------------
            # что бы узел с скрм  вкл и отк этот  сопротивление единственной ветви r+x<0.2 и pn:qn:0
            "AutoShuntForm": False,  # False нет, True сущ bsh записать в автошунт
            "AutoShuntFormSel": "(na>0|na<13)",  # строка выборка узлов
            "AutoShuntIzm": False,  # False нет, True вкл откл шунтов  autobsh
            "AutoShuntIzmSel": "(na>0|na<13)",  # строка выборка узлов
            # проверка параметров режима:
            # напряжений в узлах; дтн  в линиях(rastr.CalcIdop по GradusZ);
            # pmax pmin относительно P у генераторов и pop_zad у территорий, объединений и районов; СЕЧЕНИЯ
            # выборка в таблице узлы "na=1|na=8)"
            "control_rg2": True,
            "control_rg2_task": {'node': True, 'vetv': True, 'Gen': True, 'section': True, 'area': True
                , 'area2': True, 'darea': True, 'sel_node': "na>0"},
            # выводить данные из rastr в XL-----------------------------------------------------------------------------
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
            "block_beginning": 0,  # начало
            "block_import": 0,  # начало
            "block_end": 0,  # конец
            # ПРОЧИЕ НАСТРОЙКИ
            "folder_result": '',  # папка для сохранения результатов
            "folder_temp": '',  # папка для сохранения рабочих файлов
            "collapse": ""}

        if visual_set == 1:
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
                if ui_edit.CB_V.isChecked():
                    self.import_from_model = None
                    self.criterion_start = {"years": ui_edit.Filtr_god_V.text(),
                                            "season": ui_edit.Filtr_sez_V.currentText(),
                                            "max_min": ui_edit.Filtr_max_min_V.currentText(),
                                            "add_name": ui_edit.Filtr_dop_name_V.text()}
                    ImportFromModel.number += 1
                    self.import_from_model = ImportFromModel(import_file_name=ui_edit.file_V.text()
                                                             , criterion_start=self.criterion_start
                                                             , tables=ui_edit.tab_V.text()
                                                             , param=ui_edit.param_V.text()
                                                             , sel=ui_edit.sel_V.text()
                                                             , calc=ui_edit.tip_V.currentText())
                    ImportFromModel.number += 1
                    self.dict_import_model[ImportFromModel.number] = self.import_from_model
                    # --------------------------------------------------------------------------------
        for str_name in ["KIzFolder", "KInFolder", "excel_cor_file"]:
            if 'file:///' in self.set[str_name]:
                self.set[str_name] = self.set[str_name][8:]


def block_b():
    sel0('block_b')
    #  Del_sel ()
    rgm("block_b")


def import_model():
    """ ИД для импорта из модели(выполняется после блока начала)"""
    import_from_model = ImportFromModel(import_file_name=r"I:\rastr_add\test\импорт.rg2"
                                        , criterion_start={"years": "2026",
                                                           "season": "зим",
                                                           "max_min": "макс",
                                                           "add_name": ""}
                                        , tables="node;vetv;Generator"
                                        , param="sel"
                                        , sel="sel"
                                        , calc=2)
    ImportFromModel.number += 1
    CS.dict_import_model[ImportFromModel.number] = import_from_model
    # --------------------------------------------------------------------------------


def block_e():
    # sel0('block_e')
    rgm("block_e")


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
# if RG.test_name (array ("2020","","","")) :
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


class CurrentFile:
    # RG  RG. для хранения параметров текущего расчетного файла
    def __init__(self, rastr_file, Filter_file=False, Uslovie_file={}):
        global CS
        global GLR
        self.Name = os.path.basename(rastr_file)  # вернуть имя с расширением "2020 зим макс.rg2"
        self.Name_Base = self.Name[:-4]  # вернуть имя без расширения "2020 зим макс"
        self.tip_file = self.Name[-3:]  # rst или rg2
        self.shablon = GS.set_save["шаблон " + self.tip_file]
        self.kod_name_rg2 = 0  # 0 не распознан, 1 зим макс 2 зим мин 3 лет макс 4 лет мин 5 паводок макс
        self.temp_a_v_gost = False  # True темературы  а-в - зима + лето ПЭВТ
        self.TabRgmCount = 1  # счетчик режимов в каждой таблице
        self.txt_dop = ""
        self.GradusZ = 0
        self.Gradus = ""
        self.loadRGM = False
        self.DopNameStr = ""
        self.name_list = ["-", "-", "-"]
        pattern_name = re.compile("^(20[1-9][0-9])\s(лет\w?|зим\w?|паводок)\s?(макс|мин)?")
        match = re.search(pattern_name, self.Name_Base)
        if match:
            if match.re.groups == 3:
                self.name_list = [match[1], match[2], match[3]]
                if self.name_list[2] == None:
                    self.name_list = "-"
                if self.name_list[1] == "паводок":
                    self.kod_name_rg2 = 5
                    self.SezonName = "Паводок"
                if self.name_list[1] == "зим" and self.name_list[2] == "макс":
                    self.kod_name_rg2 = 1
                    self.SezonName = "Зимний максимум нагрузки"
                if self.name_list[1] == "зим" and self.name_list[2] == "мин":
                    self.kod_name_rg2 = 2
                    self.SezonName = "Зимний минимум нагрузки"
                if self.name_list[1] == "лет" and self.name_list[2] == "макс":
                    self.kod_name_rg2 = 3
                    self.SezonName = "Летний максимум нагрузки"
                if self.name_list[1] == "лет" and self.name_list[2] == "мин":
                    self.kod_name_rg2 = 4
                    self.SezonName = "Летний минимум нагрузки"

        self.god = self.name_list[0]
        if self.kod_name_rg2 > 0:
            self.Name_st = self.god + " " + self.name_list[1]
            if self.kod_name_rg2 < 5:
                self.Name_st += " " + self.name_list[2]

            if GS.calc_set == 2:  # расчет режимов а не корр
                if GLR.gost58670:
                    if (self.kod_name_rg2 in [1, 2]) or ("ПЭВТ" in self.Name_Base):
                        self.temp_a_v_gost = True  # зима + период экстремально высоких температур -ПЭВТ
        else:
            self.Name_st = "не подходит"  # отсеиваем файлы задание и прочее

        pattern_name = re.compile("\((.+)\)")
        match = re.search(pattern_name, self.Name_Base)
        if match:
            self.DopNameStr = match[1]

        if self.DopNameStr.replace(" ", "") != "":
            if "," in self.DopNameStr:
                self.DopName = self.DopNameStr.split(",")
            elif ";" in self.DopNameStr:
                self.DopName = self.DopNameStr.split(";")
            else:
                self.DopName = [self.DopNameStr]
        if "°C" in self.Name_Base:
            pattern_name = re.compile("(-?\d+((,|\.)\d*)?)\s?°C")  # -45.,14 °C
            match = re.search(pattern_name, self.Name_Base)
            if match:
                self.Gradus = match[1].replace(',', '.')
                self.GradusZ = float(self.Gradus)  # число
                self.txt_dop = "Расчетная температура " + self.Gradus + " °C. "

        if GS.calc_set == 2:  # расчет режимов

            if self.kod_name_rg2 > 0:
                if GLR.zad_temperatura == 1:
                    if self.name_list[1] == "зим":
                        self.GradusZ = GLR.temperatura_zima
                    else:
                        self.GradusZ = GLR.temperatura_leto

                    self.Gradus = str(self.GradusZ)
                    self.txt_dop = "Расчетная температура " + self.Gradus + " °C. "

                # for DopName_tek in self.DopName:
                #     for each ii in GLR.rg2_name_metka
                #         if trim (DopName_tek) = trim (ii (0)):
                #             txt_dop = txt_dop + ii (1)

                self.NAME_RG2_plus = self.SezonName + " " + self.god + " г"
                if self.txt_dop != "":
                    self.NAME_RG2_plus += ". " + self.txt_dop
                self.NAME_RG2_plus2 = self.SezonName + "(" + self.Gradus + " °C)"
                self.TEXT_NAME_TAB = GLR.tabl_name_OK1 + str(
                    GLR.Ntabl_OK) + GLR.tabl_name_OK2 + self.SezonName + " " + self.god + " г. " + self.txt_dop

        if Filter_file and self.kod_name_rg2 > 0:
            if (''.join(Uslovie_file.values())).replace(' ', '') != '':  # условие не пустое ["","","",""]
                if not self.test_name(Uslovie_file):
                    self.Name_st = "не подходит"

    def test_name(self, dict_uslovie, info=""):  # возвращает истина если имя режима соответствует условию
        # нр dict_uslovie = {"years":"","season":"","max_min":"","add_name":""}-всегда истина,или год ("2020,2023-2025"), зим/лет("зим" "лет,зим"), макс/мин("макс" "мин"), доп имя("МДП:ТЭ-У" "41С,МДП:ТЭ-У")

        if not dict_uslovie:
            return True
        if self.Name_st == "не подходит":
            return False
        if dict_uslovie['years']:
            fff = False
            for us in str_in_list(str(dict_uslovie['years'])):
                if int(self.god) == us:
                    fff = True
            if not fff:
                logging.debug(info + self.Name + f" Год '{self.god}' не проходит по условию: "
                              + str(dict_uslovie['years']))
                return False
        if dict_uslovie['season']:
            if dict_uslovie['season'].strip():  # ПРОВЕРКА "зим" "лет" "паводок"
                fff = False
                temp = dict_uslovie['season'].replace(' ', '')
                for us in temp.split(","):
                    if self.name_list[1] == us:
                        fff = True
                if not fff:
                    logging.debug(info + self.Name + f" Сезон '{self.name_list[1]}' не проходит по условию: "
                                  + dict_uslovie['season'])
                    return False

        if dict_uslovie['max_min']:
            if dict_uslovie['max_min'].strip():  # ПРОВЕРКА "макс" "мин"
                if self.name_list[2] != dict_uslovie['max_min'].replace(' ', ''):
                    logging.debug(info + self.Name + f" '{self.name_list[2]}' не проходит по условию: "
                                  + dict_uslovie['max_min'])
                    return False

        if dict_uslovie['add_name']:
            if dict_uslovie['add_name'].strip():  # ПРОВЕРКА (-41С;МДП:ТЭ-У)
                if ";" in dict_uslovie['add_name']:
                    temp = dict_uslovie['add_name'].split(";")
                else:
                    temp = dict_uslovie['add_name'].split(",")
                fff = False
                for us in temp:
                    for DopName_i in self.DopName:
                        if DopName_i == us:
                            fff = True
                if not fff:
                    logging.debug(
                        info + self.Name + f" Доп. имя {self.DopNameStr} не проходит по условию: " + dict_uslovie[
                            'add_name'])
                    return False
        return True


def str_in_list(id_str: str) -> list:
    """функция из строки "2021,2023-2025" делает [2021,2023,2024,2025]"""
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


def main_cor():  # головная процедура
    global CS
    global RG
    global GS
    global pxl
    pxl = None
    # определяем корректировать файл или файлы в папке по анализу "KIzFolder"
    if os.path.isdir(CS.set["KIzFolder"]):
        CS.set["folder_file"] = 'folder'  # если корр папка
    elif os.path.isfile(CS.set["KIzFolder"]):
        CS.set["folder_file"] = 'file'  # если корр файл
        RG = CurrentFile(CS.set["KIzFolder"])
    else:
        mb.showerror("Ошибка в задании", "Не найден: " + CS.set["KIzFolder"] + ", выход")
        return False
    # создать папку KInFolder
    if CS.set["KInFolder"]:
        if not os.path.exists(CS.set["KInFolder"]):
            logging.info("Создана папка: " + CS.set["KInFolder"])
            os.mkdir(CS.set["KInFolder"])

    folder_save = CS.set["KInFolder"] if CS.set["KInFolder"] else CS.set["KIzFolder"]

    CS.set['folder_result'] = folder_save + r"\result"
    if not os.path.exists(CS.set['folder_result']):
        os.mkdir(CS.set['folder_result'])  # создать папку result
    CS.set['folder_temp'] = CS.set['folder_result'] + r"\temp"
    if not os.path.exists(CS.set['folder_temp']):
        os.mkdir(CS.set['folder_temp'])  # создать папку temp

    # if visual_set == 1 :
    #     if IE_kform.CB_bloki.checked :
    #         if len (IE_kform.bloki_file.value) > 0 :
    #             logging.info( "загружен файл: " + IE_kform.bloki_file.value)
    #             executeGlobal (CreateObject("Scripting.FileSystemObject").openTextFile(IE_kform.bloki_file.value).readAll())
    #         else:
    #             logging.info( "!!!НЕ УКАЗАН АДРЕС ФАЙЛА ЗАДАНИЯ!!!" )

    # ЭКСПОРТ ИЗ МОДЕЛЕЙ
    if CS.set['block_import'] == 1 and visual_set == 0:
        import_model()  # ИД для импорта
    if CS.set["import_val_XL"]:  # задать параметры узла по значениям в таблице excel (имя книги, имя листа)
        sheets = re.findall("\[(.+?)\]", CS.set["excel_cor_sheet"])
        for sheet in sheets:
            cor_xl(CS.set["excel_cor_file"], sheet, tip='export')
    # if IE_kform.CB_ImpRg2.checked: IE_ImpRg2()  # запуск ИД для импорта из задания IE
    # if IE_kform.CB_bloki.checked: ImpRg22()  # запуск ИД для импорта из файла блоки
    if len(CS.dict_import_model) > 0:
        for import_from_model_i in CS.dict_import_model.values():
            import_from_model_i.export_csv()

    if CS.set["folder_file"] == 'folder':  # корр файлы в папке
        files = os.listdir(CS.set["KIzFolder"])  # список всех файлов в папке
        rastr_files = list(filter(lambda x: x.endswith('.rg2') | x.endswith('.rst'), files))  # фильтр файлов

        for rastr_file in rastr_files:  # цикл по файлам .rg2 .rst в папке KIzFolder
            full_name = CS.set["KIzFolder"] + '\\' + rastr_file
            full_name_new = CS.set["KInFolder"] + '\\' + rastr_file
            RG = CurrentFile(full_name, CS.set["KFilter_file"], CS.set["cor_criterion_start"])
            if not CS.set["KFilter_file"] or RG.Name_st != "не подходит":  # отключен фильтр или соответствует ему
                GS.N_rg2_File += 1
                if CS.set["KFilter_file"]:
                    if CS.set["max_file_count"] > 0:
                        CS.set["max_file_count"] -= 1
                    else:
                        break
                logging.info("Загружен файл: " + rastr_file)
                rastr.Load(1, full_name, RG.shablon)  # загрузить
                cor_file()
                if CS.set["KInFolder"]:
                    rastr.Save(full_name_new, RG.shablon)
                    logging.info("Файл сохранен: " + full_name_new)
            else:
                logging.debug("Файл отклонен, не соответствует фильтру: " + rastr_file)

    elif CS.set["folder_file"] == 'file':  # корр файл
        rastr.Load(1, CS.set["KIzFolder"], RG.shablon)  # загрузить режим
        logging.info("Загружен файл: " + CS.set["KIzFolder"])
        cor_file()
        if CS.set["KInFolder"]:
            rastr.Save(CS.set["KInFolder"] + '\\' + RG.Name, RG.shablon)
            logging.info("Файл сохранен: " + CS.set["KInFolder"])

    if CS.set['printXL']:
        pxl.finish()
    if CS.set['collapse'] != "":
        GS.result_info += f"\nВНИМАНИЕ! имеются модели которые развалились:\n[{CS.collapse}]. "


def cor_file():
    global pxl
    if CS.set['block_beginning']:
        logging.info("\t***Блок начала ***")
        block_b()
        logging.info("\t*** Конец блока начала ***")
    # if visual_set == 1:
    #    if CS.IE_bloki:
    #        logging.info( "\t" & "*** блок начала (bloki.rbs)***" )
    #        blok_n2 ()
    #        logging.info( "\t" & "*** конец блока начала (bloki.rbs)***" )

    if len(CS.dict_import_model) > 0:
        for dict_import_model_i in CS.dict_import_model.values():
            dict_import_model_i.import_csv()

    if CS.set["import_val_XL"]:  # задать параметры узла по значениям в таблице excel (имя книги, имя листа)
        sheets = re.findall("\[(.+?)\]", CS.set["excel_cor_sheet"])
        for sheet in sheets:
            cor_xl(CS.set["excel_cor_file"], sheet, tip='XL->RastrWin')
    # if CS.import_export_xl:
    #     rastr_xl_tab (CS.table , CS.export_xl  , CS.XL_table (0) , CS.XL_table (1), CS.tip_export_xl  )
    # if CS.AutoShuntForm:
    #     add_AutoBsh (CS.AutoShuntFormSel) #  процедура записывает из поля bsh в поле AutoBsh (выборка)
    # if CS.AutoShuntIzm:
    #     AutoShunt_class_rec (CS.AutoShuntIzmSel)#  процедура формирует Umin , Umax, AutoBsh , nBsh
    #     AutoShunt_class_kor ()  #  процедура меняет Bsh  и записывает GS.AutoShunt_list
    #     GS.AutoShunt_list = ""
    #
    # if visual_set = 1:
    #     if CS.IE_CB_np_zad_sub:
    #         np_zad_sub ()   #  задать номер паралельности у ветвей с одинаковым ip i iq
    #     if CS.IE_CB_name_txt_korr:
    #         name_txt_korr ()#   name_probel (r_table , r_tabl_pole), izm_bukvi(r_table , r_tabl_pole)#  удалить пробелы в начале и конце, заменить два пробела на один, английские менять на русские буквы
    #     if CS.IE_CB_uhom_korr_sub:
    #         uhom_korr_sub ("")      #  исправить номинальные напряжения в узлах для ряда 6,10,35,110,150,220,330,500,750
    #     if CS.IE_CB_SHN_ADD:
    #         SHN_ADD () #  добавить зависимость СХН
    #     if CS.IE_bloki:
    #         logging.info( "\t" & "блок конца (bloki.rbs)" )
    #         blok_k2 ()
    #         logging.info( "\t" & "*** конец блока конца *** " )
    #
    if CS.set['block_end']:
        logging.info("\t*** Блок конца ***")
        block_e()
        logging.info("\t*** Конец блока конца ***")
    if CS.set['control_rg2']:
        control_rg2(CS.set['control_rg2_task'])  # расчет и контроль параметров режима
    if CS.set['printXL']:
        if not type(pxl) == PrintXL:
            pxl = PrintXL()
        pxl.add_val()


def cor_xl(excel_file_name, sheet, tip=''):
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
                import_from_model = ImportFromModel(import_file_name=xl.cell(row, 1).value
                                                    , criterion_start={"years": xl.cell(row, 6).value,
                                                                       "season": xl.cell(row, 7).value,
                                                                       "max_min": xl.cell(row, 8).value,
                                                                       "add_name": xl.cell(row, 9).value}
                                                    , tables=xl.cell(row, 2).value
                                                    , param=xl.cell(row, 4).value
                                                    , sel=xl.cell(row, 3).value
                                                    , calc=xl.cell(row, 5).value)
                ImportFromModel.number += 1
                CS.dict_import_model[ImportFromModel.number] = import_from_model
        # --------------------------------------------------------------------------------
    elif tip == 'XL->RastrWin' and calc_val != "Параметры импорта из файлов RastrWin":
        name_files = ""
        dict_param_column = {}  # {"pn":10-столбец}
        # шаг по колонкам и запись в словарь всех столбцов для корр
        for column_name_file in range(2, xl.max_column + 1):
            if xl.cell(1, column_name_file).value not in ["", None]:
                name_files = xl.cell(1, column_name_file).value.split("|")  # list [name_file, name_file]
            if xl.cell(2, column_name_file).value:
                duct_add = False
                for name_file in name_files:
                    if name_file in [RG.Name_Base, "*"]:
                        duct_add = True
                    if "*" in name_file and len(name_file) > 7:
                        pattern_name = re.compile("\[(.*)\]\[(.*)\]\[(.*)\]\[(.*)\]")
                        match = re.search(pattern_name, name_file)
                        if match.re.groups == 4:
                            if RG.test_name(dict_uslovie={"years": match[1], "season": match[2],
                                                          "max_min": match[3], "add_name": match[4]},
                                            info=f"\tcor_xl, условие: {name_file}, "):
                                duct_add = True
                if duct_add:
                    dict_param_column[xl.cell(2, column_name_file).value] = column_name_file
        logging.debug("\t" + str(dict_param_column))

        if len(dict_param_column) == 0:
            logging.info(f"\t {RG.Name_Base} НЕ НАЙДЕН на листе {sheet} книги {excel_file_name}")
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
                                    cor(str(kkey), f"{param}={new_val}", True)
                                else:
                                    cor(str(kkey), f"{param}={param}{calc_vals[calc_val]}{new_val}", True)
                            else:
                                cor_pop(kkey, new_val)  # изменить потребление, CS.pop_save_pn


def cor_pop(zone, new_pop, task_save=None):
    """ район("na=3", "npa=2" или "no=1", значение потребления, задание на сохранение нагрузки узлов)"""
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


def cor(keys, tasks, cor_print=True):
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
                grup_cor("node", task_equally[0], set_row, task_equally[1])
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
                grup_cor("vetv", task_equally[0], set_row, task_equally[1])

            elif key_equally[0] in ["g", "Num"]:  # генератор
                set_row = "" if (key == "g" or key == "Num") else "Num=" + key_equally[1]
                grup_cor("Generator", task_equally[0], set_row, task_equally[1])

            elif key_equally[0] == "no":  # объединене
                set_row = '' if key == "no" else "no=" + key_equally[1]
                grup_cor("darea", task_equally[0], set_row, task_equally[1])

            elif key_equally[0] == "na":  # район
                set_row = '' if key == "na" else "na=" + key_equally[1]
                grup_cor("area", task_equally[0], set_row, task_equally[1])

            elif key_equally[0] == "npa":  # территория
                set_row = '' if key == "npa" else "npa=" + key_equally[1]
                grup_cor("area2", task_equally[0], set_row, task_equally[1])

            elif key_equally[0] == "nga":  # нагрузочные группы
                set_row = '' if key == "nga" else "nga=" + key_equally[1]
                grup_cor("ngroup", task_equally[0], set_row, task_equally[1])

    if cor_print:
        logging.info(f"\t cor {keys},  {tasks}")


def grup_cor(tabl, param, viborka, formula):
    """групповая коррекция (таблица, параметр корр, выборка, формула для расчета параметра)"""
    # global rastr
    if rastr.tables.Find(tabl) < 0:
        logging.error(f"\tВНИМАНИЕ! в rastrwin не загружена таблица {tabl}")
        return False
    ptabl = rastr.tables(tabl)
    if ptabl.cols.Find(param) < 0:
        logging.error(f"ВНИМАНИЕ! в таблице {tabl} нет параметра {param}")
        return False
    pparam = ptabl.cols.item(param)
    ptabl.setsel(viborka)
    pparam.Calc(formula)
    return True


def rgm(txt=""):
    GS.kod_rgm = rastr.rgm("")
    if GS.kod_rgm == 1: GS.kod_rgm = rastr.rgm("")
    if GS.kod_rgm == 1: GS.kod_rgm = rastr.rgm("")
    if GS.kod_rgm == 1: GS.kod_rgm = rastr.rgm("p")
    if GS.kod_rgm == 1: GS.kod_rgm = rastr.rgm("p")
    if GS.kod_rgm == 1:
        if GS.calc_set == 1:
            CS.set['collapse'] += f" {RG.Name_base}: {txt}/"
        logging.error(f"расчет режима: {txt} !!!РАЗВАЛИЛСЯ!!!")
    else:
        if txt:
            logging.debug(f"\tрасчет режима: {txt}")


def control_rg2(dict_task):
    """  контроль  dict_task = {'node': True, 'vetv': True, 'Gen': True, 'section': True, 'area': True, 'area2': True,
        'darea': True, 'sel_node': "na>0"}  """
    node = rastr.tables("node")
    branch = rastr.tables("vetv")
    generator = rastr.tables("Generator")
    chart_pq = rastr.tables("graphik2")
    graph_it = rastr.tables("graphikIT")

    rgm("control_rg2")
    # НАПРЯЖЕНИЯ
    if dict_task["node"]:
        logging.info("\tКонтроль напряжений.")

        uh = [6, 10, 35, 110, 220, 330, 500, 750]  # номинальные напряжения
        umin_n = [5.8, 9.7, 32, 100, 205, 315, 490, 730]  # минимальные нормальное напряжения для контроля
        unr = [7.2, 12, 42, 126, 252, 363, 525, 787]  # наибольшее работчее напряжения

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
        rastr.CalcIdop(RG.GradusZ, 0.0, "")
        logging.info("\tКонтроль токовой загрузки, расчетная температура: " + RG.Gradus)
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

    if dict_task['area']: control_pop('area')
    if dict_task['area2']: control_pop('area2')
    if dict_task['darea']: control_pop('darea')


def control_pop(zone):
    """zone =  'darea', 'area', 'area2'"""
    key_sone = {'darea': 'no', 'area': 'na', 'area2': 'npa'
        , 'darea_pop': 'pp', 'area_pop': 'pop', 'area2_pop': 'pop'
        , 'darea_name': 'обединений', 'area_name': 'районов', 'area2_name': 'территорий'}

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


def sheet_exists(сur_workbook, sh_name):  # проверка существования лист в книге
    for sheeti in сur_workbook.Sheets:
        if sheeti.name == sh_name:
            return True
    return False


class ImportFromModel:
    """ импорта данных из файлов .rg2, .rst и др.
     import_file_name = полное имя файла
     criterion_start={"years": "","season": "","max_min": "", "add_name": ""} условие выполнения
     tables= таблица для импорта, нр "node;vetv"
     param= параметры для импорта: "" все параметры или перечисление, нр 'sel,sta'(ключи не обязательно)
     sel= выборка нр "sel" или "" - все
     calc= {"обновить": 2 , "загрузить": 1, "присоединить": 0, "присоединить-обновить": 3}"""
    number = 0
    calc_str = {"обновить": 2, "загрузить": 1, "присоединить": 0, "присоединить-обновить": 3}

    def __init__(self, import_file_name='', criterion_start={}, tables='', param='', sel='', calc='2'):
        if not os.path.exists(import_file_name):
            logging.error("Ошибка в задании, не найден файл: " + import_file_name)
            self.import_file_name = ''
        else:
            self.import_file_name = import_file_name
            self.basename = os.path.basename(import_file_name)
            self.criterion_start = criterion_start
            self.tables = tables.split(";")  # разделить на ["таблицы"]
            self.param = []
            self.sel = sel if sel != None else ''
            if type(calc) == int:
                self.calc = calc
            else:
                if calc in self.calc_str:
                    self.calc = self.calc_str[calc]
                else:
                    logging.error("Ошибка в задании, не распознано задание calc ImportFromModel: " + str(calc))
                    self.import_file_name = ''
            self.file_csv = []

            for tabl in self.tables:
                self.file_csv.append(f"{CS.set['folder_temp']}\\{self.basename}_{tabl}_{str(self.number)}.csv")
                self.param.append(param)

    def export_csv(self):
        """Экспорт данных из файла в csv"""
        if self.import_file_name != '':
            rastr.Load(1, self.import_file_name, GS.set_save['шаблон ' + self.import_file_name[-3:]])
            logging.info("\tЭкспорт из файла:" + self.import_file_name)
            for index in range(len(self.tables)):
                if not self.param[index]:  # если все параметры
                    self.param[index] = all_cols(self.tables[index])
                else:
                    if rastr.Tables(self.tables[index]).Key not in self.param[index]:
                        self.param[index] += ',' + rastr.Tables(self.tables[index]).Key

                logging.info(f"\t\tТаблица: {self.tables[index]}. Выборка: {self.sel}"
                             + f"\n\t\tПараметры: {self.param[index]}"
                             + f"\n\t\tФайл CSV: {self.file_csv[index]}")
                export_CSV(self.file_csv[index], self.tables[index], self.param[index], self.sel)

    def import_csv(self):
        """Импорт данных из csv в файла"""
        if self.import_file_name != '':
            logging.info("\tИмпорт из файла:" + self.import_file_name)
            for index in range(len(self.tables)):
                if RG.test_name(self.criterion_start, info='ImportFromModel'):
                    for index in range(len(self.tables)):
                        logging.info(f"\t\tТаблица: {self.tables[index]}. Выборка: {self.sel}. тип: {str(self.calc)}"
                                     + f"\n\t\tФайл CSV: {self.file_csv[index]}"
                                     + f"\n\t\tПараметры: {self.param[index]}")
                        import_CSV(self.file_csv[index], self.tables[index], self.param[index], self.calc)


def export_CSV(file, table, param, vibor):
    tab = rastr.Tables(table)
    tab.setsel(vibor)
    tab.WriteCSV(1, file, param, ";")  # 0 дописать, 1 заменить


def import_CSV(file, table, param, type_add):
    """{"обновить": 2 , "загрузить": 1, "присоединить": 0, "присоединить-обновить": 3}"""
    tab = rastr.Tables(table)
    tab.ReadCSV(type_add, file, param, ";", "")


def all_cols(tab):
    """Возвращает все колонки таблицы: 'ny,uhom....'"""
    cls = rastr.Tables(tab).Cols
    cols_list = []
    for col in range(cls.Count):
        if cls(col).Name not in ["kkluch", "txt_zag", "txt_adtn_zag", "txt_ddtn", "txt_adtn", "txt_ddtn_zag"]:
            # print(str(cls(col).Name))
            cols_list.append(str(cls(col).Name))
    return ','.join(cols_list)


class PrintXL:
    """класс печать данных в excel"""

    #  _log  значит протокол для сводной
    #  _p  значит параметры

    def __init__(self):  # добавить листы и первая строка с названиями
        global CS
        self.list_name = ["name_rg2", "год", "лет/зим", "макс/мин", "доп_имя1", "доп_имя2", "доп_имя3"]
        self.book = Workbook()
        #  создать лист xl и присвоить ссылку на него
        for key in CS.set['set_printXL']:
            if CS.set['set_printXL'][key]['add']:
                CS.set['set_printXL'][key]["sheet"] = self.book.create_sheet(key + "_log")
                # записать первую строку параметров
                header_list = self.list_name + CS.set['set_printXL'][key]['par'].split(',')
                CS.set['set_printXL'][key]["sheet"].append(header_list)

        if CS.set['print_parameters']['add']:
            CS.set['print_parameters']["sheet"] = self.book.create_sheet('parameters')

        if CS.set['print_balance_q']['add']:
            CS.set['print_balance_q']["sheet"] = self.book.create_sheet("balance_Q")
            self.balans_Q_X0 = 5

    def add_val(self):
        logging.info("\tВывод данных из моделей в XL")
        if RG.Name_st == "не подходит":
            DopName_list = ['-'] * 3
        else:
            DopName_list = RG.DopName[:3]
            if len(DopName_list) < 3:
                DopName_list += ['-'] * (3 - len(DopName_list))
        list_name_z = [RG.Name_Base, RG.god, RG.name_list[1], RG.name_list[2]] + DopName_list

        for key in CS.set['set_printXL']:
            if not CS.set['set_printXL'][key]['add']:
                continue
            # проверка наличия таблицы
            if rastr.Tables.Find(CS.set['set_printXL'][key]['tabl']) < 0:
                logging.error("В RastrWin не загружена таблица " + CS.set['set_printXL'][key]['tabl'])
                CS.set['set_printXL'][key]['add'] = False
                continue

            # принт данных из растр в таблицу для СВОДНОЙ
            r_table = rastr.tables(CS.set['set_printXL'][key]['tabl'])
            sheet = CS.set['set_printXL'][key]["sheet"]
            param_list = CS.set['set_printXL'][key]['par'].split(',')
            param_list = [param_list[i] if r_table.cols.Find(param_list[i]) > -1 else '-' for i in
                          range(len(param_list))]

            setsel = CS.set['set_printXL'][key]['sel'] if CS.set['set_printXL'][key]['sel'] else ""
            r_table.setsel(setsel)
            index = r_table.FindNextSel(-1)
            while index >= 0:
                sheet.append(
                    list_name_z + [r_table.cols.item(val).ZN(index) if val != '-' else '-' for val in param_list])
                index = r_table.FindNextSel(index)

        if CS.set['print_parameters']['add']:
            dict_tables = {'n': 'node', 'v': 'vetv', 'g': 'Generator', 'na': 'area', 'npa': 'area2', 'no': 'darea'
                , 'nga': 'ngroup', 'ns': 'sechen'}
            sheet = CS.set['print_parameters']["sheet"]
            if sheet.max_row == 1:
                one_row_list = self.list_name[:]
            val_list = list_name_z[:]

            for task_i in CS.set['print_parameters']['sel'].replace(' ', '').split(';'):
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

        if CS.set['print_balance_q']['add']:
            pass

    def finish(self) -> None:
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
                    self.book.create_sheet(sheet_name.replace('log', 'сводая'))
                    sheet_couple[sheet_name] = sheet_name.replace('log', 'сводая')

        self.book.save("xl.xlsx")
        self.book = None

        self.excel = win32com.client.Dispatch("Excel.Application")
        self.excel.ScreenUpdating = False  # обновление экрана
        # self.excel.Calculation = -4135  # xlCalculationManual
        self.excel.EnableEvents = False  # отслеживание событий
        self.excel.StatusBar = False  # отображение информации в строке статуса excel

        self.wbook = self.excel.Workbooks.Open(os.getcwd() + '\\xl.xlsx')
        for n in range(self.wbook.sheets.count):
            if self.wbook.sheets[n].Name in sheet_couple:
                self.pivot_tables(self.wbook.sheets[n].Name, sheet_couple[self.wbook.sheets[n].Name])

        self.excel.Visible = True
        self.excel.ScreenUpdating = True  # обновление экрана
        self.excel.Calculation = -4105  # xlCalculationAutomatic
        self.excel.EnableEvents = True  # отслеживание событий
        self.excel.StatusBar = True  # отображение информации в строке статуса excel
        self.excel = None

        # excel_file = pd.read_excel('xl.xlsx')
        # report_table = excel_file.pivot_table(index='name',
        #                                       columns='name_rg2',
        #                                       values='na',
        #                                       aggfunc='max').round(0)
        #
        # report_table.to_excel('xl2.xlsx',
        #                       sheet_name='Report',
        #                       startrow=4)

        # if CS.print_balans_Q:
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
        if CS.set['folder_result']:
            now = datetime.now()
            file_name = CS.set['folder_result'] + f"\\коррекция файлов ({now.strftime('%d-%m-%Y %H-%M')}).xlsm"
            self.wbook.SaveAs(file_name, 52)  # 52 *.xlsm 51 *.xlsx

    def pivot_tables(self, s_log: str, s_pivot: str) -> None:
        """создать сводную таблицу
        :param s_log: имя листа с исходной таблицей
        :param s_pivot: имя листа для вставки сводной"""
        for n in range(self.wbook.sheets.count):
            if s_log == self.wbook.sheets[n].Name:
                tab_log = self.wbook.sheets[n].ListObjects[0]
        rows = CS.set['set_printXL'][s_log[:-4]]['rows'].split(",")
        columns = CS.set['set_printXL'][s_log[:-4]]['columns'].split(",")
        values = CS.set['set_printXL'][s_log[:-4]]['values'].split(",")

        PTCache = self.wbook.PivotCaches().add(1, tab_log)  # создать КЭШ xlDatabase, ListObjects
        PT = PTCache.CreatePivotTable(s_pivot + "!R1C1", "Сводная " + s_log[:-4])  # создать сводную таблицу
        PT.ManualUpdate = True  # не обновить сводную
        # print(s_log, s_pivot)
        PT.AddFields(RowFields=rows, ColumnFields=columns, PageFields=["name_rg2"], AddToTable=False)
        # добавить поля фильтра которых нет в столбцах и строках
        # PT.AddFields RowFields:=poleRow_arr, ColumnFields:=poleCol_arr, PageFields:=Array("name_rg", "лет/зим", "макс/мин", "доп_имя1", "доп_имя2") #  добавить поля

        for val in values:
            PT.AddDataField(PT.PivotFields(val), val + " ", -4157)  # xlSum #  добавить поле в область значений
            # Field                      Caption             def формула расчета
            PT.PivotFields(val + " ").NumberFormat = "0"

        # .PivotFields("na").ShowDetail = True #  группировка
        PT.RowAxisLayout(1)  # xlTabularRow показывать в табличной форме!!!!
        if len(values) > 0:
            PT.DataPivotField.Orientation = 1  # xlRowField"Значения" в столбцах или строках xlColumnField

        # .DataPivotField.Position = 1 #  позиция в строках
        PT.RowGrand = False  # удалить строку общих итогов
        PT.ColumnGrand = False  # удалить столбец общих итогов
        PT.MergeLabels = True  # обединять одинаковые ячейки
        PT.HasAutoFormat = False  # не обновлять ширину при обнавлении
        PT.NullString = "--"  # заменять пустые ячейки
        PT.PreserveFormatting = False  # сохранять формат ячеек при обнавлении
        PT.ShowDrillIndicators = False  # показывать кнопки свертывания
        # PT.PivotCache.MissingItemsLimit = 0 # xlMissingItemsNone
        # xlMissingItemsNone для норм отображения уникальных значений автофильтра
        for row in rows:
            PT.PivotFields(row).Subtotals = [False, False, False, False, False, False, False, False, False, False,
                                             False, False]  # промежуточные итоги и фильтры
        for column in columns:
            PT.PivotFields(column).Subtotals = [False, False, False, False, False, False, False, False, False, False,
                                                False, False]  # промежуточные итоги и фильтры
        PT.ManualUpdate = False  # обновить сводную
        PT.TableStyle2 = ""  # стиль
        if s_log[:-4] in ["area", "area2", "darea"]:
            PT.ColumnRange.ColumnWidth = 10  # ширина строк
            PT.RowRange.ColumnWidth = 9
            PT.RowRange.Columns(1).ColumnWidth = 7
            PT.RowRange.Columns(2).ColumnWidth = 20
            PT.RowRange.Columns(3).ColumnWidth = 10
            PT.RowRange.Columns(6).ColumnWidth = 20
        PT.DataBodyRange.HorizontalAlignment = -4108  # xlCenter
        # .DataBodyRange.NumberFormat = "#,##0"
        # формат
        PT.TableRange1.WrapText = True  # перенос текста в ячейке
        PT.TableRange1.Borders(7).LineStyle = 1  # лево
        PT.TableRange1.Borders(8).LineStyle = 1  # верх
        PT.TableRange1.Borders(9).LineStyle = 1  # низ
        PT.TableRange1.Borders(10).LineStyle = 1  # право
        PT.TableRange1.Borders(11).LineStyle = 1  # внутри вертикаль
        PT.TableRange1.Borders(12).LineStyle = 1  #


def sel0(txt=''):
    """ Cнять отметку узлов, ветвей и генераторов"""
    rastr.Tables("node").cols.item("sel").Calc("0")
    rastr.Tables("vetv").cols.item("sel").Calc("0")
    rastr.Tables("Generator").cols.item("sel").Calc("0")
    if txt != '':
        logging.info("\tCнять отметку узлов, ветвей и генераторов")


def start():
    global GS
    global CS
    GS = GeneralSettings()
    if GS.calc_set == 1:
        CS = CorSettings()  # CS - это глобальный класс кор
        main_cor()  # korr
    if GS.calc_set == 2:
        mainRG()  # rashot
    GS.end_gl()


if __name__ == '__main__':
    visual_set = 1  # 1 задание через QT, 0 - в коде
    # https://docs.python.org/3/library/logging.html
    logging.basicConfig(filename="log_file.log", level=logging.DEBUG, filemode='w',
                        format='%(asctime)s %(levelname)s:%(message)s')  # debug, INFO, WARNING, ERROR и CRITICAL
    rastr = win32com.client.Dispatch("Astra.Rastr")

    if visual_set == 0:
        start()
    else:
        app = QtWidgets.QApplication([])# Новый экземпляр QApplication
        app.setApplicationName("Правка моделей RastrWin")
        ui_edit = EditWindow()# Сздание инстанса класса
        ui_edit.show()
        ui_set = SetWindow()
        # ui_set.show()
        sys.exit(app.exec_())# Запуск
