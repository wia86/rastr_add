import win32com.client
from openpyxl import Workbook, load_workbook
import sys
from PyQt5.QtGui import *
from PyQt5 import QtWidgets
from PyQt5.QtCore import *
from qt_form import Ui_MainWindow  # импорт ui: pyuic5 qt_form.ui -o qt_form.py
from datetime import datetime
import time
import os
import re
import logging
import webbrowser
from tkinter import *
from tkinter import messagebox as mb
import numpy as np

class EditWindow(QtWidgets.QMainWindow):
    def __init__(self, *args, **kwargs):
        super(EditWindow, self).__init__() #*args, **kwargs
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.ui.GB_sel_file.hide() # скрыть поля при старте
        self.ui.bloki_file.hide()
        self.ui.sel_import.hide()
        self.ui.FFV.hide()
        self.ui.GB_import_val_XL.hide()
        self.ui.GB_control.hide()
        self.ui.GB_prinr_XL.hide()
        self.ui.GB_sel_tabl.hide()
        self.ui.TA_parametr_vibor.hide()
        self.ui.balans_Q_vibor.hide()

        self.ui.CB_KFilter_file.clicked.connect(lambda: self.show_hide(self.ui.CB_KFilter_file, self.ui.GB_sel_file))
        self.ui.CB_bloki.clicked.connect(lambda: self.show_hide(self.ui.CB_bloki, self.ui.bloki_file))
        self.ui.CB_ImpRg2.clicked.connect(lambda: self.show_hide(self.ui.CB_ImpRg2, self.ui.sel_import))
        self.ui.CB_import_val_XL.clicked.connect(lambda: self.show_hide(self.ui.CB_import_val_XL, self.ui.GB_import_val_XL))
        self.ui.CB_Filtr_VTV.clicked.connect(lambda: self.show_hide(self.ui.CB_Filtr_VTV, self.ui.FFV))
        self.ui.CB_kontrol_rg2.clicked.connect(lambda: self.show_hide(self.ui.CB_kontrol_rg2, self.ui.GB_control))
        self.ui.CB_printXL.clicked.connect(lambda: self.show_hide(self.ui.CB_printXL, self.ui.GB_prinr_XL))
        self.ui.CB_print_tab_log.clicked.connect(lambda: self.show_hide(self.ui.CB_print_tab_log, self.ui.GB_sel_tabl))
        self.ui.CB_print_parametr.clicked.connect(lambda: self.show_hide(self.ui.CB_print_parametr, self.ui.TA_parametr_vibor))
        self.ui.CB_print_balans_Q.clicked.connect(lambda: self.show_hide(self.ui.CB_print_balans_Q, self.ui.balans_Q_vibor))
        self.ui.run_krg2.clicked.connect(lambda: start())

    def show_hide (self, source, receiver):
        if source.isChecked () == True:
            receiver.show()
        else:
            receiver.hide()

def start():
    global GS
    global GLK
    GS = GeneralSettings()
    if GS.calc_set == 1:
        GLK = GlobalKor()  # GLK - это глобальный класс коррр
        mainKor()  # korr
    if GS.calc_set == 2:
        mainRG ()  # rashot

    GS.end_gl()

class GeneralSettings:  # GS. для хранения общих параметров
    def __init__(self):
        self.calc_set = 1  # 1 -корректировать модели GlobalKor   2-расчитать модели Global_raschot_class!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        self.N_rg2_File = 0  # счетчик расчетных файлов
        self.now = datetime.now()

        # if visual_set == 1 :
        #     self.calc_set = RG2_IE.Document.Script.calc_set

        # self.excel = win32com.client.Dispatch("Excel.Application")
        # self.word = win32com.client.Dispatch("word.Application")
        # self.excel.ScreenUpdating = False  # обновление экрана
        # excel.Calculation = -4135 # xlCalculationManual
        # self.excel.EnableEvents = False  # отслеживание событий
        # self.excel.StatusBar = False  # отображение информации в строке статуса excel

        self.time_start = time.time()
        self.now_start = self.now.strftime("%d-%m-%Y %H:%M")
        self.set_save = {
            "шаблон rg2": r"I:\rastr_add\rastr example\pattern\режим.rg2",
            "шаблон rst": r"I:\rastr_add\rastr example\pattern\режим.rg2",
        }
        # self.set = {
        #     "calc_val": {1: "ЗАМЕНИТЬ", 2: "ПРИБАВИТЬ", 3: "ВЫЧЕСТЬ", 0: "УМНОЖИТЬ"}
        # }

    def end_gl(self):  # по завершению макроса
        if self.calc_set == 2:
            if (GLR.kol_test_da + GLR.kol_test_net) > 0:
                procenti = str(round(GLR.kol_test_net / (GLR.kol_test_da + GLR.kol_test_net) * 100))
            else:
                procenti = "0"

        # if self.excel is not None:
        #     if self.excel.Workbooks.count > 0:
        #         self.excel.Visible = True
        #         self.excel.ScreenUpdating = True  # обновление экрана
        #         self.excel.Calculation = -4105  # xlCalculationAutomatic
        #         self.excel.EnableEvents = True  # отслеживание событий
        #         self.excel.StatusBar = True  # отображение информации в строке статуса excel
        #         self.excel = None

        result_info = f"РАСЧЕТ ЗАКОНЧЕН!\nНачало расчета {self.now_start} конец {self.now.strftime('%d-%m-%Y %H:%M')} \n Затрачено: {str(round(time.time() - self.time_start, 1)) } сек или {str(round((time.time() - self.time_start) / 60, 1))} мин"
        if self.calc_set == 2:
            result_info += f"\n Сочетаний отфильтровано: {str(GLR.kol_test_net)} из {str(GLR.kol_test_da + GLR.kol_test_net)} ({procenti} %)"
            result_info += f"\n Скорость расчета: {str(round(GLR.kol_test_da / (time.time() - self.time_start), 1))} сочетаний/сек."
        logging.info(result_info)
        mb.showinfo("Инфо",result_info)
        webbrowser.open("log_file.log")

class GlobalKor:  # GLK. для хранения общих параметров  - КОРРЕКЦИЯ ФАЙЛОВ
    def __init__(self):

        self.set = {
            "folder_file": '',  # 'folder' папка,  'file' файл опеделяется кодом
                # 1 ПАПКА
                "KIzFolder": r"I:\rastr_add\test",  # расчетный файл или папка с файлами
                "KInFolder": r"I:\rastr_add\test_new",  # куда сохраняем измененные файлы, если "" не сохранять
                "save_file": True,  # опеделяется кодом
                "KFilter_file": False,  # False все файлы, True в соответствии с фильтром
                    "max_file_count": 9999,  # максимальное количество расчетных файлов
                    "KUslovie_file": {"years": "2026-2027",
                                      "season": "",
                                      "max_min": "мин",
                                      "add_name": ""
                                      },  # нр("2019,2021-2027","зим","мин","1°C;МДП") (год, зим, макс, доп имя разделитель , или ;)
            # ----------------------------------------------------------------------------------------------------------
            "import_val_XL": True,  # True(False) импорт pn,qn из excel
                "excel_cor_file": r"I:\rastr_add\test\примеры.xlsx",
                "excel_cor_sheet": "[XL->RastrWin][pop]",
            # ----------------------------------------------------------------------------------------------------------------------------
            "import_export_xl": False,  # False нет, True  import или export из xl в растр
                "table": "Generator",  # нр "oborudovanie"
                "export_xl": True,  # False нет, True - export из xl в растр
                "XL_table": [r"C:\Users\User\Desktop\1.xlsx", "Generator"],  # полный адрес и имя листа
                "tip_export_xl": 1,  # 1 загрузить, 0 присоединить 2 обновить
            # ----------------------------------------------------------------------------------------------------------------------------
            "AutoShuntForm": False,  # False нет, True сущ bsh записать в автошунт
                "AutoShuntFormSel": "(na>0|na<13)",  # строка выборка узлов
            "AutoShuntIzm": False,  # False нет, True вкл откл шунтов  autobsh
                "AutoShuntIzmSel": "(na>0|na<13)",  # строка выборка узлов
            # что бы узел с скрм  вкл и отк этот  сопротивление единственной ветви r+x<0.2 и pn:qn:0
            # ----------------------------------------------------------------------------------------------------------------------------
            "kontrol_rg2": True,  # False нет, True проверка  напряжений в узлах; дтн  в линиях(rastr.CalcIdop по GradusZ); pmax pmin относительно P у генераторов и pop_zad у территорий, объединений и районов; СЕЧЕНИЯ
                "kontrol_rg2_zad": [True, True, True, False, False, False, False, "(na>0&na<13)"],
                #  False нет  True да           (наряжений 0, токов 1, генераторы 2 , сечений 3 , район 4  , территория 5 , объединение 6, выботка в таблице узлы "na:1|na:8)" 7)
            # ----------------------------------------------------------------------------------------------------------------------------
            "printXL": True,  # False нет, True
                #                             для ид сводной
                "print_sech": True,
                "setsel_sech": "",  # сечения !!!!!!!!загрузить файл сечения !!!!!!!!
                "print_zone": False,
                "setsel_area": "",
                "print_zone2": False,
                "setsel_area2": "",
                "print_darea": False,
                "setsel_darea": "",
                "print_tab_log": False,
                    "print_tab_log_ar": ["Generator", "Num,Name,sta,Node,P,Pmax,Pmin,value","Num>0"],  # для сводной из любой таблицы растр нр array("Generator" ,"P,Pmax" или ""все параметры, "Num>0" выборка)
                    "print_tab_log_row": "Num,Name",  # поля строк в сводной
                    "print_tab_log_col": "год,лет/зим,макс/мин,доп_имя1,доп_имя2",  # поля столбцов в сводной
                    "print_tab_log_val": "P,Pmax",  # поля значений в сводной

                "print_parametr": False,
                    "parametr_vibor": "vetv:42,48,0|43,49,0|27,11,3/r|x|b; ny:8|6/pg|qg|pn|qn",
                    # вывод заданных параметров в следующем формате "vetv:42,48,0|43,49,0|27,11,3/r|x|b; ny:8|6/pg|qg|pn|qn"
                    # таблица: ny-node,vetv-vetv,Num-Generator,na-area,npa-area2,no-darea,nga-ngroup,ns-sechen
                "print_balans_Q": False,
                "balans_Q_vibor": "na:3012",  # БАЛАНС PQ_kor !!!0 тоже район,даже если в районах не задан "na>13&na<201"
            # ----------------------------------------------------------------------------------------------------------------------------
            "blok_nf": 1,  # начало
            "blok_ImpRg2": 0,  # начало
            "blok_kf": 1,  # конец
            # ПРОЧИЕ НАСТРОЙКИ
            "print_save": True,  # сохранить в папку KInFolder или KIzFolder
            "print_log_xl": True,  # выводить протокол в XL
            "razval": "",
        }
        global visual_set
        global ui_edit
        if visual_set == 1:
            self.set["KIzFolder"] = ui_edit.ui.T_IzFolder.toPlainText()  # QPlainTextEdit
            self.set["KInFolder"] = ui_edit.ui.T_InFolder.toPlainText()
            # ----------------------------------------------------------------------------------------------------------
            self.set["KFilter_file"] = ui_edit.ui.CB_KFilter_file.isChecked ()  # QCheckBox
            self.set["file_count"] = ui_edit.ui.D_count_file.value()   # QSpainBox
            self.set["KUslovie_file"]["years"] = ui_edit.ui.condition_file_years.text()   # QLineEdit text()
            self.set["KUslovie_file"]["season"] = ui_edit.ui.condition_file_season.currentText()   # QComboBox
            self.set["KUslovie_file"]["max_min"] = ui_edit.ui.condition_file_max_min.currentText()   #
            self.set["KUslovie_file"]["add_name"] = ui_edit.ui.condition_file_add_name.text()   #
            # ----------------------------------------------------------------------------------------------------------
            self.set["import_val_XL"] = ui_edit.ui.CB_import_val_XL.isChecked()
            self.set["excel_cor_file"] = ui_edit.ui.T_PQN_XL_File.toPlainText()
            self.set["excel_cor_sheet"] = ui_edit.ui.T_PQN_Sheets.text()
        for str_name in ["KIzFolder", "KInFolder", "excel_cor_file"]:
            if 'file:///' in self.set[str_name]:
                self.set[str_name] = self.set[str_name][8:]
def blok_n():
    SEL0()
    #  Del_sel ()
    RGM_kor("blok_n")

def ImpRg2():  # запуск ИД для импорта---------ИМПОРТ из модели-------------- выполняется после блока начала
    ImportClass = import_class()  #
    ImportClass.uslovie_start = array("", "", "", "")
    ImportClass.import_File = "I:\ОЭС Урала ТЭ\!КПР ХМАО ЯНАО ТО\Модели2\v117\без надстройков2\temp\2027 зим макс (0°C,МДП_37_У-Т) болчары 220.rg2"
    ImportClass.tabl = "node;vetv"
    ImportClass.param = array("",
                              "")  # "node;vetv;Generator", ("пусто-все или перечислить","набор парам")параметры так же можно ";"
    ImportClass.vibor = "sel"
    ImportClass.tip = "3"  # "2" обн, "1" заг, "0" прис, "3" обн-прис
    GLK.dictImpRg2.Add(ImportClass.import_File + str(round(Rnd, 4) * 10000), ImportClass)

def blok_k():
    sel0()
    RGM_kor("blok_k")

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
# SEL0 ()                    #  снять выделение узлов и ветвей  и генераторов
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
#  rastr.Tables("com_regim").cols.item("gen_p").Z(0) = 0 #    0- "да"; 1- "да"; 2- только Р; 3- только Q ///it_max  количество расчетов///neb_p точность расчектов////
# <<<ТКЗ>>>
#    Delet_node_VL_sub () #  удалить промежкточные точки на ЛЭП при отсутствии магнитной связи

class CurrentFile:   #  RG  RG. для хранения параметров текущего расчетного файла
    def __init__(self,rastr_file, Filter_file=False, Uslovie_file={}):
        global GLK
        global GLR
        if GS.calc_set == 1 and GLK.set["folder_file"] == 'folder':  # корр
            self.full_name = GLK.set["KIzFolder"] + '\\' + rastr_file  # вернуть имя с расширением "2020 зим макс.rg2"
            self.full_name_new = GLK.set["KInFolder"] + '\\' + rastr_file  # вернуть имя с расширением "2020 зим макс.rg2"

        self.Name = rastr_file  # вернуть имя с расширением "2020 зим макс.rg2"
        self.Name_Base = rastr_file[:-4]        #  вернуть имя без расширения "2020 зим макс"
        self.tip_file = rastr_file[-3:]  # rst или rg2
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

            if GS.calc_set == 2 : #  расчет режимов а не корр
                if GLR.gost58670:
                    if (self.kod_name_rg2 in [1, 2]) or ("ПЭВТ" in self.Name_Base):
                        self.temp_a_v_gost = True  #  зима + период экстремально высоких температур -ПЭВТ
        else:
            self.Name_st = "не подходит" #  отсеиваем файлы задание и прочее

        pattern_name = re.compile("\((.+)\)")
        match = re.search( pattern_name, self.Name_Base)
        if match:
            self.DopNameStr = match [1]

        if self.DopNameStr.replace(" ","") != "":
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
                self.Gradus = match[1].replace(',','.')
                self.GradusZ = float(self.Gradus)  # число
                self.txt_dop = "Расчетная температура " + self.Gradus  + " °C. "

        if GS.calc_set == 2:  # расчет режимов а не корр
            
            if self.kod_name_rg2 > 0:
                if GLR.zad_temperatura == 1:
                    if self.name_list [1] == "зим":
                        self.GradusZ = GLR.temperatura_zima
                    else:
                        self.GradusZ = GLR.temperatura_leto

                    self.Gradus = str (self.GradusZ)
                    self.txt_dop = "Расчетная температура " + self.Gradus  + " °C. "

                # for DopName_tek in self.DopName:
                #     for each ii in GLR.rg2_name_metka
                #         if trim (DopName_tek) = trim (ii (0)):
                #             txt_dop = txt_dop + ii (1)

                self.NAME_RG2_plus = self.SezonName + " " + self.god + " г"
                if self.txt_dop != "":
                    self.NAME_RG2_plus += ". " + self.txt_dop
                self.NAME_RG2_plus2 = self.SezonName + "(" + self.Gradus + " °C)"
                self.TEXT_NAME_TAB = GLR.tabl_name_OK1 + str (GLR.Ntabl_OK) + GLR.tabl_name_OK2 + self.SezonName +  " " +  self.god + " г. " + self.txt_dop

        if Filter_file and self.kod_name_rg2 > 0:
            if (''.join(Uslovie_file.values())).replace(' ','') != '':  #  условие не пустое ["","","",""]
                if not self.test_name (Uslovie_file):
                    self.Name_st = "не подходит"

    def test_name (self, dict_uslovie, info=""):  # возвращает истина если имя режима соответствует условию
        # нр dict_uslovie = {"years":"","season":"","max_min":"","add_name":""}-всегда истина,или год ("2020,2023-2025"), зим/лет("зим" "лет,зим"), макс/мин("макс" "мин"), доп имя("МДП:ТЭ-У" "41С,МДП:ТЭ-У")

        if dict_uslovie == None:
            return True
        if self.Name_st == "не подходит":
            return False
        if dict_uslovie['years'].replace(' ','') != "":  # ПРОВЕРКА ГОД
            fff = False
            for us in str_in_range(dict_uslovie['years']):
                if int(self.god) == us:
                    fff = True
            if fff == False:
                logging.debug(info + self.Name + f" Год '{self.god}' не проходит по условию: " + dict_uslovie['years'])
                return False

        if dict_uslovie['season'].replace(' ','') != "":  # ПРОВЕРКА "зим" "лет" "паводок"
            fff = False
            temp = dict_uslovie['season'].replace(' ','')
            for us in temp.split(","):
                if self.name_list[1] == us:
                    fff = True
            if fff == False:
                logging.debug(info + self.Name+f" Сезон '{self.name_list[1]}' не проходит по условию: "+dict_uslovie['season'])
                return False

        if dict_uslovie['max_min'].replace(' ', '') != "":  # ПРОВЕРКА "макс" "мин"
            if self.name_list[2] != dict_uslovie['max_min'].replace(' ', ''):
                logging.debug(info + self.Name + f" '{self.name_list[2]}' не проходит по условию: " + dict_uslovie['max_min'])
                return False

        if  dict_uslovie['add_name'].replace(' ', '') != "":  # ПРОВЕРКА (-41С;МДП:ТЭ-У)
            if ";" in dict_uslovie['add_name']:
                temp = dict_uslovie['add_name'].split (";")
            else:
                temp = dict_uslovie['add_name'].split (",")
            fff = False
            for us in temp:
                for DopName_i in self.DopName:
                    if DopName_i == us:
                        fff = True
            if fff == False:
                logging.debug(info + self.Name + f" Доп. имя {self.DopNameStr} не проходит по условию: "+dict_uslovie['add_name'])
                return False
        return True


def str_in_range (id_str):  # функция из "2021,2023-2025" делает [2021,2023,2024,2025]  np
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
        return False

def mainKor(): #  головная процедура
    global GLK
    global RG
    global GS
    global rastr

    # определяем корректировать файл или файлы в папке по анализу "KIzFolder"
    pattern_file = re.compile(".*\.(rg2|rst)$")
    match = re.search(pattern_file, GLK.set["KIzFolder"])
    if match:
        GLK.set["folder_file"] = 'file'  # если корр файл
        RG = CurrentFile(os.path.basename(GLK.set["KIzFolder"]))
    else:
        GLK.set["folder_file"] = 'folder'  # если корр папка

    if not os.path.exists(GLK.set["KIzFolder"]):
        mb.showerror("Ошибка в задании", "Не найден: " + GLK.set["KIzFolder"] + ", выход")
        return False

    # if GLK.blok_ImpRg2  = 1 or GLK.korr_papka_file = 0  :
    #     GLK.Folder_temp    = GLK.KIzFolder    + "\temp"
    #      Folder_add_sub ( GLK.Folder_temp ) #  создать папку
    #     LogFile = objFSO.OpenTextFile( GLK.Folder_temp & "\Принт корр " & str (Day(Now)) & "_" & str (Month(Now)) & "_" & str (Year(Now)) & "г " & str (Hour(Now)) & "ч_" & str (Minute(Now)) & "м_" & str (Second(Now)) & "c.log", 8, True)#  файл для записи протокола
    #     GLK.Folder_csv_RG2 = GLK.Folder_temp + "\csv_RG2"
    #      Folder_add_sub ( GLK.Folder_csv_RG2 ) #  создать папку
    #
    # if  GLK.printXL : PKC = print_korr_class : PKC.init_pkc ()# PKC - это printXL klass korrr
    # if visual_set = 1 :
    #     if IE_kform.CB_bloki.checked :
    #         if len (IE_kform.bloki_file.value) > 0 :
    #             logging.info( "загружен файл: " + IE_kform.bloki_file.value)
    #             executeGlobal (CreateObject("Scripting.FileSystemObject").openTextFile(IE_kform.bloki_file.value).readAll())
    #         else:
    #             logging.info( "!!!НЕ УКАЗАН АДРЕС ФАЙЛА ЗАДАНИЯ!!!" )
    #
    # if GLK.blok_ImpRg2  = 1 :
    #     if visual_set = 0 :
    #         GLK.ImpRg2 ()#  запуск ИД для импорта
    #     else:
    #         if IE_kform.CB_ImpRg2.checked : IE_ImpRg2 ()#  запуск ИД для импорта из задания IE
    #         if IE_kform.CB_bloki.checked  : ImpRg22 ()#  запуск ИД для импорта из файла блоки
    #
    #     FOR EACH dictImpRg2_i IN GLK.dictImpRg2.Items : dictImpRg2_i.init (): dictImpRg2_i.export_csv () : # next
    #
    # if  GLK.tab_pop :
    #     if not objFSO.FileExists(GLK.File_pop) : msgbox ( GLK.File_pop & " - не найден файл задать потреблние, выход" ):GS.end_gl (): exit def
    #     GLK.book_pop = GS.excel.Workbooks.Open (GLK.File_pop)
    #     if not SheetExists(GLK.book_pop, GLK.sheet_pop_name ) : msgbox ( GLK.sheet_pop_name  & " - не найден лист GLK.sheet_pop_name, выход " ):GS.end_gl (): exit def
    #     GLK.sheet_pop  = GLK.book_pop.Sheets(GLK.sheet_pop_name)

    if GLK.set["folder_file"] == 'folder':  # корр файлы в папке

        if GLK.set["KInFolder"] == "":  # не сохранять если конечная папка не указана
            GLK.set["save_file"] = False
        else:
            if not os.path.exists(GLK.set["KInFolder"]):
                logging.info("Создана папка: "+ GLK.set["KInFolder"])
                os.mkdir(GLK.set["KInFolder"])  # создать папку

        files = os.listdir(GLK.set["KIzFolder"])  # список всех файлов в папке
        rastr_files = list(filter(lambda x: x.endswith('.rg2')|x.endswith('.rst'), files))  # фильтр файлов

        for rastr_file in rastr_files:  # цикл по файлам .rg2 .rst в папке KIzFolder

            RG = CurrentFile(rastr_file, GLK.set["KFilter_file"], GLK.set["KUslovie_file"])
            if not GLK.set["KFilter_file"] or RG.Name_st != "не подходит":  # отключен фильтр или соответствует ему
                GS.N_rg2_File += 1
                if GLK.set["KFilter_file"]:
                    if GLK.set["max_file_count"] > 0:
                        GLK.set["max_file_count"] -= 1
                    else:
                        return False

                logging.info("Загружен файл: " + rastr_file)

                rastr.Load(1, RG.full_name, RG.shablon)  # загрузить
                cor_file()
                rastr.Save(RG.full_name_new, RG.shablon)
                logging.info("Файл сохранен: " + RG.full_name_new)
            else:
                logging.debug("Файл отклонен, не соответствует фильтру: " + rastr_file)

    elif GLK.set["folder_file"] == 'file':  # корр файл
        logging.info("Загружен файл: " + GLK.set["KIzFolder"])
        rastr.Load(1, GLK.set["KIzFolder"], RG.shablon)  # загрузить режим
        cor_file()
        rastr.Save(GLK.set["KInFolder"], RG.shablon)
        logging.info("Файл сохранен: " + GLK.set["KInFolder"])
    #
    # if  GLK.printXL :
    #   PKC.finish  ()
    # if GLK.razval != "" :
    #   GS.result_info = GS.result_info  &  "\n" & GLK.razval & "- ВНИМАНИЕ! имеются файлы которые развалились. "

def cor_file():
    # if GLK.blok_nf == 1:
    #     logging.info( "\t" & "***блок начала *** " )
    #     GLK.blok_n ()
    #     logging.info( "\t" & "*** конец блока начала *** " )
    # if visual_set == 1:
    #    if GLK.IE_bloki:
    #        logging.info( "\t" & "*** блок начала (bloki.rbs)***" )
    #        blok_n2 ()
    #        logging.info( "\t" & "*** конец блока начала (bloki.rbs)***" )
    #
    # if GLK.blok_ImpRg2  == 1:
    #     logging.info( "\t" & "импорт из файлов" )
    #     for EACH dictImpRg2_i in GLK.dictImpRg2.Items:
    #         dictImpRg2_i.import_csv ()

    if GLK.set["import_val_XL"]:  # задать параметры узла по значениям в таблице excel (имя книги, имя листа)
        sheets = re.findall("\[(.+?)\]", GLK.set["excel_cor_sheet"])
        for sheet in sheets:
            cor_xl(GLK.set["excel_cor_file"], sheet)
    # if GLK.import_export_xl:
    #     rastr_xl_tab (GLK.table , GLK.export_xl  , GLK.XL_table (0) , GLK.XL_table (1), GLK.tip_export_xl  )
    # if GLK.AutoShuntForm:
    #     add_AutoBsh (GLK.AutoShuntFormSel) #  процедура записывает из поля bsh в поле AutoBsh (выборка)
    # if GLK.AutoShuntIzm:
    #     AutoShunt_class_rec (GLK.AutoShuntIzmSel)#  процедура формирует Umin , Umax, AutoBsh , nBsh
    #     AutoShunt_class_kor ()  #  процедура меняет Bsh  и записывает GS.AutoShunt_list
    #     GS.AutoShunt_list = ""
    #
    # if visual_set = 1:
    #     if GLK.IE_CB_np_zad_sub:
    #         np_zad_sub ()   #  задать номер паралельности у ветвей с одинаковым ip i iq
    #     if GLK.IE_CB_name_txt_korr:
    #         name_txt_korr ()#   name_probel (r_tabl , r_tabl_pole), izm_bukvi(r_tabl , r_tabl_pole)#  удалить пробелы в начале и конце, заменить два пробела на один, английские менять на русские буквы
    #     if GLK.IE_CB_uhom_korr_sub:
    #         uhom_korr_sub ("")      #  исправить номинальные напряжения в узлах для ряда 6,10,35,110,150,220,330,500,750
    #     if GLK.IE_CB_SHN_ADD:
    #         SHN_ADD () #  добавить зависимость СХН
    #     if GLK.IE_bloki:
    #         logging.info( "\t" & "блок конца (bloki.rbs)" )
    #         blok_k2 ()
    #         logging.info( "\t" & "*** конец блока конца *** " )
    #
    # if GLK.blok_kf == 1:
    #     logging.info( "\t" & "*** блок конца *** " )
    #     GLK.blok_k ()
    #     logging.info( "\t" & "*** конец блока конца *** " )
    # if GLK.kontrol_rg2:
    #     kontrol_rg2_sub   (GLK.kontrol_rg2_zad) #        расчет режима и контроль параметров режима
    # if GLK.printXL:
    #     PKC.print_start ()


def cor_xl(excel_file_name, sheet):
    # задать параметры узла по значениям в таблице excel (имя книги, имя листа)
    logging.info (f"\t Задать значения из excel, книга: {excel_file_name}, лист: {sheet}")
    if not os.path.exists(excel_file_name):
        logging.error("Ошибка в задании, не найден файл: " + excel_file_name)
        return False

    wb = load_workbook(excel_file_name)
    if sheet not in wb.sheetnames:
        logging.error(f"Ошибка в задании, не найден лист: {sheet} в файле {excel_file_name}")
        return False

    bd_xl = wb[sheet]  # bd_xl.cell(1,1).value [строки][столбцы] bd_xl.max_row bd_xl.max_column
    name_files = ""
    calc_val = bd_xl.cell(1,1).value
    dict_param_column = {}  # {"pn":10-столбец}
    # шагаем по колонкам и записываем в словарь все столбцы для корр
    for column_name_file in range(2, bd_xl.max_column + 1):
        if bd_xl.cell(1, column_name_file).value not in ["", None]:
            name_files = bd_xl.cell(1, column_name_file).value.split("|")  # list [name_file, name_file]
        if bd_xl.cell(2, column_name_file).value:
            duct_add = False
            for name_file in name_files:
                if name_file in [RG.Name_Base, "*"]:
                    duct_add = True
                if "*" in name_file and len(name_file)>7:
                    pattern_name = re.compile("\[(.*)\]\[(.*)\]\[(.*)\]\[(.*)\]")
                    match = re.search(pattern_name, name_file)
                    if match.re.groups == 4:
                        if RG.test_name(dict_uslovie={"years": match[1], "season": match[2],
                                                       "max_min": match[3], "add_name": match[4]},
                                                        info=f"\tcor_xl, условие: {name_file}, "):
                            duct_add = True
            if duct_add:
                dict_param_column[bd_xl.cell(2, column_name_file).value] = column_name_file
    logging.debug("\t" + str(dict_param_column))

    if len(dict_param_column) == 0:
        logging.info(f"\t {RG.Name_Base} НЕ НАЙДЕН на листе {sheet} книги {excel_file_name}")
    else:
        calc_vals = {1: "ЗАМЕНИТЬ", 2: "+", 3: "-", 0: "*"}
        # 1: "ЗАМЕНИТЬ", 2: "ПРИБАВИТЬ", 3: "ВЫЧЕСТЬ", 0: "УМНОЖИТЬ"
        for row in range(3, bd_xl.max_row + 1):
            for param, column in dict_param_column.items():
                kkey = bd_xl.cell(row, 1).value
                if kkey not in [None, ""]:
                    new_val = bd_xl.cell(row, column).value
                    if new_val != None:
                        if param not in ["pop", "pp"]:
                            if calc_val == 1:
                                cor(str(kkey), f"{param}={new_val}", True)
                            else:
                                cor(str(kkey), f"{param}={param}{calc_vals[calc_val]}{new_val}", True)
                        else:
                            cor_pop(kkey, new_val)  # изменить потребление, GLK.pop_save_pn


def cor_pop(zone, new_pop, task_save=None):
    """ район("na=3", "npa=2" или "no=1", значение потребления, задание на сохранение нагрузки узлов)"""
    eps = 0.003 if new_pop < 50 else 0.0003  # точность расчета, *100=%
    react_cor = True  # менять реактивное потребление пропорционально
    if '=' not in str(zone):
        logging.error(f"Ошибка в задании, cor_pop /{zone}/{str(new_pop)}/{str(task_save)}")
        return False
    zone_id =  zone.partition('=')[0]
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

    t_node = rastr.Tables("node")
    t_zone=rastr.tables(name_zone[zone_id])
    t_zone.setsel (zone)
    ndx_z = t_zone.FindNextSel(-1)
    if zone_id == "no":
        t_area = rastr.tables("area")
        t_area.setsel(zone)
    if t_zone.cols.Find("pop_zad") > 0:
        t_zone.cols.Item("pop_zad").SetZ(ndx_z, new_pop)
    name_z = t_zone.cols.item('name').ZS(ndx_z)
    pop = t_zone.cols.item(name_zone['p_'+ zone_id]).ZS(ndx_z)
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
            logging.info(f"\tпотребление {name_z}({zone}) изменено на {round(pop)} (должно быть {str(new_pop)}, {str(i+1)} ит.)")
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
                grup_cor ("area2", task_equally[0], set_row, task_equally[1])

            elif key_equally[0] == "nga":  # нагрузочные группы
                set_row = '' if key == "nga" else "nga=" + key_equally[1]
                grup_cor("ngroup", task_equally[0], set_row, task_equally[1])

    if cor_print:
        logging.info(f"\t cor {keys},  {tasks}")


def grup_cor(tabl, param, viborka, formula):
    global rastr
    #  групповая коррекция (таблица, параметр корр, выборка, формула для расчета параметра)
    if rastr.tables.Find(tabl) < 0:
        logging.error(f"\tВНИМАНИЕ! в rastrwin не загружена таблица {tabl}")
        return False
    ptabl = rastr.Tables(tabl)
    if ptabl.cols.Find(param) < 0:
        logging.error(f"ВНИМАНИЕ! в таблице {tabl} нет параметра {param}")
        return False
    pparam = ptabl.cols.item(param)
    ptabl.setsel(viborka)
    pparam.Calc(formula)
    return True


def sheet_exists(сur_workbook, sh_name):  # проверка существования лист в книге
    for sheeti in сur_workbook.Sheets:
        if sheeti.name == sh_name:
            return True
    return False


if __name__ == '__main__':
    visual_set = 1  # 1 задание через QT, 0  - в коде
    # https://docs.python.org/3/library/logging.html
    logging.basicConfig(filename="log_file.log",level=logging.DEBUG, filemode='w' ,
                        format='%(asctime)s %(levelname)s:%(message)s') # debug, INFO, WARNING, ERROR и CRITICAL
    rastr = win32com.client.Dispatch("Astra.Rastr")

    if visual_set == 0:
        start()
    else:
        app = QtWidgets.QApplication([])
        app.setApplicationName("Правка моделей RastrWin")
        ui_edit = EditWindow()
        ui_edit.show()
        sys.exit(app.exec_())


