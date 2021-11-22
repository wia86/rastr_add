# v153 Макрос автоматизации расчетов УР и корекции Ведерников Иван wia86@mail.ru 2017-2021
#  скорость 2,5-5,5 сочетаний в сек

# СДЕЛАТЬ КОРР
# форма XL для редактирования

# СДЕЛАТЬ РЕЖИМЫ
# OTKL_po_spisku_XL_sub удалить?
#  если н-1 не считается то и н-2-3 не нада считать
#  автоматика
# табл КО проверить
#  опасные сечения - выделять на графике
#  откл СШ в н-2-3? если напряжение больше  или равно 500 кВ
#  ПЗ в ворд
#  ПРОТОКОЛ СОБЫТИЙ или код расчетов iKodUR=УР6*100000+УР5*10000+УР4*1000+УР3*100+УР2*10+УР1 нада? Или как быть уверенным что поле автоматики ликвидировалась перегрузка
#  OO транс

# dim html_RG2 , ExitDo_RG2 , IE_kform,IE_rform,  RG2_IE, visual_set , objFSO, RG  # *****общее*******
# dim GLK, PKC, KSC#  ****КОРР*******
# dim GL , GLR, RGR, Komb_List , mKO    , Comb     # КЛАССЫ ******** РАСЧЕТ******, spUniquizer

import win32com.client
import datetime
import time
from tkinter import *
from tkinter import messagebox as mb
import os

def start():
    GL = Global_class()
    # if GL.calc_set == 1: mainKor ()#  korr
    # if GL.calc_set == 2: mainRG () #  rashot
    GL.end_gl()

def r_print(txt):
    global LogFile
    LogFile.write (txt + "\n")
    # if GL.calc_set == 2: # ????? значит потом доработать
    #     if GLR.protokol_XL:
    #         GLR.protokol_XL_Sheets.Cells (GLR.protokol_XL_row,1).value = txt
    #         GLR.protokol_XL_row = GLR.protokol_XL_row + 1
    # elif GL.calc_set == 1:
    #     if GLK.printXL and GLK.print_log_xl :
    #         PKC.print_protokol_XL (txt)


class Global_class:  # GL. для хранения общих параметров
    calc_set = 1  # 1 -корректировать модели Global_kor_class   2-расчитать модели Global_raschot_class!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    N_rg2_File = 0  # счетчик расчетных файлов
    excel = None
    now = datetime.datetime.now()

    def __init__(self):
        global LogFile
        LogFile = open("LogFile.txt", 'w')  # ????? значит потом доработать
        # if visual_set == 1 :
        # calc_set = RG2_IE.Document.Script.calc_set

        self.excel = win32com.client.Dispatch("Excel.Application")
        self.word = win32com.client.Dispatch("word.Application")
        self.excel.ScreenUpdating = False  # обновление экрана
        # excel.Calculation = -4135 # xlCalculationManual
        self.excel.EnableEvents = False  # отслеживание событий
        self.excel.StatusBar = False  # отображение информации в строке статуса excel

        self.time_start = time.time()
        self.now_start = self.now.strftime("%d-%m-%Y %H:%M")

    def end_gl(self):  # по завершению макроса
        global LogFile
        if self.calc_set == 2:
            if (GLR.kol_test_da + GLR.kol_test_net) > 0:
                procenti = str(round(GLR.kol_test_net / (GLR.kol_test_da + GLR.kol_test_net) * 100))
            else:
                procenti = "0"

        if self.excel is not None:
            if self.excel.Workbooks.count > 0:
                self.excel.Visible = True
                self.excel.ScreenUpdating = True  # обновление экрана
                self.excel.Calculation = -4105  # xlCalculationAutomatic
                self.excel.EnableEvents = True  # отслеживание событий
                self.excel.StatusBar = True  # отображение информации в строке статуса excel
                self.excel = None

        result_info = f"РАСЧЕТ ЗАКОНЧЕН!\nНачало расчета {self.now_start} конец {self.now.strftime('%d-%m-%Y %H:%M')} \n Затрачено: {str(round(time.time() - self.time_start, 1)) } сек или {str(round((time.time() - self.time_start) / 60, 1))} мин"
        if self.calc_set == 2:
            result_info += f"\n Сочетаний отфильтровано: {str(GLR.kol_test_net)} из {str(GLR.kol_test_da + GLR.kol_test_net)} ({procenti} %)"
            result_info += f"\n Скорость расчета: {str(round(GLR.kol_test_da / (time.time() - self.time_start), 1))} сочетаний/сек."

        r_print(result_info)
        mb.showinfo("Инфо",result_info)
        LogFile.close()
# +++++++++++++++++++++++КОРРРР+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
class Global_kor_class:  # GLK. для хранения общих параметров  - КОРРЕКЦИЯ ФАЙЛОВ
    # dim sohran, razval, Folder_temp, Folder_csv_RG2, dictImpRg2 , kontrol_rg2,  export_xl ,import_export_xl , table , XL_table, tip_export_xl
    # dim paramN_rg2, paramV_rg2, param_area_rg2,param_node_area_rg2,param_node_area2_rg2, param_area2_rg2 , param_darea_rg2,paramPQ_rg2, param_graphikIT_rg2, param_nga_rg2, param_SHN_rg2, viborPQ
    # dim file_csv_N, file_csv_V, file_csv_area, file_csv_node_area, file_csv_node_area2, file_csv_G, file_csv_area2, file_csv_darea, file_csv_PQ, file_csv_graphikIT, file_csv_nga, file_csv_SHN#  не используются в Global_kor_class
    # dim book_pop, sheet_pop, pop_DopName, tab_pop,tab_pop_z,tab_pop_name_st,  File_pop, sheet_pop_name,pop_save_pn , import_PQN_XL,  NPQ_Excel, NPQ_Sheets, import_PQN_XL_name_st
    # dim AutoShuntForm , AutoShuntIzm,AutoShuntFormSel , AutoShuntIzmSel, printXL , IE_bloki , IE_CB_np_zad_sub , IE_CB_name_txt_korr, IE_CB_uhom_korr_sub , IE_CB_SHN_ADD
    # dim korr_papka_file, KIzFolder, KVFolder, name_izm , KFiltr_file,KUslovie_file, print_log_xl , kontrol_rg2_zad, blok_nf, blok_kf, blok_ImpRg2 ,book_NPQ_Excel, AutoShuntSel , print_save
    # dim print_sech, print_area, print_area2, print_darea, print_parametr, parametr_vibor, print_balans_Q, balans_Q_vibor , print_tab_log_row, print_tab_log_col, print_tab_log_val
    # dim dict_tabl, print_tabl, print_tabl_name,print_vibor, print_param , print_tab_log , print_tab_log_ar , setsel_sech , setsel_area ,    setsel_area2,    setsel_darea , calc_PQN
    #def __init__(self):
    korr_papka_file = 1  # 1 папка,  0 текущий файл- перед запуском сохранить файл(тк этот файл перед выполнением перезагрузится и сохранится в temp)!!!!
    # 1 ПАПКА
    KIzFolder = r"I:\ОЭС Урала ТЭ\!КПР ХМАО ЯНАО ТО\Модели2\v118"  # откуда берем файлы режима.rg2
    KVFolder = r"I:\ОЭС Урала ТЭ\!КПР ХМАО ЯНАО ТО\Модели2\v118 УТ для ТО"  # куда копируем измененные файлы режима.rg2 if GLK.pop_save_DopName = 1 and GLK.tab_pop = 1 : GLK.sohran = 0
    name_izm = ["", ")",",МДП_37_У-Т)"] # "" ничего ("добавить в конце имени файла","заменить в имени","на что заменить")  (-41°C)
    KFiltr_file = 0  # 0 все файлы 5 - 5 файлов расчитываем, -1 - выборка см ниже
    KUslovie_file = ["2021-2027", "", "",  ""] # "" все или указать какие нр array ("2019,2021-2027","зим","мин","1°C,МДП") (год, зим, макс, доп имя разделитель , или ;)
    # ----------------------------------------------------------------------------------------------------------------------------
    import_PQN_XL = False  # True(False) импорт pn,qn из excel
    import_PQN_XL_name_st = False  # True стандартное имя нр 2020 зим макс ,False  иначе полное название файла без расширения
    NPQ_Excel = "H:\ОЭС Урала без ТЭ\Свердл_ЭС\КПР СЭ 2021\Модели\ТАБЛИЦА НАГРУЗОК ПО УЗЛАМ4 по районам.xlsm"
    NPQ_Sheets = "ny!"
    calc_PQN = 2  # 1 заменить, 2 прибавить, 3 вычесть, 4 умножить,
    # ----------------------------------------------------------------------------------------------------------------------------
    tab_pop = False  # True(False)  задать потребление по таблице XL и заполнить поле pop_zad
    # первая строка  название, нр "2019 зим макс", первый столбец номер района:na=1, терр npa=1, объединение: no=1
    tab_pop_z = False  # True только заполнить поле pop_zad без корректировки потребления, False корректировки потребления
    tab_pop_name_st = True  # True стандартное имя нр 2020 зим макс ,False  иначе полное название файла без расширения
    pop_DopName = True  # 1 проверить совпадение доп имени (в XL 2 строка под названием режима) в название файла - только RG.DopName (0)

    File_pop = "H:\ОЭС Урала без ТЭ\Свердл_ЭС\КПР СЭ 2021\Модели\ТАБЛИЦА НАГРУЗОК ПО УЗЛАМ3 по районам.xlsm"  # имя файла
    sheet_pop_name = "zad pop ЭС финиш"  # имя листа  zad pop гост юз
    pop_save_pn = ""  # "" - нет ; "1 2" номера узлов через пробел для отметки (или na=1 можно додедать) или "sel" если узлы уже были отмечены  - запрет на изменение pn qn
    # ----------------------------------------------------------------------------------------------------------------------------
    import_export_xl = False  # False нет, True  import или export из xl в растр
    table = "Generator"  # нр "oborudovanie"
    export_xl = True  # False нет, True - export из xl в растр
    XL_table = [r"C:\Users\User\Desktop\1.xlsx", "Generator"]  # полный адрес и имя листа
    tip_export_xl = 1  # 1 загрузить, 0 присоединить 2 обновить
    # ----------------------------------------------------------------------------------------------------------------------------
    AutoShuntForm = False  # False нет, True сущ bsh записать в автошунт
    AutoShuntFormSel = "(na>0|na<13)"  # строка выборка узлов
    AutoShuntIzm = False  # False нет, True вкл откл шунтов  autobsh
    AutoShuntIzmSel = "(na>0|na<13)"  # строка выборка узлов
    # что бы узел с скрм  вкл и отк этот  сопротивление единственной ветви r+x<0.2 и pn=qn=0
    # ----------------------------------------------------------------------------------------------------------------------------
    kontrol_rg2 = True  # False нет, True проверка  напряжений в узлах; дтн  в линиях(rastr.CalcIdop по GradusZ); pmax pmin относительно P у генераторов и pop_zad у территорий, объединений и районов; СЕЧЕНИЯ
    kontrol_rg2_zad = [True, True, True, False, False, False, False, "(na>0&na<13)"]
    #  False нет  True да           (наряжений 0, токов 1, генераторы 2 , сечений 3 , район 4  , территория 5 , объединение 6, выботка в таблице узлы "na=1|na=8)" 7)
    # ----------------------------------------------------------------------------------------------------------------------------
    printXL = True  # False нет, True
    #                             для ид сводной
    print_sech = True
    setsel_sech = ""  # сечения !!!!!!!!загрузить файл сечения !!!!!!!!
    print_area = False
    setsel_area = ""
    print_area2 = False
    setsel_area2 = ""
    print_darea = False
    setsel_darea = ""
    print_tab_log = False
    print_tab_log_ar = ["Generator", "Num,Name,sta,Node,P,Pmax,Pmin,value","Num>0"]  # для сводной из любой таблицы растр нр array("Generator" ,"P,Pmax" или ""все параметры, "Num>0" выборка)
    print_tab_log_row = "Num,Name"  # поля строк в сводной
    print_tab_log_col = "год,лет/зим,макс/мин,доп_имя1,доп_имя2"  # поля столбцов в сводной
    print_tab_log_val = "P,Pmax"  # поля значений в сводной

    print_parametr = False
    parametr_vibor = "vetv=42,48,0|43,49,0|27,11,3/r|x|b; ny=8|6/pg|qg|pn|qn"
    # вывод заданных параметров в следующем формате "vetv=42,48,0|43,49,0|27,11,3/r|x|b; ny=8|6/pg|qg|pn|qn"
    # таблица: ny-node,vetv-vetv,Num-Generator,na-area,npa-area2,no-darea,nga-ngroup,ns-sechen
    print_tabl = False  # сводная параметров из любой таблицы (например принт всех нагрузок во всех моделях)
    print_tabl_name = "vetv"  # node vetv Generator
    print_vibor = ""  # "(na<12|na=333)&(pn>0|qn>0|pg>0|qg>0)"print_vibor   = "(ny=936)&(pn>0|qn>0|pg>0|qg>0)" #
    print_param = "x"  # pn,qn печать таблиц нагрузок по всем моделям
    print_balans_Q = False
    balans_Q_vibor = "na=3012"  # БАЛАНС PQ_kor !!!0 тоже район,даже если в районах не задан "na>13&na<201"
    # ----------------------------------------------------------------------------------------------------------------------------
    blok_nf = 1  # начало
    blok_ImpRg2 = 0  # начало
    blok_kf = 1  # конец
    def init_visual_set(self):
        if visual_set == 1:  # #####################################################  HTML ###############################################################
            if IE_kform.ID_korr_file.Checked:
                self.korr_papka_file = 0  # 1 папка,  0 текущий файл- перед запуском сохранить файл(тк этот файл перед выполнением перезагрузится и сохранится в temp)!!!!
            if IE_kform.ID_korr_papka.Checked:
                self.korr_papka_file = 1  # 1 папка,  0 текущий файл- перед запуском сохранить файл(тк этот файл перед выполнением перезагрузится и сохранится в temp)!!!!
            # 1 ПАПКА
            self.KIzFolder = IE_kform.KIzFolder.value  # откуда берем файлы режима.rg2 нр 2020 зим макс, 2022 паводок мин (30°C ПЭВТ; МДП)
            self.KVFolder = IE_kform.KVFolder.value  # куда копируем измененные файлы режима.rg2
            self.name_izm = [IE_kform.name_izm_kon.value, IE_kform.name_izm_iz.value,
                             IE_kform.name_izm_na.value]  # "" ничего ("добавить в конце имени файла","заменить в имени","на что заменить")  (-41°C)
            if IE_kform.CB_KFiltr_file.checked:
                self.KFiltr_file = float(IE_kform.kol.value)  # 0 все файлы 5 - 5 файлов расчитываем, -1 - выборка см ниже
                if self.KFiltr_file == 0:
                    self.KFiltr_file = -1
                    KUslovie_file = [IE_kform.uslovie_file_god.value
                        , IE_kform.uslovie_file_zim_let.value
                        , IE_kform.uslovie_file_max_min.value
                        , IE_kform.uslovie_file_dop_name.value]  # "" все или указать какие нр array ("2019,2021-2027","зим","мин","1°C,МДП") (год, зим, макс, доп имя разделитель , или ;)
                else:
                    self.KFiltr_file = 0
            # end if
            # ----------------------------------------------------------------------------------------------------------------------------
            self.import_PQN_XL = IE_kform.CB_import_PQN_XL.checked  # True(False) импорт pn,qn из excel
            self.import_PQN_XL_name_st = IE_kform.CB_import_PQN_XL_name_st.checked  # True стандартное имя нр 2020 зим макс ,False  иначе полное название файла без расширения
            self.NPQ_Excel = IE_kform.T_NPQ_Excel.value
            self.NPQ_Sheets = IE_kform.T_NPQ_Sheets.value
            self.calc_PQN = float(IE_kform.calc_PQN_z.value)  # 1 заменить, 2 прибавить, 3 вычесть, 4 умножить,
            # ----------------------------------------------------------------------------------------------------------------------------
            # первая строка  название, нр "2019 зим макс", первый столбец номер района:na=1, терр npa=1, объединение: no=1
            tab_pop = IE_kform.CB_tab_pop.checked  # True(False)  задать потребление по таблице XL и заполнить поле pop_zad
            tab_pop_z = IE_kform.CB_tab_pop_z.checked  # True(False)  только заполнить поле pop_zad без корректировки потребления
            tab_pop_name_st = IE_kform.CB_tab_pop_name_st.checked  # # True стандартное имя нр 2020 зим макс ,False  иначе полное название файла без расширения
            pop_DopName = IE_kform.CB_pop_DopName.checked  # 1 проверить совпадение доп имени (в XL 2 строка под названием режима) в название файла - только RG.DopName (0)

            File_pop = IE_kform.T_File_pop.value  # имя файла
            sheet_pop_name = IE_kform.T_sheet_pop_name.value  # имя листа  zad pop гост юз
            pop_save_pn = IE_kform.T_pop_save_pn.value  # "" - нет ; "1 2" номера узлов через пробел для отметки (или na=1 можно додедать) или "sel" если узлы уже были отмечены  - запрет на изменение pn qn
            # ----------------------------------------------------------------------------------------------------------------------------
            import_export_xl = IE_kform.CB_import_export_xl.checked  # False нет, True  import или export из xl в растр
            table = IE_kform.T_table.value  # нр "oborudovanie"
            export_xl = IE_kform.CB_export_xl.checked  # False нет, True - export из xl в растр
            XL_table = array(IE_kform.T_XL_table1.value, IE_kform.T_XL_table2.value)  # полный адрес и имя листа
            tip_export_xl = float(
                IE_kform.tip_export_xl.value)  # 1 загрузить, 0 присоединить 2 обновить
            # ----------------------------------------------------------------------------------------------------------------------------
            AutoShuntForm = IE_kform.CB_AutoShuntForm.checked  # False нет, True сущ bsh записать в автошунт
            AutoShuntFormSel = IE_kform.AutoShuntFormSel.value  # строка выборка узлов
            AutoShuntIzm = IE_kform.CB_AutoShuntIzm.checked  # False нет, True вкл откл шунтов  autobsh
            AutoShuntIzmSel = IE_kform.AutoShuntIzmSel.value  # строка выборка узлов
            # что бы узел с скрм  вкл и отк этот  сопротивление единственной ветви r+x<0.2 и pn=qn=0
            # ----------------------------------------------------------------------------------------------------------------------------
            kontrol_rg2 = IE_kform.CB_kontrol_rg2.checked  # False нет, True проверка  напряжений в узлах; дтн  в линиях(rastr.CalcIdop по GradusZ); pmax pmin относительно P у генераторов и pop_zad у территорий, объединений и районов; СЕЧЕНИЯ
            kontrol_rg2_zad = array(IE_kform.CB_U.checked, IE_kform.CB_I.checked, IE_kform.CB_gen.checked,
                                    IE_kform.CB_s.checked, IE_kform.CB_na.checked, IE_kform.CB_npa.checked,
                                    IE_kform.CB_no.checked, IE_kform.kontrol_rg2_Sel.value)
            #  False нет  True да           (наряжений 0,        токов 1,              генераторы 2 ,           сечений 3 ,           район 4  ,          территория 5 ,                объединение 6, выботка в таблице узлы "na=1|na=8)" 7)

            printXL = IE_kform.CB_printXL.checked  # False нет, True
            #                             для ид сводной          для готовой сводной
            print_sech = IE_kform.CB_print_sech.checked
            setsel_sech = IE_kform.setsel_sech.value  # !!!!!!!!загрузить файл сечения !!!!!!!!
            print_area = IE_kform.CB_print_area.checked
            setsel_area = IE_kform.setsel_area.value
            print_area2 = IE_kform.CB_print_area2.checked
            setsel_area2 = IE_kform.setsel_area2.value
            print_darea = IE_kform.CB_print_darea.checked
            setsel_darea = IE_kform.setsel_darea.value

            print_tab_log = IE_kform.CB_print_tab_log.checked
            print_tab_log_ar = array(IE_kform.print_tab_log_ar_tab.value, IE_kform.print_tab_log_ar_cols.value,
                                     IE_kform.print_tab_log_ar_set.value)  # для сводной из любой таблицы растр нр array("Generator" ,"P,Pmax" или ""все параметры, "Num>0" выборка)
            print_tab_log_row = IE_kform.print_tab_log_rows.value  # поля строк в сводной
            print_tab_log_col = IE_kform.print_tab_log_cols.value  # поля столбцов в сводной
            print_tab_log_val = IE_kform.print_tab_log_vals.value  # поля значений в сводной

            print_parametr = IE_kform.CB_print_parametr.checked
            parametr_vibor = IE_kform.TA_parametr_vibor.value
            # вывод заданных параметров в следующем формате "vetv=42,48,0|43,49,0|27,11,3/r|x|b; ny=8|6/pg|qg|pn|qn"
            # таблица: ny-node,vetv-vetv,Num-Generator,na-area,npa-area2,no-darea,nga-ngroup,ns-sechen
            print_tabl = False  # сводная параметров из любой таблицы (например принт всех нагрузок во всех моделях)
            print_tabl_name = "vetv"  # node vetv Generator
            print_vibor = ""  # "(na<12|na=333)&(pn>0|qn>0|pg>0|qg>0)"print_vibor   = "(ny=936)&(pn>0|qn>0|pg>0|qg>0)" #
            print_param = "x"  # pn,qn печать таблиц нагрузок по всем моделям
            print_balans_Q = IE_kform.CB_print_balans_Q.checked
            balans_Q_vibor = IE_kform.balans_Q_vibor.value  # БАЛАНС PQ_kor !!!0 тоже район,даже если в районах не задан "na>13&na<201"
            # ----------------------------------------------------------------------------------------------------------------------------
            IE_CB_np_zad_sub = IE_kform.CB_np_zad_sub.Checked
            IE_CB_name_txt_korr = IE_kform.CB_name_txt_korr.Checked
            IE_CB_uhom_korr_sub = IE_kform.CB_uhom_korr_sub.Checked
            IE_CB_SHN_ADD = IE_kform.CB_SHN_ADD.Checked

            blok_nf = 0  # начало
            IE_bloki = IE_kform.CB_bloki.checked
            blok_ImpRg2 = 1 if IE_bloki or IE_kform.CB_ImpRg2.checked else blok_ImpRg2 = 0  # начало
            blok_kf = 0  # конец
        # end if
        # ----------------------------------------------------------------------------------
        # ПРОЧИЕ НАСТРОЙКИ
        print_save = True  # сохранить в папку KVFolder или KIzFolder
        print_log_xl = True  # выводить протокол в XL
        if printXL or tab_pop or import_PQN_XL:  GL.excel.Visible = True  # выбор: True False - видно после окончания работы макроса - для ускорения
        razval = ""

    #  выполняется если задание без IE, G2_IE_ON = 0
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
        GLK.dictImpRg2.Add(ImportClass.import_File & str(round(Rnd, 4) * 10000), ImportClass)

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
    # if RG.fTEST_etap (array ("2020","","","")) :
    # = otklonenie_seshen (nomer_sesh)   #   возвращает величину отклонения psech от  pmax   + превышение; - недобор
    # = rastr.Calc("sum,max,min,val","area","qn","vibor") - функция (vibor не может быть "")
    #  ПОТРЕБЛЕНИЕ CorPotrNa(raion,potr, ZadSave) CorPotrTER(raion,potr, ZadSave)#  територия CorOb(ob,potr, ZadSave)#  обединение
    #  ГЕНЕРАТОРЫ  PGen_cor ("sel")  # если мощность P больше Pmax то изменить мощность генератора  на Pmax, если P меньше Pmin но больше 0 - то на Pmin #  если P ген = 0 то отключить генератор, чтоб реактивка не выдавалась
    #  СЕЧЕНИЕ # KorSech  (ns,newp,vibor , tip, net_Pmin_zad) #  номер сеч, новая мощность в сеч (значение или "max" "min"), выбор корр узлов  (нр "sel"или "" - авто) ,  tip - "pn" или "pg", net_Pmin_zad #  1 не учитывать Pmin
    #  Qgen_node_in_gen_sub ()  #  посчитать Q ГЕН по  Q в узле
    # <<<настройки rastr>>>
    #  rastr.Tables("com_regim").cols.item("gen_p").Z(0) = 0 #    0- "да"; 1- "да"; 2- только Р; 3- только Q ///it_max  количество расчетов///neb_p точность расчектов////

    # <<<ТКЗ>>>
    #    Delet_node_VL_sub () #  удалить промежкточные точки на ЛЭП при отсутствии магнитной связи
