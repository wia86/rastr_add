class Global_raschot_class:  # GLR. GLR для хранения общих параметров  - РАСЧЕТ РЕЖИМОВ
    # dim Folder_temp, Folder_csv_RG2, Folder_wmf, Folder_rg2 ,N_rezh , kol_test_da, kol_test_net, Zad_RG2_name
    # dim IzFolder, vid_raschota, filtr_file, uslovie_file, MaxOtkl, MinOtkl,filtr_n2,Book_XL
    # dim viborka_comb, print_viborka, FilePeregruzXL,zad_temperatura, temperatura_zima , temperatura_leto
    # dim vibor_otkl,AutoShunt, FLAG_automatika, ris_tabl_add_PA , otkl_ssch,otkl_remont_shema,   max_tok_save,Zad_vse_RG2, vibor_raschot,viborka_raschot, node_pojas_analiz, node_pojas_zad, pqn_tranzit_min
    # dim node_zad, node_auto_flag, node_auto, node_zad_flag , Zad_RG2,  Zad_RG2_VIBOR_N,    protokol_XL_row, protokol_XL_Sheets ,  overload, Zad_RG2_VIBOR_V
    # dim vetv_auto_flag, vetv_auto, vetv_zad_flag, vetv_zad, TablF_const, rg2_name_metka, picture_add , risunok_rg2, risunok_word, risunok_zag, risunok_nr, risunok_par, DOC_save,DOC_visible,  tip_doc_file
    # dim protokol_XL, XL_save,XL_Visible,Tabl_otlk_kontrol, zagruz_add_tab, tabl_name_OK1, tabl_name_OK2, Ntabl_OK , EntireRow_OK, EntireColumn_OK
    # dim new_format_doc,open_doc_file,graf_shot,name_ris1, graf_load, God_1, Graf_1, God_2, Graf_2,  vivod_komb, DRVXL, dtn_uchastki
    # dim max_tok1, kol_otkl, nomer_ris_shag , number_pict,number_pict_first,file_wmf_size, orientation_doc, FileSpisok, Folder_add, Zad_RG2_name_k, word2, paramV, paramN , gost58670, TablF_zim , TablF_let
    # dim Y_VL , Y_VL_Trans , Y_VL_Trans_V, OTKL1_ndx_tek , Yn2, Xn2, MAX_X, MAX_Y, kluch_s, OTKL_zad_XL, XL_max_tok, XL_print_mKOO ,XL_print_mKOR , TablF_sheets, Peregruz_XL, Peregruz_XL_Sheets
    # dim XL_sheet , X_list, Y_list,  word_App, word_ris_in, word_ris_iz,Dict_iz, Dict_in,  naiden, naiden_ris, name_ris_zamena , zad_log , not_n2, PZ_word , ris_PZ_word

    #  = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = == = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
    def blok_sub():  # исполняется всегда

    # cor  ("3072,7136,0" , "sta=1")# снеж -кс-6
    #  sel0 ()
    #  Del_sel ()
    # def rg2_ris (tip_zad_ris) # 4 замена рисунков в ворд из rg2 # 5  замена рисунков в ворд из ворд

    #  сделать вывод адн в перегрузки!!!!!!!!!!!!!!!!!!!!!!!!
    def init():
        vid_raschota = 0  # 0 текущий файл, 1 папка, 2 ПЗ в ворд по вклвдке "перегрузки"

        # -12 файлы в ПАПКЕ
        IzFolder = "I:\ОЭС Урала ТЭ\!КПР ХМАО ЯНАО ТО\Модели2\v117\без надстройков2 Болдч"  # откуда берем файлы режима.rg2 нр 2020 зим макс, 2022 паводок мин (30°C ПЭВТ; МДП)
        filtr_file = 0  # 0 все файлы 5 - 5 файлов расчитываем, -1 - выборка см ниже
        uslovie_file = array("2020-2021", "", "",
                             "")  # "" все или указать какие нр array ("2018,2019,2020,2021,2022,2023,2024,2025,2026","зим","мин","30°C ПЭВТ;МДП") (год, зим, макс, доп имя разделитель , или ;)
        # дополнительные имена указанны в скобочках в названии файла через ; например (30°C ПЭВТ;МДП), можно выбрать одно или несколько
        # 012- обсчитать текущий файл или все файлы в ПАПКЕ
        #  УСЛОВИЯ ПЕРЕБОРА
        MinOtkl = 1
        MaxOtkl = 2  # кол-во одновременно отключенных элементов  от 1 до 3, нр от 0 до 0
        not_n2 = False
        otkl_ssch = True  # 0 не отключать СШ /1 не отключать сш при отключении 2-х элементов сети
        otkl_remont_shema = True  # False нет, True - в n-2 при ремонте ветви учитывать ремонтную схему(remont_add),а otkl_add не учитывать, при отключении учитывать доп откл (otkl_add)
        gost58670 = True  # 1 - учет ГОСТА 58670  если темпертура а-в то в н-2 не выводить перегрузку если превышен ддтн(мдн) но не превышено адтн(адн) , н-3 только для температур ГД
        filtr_n2 = True  # True вкл, False выкл
        viborka_comb = 10  # 0 нет или сочетание  учитывать если при отключении любого одного элемента из сочетания сумма изменения загрузки других более %
        print_viborka = False  # печать матрицы откл - откл

        # --2 по перегрузкам в ЭКСЕЛЕ
        FilePeregruzXL = array(
            "H:\ОЭС Урала без ТЭ\Свердл_ЭС\КПР СЭ 2021\Модели\РМ Свердловская область v22 корр по ТУ\протокол расчета n-1,2(25_10_2021г 18ч_14м_51c, zad 3 вост 12 сухолог из 2026.rg2).xlsm",
            "перегрузки")  # файл , лист

        #  = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
        #  = = = = = = = = = = = = = = = = = = = ЗАДАНИЕ ОБЩЕЕ  = = = = = = = = = = = = = = = = = = = = = = = = = = = =
        #  = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
        # 012 РАСЧЕТНЫЕ ФУНКЦИИ
        zad_temperatura = 2  # 2 температура в имени файла Tc=0 1 принудительно задать температуру Tc=0, если 0 то возьмет из Tc не обнуляем
        temperatura_zima = -5
        temperatura_leto = 25
        vibor_otkl = "otkl1"  # поле для выбора откл ветв и узлы: "otkl1" все откл , "otkl2"
        AutoShunt = True  #
        FLAG_automatika = False  # действие ПА после выявления перегрузки
        ris_tabl_add_PA = True  # для добавления рис и расчета в таб с учетом дейстся ПА даже если перегрузки исчезли

        max_tok_save = False  # запись макс токов для присоединений в массив()
        # ______________________________________________________________________________________________________________________________________________________________
        #  ЗАДАНИЕ КОНТРОЛЬ, ОТКЛ, импорт из файла
        Zad_vse_RG2 = False  # True ничто; False - обнулить КОНТРОЛ ОТКЛ
        vibor_raschot = False  # АВТО ЗДАНИЕ
        viborka_raschot = "na=4"  # "na=1"   #  выбор района или территории для расчета(>=110кВ) или ""    sel нельзя
        node_pojas_analiz = 4  # количество поясов  примыкающих к выборке viborka_raschot для анализа
        node_pojas_zad = 0  # количество поясов  примыкающих к выборке viborka_raschot для задания
        pqn_tranzit_min = 2  # МВт+МВар нагрузка посреди транзита, если больше этой величины то откл с разных концов
        Zad_RG2 = 1  # 1 брать задание для таблиц из файла (файл в папке IzFolder , должен начиналься с ! и .rg2), 0 нет  2 брать из папки zad
        Zad_RG2_VIBOR_N = vibor_otkl + "|Kontrol"
        node_auto_flag = True
        node_auto = "autosta,autoN,automatika,otkl_add,sta_otkl_add,remont_add,sta_remont_add"  # automatika
        node_zad_flag = True
        node_zad = "otkl1,otkl2,otkl3,Kontrol,N"  # задание

        Zad_RG2_VIBOR_V = vibor_otkl + "|Kontrol"
        vetv_auto_flag = True
        vetv_auto = "autosta,autoN,automatika,otkl_add,sta_otkl_add,remont_add,sta_remont_add"  # automatika
        vetv_zad_flag = True
        vetv_zad = "otkl1,otkl2,otkl3,Kontrol,N,znak-,KontrOO"  # задание
        # ______________________________________________________________________________________________________________________________________________________________
        protokol_XL = True  # 0 нет протокола XL#  1 -  есть
        XL_save = False  # 1 сохранить протокол в IzFolder, False нет
        XL_Visible = True  # False - не смотреть: 1 - смотреть заполнение таблиц
        Tabl_otlk_kontrol = 0  # ФОРМИРОВАНИЕ ТАБЛИЦ  КОНТРОЛ-ОТКЛ # 0  не заполнять, 1 по ячейкам, 2 через ДРВ
        zagruz_add_tab = 0  # 0 добавлять в таблицу все расчеты, 100 - добавлять при загрузки более указанной величины, например100  в процентах
        tabl_name_OK1 = "    Таблица "
        Ntabl_OK = 1
        tabl_name_OK2 = " - Результаты расчетов нормального и послеаварийных режимов работы электрических сетей 110 кВ и выше района размещения  с отключением одного электросетевого элемента из нормальной схемы. "  #
        EntireRow_OK = 100
        EntireColumn_OK = 65  # шапка
        # ______________________________________________________________________________________________________________________________________________________________
        picture_add = False
        risunok_zag = False  # режимы с перегрузками
        risunok_nr = True  # все нормальные режимы
        risunok_par = False  # все ПАРы, но с учетом фильтра

        risunok_rg2 = False  # сохранять rg2

        risunok_word = True  # сохранение РИСУНКОВ: 0 нет, 1 в word
        DOC_save = True  # сохранить рисунки в IzFolder
        DOC_visible = True
        tip_doc_file = False  # False новый файл, True - open_doc_file
        new_format_doc = "A3"
        orientation_doc = False  # tip_doc_file = 0 то ФОРМАТ нр "A3" "max"     ОРИЕНТАЦИЯ 1 книжная или 0 альбомная
        open_doc_file = "D:\флешка\РАБОТА\macro\word\рис max ал.docx"
        graf_shot = array(
            "10/Переток в контролируемых сечениях «Юг» и «Маныч» на уровне МДП. Режим выдачи располагаемой мощности ВИЭ. Без выдачи мощности Бондаревской ВЭС. ")  # ("номер кадра/имя для рис","10") номер кадра от 10 ctrl+0 ("ЗАПОМНИТЬ КАДР")  до "19" ctrl+9,
        # например ("11/Урайский  энергорайоны","12/Няганский энергорайоны","13/Сургутский энергорайон","14/Когалымский энергорайон","15/Нижневартовский энергорайон","16/Северный и Ноябрьский энергорайоны")
        # ("17/Нефтеюганский энергорайон","11/Ишимский энергорайон","12/Тюменский энергорайон","13/Тобольский энергорайон","14/Южный энергорайон")
        name_ris1 = "Рисунок И.5."  # нр "Рисунок  Е.1.1."
        number_pict_first = 36  # нумирация рисунков, начинается с
        nomer_ris_shag = 0  # шаг 0 нет, а если 1 то 1,3,5..
        graf_load = 0  # 0 нет , 1 загрузить графику
        God_1 = 2018
        Graf_1 = IzFolder + "\граф гпп4 2018.grf"  # равно и меньше
        God_2 = 2019
        Graf_2 = IzFolder + "\граф гпп4 2024.grf"  # равно и больше
        # ______________________________________________________________________________________________________________________________________________________________
        PZ_word = False  # добавить ПЗ в водрд
        ris_PZ_word = False  # рисунки
        # ______________________________________________________________________________________________________________________________________________________________
        # берется в названии файла,нр "2017 зим макс rg2_name_metka (-41С;МДП:ТЭ-У) rg2_name_metka.rg2"
        rg2_name_metka = array( \
            array("Вывод ТГ-7 СУГРЭС", "Вывод ТГ-7 СУГРЭС. "), \
            array("Вывод ТГ-6 СУГРЭС", "Вывод ТГ-6 СУГРЭС. "), \
            array("Вывод ТГ-6 СУГРЭС и Б1 РефтГРЭС", "Вывод ТГ-6 СУГРЭС и Б1 РефтГРЭС. "), \
            array("Вывод ТГ-6 СУГРЭС и Б1 РефтГРЭС", "Вывод ТГ-6 СУГРЭС и Б1 РефтГРЭС. "), \
            array("МДП_37_У-Т",
                  "Переток мощности в КС «ОЭС Урала - ЭСТО» в направлении «из ОЭС Урала» близкий к МДП. "), \
            array("МДП_37_Т-У", "Переток мощности в КС «ОЭС Урала - ЭСТО» в направлении «в ОЭС Урала» близкий к МДП. ")
        )

        if visual_set == 1:  # ##########################################################  HTML #####################################################################

            # 0 текущий файл, 1 папка, 2 "перегрузки"
            vid_raschota = IE_rform.script.rgm_tip

            # -12 файлы в ПАПКЕ
            IzFolder = IE_rform.IzFolder.value  # откуда берем файлы режима.rg2 нр 2020 зим макс, 2022 паводок мин (30°C ПЭВТ; МДП)
            if IE_rform.CB_filtr_file.checked:
                filtr_file = float(IE_rform.kol.value)  # 0 все файлы 5 - 5 файлов расчитываем, -1 - выборка см ниже
                if filtr_file == 0: filtr_file = -1
                uslovie_file = array(IE_rform.uslovie_file_god.value
                                     , IE_rform.uslovie_file_zim_let.value
                                     , IE_rform.uslovie_file_max_min.value
                                     , IE_rform.uslovie_file_dop_name.value)
            else:
                filtr_file = 0

                #  "" все или указать какие нр array ("2018,2019,2020,2021,2022,2023,2024,2025,2026","зим","мин","30°C ПЭВТ;МДП") (год, зим, макс,доп имя разделитель , или ;)
                # дополнительные имена указанны в скобочках в названии файла через ; например (30°C ПЭВТ;МДП), можно выбрать одно или несколько
            # 01- обсчитать текущий файл или все файлы в ПАПКЕ
            #  УСЛОВИЯ ПЕРЕБОРА
            not_n2 = False
            if not IE_rform.CB_otkl1.CHECKED and not IE_rform.CB_otkl2.CHECKED and not IE_rform.CB_otkl3.CHECKED: MinOtkl = 0; MaxOtkl = 0
            if IE_rform.CB_otkl1.CHECKED and not IE_rform.CB_otkl2.CHECKED and not IE_rform.CB_otkl3.CHECKED: MinOtkl = 1; MaxOtkl = 1
            if not IE_rform.CB_otkl1.CHECKED and IE_rform.CB_otkl2.CHECKED and not IE_rform.CB_otkl3.CHECKED: MinOtkl = 2; MaxOtkl = 2
            if not IE_rform.CB_otkl1.CHECKED and not IE_rform.CB_otkl2.CHECKED and IE_rform.CB_otkl3.CHECKED: MinOtkl = 3; MaxOtkl = 3
            if IE_rform.CB_otkl1.CHECKED and IE_rform.CB_otkl3.CHECKED: MinOtkl = 1; MaxOtkl = 3; if
            not IE_rform.CB_otkl2.CHECKED: not_n2 = True
            if not IE_rform.CB_otkl1.CHECKED and IE_rform.CB_otkl2.CHECKED and IE_rform.CB_otkl3.CHECKED: MinOtkl = 2; MaxOtkl = 3
            if IE_rform.CB_otkl1.CHECKED and IE_rform.CB_otkl2.CHECKED and not IE_rform.CB_otkl3.CHECKED: MinOtkl = 1; MaxOtkl = 2

            gost58670 = IE_rform.CB_gost58670.CHECKED  # 1 - учет ГОСТА 58670  если темпертура а-в то в н-2 не выводить перегрузку если превышен ддтн(мдн) но не превышено адтн(адн)
            filtr_n2 = IE_rform.CB_filtr_n2.CHECKED
            viborka_comb = float(
                IE_rform.viborka_comb.value)  # 0 нет или числов % меньше которого н-2 считаться не будет нр 60 (% загрузки в нр + % изм при откл 1 эл +% изм при откл 2 эл)
            # --2 по перегрузкам в ЭКСЕЛЕ

            FilePeregruzXL = array(IE_rform.T_run_pz_word_XL.value, IE_rform.T_run_pz_word_XL_list.value)  # файл , лист
            #  = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
            #  = = = = = = = = = = = = = = = = = = = ЗАДАНИЕ ОБЩЕЕ  = = = = = = = = = = = = = = = = = = = = = = = = = = = =
            #  = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
            # 012 РАСЧЕТНЫЕ ФУНКЦИИ
            if IE_rform.ID_zad_temp1.checked: zad_temperatura = 0  # 2 температура в имени файла Tc=0 1 принудительно задать температуру Tc=0, если 0 то возьмет из Tc не обнуляем
            if IE_rform.ID_zad_temp2.checked: zad_temperatura = 2  #
            if IE_rform.ID_zad_temp3.checked: zad_temperatura = 1  #
            temperatura_zima = float(IE_rform.temp_zim.value)
            temperatura_leto = float(IE_rform.temp_let.value)
            vibor_otkl = IE_rform.in_vibor_otkl.value  # поле для выбора откл ветв и узлы: "otkl1" все откл , "otkl2"
            AutoShunt = IE_rform.CB_AutoShunt.checked
            FLAG_automatika = IE_rform.CB_FLAG_automatika.checked
            ris_tabl_add_PA = True  # для добавления рис и расчета в таб с учетом дейстся ПА даже если перегрузки исчезли
            otkl_ssch = IE_rform.CB_otkl_ssch.checked  # 0 не отключать СШ /1 не отключать сш при отключении 2-х элементов сети /2 всегда отключать
            otkl_remont_shema = IE_rform.CB_remont_shema.checked  # 0 нет, 1 - в n-2 при ремонте ветви учитывать ремонтную схему(remont_add),а otkl_add не учитывать, при отключении учитывать доп откл (otkl_add)
            max_tok_save = IE_rform.CB_max_tok_save.checked  # 1 запись макс токов для присоединений в массив()
            # ______________________________________________________________________________________________________________________________________________________________
            #  ЗАДАНИЕ КОНТРОЛЬ, ОТКЛ, импорт из файла

            Zad_vse_RG2 = False if IE_rform.ID_Zad_RG2.checked or IE_rform.ID_Zad_RG3.checked else Zad_vse_RG2 = True  # 1 ничто; 0 - обнулить КОНТРОЛ ОТКЛ

            vibor_raschot = IE_rform.ID_Zad_RG2.checked  # АВТО ЗДАНИЕ
            viborka_raschot = IE_rform.viborka_raschot.value  # "na=1"   #  выбор района или территории для расчета(>=110кВ) или ""    sel нельзя
            node_pojas_analiz = 4  # количество поясов  примыкающих к выборке viborka_raschot для анализа
            node_pojas_zad = 0  # количество поясов  примыкающих к выборке viborka_raschot для задания
            pqn_tranzit_min = 2  # МВт+МВар нагрузка посреди транзита, если больше этой величины то откл с разных концов
            if IE_rform.ID_Zad_RG3.checked:
                Zad_RG2 = 1
            else:
                Zad_RG2 = 0  # 1 брать задание для таблиц из файла, 0 нет (файл в папке IzFolder , должен начиналься с ! и .rg2) 2 брать из папки zad
            if IE_rform.ID_zad_folder.checked: Zad_RG2 = 2
            Zad_RG2_VIBOR_N = vibor_otkl + "|Kontrol"
            node_auto_flag = IE_rform.CB_pole_node_af.checked
            node_zad_flag = IE_rform.CB_pole_node_kf.checked
            node_auto = IE_rform.pole_node_a.value  # automatika
            node_zad = IE_rform.pole_node_k.value  # задание

            Zad_RG2_VIBOR_V = vibor_otkl + "|Kontrol"
            vetv_auto_flag = IE_rform.CB_pole_vetv_af.checked
            vetv_zad_flag = IE_rform.CB_pole_vetv_kf.checked
            vetv_auto = IE_rform.pole_vetv_a.value  # automatika
            vetv_zad = IE_rform.pole_vetv_k.value
            # ______________________________________________________________________________________________________________________________________________________________
            protokol_XL = IE_rform.CB_protokol_XL.checked  # 0 нет протокола XL#  1 -  есть
            XL_save = IE_rform.CB_XL_save.checked  # 1 сохранить протокол в IzFolder, 0 нет
            XL_Visible = IE_rform.CB_XL_Visible.checked  # 0 - не смотреть: 1 - смотреть заполнение таблиц
            Tabl_otlk_kontrol = 1 if IE_rform.CB_Tabl_otlk_kontrol.checked else Tabl_otlk_kontrol = 0  # ФОРМИРОВАНИЕ ТАБЛИЦ со всеми контролируемыми элементами# 0  не заполнять, 1 по ячейкам, 2 через ДРВ
            zagruz_add_tab = IE_rform.zagruz_add_tab.value  # 0 добавлять в таблицу все расчеты, 100 - добавлять при загрузки более указанной величины, например100  в процентах
            tabl_name_OK1 = IE_rform.tabl_name_OK1.value
            Ntabl_OK = float(IE_rform.Ntabl_OK.value)
            tabl_name_OK2 = IE_rform.tabl_name_OK2.value  #
            EntireRow_OK = 100: EntireColumn_OK = 65  # шапка
            # ______________________________________________________________________________________________________________________________________________________________
            picture_add = IE_rform.CB_risunok_word_rg2.checked  # сохранение РИСУНКОВ
            risunok_word = IE_rform.CB_risunok_word.checked  # сохранение РИСУНКОВ: 0 нет, 1 в word
            risunok_rg2 = IE_rform.CB_risunok_rg2.checked  # сохранять rg2
            DOC_save = IE_rform.CB_DOC_save.checked  # сохранить рисунки в IzFolder
            DOC_visible = IE_rform.CB_DOC_visible.checked  # сохранить рисунки в IzFolder
            risunok_zag = IE_rform.CB_risunok_zag.checked  # 1 режимы с перегрузками; 0 нет
            risunok_nr = IE_rform.CB_risunok_nr.checked  # 1 все нормальные режимы, 0 нет
            risunok_par = IE_rform.CB_risunok_par.checked  # 1 - все ПАРы, но с учетом фильтра 0 нет

            tip_doc_file = False  # 0 новый файл, 1 - open_doc_file
            new_format_doc = IE_rform.T_new_format_doc.value
            orientation_doc = IE_rform.CB_orientation_doc.checked  # tip_doc_file = 0 то ФОРМАТ нр "A3" "max"     ОРИЕНТАЦИЯ 1 книжная или 0 альбомная
            open_doc_file = "D:\флешка\РАБОТА\macro\word\рис max ал.docx"
            graf_shot = split(IE_rform.T_graf_shot.value,
                              "")  # от 10 ctrl+0 ("ЗАПОМНИТЬ КАДР")  до "19" ctrl+9    или номера узлов для выделения графики нр (40 , 16 ,55, 5)
            # 11 Урайский  энергорайоны 12 Няганский энергорайоны 13 Сургутский энергорайон 14 Когалымский энергорайон 15 Нижневартовский энергорайон 16 Северный и Ноябрьский энергорайоны 17 Нефтеюганский энергорайон
            # 11 Ишимский энергорайон 12 Тюменский энергорайон 13 Тобольский энергорайон  14 Южный энергорайон
            name_ris1 = IE_rform.T_name_ris1.value  # нр "Рисунок  Е.1.1."
            number_pict_first = float(IE_rform.T_nomer_ris.value)  # нумирация рисунков, начинается с
            nomer_ris_shag = 0  # шаг 0 нет, а если 1 то 1,3,5..
            graf_load = 0  # 0 нет , 1 загрузить графику
            God_1 = 2018

            Graf_1 = IzFolder + "\граф гпп4 2018.grf"  # равно и меньше
            God_2 = 2019
            Graf_2 = IzFolder + "\граф гпп4 2024.grf"  # равно и больше
            # ______________________________________________________________________________________________________________________________________________________________
            # ______________________________________________________________________________________________________________________________________________________________
            PZ_word = IE_rform.CB_pz_word.checked
            ris_PZ_word = IE_rform.CB_risunok_PZ.checked

            # берется в названии файла,нр "2017 зим макс rg2_name_metka (-41С;МДП:ТЭ-У) rg2_name_metka.rg2"
            rg2_name_metka = array(
                array(IE_rform.rg2_name_metka1.value, IE_rform.rg2_name_metka11.value),
                array(IE_rform.rg2_name_metka2.value, IE_rform.rg2_name_metka21.value),
                array(IE_rform.rg2_name_metka3.value, IE_rform.rg2_name_metka31.value),
                array(IE_rform.rg2_name_metka4.value, IE_rform.rg2_name_metka41.value)
            )

        zad_log = "##################################### ЗАДАНИЕ НА РАСЧЕТ РЕЖИМОВ #####################################" + "\n"
        if vid_raschota == 0:
            zad_log = zad_log + "РАСЧЕТ ТЕКУЩЕГО ФАЙЛА" + "\n"
        else:
            if vid_raschota == 1: zad_log = zad_log + "РАСЧЕТ ФАЙЛОВ В ПАПКЕ: " + IzFolder + "\n"
            if vid_raschota == 2: zad_log = zad_log + "РАСЧЕТ по файлу с перегрузками: " + IzFolder + "\n"
            if filtr_file > 0: zad_log = zad_log + "\t" + "количество расчетных файлов: " + str(filtr_file) + "\n"
            if filtr_file == -1 and uslovie_file(
                0) != "": zad_log = zad_log + "\t" + "фильтр файлов, годы: " + uslovie_file(0) + "\n"
            if filtr_file == -1 and uslovie_file(1) != "": zad_log = zad_log + "\t" + "фильтр файлов: " + uslovie_file(
                1) + "\n"
            if filtr_file == -1 and uslovie_file(2) != "": zad_log = zad_log + "\t" + "фильтр файлов: " + uslovie_file(
                2) + "\n"
            if filtr_file == -1 and uslovie_file(
                3) != "": zad_log = zad_log + "\t" + "фильтр файлов, доп имя: " + uslovie_file(3) + "\n"

        zad_log = zad_log + "мин количество отключений в сочетании: " + str(MinOtkl) + ", максимальное: " + str(
            MaxOtkl) + "\n"
        if not_n2 == True: zad_log = zad_log + "н-2 не нада" + "\n"

        if otkl_ssch:  zad_log = zad_log + "* отключать СШ в н-1" + "\n"
        if otkl_remont_shema:   zad_log = zad_log + "* учитывать remont_add, otkl_add" + "\n"
        if filtr_n2:  zad_log = zad_log + "фильтр файлов" + "\n"
        if filtr_n2 and gost58670:  zad_log = zad_log + "\t" + "учет ГОСТА 58670" + "\n"
        if filtr_n2:  zad_log = zad_log + "\t" + "viborka_comb: " + str(viborka_comb) + "\n"

        if zad_temperatura == 2: zad_log = zad_log + "температура в имени файла (Tc=0)" + "\n"
        if zad_temperatura == 1: zad_log = zad_log + "температура (Tc=0) зима:" + str(
            temperatura_zima) + ", лето " + str(temperatura_leto) + "\n"
        if zad_temperatura == 0: zad_log = zad_log + "температура задана в файлах" + "\n"
        zad_log = zad_log + "поле для выбора откл ветвей и узлов: " + vibor_otkl + "\n"

        if AutoShunt:  zad_log = zad_log + "* AutoShunt включен" + "\n"
        if FLAG_automatika:  zad_log = zad_log + "* automatika включена" + "\n"
        if max_tok_save:      zad_log = zad_log + "* запись макс. токов" + "\n"

        if not Zad_vse_RG2:    zad_log = zad_log + "обнулить КОНТРОЛ ОТКЛ в моделях" + "\n"
        if vibor_raschot:    zad_log = zad_log + "автозадание  КОНТРОЛ ОТКЛ в моделях, выборка: " + viborka_raschot + "\n"
        if Zad_RG2 == 1:    zad_log = zad_log + "автозадание  КОНТРОЛ ОТКЛ в моделях из файла !...rg2" + "\n"
        if Zad_RG2 == 2:    zad_log = zad_log + "автозадание  КОНТРОЛ ОТКЛ в моделях из папке zad" + "\n"
        if Zad_RG2 > 0:
            zad_log = zad_log + "\t" + "выборка: " + Zad_RG2_VIBOR_N + "\n"
            if node_auto_flag: zad_log = zad_log + "\t" + "параметры: " + node_auto + "\n"
            if node_zad_flag: zad_log = zad_log + "\t" + "параметры: " + node_zad + "\n"
            zad_log = zad_log + "\t" + "выборка: " + Zad_RG2_VIBOR_V + "\n"
            if vetv_auto_flag: zad_log = zad_log + "\t" + "параметры: " + vetv_auto + "\n"
            if vetv_zad_flag: zad_log = zad_log + "\t" + "параметры: " + vetv_zad + "\n"

        if protokol_XL == 1:
            zad_log = zad_log + "протокол XL " + "\n"
            if Tabl_otlk_kontrol > 0: zad_log = zad_log + "\t" + "ФОРМИРОВАНИЕ ТАБЛИЦ КОНТОЛ-ОТКЛ с загрузкой больше " + str(
                zagruz_add_tab) + "\n"

        if picture_add:
            if risunok_word:   zad_log = zad_log + "рисунки в WORD " + "\n"
            if risunok_rg2:   zad_log = zad_log + "рисунки в RG2" + "\n"

        if PZ_word:
            if ris_PZ_word:   zad_log = zad_log + "рисунки для ПЗ" + "\n"

        zad_log = zad_log + "##################################### КОНЕЦ ЗАДАНИЯ #####################################"
        #  = = = = = = = = = = #  КОНЕЦ ОСНОВНОГО ЗАДАНИЯ= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
        #  Настройки
        number_pict = number_pict_first
        vivod_komb = 1  # 0 нет, 1 - запоминать перечень всех сочетаний и выводить в XL сразу если 2 то по завершению
        DRVXL = 1  # 0  без ссылки ДРВ"rastr.rtd"    1 с сылкой ДРВ
        dtn_uchastki = 1  # если у лэп есть groupid то выбока максимального тока  по: 0- groupid; 1- groupid dname и доп токам
        protokol_XL_row = 1
        TablF_const = 0  # 1 не меняется КОНТРОЛЬ, 0 меняется
        #  = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
        #  донастройка
        if vid_raschota != 1:  filtr_file = 0

        Folder_add = 1  # добавлять папки для CSV
        if MinOtkl > 1 or MaxOtkl < 2: filtr_n2 = False

        if not filtr_n2:
            viborka_comb = 0
            print_viborka = False

        if not protokol_XL:
            vivod_komb = 0

            Tabl_otlk_kontrol = 0
            print_viborka = False
            max_tok_save = False
            XL_save = False

        if max_tok_save:
            protokol_XL = True

        if picture_add:
            if risunok_word: redim
            file_wmf_size(ubound(graf_shot))
        else:
            risunok_rg2 = False
            risunok_word = False

        #  = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
        N_rezh = 1  # порядковый номер всех выполненных расчетов  - те запуска  Do Rgm def
        kol_test_da = 0
        kol_test_net = 0
        # НАСТРОЙКИ

        if vid_raschota == 0: IzFolder = objFSO.GetParentFolderName(
            rastr.SendCommandMain(6, "режим.rg2", "", 0))  # возвращает имя загруженного файла E:\41.rg2

        if Folder_add == 1 and objFSO.FolderExists(IzFolder):
            Folder_temp = IzFolder + "\temp"
            Folder_add_sub(Folder_temp)  # создать папку
            LogFile = objFSO.OpenTextFile(
                Folder_temp + "\Принт растет " + str(Day(Now)) + "_" + str(Month(Now)) + "_" + str(
                    Year(Now)) + "г " + str(Hour(Now)) + "ч_" + str(Minute(Now)) + "м_" + str(
                    Second(Now)) + "c" + ".log", 8, True)  # файл для записи ошибок
            Folder_csv_RG2 = Folder_temp + "\csv_RG2"
            Folder_add_sub(Folder_csv_RG2)  # создать папку
            Folder_wmf = Folder_temp + "\wmf"
            Folder_add_sub(Folder_wmf)  # создать папку
            Folder_rg2 = Folder_temp + "\ris_rg2"
            Folder_add_sub(Folder_rg2)  # создать папку


def IE_ImpRg2():  # запуск ИД для импорта# ---------ИМПОРТ из модели-------------
    if IE_kform.CB_PQ.Checked:  #
        ImportClass = ImportRG2()  #
        ImportClass.uslovie_start = array(IE_kform.Filtr_PQ_god.value, IE_kform.Filtr_PQ_zim_let.value,
                                          IE_kform.Filtr_PQ_max_min.value, IE_kform.Filtr_PQ_dop_name.value)
        ImportClass.import_File = IE_kform.file_PQFolder.value  # "I:\ОЭС Юга\КПР КУБАНИ\модели\sel уд гр айди.rg2"
        ImportClass.tabl = IE_kform.tab_PQFolder.value  # "vetv"
        ImportClass.param = array(IE_kform.param_PQFolder.value,
                                  "")  # "node;vetv;Generator", ("пусто-все или перечислить","набор парам")параметры так же можно ";"
        ImportClass.vibor = IE_kform.vibor_PQFolder.value  # "sel"
        ImportClass.tip = IE_kform.tip_PQFolder.value  # "2" # "2" обн, "1" заг, "0" прис, "3" обн-прис
        CS.dictImpRg2.Add(ImportClass.import_File + str(round(Rnd, 4) * 10000), ImportClass)

    if IE_kform.CB_SHN.Checked:  #
        ImportClass = ImportRG2()  #
        ImportClass.uslovie_start = array(IE_kform.Filtr_SHN_god.value, IE_kform.Filtr_SHN_zim_let.value,
                                          IE_kform.Filtr_SHN_max_min.value, IE_kform.Filtr_SHN_dop_name.value)
        ImportClass.import_File = IE_kform.file_SHNFolder.value  # "I:\ОЭС Юга\КПР КУБАНИ\модели\sel уд гр айди.rg2"
        ImportClass.tabl = IE_kform.tab_SHNFolder.value  # "vetv"
        ImportClass.param = array(IE_kform.param_SHNFolder.value,
                                  "")  # "node;vetv;Generator", ("пусто-все или перечислить","набор парам")параметры так же можно ";"
        ImportClass.vibor = IE_kform.vibor_SHNFolder.value  # "sel"
        ImportClass.tip = IE_kform.tip_SHNFolder.value  # "2" # "2" обн, "1" заг, "0" прис, "3" обн-прис
        CS.dictImpRg2.Add(ImportClass.import_File + str(round(Rnd, 4) * 10000), ImportClass)

    if IE_kform.CB_I_T.Checked:  #
        ImportClass = ImportRG2  #
        ImportClass.uslovie_start = array(IE_kform.Filtr_I_T_god.value, IE_kform.Filtr_I_T_zim_let.value,
                                          IE_kform.Filtr_I_T_max_min.value, IE_kform.Filtr_I_T_dop_name.value)
        ImportClass.import_File = IE_kform.file_I_TFolder.value  # "I:\ОЭС Юга\КПР КУБАНИ\модели\sel уд гр айди.rg2"
        ImportClass.tabl = IE_kform.tab_I_TFolder.value  # "vetv"
        ImportClass.param = array(IE_kform.param_I_TFolder.value,
                                  "")  # "node;vetv;Generator", ("пусто-все или перечислить","набор парам")параметры так же можно ";"
        ImportClass.vibor = IE_kform.vibor_I_TFolder.value  # "sel"
        ImportClass.tip = IE_kform.tip_I_TFolder.value  # "2" # "2" обн, "1" заг, "0" прис, "3" обн-прис
        CS.dictImpRg2.Add(ImportClass.import_File + str(round(Rnd, 4) * 10000), ImportClass)

    if IE_kform.CB_NGA.Checked:  #
        ImportClass = ImportRG2  #
        ImportClass.uslovie_start = array(IE_kform.Filtr_NGA_god.value, IE_kform.Filtr_NGA_zim_let.value,
                                          IE_kform.Filtr_NGA_max_min.value, IE_kform.Filtr_NGA_dop_name.value)
        ImportClass.import_File = IE_kform.file_NGAFolder.value  # "I:\ОЭС Юга\КПР КУБАНИ\модели\sel уд гр айди.rg2"
        ImportClass.tabl = IE_kform.tab_NGAFolder.value  # "vetv"
        ImportClass.param = array(IE_kform.param_NGAFolder.value,
                                  "")  # "node;vetv;Generator", ("пусто-все или перечислить","набор парам")параметры так же можно ";"
        ImportClass.vibor = IE_kform.vibor_NGAFolder.value  # "sel"
        ImportClass.tip = IE_kform.tip_NGAFolder.value  # "2" # "2" обн, "1" заг, "0" прис, "3" обн-прис
        CS.dictImpRg2.Add(ImportClass.import_File + str(round(Rnd, 4) * 10000), ImportClass)

    if IE_kform.CB_RAJ.Checked:  #
        ImportClass = ImportRG2  #
        ImportClass.uslovie_start = array(IE_kform.Filtr_RAJ_god.value, IE_kform.Filtr_RAJ_zim_let.value,
                                          IE_kform.Filtr_RAJ_max_min.value, IE_kform.Filtr_RAJ_dop_name.value)
        ImportClass.import_File = IE_kform.file_RAJFolder.value  # "I:\ОЭС Юга\КПР КУБАНИ\модели\sel уд гр айди.rg2"
        ImportClass.tabl = IE_kform.tab_RAJFolder.value  # "vetv"
        ImportClass.param = array(IE_kform.param_RAJFolder.value,
                                  "")  # "node;vetv;Generator", ("пусто-все или перечислить","набор парам")параметры так же можно ";"
        ImportClass.vibor = IE_kform.vibor_RAJFolder.value  # "sel"
        ImportClass.tip = IE_kform.tip_RAJFolder.value  # "2" # "2" обн, "1" заг, "0" прис, "3" обн-прис
        CS.dictImpRg2.Add(ImportClass.import_File + str(round(Rnd, 4) * 10000), ImportClass)

    if IE_kform.CB_TERR.Checked:  #
        ImportClass = ImportRG2  #
        ImportClass.uslovie_start = array(IE_kform.Filtr_TERR_god.value, IE_kform.Filtr_TERR_zim_let.value,
                                          IE_kform.Filtr_TERR_max_min.value, IE_kform.Filtr_TERR_dop_name.value)
        ImportClass.import_File = IE_kform.file_TERRFolder.value  # "I:\ОЭС Юга\КПР КУБАНИ\модели\sel уд гр айди.rg2"
        ImportClass.tabl = IE_kform.tab_TERRFolder.value  # "vetv"
        ImportClass.param = array(IE_kform.param_TERRFolder.value,
                                  "")  # "node;vetv;Generator", ("пусто-все или перечислить","набор парам")параметры так же можно ";"
        ImportClass.vibor = IE_kform.vibor_TERRFolder.value  # "sel"
        ImportClass.tip = IE_kform.tip_TERRFolder.value  # "2" # "2" обн, "1" заг, "0" прис, "3" обн-прис
        CS.dictImpRg2.Add(ImportClass.import_File + str(round(Rnd, 4) * 10000), ImportClass)

    if IE_kform.CB_OBED.Checked:  #
        ImportClass = ImportRG2  #
        ImportClass.uslovie_start = array(IE_kform.Filtr_OBED_god.value, IE_kform.Filtr_OBED_zim_let.value,
                                          IE_kform.Filtr_OBED_max_min.value, IE_kform.Filtr_OBED_dop_name.value)
        ImportClass.import_File = IE_kform.file_OBEDFolder.value  # "I:\ОЭС Юга\КПР КУБАНИ\модели\sel уд гр айди.rg2"
        ImportClass.tabl = IE_kform.tab_OBEDFolder.value  # "vetv"
        ImportClass.param = array(IE_kform.param_OBEDFolder.value,
                                  "")  # "node;vetv;Generator", ("пусто-все или перечислить","набор парам")параметры так же можно ";"
        ImportClass.vibor = IE_kform.vibor_OBEDFolder.value  # "sel"
        ImportClass.tip = IE_kform.tip_OBEDFolder.value  # "2" # "2" обн, "1" заг, "0" прис, "3" обн-прис
        CS.dictImpRg2.Add(ImportClass.import_File + str(round(Rnd, 4) * 10000), ImportClass)

    if IE_kform.CB_GNR.Checked:  #
        ImportClass = ImportRG2  #
        ImportClass.uslovie_start = array(IE_kform.Filtr_GNR_god.value, IE_kform.Filtr_GNR_zim_let.value,
                                          IE_kform.Filtr_GNR_max_min.value, IE_kform.Filtr_GNR_dop_name.value)
        ImportClass.import_File = IE_kform.file_GNRFolder.value  # "I:\ОЭС Юга\КПР КУБАНИ\модели\sel уд гр айди.rg2"
        ImportClass.tabl = IE_kform.tab_GNRFolder.value  # "vetv"
        ImportClass.param = array(IE_kform.param_GNRFolder.value,
                                  "")  # "node;vetv;Generator", ("пусто-все или перечислить","набор парам")параметры так же можно ";"
        ImportClass.vibor = IE_kform.vibor_GNRFolder.value  # "sel"
        ImportClass.tip = IE_kform.tip_GNRFolder.value  # "2" # "2" обн, "1" заг, "0" прис, "3" обн-прис
        CS.dictImpRg2.Add(ImportClass.import_File + str(round(Rnd, 4) * 10000), ImportClass)

    if IE_kform.CB_NDE.Checked:  #
        ImportClass = ImportRG2  #
        ImportClass.uslovie_start = array(IE_kform.Filtr_NDE_god.value, IE_kform.Filtr_NDE_zim_let.value,
                                          IE_kform.Filtr_NDE_max_min.value, IE_kform.Filtr_NDE_dop_name.value)
        ImportClass.import_File = IE_kform.file_NDEFolder.value  # "I:\ОЭС Юга\КПР КУБАНИ\модели\sel уд гр айди.rg2"
        ImportClass.tabl = IE_kform.tab_NDEFolder.value  # "vetv"
        ImportClass.param = array(IE_kform.param_NDEFolder.value,
                                  "")  # "node;vetv;Generator", ("пусто-все или перечислить","набор парам")параметры так же можно ";"
        ImportClass.vibor = IE_kform.vibor_NDEFolder.value  # "sel"
        ImportClass.tip = IE_kform.tip_NDEFolder.value  # "2" # "2" обн, "1" заг, "0" прис, "3" обн-прис
        CS.dictImpRg2.Add(ImportClass.import_File + str(round(Rnd, 4) * 10000), ImportClass)

    if IE_kform.CB_VTV.Checked:  #
        ImportClass = ImportRG2  #
        ImportClass.uslovie_start = array(IE_kform.Filtr_VTV_god.value, IE_kform.Filtr_VTV_zim_let.value,
                                          IE_kform.Filtr_VTV_max_min.value, IE_kform.Filtr_VTV_dop_name.value)
        ImportClass.import_File = IE_kform.file_VTVFolder.value  # "I:\ОЭС Юга\КПР КУБАНИ\модели\sel уд гр айди.rg2"
        ImportClass.tabl = IE_kform.tab_VTVFolder.value  # "vetv"
        ImportClass.param = array(IE_kform.param_VTVFolder.value,
                                  "")  # "node;vetv;Generator", ("пусто-все или перечислить","набор парам")параметры так же можно ";"
        ImportClass.vibor = IE_kform.vibor_VTVFolder.value  # "sel"
        ImportClass.tip = IE_kform.tip_VTVFolder.value  # "2" # "2" обн, "1" заг, "0" прис, "3" обн-прис
        CS.dictImpRg2.Add(ImportClass.import_File + str(round(Rnd, 4) * 10000), ImportClass)


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def name_step():  # заполняет поле name_step таблицы AutoZad
    # dim tAZad , tN46, tV46 , dict_N_PA ,  zadanie1, otkl11  ,  rem11
    if rastr.Tables.Find("AutoZad") > -1:
        tN46 = rastr.Tables("node")
        tV46 = rastr.Tables("vetv")
        dict_N_PA = CreateObject("Scripting.Dictionary")

        for i = 0 to tV46.size-1
        zadanie1 = replace(tV46.cols.item("automatika").Z(i), " ", "")
        otkl11 = replace(tV46.cols.item("otkl_add").Z(i), " ", "")
        rem11 = replace(tV46.cols.item("remont_add").Z(i), " ", "")

        if not zadanie1 == "": name_stepVN(dict_N_PA, zadanie1,
                                           "ПА[" + tV46.cols.item("kkluch").Z(i) + "] " + tV46.cols.item("dname").ZS(i))
        if not otkl11 == "": name_stepVN(dict_N_PA, otkl11,
                                         "откл[" + tV46.cols.item("kkluch").Z(i) + "] " + tV46.cols.item("dname").ZS(i))
        if not rem11 == "": name_stepVN(dict_N_PA, rem11,
                                        "ремонт[" + tV46.cols.item("kkluch").Z(i) + "] " + tV46.cols.item("dname").Z(i))

    for i = 0 to tN46.size-1
    zadanie1 = replace(tN46.cols.item("automatika").Z(i), " ", "")
    otkl11 = replace(tN46.cols.item("otkl_add").Z(i), " ", "")
    rem11 = replace(tN46.cols.item("remont_add").Z(i), " ", "")

    if not zadanie1 == "": name_stepVN(dict_N_PA, zadanie1,
                                       "ПА[" + tN46.cols.item("ny").Z(i) + "] " + tN46.cols.item("name").Z(i))
    if not otkl11 == "": name_stepVN(dict_N_PA, otkl11,
                                     "откл[" + tN46.cols.item("ny").Z(i) + "] " + tN46.cols.item("name").Z(i))
    if not rem11 == "": name_stepVN(dict_N_PA, rem11,
                                    "ремонт[" + tN46.cols.item("ny").Z(i) + "] " + tN46.cols.item("name").Z(i))


tAZad = rastr.Tables("AutoZad")

for i = 0 to tAZad.size-1
tAZad.cols.item("name_step").Z(i) = ""
if tAZad.cols.item("N").Z(i) > 0:
    if dict_N_PA.Exists(tAZad.cols.item("N").ZS(i)):
        tAZad.cols.item("name_step").Z(i) = dict_N_PA.Item(tAZad.cols.item("N").ZS(i)) + " " else:
        tAZad.cols.item("name_step").Z(i) = " - "
    if tAZad.cols.item("tabl").Z(i) = 0:  # узел
        if fNDX("node", tAZad.cols.item("kluch").ZS(i)) > - 1:
            dopiska = "(" + tN46.cols.item("name").ZS(fNDX("node", tAZad.cols.item("kluch").ZS(i))) + ")"  else:
            dopiska = "(не найден)"
        tAZad.cols.item("name_step").Z(i) = tAZad.cols.item("name_step").ZS(i) + dopiska
    elif tAZad.cols.item("tabl").Z(i) = 1:  # ветвь
        if fNDX("vetv", tAZad.cols.item("kluch").ZS(i)) > - 1:
            dopiska = "(" + tV46.cols.item("dname").ZS(fNDX("vetv", tAZad.cols.item("kluch").ZS(i))) + ")"  else:
            dopiska = "(не найден)"
        tAZad.cols.item("name_step").Z(i) = tAZad.cols.item("name_step").ZS(i) + dopiska
    elif tAZad.cols.item("tabl").Z(i) = 2:  # район
        if fNDX("area", tAZad.cols.item("kluch").ZS(i)) > - 1:
            dopiska = "(" + rastr.tables("area").cols.item("name").ZS(
                fNDX("area", tAZad.cols.item("kluch").ZS(i))) + ")"  else:
            dopiska = "(не найден)"
        tAZad.cols.item("name_step").Z(i) = tAZad.cols.item("name_step").ZS(i) + dopiska
    elif tAZad.cols.item("tabl").Z(i) = 3:  # терр
        if fNDX("area2", tAZad.cols.item("kluch").ZS(i)) > - 1:
            dopiska = "(" + rastr.tables("area2").cols.item("name").ZS(
                fNDX("area2", tAZad.cols.item("kluch").ZS(i))) + ")"  else:
            dopiska = "(не найден)"
        tAZad.cols.item("name_step").Z(i) = tAZad.cols.item("name_step").ZS(i) + dopiska
    elif tAZad.cols.item("tabl").Z(i) = 4:  # нагр груп
        if fNDX("ngroup", tAZad.cols.item("kluch").ZS(i)) > - 1:
            dopiska = "(" + rastr.tables("ngroup").cols.item("name").ZS(
                fNDX("ngroup", tAZad.cols.item("kluch").ZS(i))) + ")"  else:
            dopiska = "(не найден)"
        tAZad.cols.item("name_step").Z(i) = tAZad.cols.item("name_step").ZS(i) + dopiska
    elif tAZad.cols.item("tabl").Z(i) = 5:  # ген
        if fNDX("Generator", tAZad.cols.item("kluch").ZS(i)) > - 1:
            dopiska = "(" + rastr.tables("Generator").cols.item("Name").ZS(
                fNDX("Generator", tAZad.cols.item("kluch").ZS(i))) + ")"  else:
            dopiska = "(не найден)"
        tAZad.cols.item("name_step").Z(i) = tAZad.cols.item("name_step").ZS(i) + dopiska
    # if tAZad.cols.item("tabl").Z(i) = 6 #  изм

else:
logging.info("!!! не загружен файл автомитики !!!")

# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в  коррррр++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def arr_numeric(ar_id):  # заменить значения в массиве строки на числа если можно
    for i = 0 to ubound (ar_id)
    if isnumeric(ar_id(i)): ar_id(i) = float(ar_id(i))


arr_numeric = ar_id


# End def return

# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def mainRG():  # самая головная процедура расчета
    GLR = Global_raschot_class
    GLR.init()  #
    if GLR.Zad_RG2 == 2:
        msgbox("в папке \zad " + str(objFSO.GetFolder(GLR.IzFolder + "\zad").Files.Count) + " файлов")
        for Each objFile_zad in objFSO.GetFolder(GLR.IzFolder + "\zad").Files
            GLR.Zad_RG2_name = objFile_zad.Path
            GLR.Zad_RG2_name_k = objFile_zad.name
            if objFile_zad.type == "Файл режима rg2":
                GLR.protokol_XL_row = 1
                mainR()

    else:
        mainR()


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def mainR():  # головная процедура
    # dim format_list , tekk , str , temp , BDP , array_str , OLC , dictVal1
    # ПАПКИ и ФАЙЛЫ

    if GLR.vid_raschota == 1 OR (GLR.vid_raschota = 2 and GLR.PZ_word): if
    not objFSO.FolderExists(GLR.IzFolder): msgbox(GLR.IzFolder + " - не найден"):GS.end_gl(): exit

    def

    if GLR.Zad_RG2 == 1:
        GLR.Zad_RG2_name = fSCAN_FOLDER(GLR.IzFolder, "!", "Файл режима rg2")  # функция возвращает полное имя файл  "!"
        if GLR.Zad_RG2_name == "не найден":
            msgbox(" не найден файл !задание.rg2")
            GS.end_gl()
            exit

            def

    GLR.TablF_zim = 0
    GLR.TablF_let = 0
    # if GLR.risunok_word or:

    if GLR.risunok_word:
        format_list = array(_
        array("max", 42, 55.8), _
        array("A0", 84.1, 118.9), _
        array("A1", 59.4, 84.1), _
        array("A2", 42.0, 59.4), _
        array("A3", 29.7, 42.0), _
        array("A4", 21.0, 29.7))

        Del_v_Folder(GLR.Folder_wmf)

        GS.word.ScreenUpdating = False
        if GLR.DOC_visible: GS.word.Visible = True
        if not GLR.tip_doc_file:
            GLR.word2 = GS.word.Documents.Add()  # новый док
            GLR.new_format_doc = replace(GLR.new_format_doc, "А", "A")
            for each format_list_i in format_list
                if format_list_i(0) == GLR.new_format_doc:
                    GLR.word2.PageSetup.PageWidth = format_list_i(
                        2) * 28.35  # CentimetersToPoints( format_list_i (2) ) 1 см = 28,35
                    GLR.word2.PageSetup.PageHeight = format_list_i(
                        1) * 28.35  # CentimetersToPoints( format_list_i (1) )
                    exit
                    for

            GLR.word2.PageSetup.Orientation = GLR.orientation_doc
            GLR.word2.PageSetup.TopMargin = 1.5 * 28.35
            GLR.word2.PageSetup.BottomMargin = 1.5 * 28.35
            GLR.word2.PageSetup.LeftMargin = 2 * 28.35
            GLR.word2.PageSetup.RightMargin = 1.5 * 28.35
        else:
            if objFSO.FileExists(GLR.open_doc_file):
                GLR.word2 = GS.word.Documents.Open(GLR.open_doc_file)  # открыть сущ док
            else:
                msgbox(" не найден файл:" + GLR.open_doc_file)
                GS.end_gl()
                exit

                def
    if GLR.risunok_rg2: Del_v_Folder(GLR.Folder_rg2)
    #  = = = = = = = = = = = = = = = = = = = = = = = =
    if GLR.protokol_XL:
        if GLR.XL_Visible: GS.excel.Visible = True: GS.excel.ScreenUpdating = True  # разблокировки

    if GLR.protokol_XL:  # новая книга
        GLR.Book_XL = GS.excel.Workbooks.Add
        GLR.protokol_XL_Sheets = GLR.Book_XL.Worksheets(1)
        GLR.Book_XL.Worksheets(1).Name = "протокол"
        GLR.protokol_XL_Sheets.Columns(1).ColumnWidth = 200

    if GLR.protokol_XL: GLR.overload = overload_class: GLR.overload.init_p
    if GLR.vivod_komb > 0: Komb_List = Komb_All_List_class: Komb_List.init
    #  добавить лист
    if GLR.max_tok_save:  Sheets_add(GLR.Book_XL, temp, "Imax"): GLR.XL_max_tok = temp
    if GLR.viborka_comb > 0 and GLR.protokol_XL:  Sheets_add(GLR.Book_XL, temp, "сечения"): GLR.XL_print_mKOO = temp
    if GLR.viborka_comb > 0 and GLR.protokol_XL and GLR.print_viborka:  Sheets_add(GLR.Book_XL, temp,
                                                                                   "сеч_ремонт"): GLR.XL_print_mKOR = temp
    if GLR.TablF_const == 1 and GLR.protokol_XL:  Sheets_add(GLR.Book_XL, temp,
                                                             "TablF"): GLR.TablF_sheets = temp: GLR.TablF_zim = 1:  GLR.TablF_let = 1

    logging.info("НАЧАЛО РАСЧЕТА: " + str(Now()))  # дата время
    tekk = split(GLR.zad_log, "\n")
    for each str in tekk
        logging.info(str)

    if GLR.max_tok_save:  # СОХРАНЕНИЕ МАКС ТОКОВ
        GLR.max_tok1 = max_tok_class
        GLR.max_tok1.init_sub()  # класс для записи максимальных токов по присоединениям

    tekk = fTipFile()(1)  # полное имя загруженного файла "С:\1\2020 зим макс !.rg2"
    if GLR.Zad_RG2 > 0:
        logging.info("файл задание: " + GLR.Zad_RG2_name)
        export_RG2()  # загрузить файл Zad_RG2 и из него выгрузка ид для таблицы из файла в CSV

    if GLR.vid_raschota == 0:  # 0 текущий файл
        if GLR.Zad_RG2 > 0:
            rastr.Load(1, tekk, fshablon(tekk))  # загрузить режим
            logging.info("перезагружен файл: " + fTipFile()(1))

        RG = CurrentFile
        RG.file_path = tekk  # полное имя загруженного файла "С:\1\2020 зим макс !.rg2"
        RG.initRG(0, array("", "", "", ""))  # разбирает file_path и тд
        logging.info("расчет текущего файла: " + RG.file_path)

        if not RG.Name_st == "не подходит":
            rg2_raschot()  # общие действия выполняемые с файлом
        else:
            logging.info("имя текущего файла не проходит: " + RG.Name_base)

    elif GLR.vid_raschota == 1:  # 1 - обычный цикл по файлам в папке

        for Each objFile in objFSO.GetFolder(GLR.IzFolder).Files  # цикл по файлам в  указанной папке
            if objFile.type == "Файл режима rg2":
                RG = CurrentFile
                RG.file_path = objFile.Path
                RG.initRG(GLR.filtr_file, GLR.uslovie_file)  # разбирает file_path и тд
                if not RG.Name_st == "не подходит":
                    GS.N_rg2_File = GS.N_rg2_File + 1
                    rastr.Load(1, objFile.Path, RG.shablon)  # загрузить режим
                    if GLR.risunok_word and GLR.graf_load == 1: GrfLoad()
                    rg2_raschot()  # общие действия выполняемые с файлом
                    if GLR.filtr_file == 1: exit
                    for
                        if GLR.filtr_file > 1: GLR.filtr_file = GLR.filtr_file - 1
                else:
                    logging.info(RG.file_path + "- файл отклонен")

    elif GLR.vid_raschota = 2:  # 2 - цикл по режимам в "перегрузки"

        if not objFSO.FileExists(GLR.FilePeregruzXL(0))
            msgbox(GLR.FilePeregruzXL(0) + " - не найден файл GLR.FilePeregruzXL")
            GS.end_gl()
            exit

            def

        GLR.Peregruz_XL = GS.excel.Workbooks.Open(GLR.FilePeregruzXL(0))
        if not SheetExists(GLR.Peregruz_XL, GLR.FilePeregruzXL(1))
            msgbox(GLR.FilePeregruzXL(1) + " - не найден лист перегрузки ")
            GS.end_gl()
            exit

            def
        GLR.Peregruz_XL_Sheets = GLR.Peregruz_XL.Sheets(GLR.FilePeregruzXL(1))
        dictVal1 = CreateObject("Scripting.Dictionary")  # для хранения
        BDP = GLR.Peregruz_XL_Sheets.UsedRange.Value  # база данныз перегрузок
        redim
        array_str(ubound(BDP, 2))

        for i = 2 to ubound (BDP, 1)  # цикул по строкам
        for iui = 1 to ubound (BDP, 2)  # цикул по строкам
        if isempty(BDP(i, iui)):
            array_str(iui - 1) = "" else:
            array_str(iui - 1) = BDP(i, iui)

    dictVal1.add
    i - 1, array_str


OLC = overload_class
OLC.init_c()
OLC.dictVal = dictVal1
OLC.print_end_p(1)
# if GLR.XL_save: GLR.Book_XL.SaveAs (GLR.IzFolder + "\протокол расчета по выборке"  + "(" + str (Day(Now)) + "_" + str (Month(Now)) + "_" + str (Year(Now)) + "г " + str (Hour(Now)) + "ч_" + str (Minute(Now)) + "м_" + str (Second(Now)) + "c, " + GLR.Zad_RG2_name_k + ").xlsm" , 52)

if GLR.risunok_word:
    if GLR.DOC_save: GLR.word2.SaveAs2(left(GLR.IzFolder + "\" + GLR.name_ris1 + "(" + str(Day(Now)) + "
    _
    " + str(Month(Now)) + "
    _
    " + str(Year(Now)) + "
    г
    " + str(Hour(Now)) + "
    ч_
    " + str(Minute(Now)) + "
    м_
    " + str(Second(Now)) + "
    c)" , 250) + ".docx
    ", 12)
    GS.word.Visible = True
    GS.word.ScreenUpdating = True
    GS.word = None

    if GLR.max_tok_save:
        GLR.max_tok1.print_max_tok()
    # вывод мах токов в XL
    if GLR.protokol_XL:
        GLR.overload.print_end_p(0)
    if GLR.vivod_komb > 0: Komb_List.print_end_KL()

    GS.kod_rgm = rastr.rgm("")
    if GS.kod_rgm = 1:
    GS.kod_rgm = rastr.rgm("")
    if GS.kod_rgm = 1:
    GS.kod_rgm = rastr.rgm("p")
    if GS.kod_rgm = 1:
    GS.kod_rgm = rastr.rgm("p")
    if GS.kod_rgm = 1:
    GS.kod_rgm = rastr.rgm("p")

    if GLR.XL_save:
        GLR.Book_XL.SaveAs(GLR.IzFolder + "\" + f_txt_name_good ( "
    протокол
    расчета
    n - " + str (GLR.MinOtkl) + ", " + str (GLR.MaxOtkl) + "(" + str (Day(Now)) + "
    _
    " + str (Month(Now)) + "
    _
    " + str (Year(Now)) + "
    г
    " + str (Hour(Now)) + "
    ч_
    " + str (Minute(Now)) + "
    м_
    " + str (Second(Now)) + "
    c, " + GLR.Zad_RG2_name_k) + ").xlsm
    " , 52)
    if GLR.XL_save:
        GLR.Book_XL.SaveAs(GLR.IzFolder + "\" + f_txt_name_good ( "
    протокол
    расчета
    по
    списку(" + str (Day(Now)) + "
    _
    " + str (Month(Now)) + "
    _
    " + str (Year(Now)) + "
    г
    " + str (Hour(Now)) + "
    ч_
    " + str (Minute(Now)) + "
    м_
    " + str (Second(Now)) + "
    c, " + GLR.Zad_RG2_name_k) + ").xlsm
    " , 52)

    # *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************


def rg2_raschot():  # общие действия выполняемые с файлом
    # dim i_comb
    GLR.kol_otkl = 0  # тек число одновременных отключений
    logging.info(RG.Name_Base)
    GS.kod_rgm = rastr.rgm("")
    if GS.kod_rgm = 1:
        GS.kod_rgm = rastr.rgm("")
    if GS.kod_rgm = 1:
        GS.kod_rgm = rastr.rgm("p")
    if GS.kod_rgm = 1:
        GS.kod_rgm = rastr.rgm("p")
    if GS.kod_rgm = 1:
        GS.kod_rgm = rastr.rgm("p")
    if GS.kod_rgm = 1: logging.info(
        "\t" + "НОРМАЛЬНЫЙ РЕЖИМ НЕ МОДЕЛИРУЕТСЯ!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!, выход"): exit

    def

        if GLR.Tabl_otlk_kontrol > 0: XL_sheet_add()  # добавить лист для таблицы расчета текущего режима

    GLR.blok_sub()
    VIBOR_KONTROL_OTKL()  # процедура для отметки узлов и ветвей КОНТРОЛЬ и ОТКЛ

    RG.Dict_StoredOTKLUCH = CreateObject("Scripting.Dictionary")  # для хранения перечня откл
    # spUniquizer = Uniquizer
    if GLR.AutoShunt:
        AutoShunt_class_rec("")  # процедура формирует Umin , Umax, AutoBsh , nBsh
        AutoShunt_class_kor()  # процедура меняет Bsh  и записывает GS.AutoShunt_list
        if GLR.Tabl_otlk_kontrol > 0 and GS.AutoShunt_list != "": GLR.XL_sheet.cell(1,
                                                                                     17).Value = GS.AutoShunt_list: GS.AutoShunt_list = ""

    if GLR.zad_temperatura > 0: Tc_0_Sub()  # обнулить расчетную температуру в ветвях районах и тд
    rastr.CalcIdop(RG.GradusZ, float(0), ""): logging.info(
        "\t" + "расчетная температура(rg2_raschot): " + str(RG.GradusZ))  #
    TopologyStore(): PnQnPgStore()  # ЗАПИСАТЬ параметры режима
    RG.Kontrol_init()  # формирует KontrolVL , KontrolTrans , KontrolNode
    RG.OTKL1_masiv()  # контр   - формирует RG.OTKL_masiv

    if GLR.Tabl_otlk_kontrol > 0:  TablF_init()  # инициализация TablF_Sub ()  и запуск
    RG.redim_Otkl_Comb_tek(3,
                           1)  # (0-ndx 1-"node"/"vetv" 2 -kluch-SelString(ndx) 3 "otkl_add"/"remont_add"-имя поля * кол элементов)

    RGR = raschot_tek_comb  # НР
    Comb = Combinator
    Comb.tip_comb = -1
    RGR.init_new()  #
    DoRgm()  # НР
    #  ФИЛЬТР КОМБИНАЦИЙ
    if GLR.viborka_comb > 0:
        mKO = mKontrol_Otkl
        mKO.KontrolOtkl_init()  # запись нр
        if GLR.protokol_XL: mKO.Print_XL_mKO(1)  # печать mKO 1 загаловки  и нр, 2 рез н-1

    if GLR.MaxOtkl > UBound(RG.OTKL_masiv, 2) + 1:
        GLR.MaxOtkl = UBound(RG.OTKL_masiv, 2) + 1
        logging.info("!!!Количество элементов в сочетаниях уменьшено до количества отключаемых элементов равных " + str(
            GLR.MaxOtkl) + "!!!")

    for i_comb = GLR.MinOtkl to GLR.MaxOtkl  # цикл  ОТКЛЮЧЕНИЯ ПО ПОРЯДКУ  формирует RG.Otkl_Comb_tek из RG.OTKL_masiv         и запускает OTKL_Comb_tip
    if not (GLR.not_n2 and i_comb=2):  # чтоб можно  было посчитать н-1 и н-3
        if i_comb = 3 and RG.temp_a_v_gost: logging.info("\t" + "n-3 по ГОСТу не считаем"): exit
        for
            if i_comb > 0:
                GLR.kol_otkl = i_comb
                if Comb.f_Init(RG.OTKL_masiv, GLR.kol_otkl):
                    if Comb.fFirstCombination():
                        do
                        OTKL_Comb_tip()
                    loop
                    while Comb.fNextCombination()
                        OTKL_Comb_tip()
                else:  # если единственное сочетание
                    OTKL_Comb_tip()

                if GLR.kol_otkl = 1 and GLR.viborka_comb > 0 and GLR.protokol_XL:  mKO.Print_XL_mKO(
                    2)  # печать mKO 1 загаловки + нр, 2 рез н-1


Comb = None

if GLR.Tabl_otlk_kontrol > 0:  XL_sheet_oform()  # оформление таблицы по годам


# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в расчет +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
class rg2_tek_file:  # RG  RG. для хранения параметров текущего расчетного файла

    def Kontrol_init():
        # dim CountVL , ndx , Uip, Uiq , idx , i , j , key, CountTrans, NodeCount , tV11  , tN12
        if rastr.tables("vetv").cols.Find("_SortKey") < 1 and GLR.Tabl_otlk_kontrol > 0: rastr.tables("vetv").Cols.Add
        "_SortKey", 0  # добавить столбцы

        tN12 = rastr.Tables("node")
        tV11 = rastr.Tables("vetv")
        tV11.setsel("Kontrol!=0 + tip!=1")  # контр  ВЛ + без транс
        ReDim
        KontrolVL(tV11.Count - 1)
        CountVL = 0
        ndx = tV11.FindNextSel(-1)
        while ndx >= 0
            KontrolVL(CountVL) = ndx
            if tV11.count > 1 and GLR.Tabl_otlk_kontrol > 0:
                Uip = -1
                Uiq = -1
                for idx = 0 to tN12.Size-1
                if tN12.cols.item("ny").Z(idx) = tV11.cols.item("ip").Z(ndx): Uip = tN12.cols.item("uhom").Z(idx)
                if tN12.cols.item("ny").Z(idx) = tV11.cols.item("iq").Z(ndx): Uiq = tN12.cols.item("uhom").Z(idx)
                if Uip >= 0 And Uiq >= 0:
                    if Uip > Uiq:
                        tV11.cols.item("_SortKey").Z(ndx) = Uip * 10000 + Uiq
                    else:
                        tV11.cols.item("_SortKey").Z(ndx) = Uiq * 10000 + Uip

                    Exit
                    for

        DName_sub("vetv", ndx)
        CountVL = CountVL + 1
        ndx = tV11.FindNextSel(ndx)

    wend

    if tV11.count > 1 and GLR.Tabl_otlk_kontrol > 0:
        if tV11.cols.item("N").Z(KontrolVL(0)) > 0:
            for i = 1 to CountVL-1  # цикл сорт  по Н
            key = KontrolVL(i)
            j = i - 1
            do
            while j >= 0
                if fSort_N(KontrolVL(j)) < fSort_N(key): Exit
                Do
                KontrolVL(j + 1) = KontrolVL(j)
                j = j - 1
            loop
            KontrolVL(j + 1) = key

    else:
        for i = 1 to CountVL-1  # цикл сорт ВЛ
        key = KontrolVL(i)
        j = i - 1
        do
        while j >= 0
            if fVetvKey(KontrolVL(j)) > fVetvKey(key): Exit
            Do
            KontrolVL(j + 1) = KontrolVL(j)
            j = j - 1
        loop
        KontrolVL(j + 1) = key

        # контр ТРАНСЫ


tV11.setsel("Kontrol!=0 + tip=1")
ReDim
KontrolTrans(tV11.Count - 1)

CountTrans = 0
ndx = tV11.FindNextSel(-1)
while ndx >= 0
    KontrolTrans(CountTrans) = ndx
    if tV11.count > 1 and GLR.Tabl_otlk_kontrol > 0:
        Uip = -1
        Uiq = -1
        for idx = 0 to tN12.Size-1
        if tN12.cols.item("ny").Z(idx) = tV11.cols.item("ip").Z(ndx): Uip = tN12.cols.item("uhom").Z(idx)
        if tN12.cols.item("ny").Z(idx) = tV11.cols.item("iq").Z(ndx): Uiq = tN12.cols.item("uhom").Z(idx)
        if Uip >= 0 And Uiq >= 0:
            if Uip > Uiq:
                tV11.cols.item("_SortKey").Z(ndx) = Uip * 10000 + Uiq
            else:
                tV11.cols.item("_SortKey").Z(ndx) = Uiq * 10000 + Uip

            Exit
            for

DName_sub("vetv", ndx)
CountTrans = CountTrans + 1
ndx = tV11.FindNextSel(ndx)
wend
if tV11.count > 1 and GLR.Tabl_otlk_kontrol > 0:
    if tV11.cols.item("N").Z(KontrolTrans(0)) > 0:
        for i = 1 to CountTrans-1  # цикл сорт Транс по Н
        key = KontrolTrans(i)
        j = i - 1
        do
        while j >= 0
            if fSort_N(KontrolTrans(j)) < fSort_N(key): Exit
            Do
            KontrolTrans(j + 1) = KontrolTrans(j)
            j = j - 1
        loop
        KontrolTrans(j + 1) = key

else:
    for i = 1 to CountTrans - 1  # цикл сорт Транс по КЕЙ
    key = KontrolTrans(i)
    j = i - 1
    do
    while j >= 0
        if fVetvKey(KontrolTrans(j)) > fVetvKey(key): Exit
        Do
        KontrolTrans(j + 1) = KontrolTrans(j)
        j = j - 1
    loop
    KontrolTrans(j + 1) = key

tV11.setsel("Kontrol!=0+tip=1+KontrOO!=0")  # общая обмотка
ReDim
KontrolTransOO(tV11.Count - 1)

CountTrans = 0
ndx = tV11.FindNextSel(-1)
while ndx >= 0
    KontrolTransOO(CountTrans) = ndx
    if tV11.count > 1 and GLR.Tabl_otlk_kontrol > 0:
        Uip = -1
        Uiq = -1
        for idx = 0 to tN12.Size-1
        if tN12.cols.item("ny").Z(idx) = tV11.cols.item("ip").Z(ndx): Uip = tN12.cols.item("uhom").Z(idx)
        if tN12.cols.item("ny").Z(idx) = tV11.cols.item("iq").Z(ndx): Uiq = tN12.cols.item("uhom").Z(idx)
        if Uip >= 0 And Uiq >= 0:
            if Uip > Uiq:
                tV11.cols.item("_SortKey").Z(ndx) = Uip * 10000 + Uiq
            else:
                tV11.cols.item("_SortKey").Z(ndx) = Uiq * 10000 + Uip

            Exit
            for

CountTrans = CountTrans + 1
ndx = tV11.FindNextSel(ndx)
wend
if tV11.count > 1 and GLR.Tabl_otlk_kontrol > 0:
    for i = 1 to CountTrans - 1  # цикл сорт Транс по КЕЙ
    key = KontrolTransOO(i)
    j = i - 1
    do
    while j >= 0
        if fVetvKey(KontrolTransOO(j)) > fVetvKey(key): Exit
        Do
        KontrolTransOO(j + 1) = KontrolTransOO(j)
        j = j - 1
    loop
    KontrolTransOO(j + 1) = key

    #     КОНТРОЛ УЗЛЫ
uhom_korr_sub("Kontrol")
tN12.setsel("Kontrol")
ReDim
KontrolNode(tN12.Count - 1)
NodeCount = 0
ndx = tN12.FindNextSel(-1)
while ndx >= 0  # ЗАПИСЬ KontrolNode
    KontrolNode(NodeCount) = ndx
    DName_sub("node", ndx)

    if tN12.cols.item("uhom").Z(ndx) > 100:
        if tN12.cols.item("umin").Z(ndx) = 0:  tN12.cols.item("umin").Z(ndx) = tN12.cols.item("uhom").Z(
            ndx) * 1.15 * 0.7
        if tN12.cols.item("umin_av").Z(ndx) = 0:  tN12.cols.item("umin_av").Z(ndx) = tN12.cols.item("uhom").Z(
            ndx) * 1.1 * 0.7
        if tN12.cols.item("umax").Z(ndx) = 0:  tN12.cols.item("umin_av").Z(ndx) = tN12.cols.item("uhom").Z(
            ndx) * 1.1 * 0.7

    ndx = tN12.FindNextSel(ndx)
    NodeCount = NodeCount + 1
wend
if tN12.count > 1 and GLR.Tabl_otlk_kontrol > 0:
    if tN12.cols.item("N").Z(KontrolNode(0)) > 0:  # СОРТИРОВКА
        for i = 1 to NodeCount - 1  # цикл сорт N
        key = KontrolNode(i)
        j = i - 1
        do
        while j >= 0
            if fSort_NNod(KontrolNode(j), "N") < fSort_NNod(key, "N"): Exit
            Do
            KontrolNode(j + 1) = KontrolNode(j)
            j = j - 1
        loop
        KontrolNode(j + 1) = key

else:
    for i = 1 to NodeCount-1  # цикл сорт U
    key = KontrolNode(i)
    j = i - 1
    do
    while j >= 0
        if fSort_NNod(KontrolNode(j), "uhom") > fSort_NNod(key, "uhom"): Exit
        Do
        KontrolNode(j + 1) = KontrolNode(j)
        j = j - 1
    loop
    KontrolNode(j + 1) = key

logging.info(
    "\t" + "контролитуемых ветвей + узлов:" + str(UBound(KontrolVL, 1) + UBound(KontrolTrans, 1) + 2) + " + " + str(
        UBound(KontrolNode, 1) + 1) + " = " + str(
        UBound(KontrolVL, 1) + UBound(KontrolTrans, 1) + 2 + UBound(KontrolNode, 1) + 1))


def OTKL1_masiv():  # ОТКЛЮЧАЕМЫЕ ЭЛЕМЕНТЫ в массив
    # dim n_otkl_v , n_otkl_n , n_otkl , VetvCount , i , j , ndx , key , key1 , key2 , auto_est , tV12  , tN13
    auto_est = False  # наличие задания в полях remont_add otkl_add
    tN13 = rastr.Tables("node")
    tV12 = rastr.Tables("vetv")
    tV12.setsel(GLR.vibor_otkl + "+!sta")
    n_otkl_v = tV12.Count
    if GLR.otkl_ssch:
        tN13.setsel(GLR.vibor_otkl)
        n_otkl_n = tN13.Count
    else:
        n_otkl_n = 0

    n_otkl = n_otkl_v + n_otkl_n - 1  # -1 тк с нуля

    ReDim
    OTKL_masiv(2, n_otkl)  # ndx,  "vetv" и "ip=+iq=np="
    VetvCount = 0

    tV12.setsel(GLR.vibor_otkl + "+!sta")  # ВЕТВИ
    ndx = tV12.FindNextSel(-1)
    while ndx >= 0
        OTKL_masiv(0, VetvCount) = ndx
        OTKL_masiv(1, VetvCount) = "vetv"
        OTKL_masiv(2, VetvCount) = rastr.Tables("vetv").SelString(ndx)
        VetvCount = VetvCount + 1
        DName_sub("vetv", ndx)
        if tV12.cols.Find("remont_add") > -1 or tV12.cols.Find("otkl_add") > -1:
            tV12.cols.item("remont_add").Z(ndx) = replace(tV12.cols.item("remont_add").Z(ndx), " ", "")
            tV12.cols.item("otkl_add").Z(ndx) = replace(tV12.cols.item("otkl_add").Z(ndx), " ", "")
            if tV12.cols.item("remont_add").Z(ndx) != "" or tV12.cols.item("otkl_add").Z(ndx) != "": auto_est = True

        ndx = tV12.FindNextSel(ndx)
    wend

    for i = 1 to VetvCount - 1  # цикл сорт
    key = OTKL_masiv(0, i)
    key1 = OTKL_masiv(1, i)
    key2 = OTKL_masiv(2, i)
    j = i - 1
    do
    while j >= 0
        if fVetvKey(OTKL_masiv(0, j)) > fVetvKey(key): Exit
        Do
        OTKL_masiv(0, j + 1) = OTKL_masiv(0, j)
        OTKL_masiv(1, j + 1) = OTKL_masiv(1, j)
        OTKL_masiv(2, j + 1) = OTKL_masiv(2, j)
        j = j - 1
    loop
    OTKL_masiv(0, j + 1) = key
    OTKL_masiv(1, j + 1) = key1
    OTKL_masiv(2, j + 1) = key2


if GLR.otkl_ssch:
    tN13.setsel(GLR.vibor_otkl)  # УЗЛЫ
    ndx = tN13.FindNextSel(-1)

    while ndx >= 0
        DName_sub("node", ndx)
        OTKL_masiv(0, VetvCount) = ndx
        OTKL_masiv(1, VetvCount) = "node"
        OTKL_masiv(2, VetvCount) = rastr.Tables("node").SelString(ndx)
        VetvCount = VetvCount + 1

        if tN13.cols.Find("remont_add") > -1 or tN13.cols.Find("otkl_add") > -1:
            tN13.cols.item("remont_add").Z(ndx) = replace(tN13.cols.item("remont_add").Z(ndx), " ", "")
            tN13.cols.item("otkl_add").Z(ndx) = replace(tN13.cols.item("otkl_add").Z(ndx), " ", "")
            if tN13.cols.item("remont_add").Z(ndx) != "" or tN13.cols.item("otkl_add").Z(ndx) != "": auto_est = True

        ndx = tN13.FindNextSel(ndx)
    wend

logging.info(
    "\t" + "отключаемых ветвей + узлов:" + str(n_otkl_v) + " + " + str(n_otkl_n) + " = " + str(n_otkl_v + n_otkl_n))
if auto_est = True:
    if rastr.Tables.Find("AutoZad") < 0:
        logging.info("!!! НЕ ЗАГРУЖЕН ШАБЛОН АВТОМАТИКИ с таблицей  AutoZad (.amt) !!")
    else:
        if rastr.Tables("AutoZad").size = 0:  logging.info("!!! НЕ ЗАГРУЖЕН ФАЙЛ АВТОМАТИКИ (.amt) !!")


def redim_Otkl_Comb_tek(kol1, kol2):
    ReDim
    Otkl_Comb_tek(kol1, kol2)


# end class# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
class raschot_tek_comb:  # RGR. RGR для хранения параметров текущего расчета (нр или сочетания) RGR.
    # dim raschot_name , remont_name1 , remont_name2 , otkl_name , NodeNetReserv , NodeRezerv
    # dim AutoZad #  формирует Dorgm
    # dim autoTXT_fact_Otkl_Remont , autoTXT_fact_Otkl_Remont_tek ,  autoTXT_factPA  #  "Действие на " - формирует атоматика
    #  по факту отключения, загрузки
    # dim AutoKontrol , autoNDX_listV_info , autoNDX_listN_info
    # dim FLAG_ris_tabl_add_PA#  для добавления рис и расчета в таб с учетом дейстся ПА даже если перегрузки исчезли
    # dim add_rgm_PA, autoTXT_fPA, add_risunok , name_ris, name_ris_info,  otkl_key  , remont_key1, remont_key2
    Private
    txt_temp

    def init_new():
        redim
        AutoKontrol(1)
        raschot_name = "Нормальный режим"  # полное название: откл в схеме ремонта
        otkl_name = "-"  # отключаемый элемент
        remont_name1 = "-"  # ремонтируемый элемент1
        remont_name2 = "-"  # ремонтируемый элемент2

        otkl_key = "-"  # отключаемый элемент
        remont_key1 = "-"  # ремонтируемый элемент1
        remont_key2 = "-"  # ремонтируемый элемент2

        autoNDX_listV_info = ""
        autoNDX_listN_info = ""
        NodeNetReserv = ""
        NodeRezerv = ""
        AutoZad = ""
        autoTXT_fPA = ""
        FLAG_ris_tabl_add_PA = 0
        autoTXT_fact_Otkl_Remont = ""
        autoTXT_fact_Otkl_Remont_tek = ""  # [действие auto_run]
        add_rgm_PA = False  # тут не менять если присвоено значение 1 то добавляем режим с ПА

    def init_PA():
        AutoZad = ""
        autoNDX_listV_info = ""
        autoNDX_listN_info = ""
        AutoKontrol(0) = ""
        AutoKontrol(1) = 0
        add_rgm_PA = False  # если присвоено значение 1 то добавляем режим с ПА
        if autoTXT_fPA != "": raschot_name = raschot_name + ". " + autoTXT_fPA  #


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
class Uniquizer:  # запись отключений
    public

    def fReadOtkl:
        fReadOtkl = True  # новое сочетание
        # dim ndx , Otkl_Print , tab_vetv

        tab_vetv = rastr.Tables("vetv")
        tab_vetv.setsel("sta")
        ndx = tab_vetv.FindNextSel(-1)
        Otkl_Print = ""
        while ndx >= 0
            Otkl_Print = Otkl_Print + str(ndx)
            Otkl_Print = Otkl_Print + ";"
            ndx = tab_vetv.FindNextSel(ndx)
        wend
        if scan_otkl_remont != "" and GLR.otkl_remont_shema:
            fReadOtkl = True
        else:
            if RG.Dict_StoredOTKLUCH.Exists(Otkl_Print):  # проверка веречня отключений на уникальность

                fReadOtkl = False  # если такой набор откл уже был
                logging.info("\t" + "отклонено повторяющееся сочетание: " + RGR.otkl_name)
            else:
                RG.Dict_StoredOTKLUCH.Add(Otkl_Print, 1)
                fReadOtkl = True  # End def return


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
class Combinator:
    # dim Counter(), Vcopy(), tip_comb#  (tip): -1 - простое откл,  0 - откл +ремонты, 1 - ремонт + откл +ремонт , 2 - ремонт + ремонт +откл (номер указывает положение отключаемого элемента в Otkl_Comb_tek)
    Private
    m, n
    public

    def f_Init(Vsource, Em):  # (массив отключений, кол откл)
        f_Init = True
        m = Em  # кол откл
        n = Ubound(Vsource, 2) + 1  # количество элеметов в массиве RG.OTKL_masiv
        Redim
        Vcopy(2, n)  # копируем массив отключаемых элементов
        for i = 0 to n-1
        Vcopy(0, i) = Vsource(0, i)
        Vcopy(1, i) = Vsource(1, i)
        Vcopy(2, i) = Vsource(2, i)

    if m > n + 1:
        logging.info("!!!Количество элементов в сочетаниях меньше количества отключаемых элементов!!!")
        f_Init = False
    else:
        Redim
        Counter(n - 1)  # создаем массив размером с массив отключений - 1   #  ????????????? может не нада -1?????????
        for i = 0 to Ubound(Counter)
        Counter(i) = i

        # End def return


public


def fFirstCombination():  # первая комбинация из первого элемента в массиве откл
    RG.redim_Otkl_Comb_tek(3, m - 1)
    GLR.OTKL1_ndx_tek = ""  # строка c текущими индексами отключаемых элементов массива отключений через запятую
    for i = 0 to m-1
    RG.Otkl_Comb_tek(0, i) = Vcopy(0, Counter(i))
    RG.Otkl_Comb_tek(1, i) = Vcopy(1, Counter(i))
    RG.Otkl_Comb_tek(2, i) = Vcopy(2, Counter(i))
    if i = 0: GLR.OTKL1_ndx_tek = str(Counter(i))  else:
        GLR.OTKL1_ndx_tek = GLR.OTKL1_ndx_tek + "," + str(Counter(i))  #


# logging.info( GLR.OTKL1_ndx_tek )    #  del
fFirstCombination = m < n  # кол_откл элементов больше количества элементов в сочетании
# End def return

public


def fNextCombination():
    if Counter(m - 1) < n - 1:
        Counter(m - 1) = Counter(m - 1) + 1  # m=0 те  н-1 каждый проход записываем Counter(0)=+1
    else:
        for i = m-2 to 0 step -1
        if Counter(i) < n - m + 1:
            Counter(i) = Counter(i) + 1
            for j = i + 1 to m - 1
            Counter(j) = Counter(j - 1) + 1

        exit
        for  # не хватало


RG.redim_Otkl_Comb_tek(3, m - 1)
GLR.OTKL1_ndx_tek = ""

for i = 0 to m-1
RG.Otkl_Comb_tek(0, i) = Vcopy(0, Counter(i))
RG.Otkl_Comb_tek(1, i) = Vcopy(1, Counter(i))
RG.Otkl_Comb_tek(2, i) = Vcopy(2, Counter(i))
if i = 0: GLR.OTKL1_ndx_tek = str(Counter(i))  else:
    GLR.OTKL1_ndx_tek = GLR.OTKL1_ndx_tek + "," + str(Counter(i))

# logging.info( GLR.OTKL1_ndx_tek )  #  del
if Counter(0) = n - m:
    fNextCombination = False
else:
    fNextCombination = True

# End def return

# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
class mKontrol_Otkl:  # класс для создания массивов контрол - н-1, один обект класса для вл и один для тр и один для напряжений
    Private
    i, ii, x, xx

    # dim Otkl_zag_nr() , mOtkl_Otkl (), mOtkl_remont () ,  section () ,  otkl_count , arrNDX , tV1 , tN1 , N_pervoe , N_zamena
    # dim N_section , search_section () , dname0 , dname1 , dname2, delta_0, delta_1 , delta_01, delta_10, delta_20, delta_02, delta_12, delta_21
    def KontrolOtkl_init():  # запись загрузки в нр и инициализация массивов откл - контрол
        tV1 = rastr.Tables("vetv")
        tN1 = rastr.Tables("node")
        otkl_count = ubound(RG.OTKL_masiv, 2) + 1  # кол откл элементов
        ReDim
        Otkl_zag_nr(otkl_count - 1)
        for i = 0 to otkl_count -1  # цикл записи загрузки в нр
        if RG.OTKL_masiv(1, i) = "vetv":  # ветвь
            Otkl_zag_nr(i) = round(tV1.cols.item("i_zag").Z(RG.OTKL_masiv(0, i)) * 1000)
        else:  # узел

    # Otkl_zag_nr (i) = round ( tN1.cols.item("otv_min").Z(RG.OTKL_masiv(0 , i ))   , 1 )

    redim
    mOtkl_Otkl(otkl_count - 1, otkl_count - 1)  # ( вертикаль - загрузка , горизонт - откл)
    redim
    mOtkl_remont(otkl_count - 1, otkl_count - 1)  # ( вертикаль - загрузка , горизонт - откл)


def mKontrol_Otkl_RecN1(tip_otkl):  # ("откл","ремонт") запись изм зарузки в н-1 относительно нр

    for ii = 0 to otkl_count - 1  #
    if RG.OTKL_masiv(1, ii) = "vetv":  # ветвь
        if tip_otkl = "откл" or tip_otkl = "":  #
            if tV1.cols.item("i_zag").Z(RG.OTKL_masiv(0, ii)) = 0: mOtkl_Otkl(ii, float(GLR.OTKL1_ndx_tek)) = 0
            if tV1.cols.item("i_zag").Z(RG.OTKL_masiv(0, ii)) > 0: mOtkl_Otkl(ii, float(GLR.OTKL1_ndx_tek)) = round(
                tV1.cols.item("i_zag").Z(RG.OTKL_masiv(0, ii)) * 1000 - Otkl_zag_nr(ii))

        if tip_otkl = "ремонт" or tip_otkl = "":  #
            if tV1.cols.item("i_zag").Z(RG.OTKL_masiv(0, ii)) = 0: mOtkl_remont(ii, float(GLR.OTKL1_ndx_tek)) = 0
            if tV1.cols.item("i_zag").Z(RG.OTKL_masiv(0, ii)) > 0: mOtkl_remont(ii, float(GLR.OTKL1_ndx_tek)) = round(
                tV1.cols.item("i_zag").Z(RG.OTKL_masiv(0, ii)) * 1000 - Otkl_zag_nr(ii))

    else:  # узел
# if tN1.cols.item("vras").Z(RG.OTKL_masiv (0 , ii ) ) > 0 and tN1.cols.item("umin").Z(RG.OTKL_masiv (0 , ii )) > 0:
#     mOtkl_Otkl ( ii , float (GLR.OTKL1_ndx_tek ))  = round ( tN1.cols.item("otv_min").Z(RG.OTKL_masiv (0 , ii ) ) - Otkl_zag_nr (ii) , 1 )
# else:
#     mOtkl_Otkl ( ii , float (GLR.OTKL1_ndx_tek ))  = 0
# # end if

public


def fTest_Comb():  # возвращает лож если не нужно считать режим и истина если нужно
    #   (comb.tip_comb): -1 - простое откл,  0 - откл +ремонты, 1 - ремонт + откл +ремонт , 2 - ремонт + ремонт +откл (номер указывает положение отключаемого элемента в Otkl_Comb_tek)
    #  для узлов нет
    #  сочетание  рассматривать, если при отключении любого элемента из сочетания сумма изменения загрузки других больше заданного
    arrNDX = split(GLR.OTKL1_ndx_tek, ",")
    for i = 0 to ubound (arrNDX)
    arrNDX(i) = float(arrNDX(i))


dname0 = tV1.cols.item("dname").ZS(RG.OTKL_masiv(0, arrNDX(0)))
dname1 = tV1.cols.item("dname").ZS(RG.OTKL_masiv(0, arrNDX(1)))
if comb.tip_comb < 1:
    delta_0 = mOtkl_Otkl(arrNDX(1), arrNDX(0)) else:
    delta_0 = mOtkl_remont(arrNDX(1), arrNDX(0))
if comb.tip_comb = 1: delta_1 = mOtkl_Otkl(arrNDX(0), arrNDX(1)) else:
    delta_1 = mOtkl_remont(arrNDX(0), arrNDX(1))
fTest_Comb = False
if ubound(arrNDX) = 1:  # н-2
    if abs(delta_0) > GLR.viborka_comb or abs(delta_1) > GLR.viborka_comb:
        fTest_Comb = True
        if GLR.print_viborka: logging.info(
            "\t" + "н-2 сочетание проходит: " + dname0 + " (" + str(delta_1) + "%); " + dname1 + " (" + str(
                delta_0) + "%)[" + str(comb.tip_comb) + "]")  # имя (его загрузка при отключении других элементов)
    else:
        if GLR.print_viborka: logging.info(
            "\t" + "н-2 сочетание отклонено:" + dname0 + " (" + str(delta_1) + "%); " + dname1 + " (" + str(
                delta_0) + "%)[" + str(comb.tip_comb) + "]")

elif ubound(arrNDX) = 2:  # н-3
    dname2 = tV1.cols.item("dname").ZS(RG.OTKL_masiv(0, arrNDX(2)))

    if comb.tip_comb = 1: delta_01 = mOtkl_Otkl(arrNDX(0), arrNDX(1)) else:
        delta_01 = mOtkl_remont(arrNDX(0), arrNDX(1))
    if comb.tip_comb < 1:
        delta_10 = mOtkl_Otkl(arrNDX(1), arrNDX(0)) else:
        delta_10 = mOtkl_remont(arrNDX(1), arrNDX(0))
    if comb.tip_comb < 1:
        delta_20 = mOtkl_Otkl(arrNDX(2), arrNDX(0)) else:
        delta_20 = mOtkl_remont(arrNDX(2), arrNDX(0))
    if comb.tip_comb = 2: delta_02 = mOtkl_Otkl(arrNDX(0), arrNDX(2)) else:
        delta_02 = mOtkl_remont(arrNDX(0), arrNDX(2))
    if comb.tip_comb = 2: delta_12 = mOtkl_Otkl(arrNDX(1), arrNDX(2)) else:
        delta_12 = mOtkl_remont(arrNDX(1), arrNDX(2))
    if comb.tip_comb = 1: delta_21 = mOtkl_Otkl(arrNDX(2), arrNDX(1)) else:
        delta_21 = mOtkl_remont(arrNDX(2), arrNDX(1))

    if (abs(delta_01) + abs(delta_02)) > GLR.viborka_comb and (abs(delta_10) + abs(delta_12)) > GLR.viborka_comb and (
            abs(delta_20) + abs(delta_21)) > GLR.viborka_comb:
        fTest_Comb = True
        if GLR.print_viborka: logging.info(
            "\t" + "н-3 сочетание проходит: " + dname0 + " (" + str(delta_01) + "% + " + str(
                delta_02) + "%); " + dname1 + " (" + str(delta_10) + "% + " + str(
                delta_12) + "%)" + dname2 + " (" + str(delta_20) + "% + " + str(delta_21) + "%)[" + str(
                tip_comb) + "]")  # имя (его загрузка при отключении других элементов)
    else:
        if GLR.print_viborka: logging.info(
            "\t" + "н-3 сочетание отклонено: " + dname0 + " (" + str(delta_01) + "% + " + str(
                delta_02) + "%); " + dname1 + " (" + str(delta_10) + "% + " + str(
                delta_12) + "%)" + dname2 + " (" + str(delta_20) + "% + " + str(delta_21) + "%)[" + str(
                tip_comb) + "]")  # End def return


def find_section():  # поиск сечения
    redim
    search_section(otkl_count - 1, otkl_count - 1)  # = mOtkl_Otkl записывать номера сечений
    redim
    section(otkl_count)

    N_section = 1
    for i = 0 to otkl_count - 1
    for ii = 0 to otkl_count - 1
    if abs(mOtkl_Otkl(i, ii)) > GLR.viborka_comb:
        search_section(i, ii) = N_section
        search_section(ii, i) = N_section
        N_section = N_section + 1
        for i = 0 to otkl_count - 1


N_pervoe = 0
for ii = 0 to otkl_count - 1
if not isempty(search_section(i, ii)) and N_pervoe = 0: N_pervoe = search_section(i, ii)
if not isempty(search_section(i, ii)) and N_pervoe > 0:
    N_zamena = search_section(i, ii)

    for x = 0 to otkl_count - 1
    for xx = 0 to otkl_count - 1
    if search_section(x, xx) = N_zamena:  search_section(x, xx) = N_pervoe

for i = 0 to otkl_count - 1
section(i) = 0
for ii = 0 to otkl_count - 1
if not isempty(search_section(i, ii)):
    section(i) = search_section(i, ii)  # записать     def Print_XL_mKO (tip_print):
if tip_print = 1:
    #  печать загаловки
    GLR.XL_print_mKOO.cell(4, 1).Value = "N сечения"
    GLR.XL_print_mKOO.cell(4, 2).Value = "index"
    GLR.XL_print_mKOO.cell(4, 3).Value = "таблица"
    GLR.XL_print_mKOO.cell(4, 4).Value = "ключ"
    GLR.XL_print_mKOO.cell(4, 5).Value = "имя"
    Print_XL_otklNAME(GLR.XL_print_mKOO, 5, 5, RG.OTKL_masiv, "верт", "dname")  #
    Print_XL(GLR.XL_print_mKOO, 2, 5, RG.OTKL_masiv, 2, "гор", "", "", "")  #

    if GLR.print_viborka:
        Print_XL(GLR.XL_print_mKOO, 7, 1, RG.OTKL_masiv, 2, "верт", "", "", "")  #

        Print_XL_otklNAME(GLR.XL_print_mKOO, 7, 4, RG.OTKL_masiv, "гор", "dname")  #
        Print_XL(GLR.XL_print_mKOO, 6, 5, Otkl_zag_nr, 1, "верт", "", "", "")  # печать нр#

elif tip_print = 2:  # печать mOtkl_Otkl
    find_section()
    Print_XL(GLR.XL_print_mKOO, 1, 5, section, 1, "верт", "", "", "")  #
    if GLR.print_viborka:
        Print_XL(GLR.XL_print_mKOO, 7, 5, mOtkl_Otkl, 2, "верт", "", "", "")  #
    else:
        if GS.N_rg2_File = 0: GLR.XL_print_mKOO.ListObjects.Add(1, GLR.XL_print_mKOO.Range(
            GLR.XL_print_mKOO.UsedRange.address))  # таблица        if GLR.print_viborka:
    if tip_print = 1:
        #  печать загаловки
        GLR.XL_print_mKOR.cell(4, 1).Value = "N сечения"
        GLR.XL_print_mKOR.cell(4, 2).Value = "index"
        GLR.XL_print_mKOR.cell(4, 3).Value = "таблица"
        GLR.XL_print_mKOR.cell(4, 4).Value = "ключ"
        GLR.XL_print_mKOR.cell(4, 5).Value = "имя"
        Print_XL_otklNAME(GLR.XL_print_mKOR, 5, 5, RG.OTKL_masiv, "верт", "dname")  #
        Print_XL(GLR.XL_print_mKOR, 2, 5, RG.OTKL_masiv, 2, "гор", "", "", "")  #
        Print_XL(GLR.XL_print_mKOR, 7, 1, RG.OTKL_masiv, 2, "верт", "", "", "")  #
        Print_XL_otklNAME(GLR.XL_print_mKOR, 7, 4, RG.OTKL_masiv, "гор", "dname")  #
        Print_XL(GLR.XL_print_mKOR, 6, 5, Otkl_zag_nr, 1, "верт", "", "", "")  # печать нр#
    elif tip_print = 2:  # печать mOtkl_remont
        Print_XL(GLR.XL_print_mKOR, 1, 5, section, 1, "верт", "", "", "")  #
        Print_XL(GLR.XL_print_mKOR, 7, 5, mOtkl_remont, 2, "верт", "", "", "")  #


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
class max_tok_class:  # класс для записи максимальных токов по присоединениям
    # dim dict_Imax_kontrol , temp#  для хранения ключей контроль со значениями токовой загрузки

    def init_sub():
        dict_Imax_kontrol = CreateObject(
            "Scripting.Dictionary")  # для хранения keys - ip=iq=np= Item - объект класса max_tok_ikontrol с инфо о одной ветви

    def test_max_tok(zad_ndx, zad_sez, zad_I, zad_god, zad_soch,
                     zad_n_otkl):  # zad_max_tok (0-ndx   ,  1-зим макс -10  , 600 ток, год , нр/откл)
        temp = rastr.tables("vetv").SelString(zad_ndx)
        if dict_Imax_kontrol.Exists(temp):  # есть, просерить и перезаписать
            dict_Imax_kontrol.Item(temp).test_add(zad_sez, zad_I, zad_god, zad_soch, zad_n_otkl)
        else:  # нет - добавить записать
            max_tok_ikontrol1 = max_tok_ikontrol
            max_tok_ikontrol1.init_max_toki(zad_ndx)
            max_tok_ikontrol1.test_add(zad_sez, zad_I, zad_god, zad_soch, zad_n_otkl)
            dict_Imax_kontrol.add(temp, max_tok_ikontrol1)
        # end if#

    def print_max_tok():
        # dim yo , xo #  верхний левый угол таблицы
        # dim rng,  PTCache, PT , List
        yo = 1: xo = 2
        # ШАПКА
        GLR.XL_max_tok.cell(yo, 1).Value = "Количество отключаемых элементов"
        GLR.XL_max_tok.cell(yo, 2).Value = "Наименование элемента"
        GLR.XL_max_tok.cell(yo, 3).Value = "ключ"
        GLR.XL_max_tok.cell(yo, 4).Value = "Сезон"
        GLR.XL_max_tok.cell(yo, 5).Value = "Ток, А"
        GLR.XL_max_tok.cell(yo, 6).Value = "Год"
        GLR.XL_max_tok.cell(yo, 7).Value = "Наименование режима"
        yo = yo + 1
        for EACH vetv_kontrol in dict_Imax_kontrol.Items
            disName = vetv_kontrol.disName
            kluch = vetv_kontrol.kluch
            for EACH vetv_kontrol_sez in vetv_kontrol.dict_Imax_ikontrol.Items
                GLR.XL_max_tok.cell(yo, 1).Value = vetv_kontrol_sez.zad_n_otkl1
                GLR.XL_max_tok.cell(yo, 2).Value = disName
                GLR.XL_max_tok.cell(yo, 3).Value = kluch
                GLR.XL_max_tok.cell(yo, 4).Value = vetv_kontrol_sez.zad_sez1
                GLR.XL_max_tok.cell(yo, 5).Value = vetv_kontrol_sez.zad_I1
                GLR.XL_max_tok.cell(yo, 6).Value = vetv_kontrol_sez.zad_god1
                GLR.XL_max_tok.cell(yo, 7).Value = vetv_kontrol_sez.zad_soch1
                yo = yo + 1
                if GLR.XL_max_tok.UsedRange.rows.count > 1:
            GLR.XL_max_tok.ListObjects.Add(1, GLR.XL_max_tok.Range(
                GLR.XL_max_tok.UsedRange.address))  # используемы диапозон листа

            GLR.XL_max_tok.ListObjects(1).Name = "I_max"
            tabl_I_max = GLR.XL_max_tok.ListObjects("I_max")
            GLR.XL_max_tok.Columns("A:AA").AutoFit

            GLR.Book_XL.Worksheets(1).Activate
            Sheets_add(GLR.Book_XL, List, "свод_Imax")

            PTCache = GLR.Book_XL.PivotCaches.Create(1, tabl_I_max)  # создать КЭШ
            PT = PTCache.CreatePivotTable("свод_Imax!R1C1", "stImax")  # создать сводную таблицу

            With
            PT
            .ManualUpdate = True  # не обновить сводную
            .AddFields
            Array("Наименование элемента", "Количество отключаемых элементов", "Наименование режима", "Год",
                  "Сезон"),, , False
            .AddDataField.PivotFields("Ток, А"), "Iрасч.,A ", -4136  #
            .RowAxisLayout
            1  # 1 xlTabularRow показывать в табличной форме!!!!
            .RowGrand = False  # удалить строку общих итогов
            .ColumnGrand = False  # удалить столбец общих итогов
            .MergeLabels = True  # обединять одинаковые ячейки
            .HasAutoFormat = False  # не обновлять ширину при обнавлении
            .NullString = "--"  # заменять пустые ячейки
            .PreserveFormatting = False  # сохранять формат ячеек при обнавлении
            .ShowDrillIndicators = False  # показывать кнопки свертывания
            .PivotCache.MissingItemsLimit = xlMissingItemsNone  # для норм отображения уникальных значений автофильтра ???????
            .PivotFields("Наименование элемента").Subtotals = Array(False, False, False, False, False, False, False,
                                                                    False, False, False, False,
                                                                    False)  # промежуточные итоги и фильтры
            .PivotFields("Год").Subtotals = Array(False, False, False, False, False, False, False, False, False, False,
                                                  False, False)  # промежуточные итоги и фильтры
            .PivotFields("Наименование режима").Subtotals = Array(False, False, False, False, False, False, False,
                                                                  False, False, False, False,
                                                                  False)  # промежуточные итоги и фильтры
            .PivotFields("Сезон").Subtotals = Array(False, False, False, False, False, False, False, False, False,
                                                    False, False, False)  # промежуточные итоги и фильтры
            .PivotFields("Количество отключаемых элементов").Subtotals = Array(False, False, False, False, False, False,
                                                                               False, False, False, False, False,
                                                                               False)  # промежуточные итоги и фильтры
            .ManualUpdate = False  # обновить сводную
            .TableStyle2 = ""  # стиль
            .DataBodyRange.HorizontalAlignment = -4108  # xlCenter = -4108
            .DataBodyRange.NumberFormat = "#,##0"

            With.TableRange1  # формат
            .WrapText = True  # перенос текста в ячейке
            .Borders(7).LineStyle = 1  # лево
            .Borders(8).LineStyle = 1  # верх
            .Borders(9).LineStyle = 1  # низ
            .Borders(10).LineStyle = 1  # право
            .Borders(11).LineStyle = 1  # внутри вертикаль
            .Borders(12).LineStyle = 1  #
        End
        With

    End
    With


class max_tok_ikontrol:  # класс для записи макс значения для конкретного элемента контроль периода и температуры
    # dim dict_Imax_ikontrol , disName , tip_vetv , kluch # ЛЭП-тр|dname|

    def init_max_toki(ndxi):
        dict_Imax_ikontrol = CreateObject(
            "Scripting.Dictionary")  # # для хранения keys - нр лнт макс ПЭВТ   Item - объект класса max_tok_z с инфо о I в разние сезоны - обекты
        disName = rastr.tables("vetv").cols.item("dname").Z(ndxi)
        tip_vetv = rastr.tables("vetv").cols.item("tip").Z(ndxi)
        kluch = rastr.tables("vetv").SelString(ndxi)

    def test_add(zad_sez, zad_I, zad_god, zad_soch, zad_n_otkl):  # добавить  если нет
        if dict_Imax_ikontrol.Exists(zad_sez + "|" + str(zad_n_otkl)):  # есть сравниваем
            dict_Imax_ikontrol.Item(zad_sez + "|" + str(zad_n_otkl)).test(zad_I, zad_god, zad_soch, zad_sez, zad_n_otkl)
        else:  # нет - добавляем
            max_tok_z1 = max_tok_z
            max_tok_z1.init()
            max_tok_z1.test(zad_I, zad_god, zad_soch, zad_sez, zad_n_otkl)
            dict_Imax_ikontrol.add(zad_sez + "|" + str(zad_n_otkl), max_tok_z1)


class max_tok_z:  # класс для записи макс значения для конкретного периода и температуры
    # dim zad_I1,zad_god1 ,zad_soch1 , zad_sez1 , zad_n_otkl1#   ток/этап/ откл
    def init():
        zad_I1 = 0

    def test(zad_I, zad_god, zad_soch, zad_sez, zad_n_otkl):
        if zad_I1 < zad_I:
            zad_I1 = zad_I
            zad_god1 = zad_god
            zad_soch1 = zad_soch
            zad_sez1 = zad_sez
            zad_n_otkl1 = zad_n_otkl


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def OTKL_Comb_tip():  # процедура определяет одно сочетание или два (если есть remont_add otkl_add), фильтра mKO
    # dim test_ssch , scan_otkl_remontz, raschot_da , otkl_element , iComb
    scan_otkl_remontz = scan_otkl_remont
    if GLR.kol_otkl = 1:
        if scan_otkl_remontz != "" and GLR.otkl_remont_shema:  # истина то схема ремонтная или откл адд
            #  две комбинации
            Comb.tip_comb = 0  # (tip): -1 - простое откл,  0 - откл +ремонты, 1 - ремонт + откл +ремонт , 2 - ремонт + ремонт +откл (номер указывает положение отключаемого элемента в Otkl_Comb_tek)
            PlaceOutages()
            if GLR.viborka_comb > 0: mKO.mKontrol_Otkl_RecN1("откл")  # запись в н-1 изм загрузки в н-1 относительно нр
            TopologyRestore()

            Comb.tip_comb = 1
            PlaceOutages()
            if GLR.viborka_comb > 0: mKO.mKontrol_Otkl_RecN1(
                "ремонт")  # запись в н-1 изм загрузки в н-1 относительно нр
            TopologyRestore()
            GLR.kol_test_da = GLR.kol_test_da + 2
        else:  # одна комбинация
            Comb.tip_comb = -1
            PlaceOutages()
            if GLR.viborka_comb > 0: mKO.mKontrol_Otkl_RecN1("")  # запись в н-1 изм загрузки в н-1 относительно нр
            TopologyRestore()
            GLR.kol_test_da = GLR.kol_test_da + 1

    else:  # GLR.kol_otkl > 1
        test_ssch = True
        if GLR.otkl_ssch:
            for i = 0 to ubound(RG.Otkl_Comb_tek, 2)
            if RG.Otkl_Comb_tek(1, i)  = "node": test_ssch = False  # не берем

    if test_ssch:
        otkl_element = array(
            -1)  # (tip): -1 - простое откл,  0 - откл +ремонты, 1 - ремонт + откл +ремонт , 2 - ремонт + ремонт +откл (номер указывает положение отключаемого элемента в Otkl_Comb_tek)
        if scan_otkl_remontz != "" and GLR.otkl_remont_shema:  # истина то схема ремонтная или откл адд
            if GLR.kol_otkl = 2:
                otkl_element = array(0, 1)  # две комбинации
            elif GLR.kol_otkl = 3:
                if scan_otkl_remontz = "012":
                    otkl_element = array(0, 1, 2)  # три комбинации
                else:
                    if scan_otkl_remontz = "01" or scan_otkl_remontz = "1" or scan_otkl_remontz = "0":
                        otkl_element = array(0, 1)  # две комбинации
                    elif scan_otkl_remontz = "12" or scan_otkl_remontz = "2":
                        otkl_element = array(1, 2)  # две комбинации
                    elif instr(scan_otkl_remontz, "02") > 0:
                        otkl_element = array(0, 2)  # две комбинации

        for each iComb in otkl_element
            raschot_da = True
            if GLR.viborka_comb > 0:
                Comb.tip_comb = iComb
                if not mKO.fTest_Comb():
                    raschot_da = False
                    GLR.kol_test_net = GLR.kol_test_net + 1
                    if raschot_da:
                Comb.tip_comb = iComb
                PlaceOutages(): TopologyRestore()
                GLR.kol_test_da = GLR.kol_test_da + 1


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def scan_otkl_remont():  # возвращает ложь если простое откл 2-х элементов и истина если нада простчиать два сочетания
    # dim tab
    scan_otkl_remont = ""
    for i = 0 to ubound (RG.Otkl_Comb_tek, 2)  # кол элементов
    tab = rastr.Tables(RG.Otkl_Comb_tek(1, i))

    if (tab.cols.item("remont_add").Z(RG.Otkl_Comb_tek(0, i)) != "" and tab.cols.item("sta_remont_add").Z(
            RG.Otkl_Comb_tek(0, i)) = 0) or (
            tab.cols.item("otkl_add").Z(RG.Otkl_Comb_tek(0, i)) != "" and tab.cols.item("sta_otkl_add").Z(
            RG.Otkl_Comb_tek(0, i)) = 0):
        if (tab.cols.item("otkl_add").Z(RG.Otkl_Comb_tek(0, i)) != tab.cols.item("remont_add").Z(
                RG.Otkl_Comb_tek(0, i)) or tab.cols.item("sta_remont_add").Z(RG.Otkl_Comb_tek(0, i)) != tab.cols.item(
                "sta_otkl_add").Z(RG.Otkl_Comb_tek(0, i))):
            scan_otkl_remont = scan_otkl_remont + str(i)


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def PlaceOutages():  # формирует otkl_add remont_add в RG.Otkl_Comb_tek Do Rgm
    #  (Comb.tip_comb): -1 - простое откл,  0 - откл +ремонты, 1 - ремонт + откл +ремонт , 2 - ремонт + ремонт +откл (номер указывает положение отключаемого элемента в Otkl_Comb_tek)
    # dim tablic , ny_otkl , rez , izm_pole , pole_sta , off_index , tablica, tablicaV , tip_auto_run , dname_i
    #  формирует RG.Otkl_Comb_tek (3,i) = "otkl_add"/"remont_add"-имя поля
    if GLR.kol_otkl = 1:  # n-1
        #  в RG.Otkl_Comb_tek (3,x)  записываем поле otkl_add или remont_add
        tablic = RG.Otkl_Comb_tek(1, 0)
        if Comb.tip_comb = 0: RG.Otkl_Comb_tek(3, 0) = rastr.tables(tablic).cols.item("otkl_add").ZS(
            RG.Otkl_Comb_tek(0, 0))   else:
            RG.Otkl_Comb_tek(3, 0) = None
        if Comb.tip_comb = 1: RG.Otkl_Comb_tek(3, 0) = rastr.tables(tablic).cols.item("remont_add").ZS(
            RG.Otkl_Comb_tek(0, 0)) else:
            RG.Otkl_Comb_tek(3, 0) = None

    elif GLR.kol_otkl = 2:  # n-2

        tablic = RG.Otkl_Comb_tek(1, 0)  # элемент 0 в Otkl_Comb_tek
        if Comb.tip_comb = 0: RG.Otkl_Comb_tek(3, 0) = rastr.tables(tablic).cols.item("otkl_add").ZS(
            RG.Otkl_Comb_tek(0, 0))   else:
            RG.Otkl_Comb_tek(3, 0) = None
        if Comb.tip_comb = 1: RG.Otkl_Comb_tek(3, 0) = rastr.tables(tablic).cols.item("remont_add").ZS(
            RG.Otkl_Comb_tek(0, 0)) else:
            RG.Otkl_Comb_tek(3, 0) = None

        tablic = RG.Otkl_Comb_tek(1, 1)  # элемент 1 в Otkl_Comb_tek
        if Comb.tip_comb = 0: RG.Otkl_Comb_tek(3, 1) = rastr.tables(tablic).cols.item("remont_add").ZS(
            RG.Otkl_Comb_tek(0, 1)) else:
            RG.Otkl_Comb_tek(3, 1) = None
        if Comb.tip_comb = 1: RG.Otkl_Comb_tek(3, 1) = rastr.tables(tablic).cols.item("otkl_add").ZS(
            RG.Otkl_Comb_tek(0, 1))   else:
            RG.Otkl_Comb_tek(3, 1) = None

    elif GLR.kol_otkl = 3:  # n-3

        tablic = RG.Otkl_Comb_tek(1, 0)  # элемент 0 в Otkl_Comb_tek
        if Comb.tip_comb = 0: RG.Otkl_Comb_tek(3, 0) = rastr.tables(tablic).cols.item("otkl_add").ZS(
            RG.Otkl_Comb_tek(0, 0))   else:
            RG.Otkl_Comb_tek(3, 0) = None
        if Comb.tip_comb > 0:
            RG.Otkl_Comb_tek(3, 0) = rastr.tables(tablic).cols.item("remont_add").ZS(RG.Otkl_Comb_tek(0, 0)) else:
            RG.Otkl_Comb_tek(3, 0) = None

        tablic = RG.Otkl_Comb_tek(1, 1)  # элемент 1 в Otkl_Comb_tek
        if Comb.tip_comb != 1:
            RG.Otkl_Comb_tek(3, 1) = rastr.tables(tablic).cols.item("remont_add").ZS(RG.Otkl_Comb_tek(0, 1)) else:
            RG.Otkl_Comb_tek(3, 1) = None
        if Comb.tip_comb = 1: RG.Otkl_Comb_tek(3, 1) = rastr.tables(tablic).cols.item("otkl_add").ZS(
            RG.Otkl_Comb_tek(0, 1))   else:
            RG.Otkl_Comb_tek(3, 1) = None

        tablic = RG.Otkl_Comb_tek(1, 2)  # элемент 2 в Otkl_Comb_tek
        if Comb.tip_comb < 2:
            RG.Otkl_Comb_tek(3, 2) = rastr.tables(tablic).cols.item("remont_add").ZS(RG.Otkl_Comb_tek(0, 2)) else:
            RG.Otkl_Comb_tek(3, 2) = None
        if Comb.tip_comb = 2: RG.Otkl_Comb_tek(3, 2) = rastr.tables(tablic).cols.item("otkl_add").ZS(
            RG.Otkl_Comb_tek(0, 2))   else:
            RG.Otkl_Comb_tek(3, 2) = None

    if RG.loadRGM:
        PnQnPgKtrRestore()  # загрузить режим
        RG.loadRGM = False

    RGR = raschot_tek_comb  # сочетание
    RGR.init_new()

    #  отключаем элементы
    for off_index = 0 to GLR.kol_otkl - 1

    tablica = rastr.Tables(RG.Otkl_Comb_tek(1, off_index))  # "vetv" или "node"
    if RG.Otkl_Comb_tek(1, off_index) = "vetv":  # ОТКЛЧЕНИЕ ВЕТВЬ

        if tablica.cols.item("groupid").Z(RG.Otkl_Comb_tek(0, off_index)) > 0:
            rez = fVetv_Sta("groupid", tablica.cols.item("groupid").Z(RG.Otkl_Comb_tek(0, off_index)),
                            1)  # "ndx"/"groupid"/"kluch"; "ip=1,iq=2,np=0"; vkl_otkl= 1 отключить/ 0 включить)
        else:
            rez = fVetv_Sta("ndx", RG.Otkl_Comb_tek(0, off_index),
                            1)  # "ndx"/"groupid"/"kluch"; "ip=1,iq=2,np=0"; vkl_otkl= 1 отключить/ 0 включить)

        if not rez:
            logging.info("\t" + "PlaceOutages: Комбинация отклонена, тк ветвь уже была отключена: " + tablica.SelString(
                (RG.Otkl_Comb_tek(0, off_index))) + " {" + tablica.cols.item("dname").Z(
                RG.Otkl_Comb_tek(0, off_index)) + "}, N комб: " + str(GLR.N_rezh))
            exit

            def

    elif RG.Otkl_Comb_tek(1, off_index) = "node":  # ОТКЛЧЕНИЕ УЗЛА
        tablicaV = rastr.Tables("vetv")
        tablica.cols.item("sta").Z(RG.Otkl_Comb_tek(0, off_index)) = 1  # отключаем узел
        ny_otkl = tablica.cols.item("ny").Z(RG.Otkl_Comb_tek(0, off_index))
        tablicaV.setsel("ip=" + str(ny_otkl) + "|iq=" + str(ny_otkl))  # отключаем примыкающие к узлу ветви
        tablicaV.cols.item("sta").Calc(1)

    dname_i = tablica.cols.item("dname").ZS(RG.Otkl_Comb_tek(0, off_index))
    if instr(dname_i, ",") > 0: dname_i = mid(dname_i, 1, instr(dname_i, ",") - 1)
    if instr(dname_i, "(") > 0: dname_i = mid(dname_i, 1, instr(dname_i, "(") - 1)

    #  доп откл  или ремонт
    if GLR.otkl_remont_shema:
        if off_index = 0:
            if Comb.tip_comb = 0: izm_pole = "otkl_add": pole_sta = "sta_otkl_add": tip_auto_run = 0  #
            if Comb.tip_comb != 0: izm_pole = "remont_add": pole_sta = "sta_remont_add": tip_auto_run = 1
        elif off_index = 1:
            if Comb.tip_comb = 1: izm_pole = "otkl_add": pole_sta = "sta_otkl_add": tip_auto_run = 0  #
            if Comb.tip_comb != 1: izm_pole = "remont_add": pole_sta = "sta_remont_add": tip_auto_run = 1
        elif off_index = 2:
            if Comb.tip_comb = 2: izm_pole = "otkl_add": pole_sta = "sta_otkl_add": tip_auto_run = 0  #
            if Comb.tip_comb != 2: izm_pole = "remont_add": pole_sta = "sta_remont_add": tip_auto_run = 1

        if trim(tablica.cols.item(izm_pole).ZS(RG.Otkl_Comb_tek(0, off_index))) != "" and tablica.cols.item(pole_sta).Z(
                RG.Otkl_Comb_tek(0, off_index)) = 0:
            # redim kontr0 (1): kontr0 (0) = ""
            auto_run(tablica.cols.item(izm_pole).ZS(RG.Otkl_Comb_tek(0, off_index)), array("", 0),
                     tip_auto_run)  # tip_auto_run     0  - действие по факту отключения адд откл,         1 - действие при ремонте ,       2 действие по факту перегрузки
            if RGR.autoTXT_fact_Otkl_Remont_tek != "": dname_i = dname_i + " " + RGR.autoTXT_fact_Otkl_Remont_tek:  RGR.autoTXT_fact_Otkl_Remont_tek = ""
            # RG.loadRGM = True   # ????????        if Comb.tip_comb > -1:
        if off_index  = Comb.tip_comb:
            RGR.otkl_name = dname_i
            RGR.otkl_key = RG.Otkl_Comb_tek(2, off_index)
        else:
            if RGR.remont_name1 = "-":
                RGR.remont_name1 = dname_i
                RGR.remont_key1 = RG.Otkl_Comb_tek(2, off_index)
            else:
                RGR.remont_name2 = dname_i
                RGR.remont_key2 = RG.Otkl_Comb_tek(2, off_index) else:
            if off_index = 0: RGR.otkl_name = dname_i: RGR.otkl_key = RG.Otkl_Comb_tek(2, off_index)
            if off_index = 1: RGR.remont_name1 = dname_i: RGR.remont_key1 = RG.Otkl_Comb_tek(2, off_index)
            if off_index = 2: RGR.remont_name2 = dname_i: RGR.remont_key2 = RG.Otkl_Comb_tek(2, off_index)

    if RGR.otkl_name != "-":
        RGR.raschot_name = "Отключение " + RGR.otkl_name
        if RGR.remont_name1 != "-": RGR.raschot_name = RGR.raschot_name + " в схеме ремонта " + RGR.remont_name1
        if RGR.remont_name2 != "-": RGR.raschot_name = RGR.raschot_name + ", " + RGR.remont_name2
    else:
        RGR.raschot_name = "Ремонт  " + RGR.remont_name1

    # if spUniquizer.fReadOtkl: #  проверяем если это отключение уже было то ложь
    GS.kod_rgm = rastr.rgm("")
    if GS.kod_rgm = 1:
        GS.kod_rgm = rastr.rgm("")
    if GS.kod_rgm = 1:
        GS.kod_rgm = rastr.rgm("p")
    if GS.kod_rgm = 1:
        GS.kod_rgm = rastr.rgm("p")
    if GS.kod_rgm = 1:
        GS.kod_rgm = rastr.rgm("p")

        NodeTest()  # восстановление питания узлов
        if GLR.AutoShunt: AutoShunt_class_kor()  # процедура меняет Bsh  и записывает GS.AutoShunt_list

        DoRgm()  # сочетание
        GLR.N_rezh = GLR.N_rezh + 1
        while RGR.add_rgm_PA
            auto_run(RGR.AutoZad, RGR.AutoKontrol,
                     2)  # (zadanie , Kontrol , tip_auto_run) tip_auto_run 1 по факту отключения, 2 по факту перегрузки
            RGR.init_PA()
            DoRgm()  # сочетание с ПА
            GLR.N_rezh = GLR.N_rezh + 1
        wend


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fVetv_Sta(tip, znachenie,
              vkl_otkl):  # "ndx"/"groupid"/"kluch"; "ip=1+iq=2+np=0"; vkl_otkl= 1 отключить/ 0 включить)
    # dim tV2
    tV2 = rastr.Tables("vetv")
    fVetv_Sta = True
    if tip = "ndx":
        if znachenie > -1:
            if tV2.cols.item("sta").Z(znachenie) = vkl_otkl:
                fVetv_Sta = False
                Exit

                def
            else:
                tV2.cols.item("sta").Z(znachenie) = vkl_otkl

        else:
            logging.info("\t" + "ОШИбКА  fVetv_Sta ndx=-1")
            fVetv_Sta = False
            Exit

            def

    elif tip = "groupid":
        tV2.setsel("groupid=" + str(znachenie))
        #  ndxxx = tV2.FindNextSel(-1)
        #  if tV2.cols.item("sta").Z(ndxxx) = vkl_otkl:
        #      fVetv_Sta = False
        #      Exit def
        #  else:
        tV2.cols.item("sta").Calc(vkl_otkl)
        #  # end if

elif tip = "kluch":
tV2.setsel(znachenie)
ndxxx = tV2.FindNextSel(-1)
if tV2.cols.item("sta").Z(ndxxx) = vkl_otkl:
    fVetv_Sta = False
    Exit


    def
else:
    tV2.cols.item("sta").Calc(vkl_otkl)

else:
logging.info("\t" + "ОШИбКА  fVetv_Sta tip: " + tip)


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fNodeSta(tip, znachenie, vkl_otkl):  # "ndx"/"ny"; 121 ; vkl_otkl= 1 отключить/ 0 включить)
    # dim tN2
    tN2 = rastr.Tables("node")
    fNodeSta = True
    if tip = "ndx":
        if tN2.cols.item("sta").Z(znachenie) = vkl_otkl:
            fNodeSta = False
        Exit

        def
    else:
        tN2.cols.item("sta").Z(znachenie) = vkl_otkl

elif tip = "ny":
tN2.setsel("ny=" + str(znachenie))

if tN2.cols.item("sta").Z(tN2.FindNextSel(-1)) = vkl_otkl:
    fNodeSta = False
    Exit


    def
else:
    tN2.cols.item("sta").Calc(vkl_otkl)

else:
logging.info("\t" + "ошибка  fNodeSta tip: " + tip)


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def XL_sheet_oform():  # оформление таблицы по годам

    With
    GLR.XL_sheet.Range(GLR.XL_sheet.cell(3, 11), GLR.XL_sheet.cell(5, GLR.X_list - 1))  # формат шапки
    .WrapText = True  # перенос текста в ячейке
    .HorizontalAlignment = -4108  # выравнивание по центру
    .VerticalAlignment = -4108
    .Borders(7).LineStyle = 1  # лево
    .Borders(8).LineStyle = 1  # верх
    .Borders(9).LineStyle = 1  # низ
    .Borders(10).LineStyle = 1  # право
    .Borders(11).LineStyle = 1  # внутри вертикаль
    .Borders(12).LineStyle = 1  #


End
With
With
GLR.XL_sheet.Range(GLR.XL_sheet.cell(3, 11), GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 1,
                                                                 GLR.X_list - 1))  # рамка периметр  и весь текст
.VerticalAlignment = -4108
.Borders(7).LineStyle = 1  # лево
.Borders(8).LineStyle = 1  # верх
.Borders(9).LineStyle = 1  # низ
.Borders(10).LineStyle = 1  # право
.HorizontalAlignment = -4108
.VerticalAlignment = -4108
.WrapText = True  # перенос текста в ячейке
End
With
With
GLR.XL_sheet.cell(6, 15)
.WrapText = False
.HorizontalAlignment = -4131  # выравнивание по лево
End
With
GLR.XL_sheet.Columns("A:I").Hidden = True  # скрыть столбцы
GLR.XL_sheet.Columns("K:K").HorizontalAlignment = -4131  # выравнивание по лево
GLR.XL_sheet.PageSetup.PrintArea = "" + GLR.XL_sheet.cell(2, 11).Address + ":" + GLR.XL_sheet.cell(
    GLR.Y_list + GLR.Y_VL_Trans_V + 1, GLR.X_list - 1).Address + ""  # задать область печати
GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 2, 11),
                   GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 200,
                                      GLR.X_list - 1)).WrapText = True  # перенос текста в ячейке
DDD = split(GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 2, 1).Address, "$")
DDD1 = float(DDD(2))
DDD2 = str(DDD1) + ":" + str(DDD1 + 100)
GLR.XL_sheet.Rows(DDD2).EntireRow.AutoFit
GLR.Ntabl_OK = GLR.Ntabl_OK + 1


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def TablF_init():  # инициализация TablF_Sub ()
    # dim Kol_Tr_OO
    # dim tV3
    tV3 = rastr.Tables("vetv")
    GLR.Y_VL = (Ubound(RG.KontrolVL) - LBound(RG.KontrolVL)) * 2 + 2  # размер VL
    Kol_Tr_OO = 0
    tV3.setsel("KontrOO")
    Kol_Tr_OO = tV3.Count
    GLR.Y_VL_Trans = GLR.Y_VL + (
                Ubound(RG.KontrolTrans) - LBound(RG.KontrolTrans)) * 2 + Kol_Tr_OO + 2  # размер VL+Trans
    GLR.Y_VL_Trans_V = GLR.Y_VL_Trans + (Ubound(RG.KontrolNode) - LBound(RG.KontrolNode)) + 1
    GLR.XL_sheet.cell(2, 11).Value = RG.TEXT_NAME_TAB
    GLR.XL_sheet.Rows("7:435").RowHeight = 15  # высота строки
    if GLR.TablF_const = 0 and GLR.Tabl_otlk_kontrol > 0:
        TablF_Sub()
    elif GLR.TablF_const = 1 and GLR.Tabl_otlk_kontrol > 0:  # КОПИ ПАСТ табл Ф
        if RG.name_list(1) = "зим":
            if GLR.TablF_const = 1 and GLR.TablF_zim = 1:
                TablF_Sub()
                GLR.XL_sheet.Range(GLR.XL_sheet.cell(3, 1),
                                   GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 1, 14)).Copy
                GLR.TablF_sheets.cell(3, 1).PasteSpecial()
                GLR.TablF_zim = 0
            else:
                GLR.TablF_sheets.Range(GLR.TablF_sheets.cell(3, 1),
                                       GLR.TablF_sheets.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 1, 14)).Copy
                GLR.XL_sheet.cell(3, 1).PasteSpecial()

        else:  # if RG.name_list(1) = "лет":
            if GLR.TablF_const = 1 and GLR.TablF_let = 1:
                TablF_Sub()
                GLR.XL_sheet.Range(GLR.XL_sheet.cell(3, 1),
                                   GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 1, 14)).Copy
                GLR.TablF_sheets.cell(3, 15).PasteSpecial()
                GLR.TablF_let = 0
            else:
                GLR.TablF_sheets.Range(GLR.TablF_sheets.cell(3, 15),
                                       GLR.TablF_sheets.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 1, 28)).Copy
                GLR.XL_sheet.cell(3, 1).PasteSpecial()
    GLR.XL_sheet.Columns("L").Hidden = True: GLR.XL_sheet.Columns("N").Hidden = True  # скрыть столбцы


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def XL_sheet_add():  # добавить лист для таблицы расчета текущего режима

    GLR.Book_XL.Sheets.Add(, GLR.Book_XL.Sheets(GLR.Book_XL.Sheets.count))
    GLR.XL_sheet = GLR.Book_XL.Worksheets(GLR.Book_XL.Sheets.count)
    GLR.XL_sheet.Name = left(RG.Name_base, 31)
    # GLR.XL_sheet.Move ,GLR.Book_XL.Sheets(GLR.Book_XL.Sheets.count)

    GLR.XL_sheet.Columns(13).ColumnWidth = 16  # ширина столбца
    GLR.XL_sheet.Columns(14).ColumnWidth = 16
    GLR.XL_sheet.Rows(3).RowHeight = 50  # высота строки
    GLR.XL_sheet.Rows(4).RowHeight = GLR.EntireRow_OK  # высота строки
    GLR.XL_sheet.Columns(11).ColumnWidth = GLR.EntireColumn_OK  # ширина столбца
    GLR.XL_sheet.Columns(12).ColumnWidth = 12  # ширина столбца

    GLR.X_list = 15  # базовый столбец
    GLR.Y_list = 7  # базовая строка
    GLR.Y_VL = 0
    GLR.Y_VL_Trans = 0
    GLR.Y_VL_Trans_V = 0


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def GrfLoad():  # обнулить расчетную температуру в ветвях районах и тд
    if float(RG.god) = < GLR.God_1:  rastr.Load(1, GLR.Graf_1, fshablon(GLR.Graf_1))
    if float(RG.god) = > GLR.God_2:  rastr.Load(1, GLR.Graf_2, fshablon(GLR.Graf_2))


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def VIBOR_KONTROL_OTKL():  # главная процедура АВТОЗАД КОНТРОЛ ОТКЛ
    # dim tV4 , tV40 , tN3 , ndx_tVetvZ
    tN3 = rastr.Tables("node")
    tV4 = rastr.Tables("vetv")
    tV40 = rastr.Tables("vetv")

    if rastr.tables("vetv").cols.Find("Kontrol_all") < 1: rastr.tables("vetv").Cols.Add
    "Kontrol_all", 3  # для отметки всех участков ветвей отмеченных Kontrol

    if not GLR.Zad_vse_RG2:
        rastr.Tables("node").cols.item("Kontrol").Calc("0")
        rastr.Tables("node").cols.item("otkl1").Calc("0")
        rastr.Tables("node").cols.item("otkl2").Calc("0")
        rastr.Tables("node").cols.item("otkl3").Calc("0")

        rastr.Tables("vetv").cols.item("Kontrol").Calc("0")
        rastr.Tables("vetv").cols.item("KontrOO").Calc("0")
        rastr.Tables("vetv").cols.item("otkl1").Calc("0")
        rastr.Tables("vetv").cols.item("otkl2").Calc("0")
        rastr.Tables("vetv").cols.item("otkl3").Calc("0")

    if GLR.vibor_raschot:
        # # # #  выделить сеть для анализа # # #
        sel0()

        tN3.setsel(GLR.viborka_raschot)  # + "+uhom>90" #  выбираем все узлы района с напряжением 110 и более
        tN3.cols.item("sel").Calc("1")

        tV4.setsel("ip.sel|iq.sel")  # выбираем все ветви связанные с  выбранными узлами
        tV4.cols.item("sel").Calc("1")
        #  отмечаем узлы связаные с нашей выборкой ветвью
        if GLR.node_pojas_analiz > 0:  # количество поясов для отметки узлов примыкающих к выборке GLR.viborka_raschot
            Node_pojas_sel(GLR.node_pojas_analiz)  # отмкетить узлы  и ветви привыкающие к отмеченным узлам n раз

        # # # #  АНАЛИЗ СХЕМЫ  # # #
        rastr.Tables("node").cols.item("tupik").Calc("0")
        rastr.Tables("node").cols.item("tranzit").Calc("0")
        rastr.Tables("node").cols.item("uzlovaja").Calc("0")
        rastr.Tables("vetv").cols.item("tupik").Calc("0")
        rastr.Tables("vetv").cols.item("tranzit").Calc("0")

        tupik_auto("sel")  # процедура для определения тупиковых узлов среди отмеченых и отметка поля tupik узла

        tN3.setsel("tupik")  # убираем тупика из Sel
        tN3.cols.item("sel").Calc("0")

        tranzit_auto("sel")  # определяем транзит

        tip_vetv_auto()  # задать тип ветви    - тупик или транзит

        # # # #  выделить сеть для задания # # #
        sel0()

        tN3.setsel(GLR.viborka_raschot + "+uhom>90")  # выбираем все узлы района с напряжением 110 и более
        tN3.cols.item("sel").Calc("1")

        tV4.setsel("ip.sel|iq.sel")  # выбираем все ветви связанные с  выбранными узлами
        tV4.cols.item("sel").Calc("1")
        #  отмечаем узлы связаные с нашей выборкой ветвью
        if GLR.node_pojas_zad > 0:  Node_pojas_sel(
            GLR.node_pojas_zad)  # отмкетить узлы  и ветви привыкающие к отмеченным узлам n раз

        # # #  ФОРМИРОВАНИЕ ЗАДАНИЯ # # #

        Dict_vetv_tranzit = CreateObject("Scripting.Dictionary")  # колекция транзитов
        Dict_unik_value_sub("vetv", "tranzit", "sel",
                            Dict_vetv_tranzit)  # (нр "vetv","tranzit","sel",dict) наполняет Dictionary  уникальными значениями столбца таблицы
        # print_dic (Dict_vetv_tranzit)

        for Each value_tranzit in Dict_vetv_tranzit.Keys  # цикл по tranzit

            if value_tranzit > 0:
                tV4.setsel("tranzit=" + str(value_tranzit))
                if tV4.count = 1:  # если в транзите 1 ветвь
                    tV4.setsel("tranzit=" + str(
                        value_tranzit) + "+tip<2+!sta+ip.uhom>90+iq.uhom>90")  # если включена, 110 кВ и выше и не выключатель
                    tV4.cols.item("Kontrol").calc("1")
                    tV4.cols.item(GLR.vibor_otkl).calc("1")
                else:
                    Dict_vetv_groupid = CreateObject("Scripting.Dictionary")  # колекция groupid  в тек транзите
                    Dict_unik_value_sub("vetv", "groupid", "tranzit=" + str(value_tranzit),
                                        Dict_vetv_groupid)  # в    Dict_vetv_groupid добавить номера groupid в текущем транзите
                    # отметить КОНТРОЛ ветви
                    for Each value_groupid in Dict_vetv_groupid.Keys  # цикл по groupid в tranzit
                        # tV4.setsel ("tranzit=" + str (value_tranzit) + "+groupid=" + str (value_groupid))
                        if value_groupid > 0:
                            groupid_dtn_kontrol(
                                value_groupid)  # определяет количество уникальных значений   dtn   и отмечает КОНТРОЛ
                        else:  # если участки транзита без groupid
                            tV40.setsel("tranzit=" + str(value_tranzit) + "+groupid=" + str(
                                value_groupid) + "+tip<2+!sta+ip.uhom>90+iq.uhom>90")  # tip<2  не выключатель
                            ndx_tVetvZ = tV40.FindNextSel(-1)
                            while ndx_tVetvZ >= 0  #
                                tV40.cols.item("Kontrol").Z(ndx_tVetvZ) = 1
                                ndx_tVetvZ = tV40.FindNextSel(ndx_tVetvZ)
                            wend

                            # logging.info ( "\t" + "ключ:" + str (value_tranzit) + ", значение:" +str (Dict_vetv_tranzit.Item( value_tranzit )))
                    # print_dic (Dict_vetv_groupid)

                    # отметить ОТКЛЮЧИТЬ
                    Otkl_tranzit(value_tranzit)  # отметит ОТКЛЮЧ участки транзита
                    Dict_vetv_groupid.RemoveAll()

        Otkl_Kontrol_node()  # отметит КОНТРОЛ и ОТКЛ узлы из uzlovaja если примыкает хотя бы 1 выкл

    if GLR.Zad_RG2 > 0: import_RG2()  # загрузка ид из файла
    #  для отметки всех участков ветвей отмеченных Kontrol
    tV4.setsel("")
    tV4.cols.item("Kontrol_all").calc("Kontrol")
    tV4.setsel("Kontrol_all")
    ndx_tVetvZ = tV4.FindNextSel(-1)
    while ndx_tVetvZ >= 0  #
        if tV4.cols.item("groupid").Z(ndx_tVetvZ) > 0:
            tV40.setsel("groupid=" + tV4.cols.item("groupid").ZS(ndx_tVetvZ))
            tV40.cols.item("Kontrol_all").calc("1")

        ndx_tVetvZ = tV4.FindNextSel(ndx_tVetvZ)
    wend


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def Otkl_Kontrol_node():  # отметит КОНТРОЛ и ОТКЛ узлы из uzlovaja если примыкает хотя бы 1 выкл
    # dim ndx , tV5  , tN4
    tN4 = rastr.Tables("node")
    tV5 = rastr.Tables("vetv")
    tN4.setsel(GLR.viborka_raschot + "+uhom>90+uzlovaja")  # выбираем все узлы района с напряжением 110 и более
    ndx = tN4.FindNextSel(-1)

    while ndx >= 0  #
        ny = tN4.cols.item("ny").Z(ndx)
        tV5.setsel("(ip=" + str(ny) + "|iq=" + str(ny) + ")+tip=2")
        if tV5.count > 0:
            tN4.cols.item("Kontrol").Z(ndx) = 1
            if tN4.cols.item("uhom").Z(ndx) < 300: tN4.cols.item(GLR.vibor_otkl).Z(ndx) = 1

        ndx = tN4.FindNextSel(ndx)
    wend


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def Otkl_tranzit(value_tranzit):  # отметит ОТКЛЮЧ участки транзита
    # dim ndx, tV6  , tN5 , Vetv1 , ndx1 , ndx_v
    tN5 = rastr.Tables("node")
    tV6 = rastr.Tables("vetv")
    # выбор начала и конца транзита
    tV6.setsel("tip<2+ip.uhom>90+iq.uhom>90+" + "tranzit=" + str(value_tranzit) + "+!sta+((ip.tranzit!=" + str(
        value_tranzit) + ")|(iq.tranzit!=" + str(value_tranzit) + "))")
    ndx = tV6.FindNextSel(-1)

    pn_tranzit = fSum_pn_tranzit(value_tranzit)  # сумма нагрузки внутри транзита
    redim
    ndx_v(tV6.count - 1)  # массив ветвей начала и конца транзита
    counter = 0

    tN5.setsel("")
    while ndx >= 0  # цикл по ветвям

        if tV6.cols.item("tip").Z(ndx) = 2:  # выключатель , нада искать предыдущую ветвь
            Vetv1 = rastr.Tables("vetv")
            if tN5.cols.item("tranzit").Z(fNDX("node", tV6.cols.item("ip").Z(ndx))) = value_tranzit:
                Vetv1.setsel("tranzit=" + str(value_tranzit) + "+(ip=" + str(tV6.cols.item("ip").Z(ndx)) + "|iq=" + str(
                    tV6.cols.item("ip").Z(ndx)) + ")")
            else:
                Vetv1.setsel("tranzit=" + str(value_tranzit) + "+(ip=" + str(tV6.cols.item("iq").Z(ndx)) + "|iq=" + str(
                    tV6.cols.item("iq").Z(ndx)) + ")")

            ndx1 = tV6.FindNextSel(-1)
            ndx_v(counter) = ndx1
        else:  # вл или транс
            ndx_v(counter) = ndx

        counter = counter + 1
        ndx = tV6.FindNextSel(ndx)
    wend

    if pn_tranzit > GLR.pqn_tranzit_min or (
    ubound(ndx_v)) > 1:  # если нагрузка в транзите меньше 1 МВт то отключаем транзит в любом месте, иначе в начале и конце

        Dict_groupid = CreateObject("Scripting.Dictionary")  # колекция groupid

        for Each ndx_i in ndx_v  # отмечаем откл
            if tV6.cols.item("groupid").Z(ndx_i) > 0:  # чтоб 2 раза одну группу не отключать
                if tV6.cols.item("tip").Z(ndx_i) = 0:  # если лэп
                    if not Dict_groupid.Exists(tV6.cols.item("groupid").Z(ndx_i)):  # Exists проверяет наличие ключа
                        Dict_groupid.Add(tV6.cols.item("groupid").Z(ndx_i), 0)  #
                        tV6.cols.item(GLR.vibor_otkl).Z(ndx_i) = 1

                elif tV6.cols.item("tip").Z(ndx_i) = 1:  # если транс
                    if tV6.cols.item("ktr").Z(ndx_i) = 1: tV6.cols.item(GLR.vibor_otkl).Z(ndx_i) = 1

            else:
                tV6.cols.item(GLR.vibor_otkl).Z(ndx_i) = 1

        Dict_groupid.RemoveAll()
    else:
        if ubound(ndx_v) = 0:
            tV6.cols.item(GLR.vibor_otkl).Z(ndx_v(0)) = 1
        elif ubound(ndx_v) = 1:
            if tV6.cols.item("tip").Z(ndx_v(0)) = 0: tV6.cols.item(GLR.vibor_otkl).Z(ndx_v(0)) = 1  # если лэп
            if tV6.cols.item("tip").Z(ndx_v(0)) = 1 and tV6.cols.item("ktr").Z(ndx_v (0)) = 1: tV6.cols.item(
                GLR.vibor_otkl).Z(ndx_v(0)) = 1  # если тр
            if tV6.cols.item("tip").Z(ndx_v(1)) = 1 and tV6.cols.item("ktr").Z(ndx_v (1)) = 1: tV6.cols.item(
                GLR.vibor_otkl).Z(ndx_v(1)) = 1  # если тр


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def Node_pojas_sel(n):  # отмкетить узлы  и ветви привыкающие к отмеченным узлам n раз
    # dim tV7  , tN6
    tN6 = rastr.Tables("node")
    tV7 = rastr.Tables("vetv")

    for i = 1 to n
    tV7.setsel("sel+!sta+(!ip.sel|!iq.sel)")
    tN6.setsel("")
    ndx_tVetv = tV7.FindNextSel(-1)
    while ndx_tVetv >= 0  #
        # if tN6.cols.item("uhom").Z(fNDX("node",tV7.cols.item("ip").Z(ndx_tVetv))) > 90:
        tN6.cols.item("sel").Z(fNDX("node", tV7.cols.item("ip").Z(ndx_tVetv))) = 1
        # if tN6.cols.item("uhom").Z(fNDX("node",tV7.cols.item("iq").Z(ndx_tVetv))) > 90:
        tN6.cols.item("sel").Z(fNDX("node", tV7.cols.item("iq").Z(ndx_tVetv))) = 1

        ndx_tVetv = tV7.FindNextSel(ndx_tVetv)
    wend

    tV7.setsel("ip.sel|iq.sel")  # выбираем все ветви связанные с  выбранными узлами
    tV7.cols.item("sel").Calc("1")

# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fSum_pn_tranzit(NNtranzit):  # возвращает сумму нагрузок в узлах  транзита
    # dim tVetv1 , tN8
    tN8 = rastr.Tables("node")
    tVetv1 = rastr.Tables("vetv")
    tN8.setsel("tranzit=" + str(NNtranzit))  # цикл по узлам выбранного транзита
    ndx_tNode = tN8.FindNextSel(-1)
    while ndx_tNode >= 0  #
        fSum_pn_tranzit = fSum_pn_tranzit + tN8.cols.item("pn").Z(ndx_tNode) + tN8.cols.item("qn").Z(
            ndx_tNode)  # сумируем нагрузку в самом узле

        #  ищем тупиковые ответвления от транзита
        # tVetv1.setsel ("(ip=" + str (tN8.cols.item("ny").ZS(ndx_tNode)) + "|iq=" + str (tN8.cols.item("ny").ZS(ndx_tNode))+ ")+tupik")
        tVetv1.setsel("(ip=" + str(tN8.cols.item("ny").ZS(ndx_tNode)) + "|iq=" + str(
            tN8.cols.item("ny").ZS(ndx_tNode)) + ")+tranzit!=" + str(NNtranzit))
        if tVetv1.count > 0:
            ndx_tVetv = tVetv1.FindNextSel(-1)
            while ndx_tVetv >= 0  #
                if tVetv1.cols.item("ip").Z(
                    ndx_tVetv) = tN8.cols.item("ny").Z(ndx_tNode): fSum_pn_tranzit = fSum_pn_tranzit - tVetv1.cols.item(
                    "pl_ip").Z(ndx_tVetv) - tVetv1.cols.item("ql_ip").Z(ndx_tVetv)
                if tVetv1.cols.item("iq").Z(
                    ndx_tVetv) = tN8.cols.item("ny").Z(ndx_tNode): fSum_pn_tranzit = fSum_pn_tranzit + tVetv1.cols.item(
                    "pl_iq").Z(ndx_tVetv) + tVetv1.cols.item("ql_iq").Z(ndx_tVetv)

                ndx_tVetv = tVetv1.FindNextSel(ndx_tVetv)
            wend

        ndx_tNode = tN8.FindNextSel(ndx_tNode)
    wend


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def groupid_dtn_kontrol(
        NNgroupid):  # определяет количество уникальных значений   dtn   и отмечает КОНТРОЛ ветви с уникальным groupid i_dop и dname
    tVetv_s = rastr.Tables("vetv")
    Dic_groupid_dtn = CreateObject("Scripting.Dictionary")  # колекция

    tVetv_s.setsel("groupid=" + str(NNgroupid) + "+!tupik+!sta+ip.uhom>90+iq.uhom>90+tip<2")  # не тупик
    ndx_tVetv_s = tVetv_s.FindNextSel(-1)

    while ndx_tVetv_s >= 0  #
        if tVetv_s.cols.item("tip").Z(ndx_tVetv_s) = 0:  # ЛЭП
            Dict_kluch = f_kluch_dtn(ndx_tVetv_s)
            if not Dic_groupid_dtn.Exists(Dict_kluch):  # Exists проверяет наличие ключа, если нет добавляем его
                Dic_groupid_dtn.Add(Dict_kluch, 0)  #
                tVetv_s.cols.item("Kontrol").Z(ndx_tVetv_s) = 1

        elif tVetv_s.cols.item("tip").Z(ndx_tVetv_s) = 1:  # ТРАНСФОРМАОР
            if tVetv_s.cols.item("ktr").Z(ndx_tVetv_s) = 1: tVetv_s.cols.item("Kontrol").Z(ndx_tVetv_s) = 1

        ndx_tVetv_s = tVetv_s.FindNextSel(ndx_tVetv_s)
    wend
    Dic_groupid_dtn.RemoveAll


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def f_kluch_dtn(ndx_dtn):  # формирует ключ дтн
    tVetv_f = rastr.Tables("vetv")
    f_kluch_dtn = "i_dop_ob=" + tVetv_f.cols.item("i_dop_ob").ZS(ndx_dtn) + _
    ",i_dop=" + tVetv_f.cols.item("i_dop").ZS(ndx_dtn) + _
    ",n_it=" + tVetv_f.cols.item("n_it").ZS(ndx_dtn) + _
    ",n_it_av=" + tVetv_f.cols.item("n_it_av").ZS(ndx_dtn) + _
    ",i_dop_ob_av=" + tVetv_f.cols.item("i_dop_ob_av").ZS(ndx_dtn) + _
    ",i_dop_av=" + tVetv_f.cols.item("i_dop_av").ZS(ndx_dtn) + _
    ",dname=" + tVetv_f.cols.item("dname").ZS(ndx_dtn)
# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def Dict_unik_value_sub(tabl, param, vibor,
                        Dict_unik_value):  # (нр "vetv","tranzit","sel",dict) наполняет Dictionary  уникальными значениями столбца таблицы
    Dict_unik_value.RemoveAll()
    tTabl = rastr.Tables(tabl)
    tParam = tTabl.cols.item(param)

    tTabl.setsel(vibor)
    ndx_tTabl = tTabl.FindNextSel(-1)
    while ndx_tTabl >= 0  #
        Dict_kluch = tParam.Z(ndx_tTabl)
        # if not Dict_unik_value.Exists ( Dict_kluch ) and Dict_kluch > 0:    #  Exists проверяет наличие ключа
        if not Dict_unik_value.Exists(Dict_kluch):  # Exists проверяет наличие ключа
            Dict_unik_value.Add(Dict_kluch, tabl + " / " + str(param))

        ndx_tTabl = tTabl.FindNextSel(ndx_tTabl)
    wend


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def print_dic(Dic):  # печатать колекцию
    for Each varKey in Dic.Keys  # цикл по ключам
        logging.info("\t" + "print_dic     ключ:" + str(varKey) + ", значение:" + str(Dic.Item(varKey)))


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def tip_vetv_auto():  # задать тип ветви    - тупик или транзит
    # dim tV8 , tVetv2 , tN9
    tN9 = rastr.Tables("node")
    tV8 = rastr.Tables("vetv")
    # задать tranzit и tupik по анализу  узлов с соотв. отместкой
    tV8.setsel("(ip.tranzit>0)|(iq.tranzit>0)+!sta")
    ndx_vetv = tV8.FindNextSel(-1)
    NNtranzit = 0
    while ndx_vetv >= 0  # цикл по ветвям
        tN9.setsel("ny=" + str(tV8.cols.item("ip").Z(ndx_vetv)))
        ndx_node = tN9.FindNextSel(-1)
        if tN9.cols.item("tranzit").Z(ndx_node) > 0:
            NNtranzit = tN9.cols.item("tranzit").Z(ndx_node)
        else:
            tN9.setsel("ny=" + str(tV8.cols.item("iq").Z(ndx_vetv)))
            ndx_node = tN9.FindNextSel(-1)
            if tN9.cols.item("tranzit").Z(ndx_node) > 0: NNtranzit = tN9.cols.item("tranzit").Z(ndx_node)

        tV8.cols.item("tranzit").Z(ndx_vetv) = NNtranzit
        ndx_vetv = tV8.FindNextSel(ndx_vetv)
    wend
    tV8.setsel("(ip.tupik|iq.tupik)")
    tV8.cols.item("tranzit").calc("0")
    tV8.cols.item("tupik").calc("1")
    # задать tranzit прочих ветвей и присвоить номер транзита
    tV8.setsel("sel+ip.uhom>90+iq.uhom>90+!sta+!tranzit+!tupik+tip<2")
    # tV8.setsel ("sel+ip.uhom>90+iq.uhom>90+!sta+!tranzit+!tupik")
    ndx_vetv = tV8.FindNextSel(-1)
    NNtranzit = rastr.Calc("max", "vetv", "tranzit", "ip>0") + 1
    while ndx_vetv >= 0  # цикл по ветвям
        if tV8.cols.item("groupid").Z(ndx_vetv) > 0:
            tVetv2 = rastr.Tables("vetv")
            tVetv2.setsel("groupid=" + tV8.cols.item("groupid").ZS(ndx_vetv))
            tVetv2.cols.item("tranzit").calc(NNtranzit)
        else:
            tV8.cols.item("tranzit").Z(ndx_vetv) = NNtranzit

        NNtranzit = NNtranzit + 1
        ndx_vetv = tV8.FindNextSel(ndx_vetv)
    wend


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def tupik_auto(vibor):  # процедура для определения тупиковых узлов среди отмеченых и отметка поля tupik узла
    # dim ny , ndx_node , ndx_vetv
    # dim tV9  , tN10
    tN10 = rastr.Tables("node")
    tV9 = rastr.Tables("vetv")
    tN10.setsel(vibor + "+!tupik")  # выбрать отмеченные узлы не отмеченные тупик
    ndx_node = tN10.FindNextSel(-1)

    while ndx_node >= 0
        if tN10.cols.item("tupik").Z(ndx_node) = 0:
            ny = tN10.cols.item("ny").Z(ndx_node)
            tV9.setsel(
                "(ip=" + str(ny) + "|iq=" + str(ny) + ")+(ip.uhom>90)+(iq.uhom>90)+!sta")  # выбор примыкающих ветвей

            if tV9.count < 2:  # если истина это самый конец тупика
                tN10.cols.item("tupik").Z(ndx_node) = 1
                ndx_vetv = tV9.FindNextSel(-1)
                if tV9.count = 1:
                    ny_next = 0
                    if ny = tV9.cols.item("ip").Z(ndx_vetv): ny_next = tV9.cols.item("iq").Z(ndx_vetv)
                    if ny = tV9.cols.item("iq").Z(ndx_vetv): ny_next = tV9.cols.item("ip").Z(ndx_vetv)
                    # cikl = 1
                    while ny_next > 0  # цикл по ветвям рассматриваемого узла
                        ny_next = fTupik_analiz(ny_next)
                        # if ny_next = 0: cikl = 0
                    wend
        ndx_node = tN10.FindNextSel(ndx_node)
    wend


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fTupik_analiz(nyT):  # определяет тупиковый узел или нет
    # визвращает ложь или номер узла если все примыкающие к узлу узлы кроме этого одного отмечены тупик
    tNodeT = rastr.Tables("node")
    tVetvT = rastr.Tables("vetv")
    ndx_node1 = fNDX("node", nyT)
    fTupik_analiz = 0
    Kol_Ne_Tupik = 0
    # tVetvT.setsel ("(ip=" +  str (nyT)  + "|iq=" +  str (nyT) + ")+((ip.uhom>90)+(iq.uhom>90))") #  выбор примыкающих ветвей
    # tVetvT.setsel ("(ip=" +  str (nyT)  + "|iq=" +  str (nyT) + ")+((ip.uhom>90)+(iq.uhom>90))+!sta") #  выбор примыкающих ветвей
    tVetvT.setsel("(ip=" + str(nyT) + "|iq=" + str(nyT) + ")+!sta")  # выбор примыкающих ветвей
    ndx_vetvT = tVetvT.FindNextSel(-1)

    while ndx_vetvT >= 0  # цикл по ветвям рассматриваемого узла

        if nyT = tVetvT.cols.item("ip").Z(ndx_vetvT): nyT2 = tVetvT.cols.item("iq").Z(ndx_vetvT)
        if nyT = tVetvT.cols.item("iq").Z(ndx_vetvT): nyT2 = tVetvT.cols.item("ip").Z(
            ndx_vetvT)  # nyT2 примыкающий узел
        ndx_node2 = fNDX("node", nyT2)
        if tNodeT.cols.item("tupik").Z(
            ndx_node2) = 0: Kol_Ne_Tupik = Kol_Ne_Tupik + 1: nyT3 = nyT2  # примыкающий узел для возврата функции
        ndx_vetvT = tVetvT.FindNextSel(ndx_vetvT)
    wend
    if Kol_Ne_Tupik = 1:
        tNodeT.cols.item("tupik").Z(ndx_node1) = 1
        fTupik_analiz = nyT3


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def tranzit_auto(
        vibor):  # процедура для опредеоения транзитных цепочек, запускать после определения тупиков (среди выборки)
    #  также определяет "узловые узлы"
    # dim ny , ndx_node , ndx_vetv , NNtranzit
    # dim tV10  , tN11
    tN11 = rastr.Tables("node")
    tV10 = rastr.Tables("vetv")
    NNtranzit = rastr.Calc("max", "node", "tranzit", "ny>0") + 1
    tN11.setsel(vibor + "+!tupik")  # выбрать отмеченные узлы не отмеченные тупик
    ndx_node = tN11.FindNextSel(-1)
    while ndx_node >= 0
        if tN11.cols.item("tranzit").Z(ndx_node) < 1:
            ny = tN11.cols.item("ny").Z(ndx_node)
            tV10.setsel("(ip=" + str(ny) + "|iq=" + str(
                ny) + ")+((ip.uhom>90)+(iq.uhom>90))+((ip.tupik=0)+(iq.tupik=0))+!sta")  # выбор примыкающих ветвей
            if tV10.count > 2:  # узловая похоже
                tN11.cols.item("uzlovaja").Z(ndx_node) = 1

            elif tV10.count = 2:  # транзитный узел
                tN11.cols.item("tranzit").Z(ndx_node) = NNtranzit  # присваиваем уникальный номер транзита

                ndx_vetv = tV10.FindNextSel(-1)
                while ndx_vetv >= 0
                    ny_next = 0
                    if ny = tV10.cols.item("ip").Z(ndx_vetv): ny_next = tV10.cols.item("iq").Z(ndx_vetv)
                    if ny = tV10.cols.item("iq").Z(ndx_vetv): ny_next = tV10.cols.item("ip").Z(ndx_vetv)

                    while ny_next > 0  # цикл по ветвям рассматриваемого узла
                        ny_next = fTranzit_analiz(ny_next, NNtranzit)
                    wend
                    ndx_vetv = tV10.FindNextSel(ndx_vetv)
                wend

                NNtranzit = NNtranzit + 1
                ndx_node = tN11.FindNextSel(ndx_node)
    wend


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fTranzit_analiz(nyT, NN):  # определяет транзитный узел или нет (проверяемый узел, порядковый номер транзита)
    # визвращает 0 или номер узла если все кроме 1 примыкающие к узлу узлы отмечены тупик, транзит или узловая
    tNodeT = rastr.Tables("node")
    tVetvT = rastr.Tables("vetv")
    ndx_node1 = fNDX("node", nyT)
    fTranzit_analiz = 0
    Kol_Tranzit_new = 0  # количество примыкающих не тупиковых узлов и не отмеченных транзитом
    Kol_ne_tupik = 0  # количество примыкающих не тупиковых узлов

    tVetvT.setsel(
        "(ip=" + str(nyT) + "|iq=" + str(nyT) + ")+((ip.uhom>90)+(iq.uhom>90))+!sta")  # выбор примыкающих ветвей
    ndx_vetvT = tVetvT.FindNextSel(-1)

    while ndx_vetvT >= 0  # цикл по ветвям рассматриваемого узла

        if nyT = tVetvT.cols.item("ip").Z(ndx_vetvT): nyT2 = tVetvT.cols.item("iq").Z(ndx_vetvT)
        if nyT = tVetvT.cols.item("iq").Z(ndx_vetvT): nyT2 = tVetvT.cols.item("ip").Z(
            ndx_vetvT)  # nyT2 примыкающий узел
        ndx_node2 = fNDX("node", nyT2)
        if tNodeT.cols.item("tupik").Z(
            ndx_node2) = 0 and tNodeT.cols.item("tranzit").Z(ndx_node2) < 1: Kol_Tranzit_new = Kol_Tranzit_new + 1: nyT3 = nyT2  # примыкающий узел для возврата функции
        if tNodeT.cols.item("tupik").Z(ndx_node2) = 0: Kol_ne_tupik = Kol_ne_tupik + 1  #

        ndx_vetvT = tVetvT.FindNextSel(ndx_vetvT)
    wend
    if Kol_Tranzit_new = 1 and Kol_ne_tupik = 2:  # транзит продолжается
        tNodeT.cols.item("tranzit").Z(ndx_node1) = NN
        fTranzit_analiz = nyT3
    elif Kol_ne_tupik > 2:  # узловая встретилась
        tNodeT.cols.item("uzlovaja").Z(ndx_node1) = 1
    elif Kol_ne_tupik = 1 and Kol_Tranzit_new = 0:  # встретился тупик
        tNodeT.cols.item("tupik").Z(ndx_node1) = 1
        ny_next = nyT
        while ny_next > 0  # цикл по ветвям рассматриваемого узла
            ny_next = fTupik_analiz(ny_next)
            # if ny_next = 0: cikl = 0
        wend
    else:
        logging.info("\t" + "тип узла не распознан ny:" + str(nyT))

    # End def return


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fBukva_colunm(col):
    fBukva_colunm = GS.excel.ConvertFormula("r1c" + str(col), -4150, 1)
    fBukva_colunm = Replace(Replace(Mid(fBukva_colunm, 2), "$", ""), "1", "")


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def Print_XL(xl, x0, y0, masiv, nn, raspolozhenie, tab, par,
             plus_txt):  # печать массива: лист ХL ,по X , по Y , массив , кол изм массива 1 или 2 , "гор" "верт" , "" или "vetv" ,"" или "name" , "" или "орвыаи " - произвольный текст
    if nn = 2:  # если матрица
        if raspolozhenie = "гор":
            for po_x = 0 to ubound (masiv, 1)
            for po_y = 0 to ubound (masiv, 2)
            if tab = "":
                xl.cell(y0 + po_y, x0 + po_x).Value = masiv(po_x, po_y)
                if plus_txt != "": xl.cell(y0 + po_y, x0 + po_x).Value = str(masiv(po_x, po_y).Value) + plus_txt
            else:
                xl.cell(y0 + po_y, x0 + po_x).Value = rastr.Tables(tab).cols.item(par).Z(masiv(po_x, po_y))
                if plus_txt != "":
                    xl.cell(y0 + po_y, x0 + po_x).Value = str(rastr.Tables(tab).cols.item(par).Z(
                        masiv(po_x, po_y))) + plus_txt        elif raspolozhenie = "верт":
    for po_x = 0 to ubound (masiv, 2)
    for po_y = 0 to ubound (masiv, 1)
    if tab = "":
        xl.cell(y0 + po_y, x0 + po_x).Value = masiv(po_y, po_x)
        if plus_txt != "": xl.cell(y0 + po_y, x0 + po_x).Value = str(masiv(po_y, po_x)) + plus_txt
    else:
        xl.cell(y0 + po_y, x0 + po_x).Value = rastr.Tables(tab).cols.item(par).Z(masiv(po_y, po_x))
        if plus_txt != "":
            xl.cell(y0 + po_y, x0 + po_x).Value = str(
                rastr.Tables(tab).cols.item(par).Z(masiv(po_y, po_x))) + plus_txt    elif nn = 1:  # если моссив


if raspolozhenie = "гор":
    for po_x = 0 to ubound (masiv )
    if tab = "":
        xl.cell(y0, x0 + po_x).Value = masiv(po_x)
        if plus_txt != "": xl.cell(y0 + po_y, x0 + po_x).Value = str(masiv(po_x)) + plus_txt
    else:
        xl.cell(y0, x0 + po_x).Value = rastr.Tables(tab).cols.item(par).Z(masiv(po_x))
        if plus_txt != "": xl.cell(y0 + po_y, x0 + po_x).Value = str(
            rastr.Tables(tab).cols.item(par).Z(masiv(po_x))) + plus_txt

elif raspolozhenie = "верт":
for po_x = 0 to ubound (masiv )
if tab = "":
    xl.cell(y0 + po_x, x0).Value = masiv(po_x)
    if plus_txt != "": xl.cell(y0 + po_y, x0 + po_x).Value = str(masiv(po_x)) + plus_txt
else:
    xl.cell(y0 + po_x, x0).Value = rastr.Tables(tab).cols.item(par).Z(masiv(po_x))
    if plus_txt != "": xl.cell(y0 + po_x, x0).Value = str(rastr.Tables(tab).cols.item(par).Z(masiv(po_x))) + plus_txt


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def Print_XL_otklNAME(xl, x0, y0, masiv, raspolozhenie,
                      par):  # печать массива по даным из массива: лист ХL ,по X , по Y , массив , кол изм массива 1 или 2 , "гор" "верт" ,  "name" параметр
    if raspolozhenie = "гор":
        for po_x = 0 to ubound (masiv, 2)
        if not isempty(masiv(0, po_x)): xl.cell(y0, x0 + po_x).Value = rastr.Tables(masiv(1, po_x)).cols.item(par).Z(
            masiv(0, po_x))

elif raspolozhenie = "верт":
for po_y = 0 to ubound (masiv, 2)
if not isempty(masiv(1, po_y)): xl.cell(y0 + po_y, x0).Value = rastr.Tables(masiv(1, po_y)).cols.item(par).Z(
    masiv(0, po_y))

if plus_txt != "": xl.cell(y0 + po_y, x0 + po_x).Value = str(xl.cell(y0 + po_y, x0 + po_x).Value) + plus_txt


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def export_RG2():  # экспорт парам в CSV
    # dim tND , tVT
    rastr.Load(1, GLR.Zad_RG2_name, rastr.SendCommandMain(3, "", "", 0) + "shablon\режим.rg2")
    tND = rastr.Tables("node")
    tVT = rastr.Tables("vetv")
    # NODE
    arr_paramN_rg2_flag = array(True, GLR.node_auto_flag, GLR.node_zad_flag)
    arr_paramN_rg2_z = array("ny", GLR.node_auto, GLR.node_zad)
    GLR.paramN = fParam_str(arr_paramN_rg2_flag, arr_paramN_rg2_z)

    tND.setsel(GLR.Zad_RG2_VIBOR_N)
    tND.WriteCSV(1, GLR.Folder_csv_RG2 + "\uzli_Zad_RG2.csv", GLR.paramN, ";")  # 0 дополнить, 1 заменить
    # VETV
    arr_paramV_rg2_flag = array(True, GLR.vetv_auto_flag, GLR.vetv_zad_flag)
    arr_paramV_rg2_z = array("ip,iq,np", GLR.vetv_auto, GLR.vetv_zad)
    GLR.paramV = fParam_str(arr_paramV_rg2_flag, arr_paramV_rg2_z)

    tVT.setsel(GLR.Zad_RG2_VIBOR_V)
    tVT.WriteCSV(1, GLR.Folder_csv_RG2 + "\vetki_Zad_RG2.csv", GLR.paramV, ";")  #


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def import_RG2():  # обнуление и импорт парам из CSV
    tVT = rastr.Tables("vetv")  #
    tND = rastr.Tables("node")
    tND.ReadCSV(2, GLR.Folder_csv_RG2 + "\uzli_Zad_RG2.csv", GLR.paramN, ";", "")  #
    tVT.ReadCSV(2, GLR.Folder_csv_RG2 + "\vetki_Zad_RG2.csv", GLR.paramV, ";", "")  #


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fSCAN_FOLDER(Folder, scan_str, type_file):  # функция возвращает файл для импорта - имя начинается с "!"
    # dim objFile_z
    fSCAN_FOLDER = "не найден"
    if objFSO.FolderExists(Folder):
        for Each objFile_z in objFSO.GetFolder(Folder).Files  # цикл по файлам в  указанной папке
            if LEFT(objFile_z.name, Len(scan_str)) = scan_str:
                if objFile_z.type = type_file or type_file = "-":
                    fSCAN_FOLDER = objFile_z.Path
                    GLR.Zad_RG2_name_k = objFile_z.name
                    exit
                    for


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def Del_v_Folder(Folder):  # удалить все из папки
    # dim sDirectoryPath , oFolder , oDelFolder , oFileCollection
    # dim oFile , oFolderCollection

    sDirectoryPath = Folder
    oFolder = objFSO.GetFolder(sDirectoryPath)
    oFolderCollection = oFolder.SubFolders
    oFileCollection = oFolder.Files

    for each oFile in oFileCollection
        oFile.Delete(True)

    for each oDelFolder in oFolderCollection
        oDelFolder.Delete(True)

    oFolder = Nothing
    oFileCollection = Nothing
    oFile = Nothing


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def TopologyStore():
    if rastr.tables("node").cols.Find("staRes") < 1: rastr.tables("node").Cols.Add
    "staRes", 3  # добавить столбцы
    if rastr.tables("node").cols.Find("bshRes") < 1: rastr.tables("node").Cols.Add
    "bshRes", 1  # добавить столбцы
    if rastr.tables("node").cols.Find("autoStaRes") < 1: rastr.tables("node").Cols.Add
    "autoStaRes", 3  # добавить столбцы

    if rastr.tables("vetv").cols.Find("staRes") < 1: rastr.tables("vetv").Cols.Add
    "staRes", 0  # добавить столбцы
    if rastr.tables("vetv").cols.Find("autoStaRes") < 1: rastr.tables("vetv").Cols.Add
    "autoStaRes", 3  # добавить столбцы

    if rastr.tables("Generator").cols.Find("staRes") < 1: rastr.tables("Generator").Cols.Add
    "staRes", 3  # добавить столбцы

    rastr.Tables("node").cols.item("staRes").Calc("sta")
    rastr.Tables("node").cols.item("bshRes").Calc("bsh*1000000")
    rastr.Tables("node").cols.item("autoStaRes").Calc("autosta")
    rastr.Tables("vetv").cols.item("staRes").Calc("sta")
    rastr.Tables("vetv").cols.item("autoStaRes").Calc("autosta")
    rastr.Tables("Generator").cols.item("staRes").Calc("sta")
    # PR_INT 0 Целый;  PR_REAL 1 Вещественный;  PR_STRING 2 Строка;  PR_BOOL 3 Переключатель;  PR_ENUM 4 Перечисление;  PR_ENPIC 5 Перечисление картинок (не используется);  PR_COLOR 6 Цвет;


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def TopologyRestore():
    #  if rastr.tables("node").cols.Find("staRes") < 1: logging.info( " пропало поле staRes")
    #  if rastr.tables("node").cols.Find("bshRes") < 1: logging.info( " пропало поле bshRes")
    #  if rastr.tables("node").cols.Find("autoStaRes") < 1: logging.info( " пропало поле autoStaRes")
    #  if rastr.tables("vetv").cols.Find("staRes") < 1: logging.info( " пропало поле staRes")
    #  if rastr.tables("vetv").cols.Find("autoStaRes") < 1: logging.info( " пропало поле autoStaRes")
    #  if rastr.tables("Generator").cols.Find("staRes") < 1: logging.info( " пропало поле staRes")
    rastr.Tables("node").cols.item("sta").Calc("staRes")
    rastr.Tables("node").cols.item("bsh").Calc("bshRes/1000000")
    rastr.Tables("node").cols.item("autosta").Calc("autoStaRes")

    rastr.Tables("vetv").cols.item("sta").Calc("staRes")
    rastr.Tables("vetv").cols.item("autosta").Calc("autoStaRes")

    rastr.Tables("Generator").cols.item("sta").Calc("staRes")


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def PnQnPgStore():
    if rastr.tables("node").cols.Find("pnRes") < 1: rastr.tables("node").Cols.Add
    "pnRes", 1  # добавить столбцы
    if rastr.tables("node").cols.Find("qnRes") < 1: rastr.tables("node").Cols.Add
    "qnRes", 1  # добавить столбцы
    if rastr.tables("node").cols.Find("pgRes") < 1: rastr.tables("node").Cols.Add
    "pgRes", 1  # добавить столбцы
    if rastr.tables("Generator").cols.Find("PRes") < 1: rastr.tables("Generator").Cols.Add
    "PRes", 1  # добавить столбцы
    if rastr.tables("vetv").cols.Find("ktrRes") < 1: rastr.tables("vetv").Cols.Add
    "ktrRes", 1  # добавить столбцы

    rastr.Tables("node").cols.item("pnRes").Calc("pn")
    rastr.Tables("node").cols.item("qnRes").Calc("qn")
    rastr.Tables("node").cols.item("pgRes").Calc("pg")
    rastr.Tables("Generator").cols.item("PRes").Calc("P")
    rastr.Tables("vetv").cols.item("ktrRes").Calc("ktr")


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def PnQnPgKtrRestore():
    #  if rastr.tables("node").cols.Find("pnRes") < 1: logging.info( " пропало поле node")
    #  if rastr.tables("node").cols.Find("qnRes") < 1: logging.info( " пропало поле node")
    #  if rastr.tables("node").cols.Find("pgRes") < 1: logging.info( " пропало поле node")
    #  if rastr.tables("Generator").cols.Find("PRes") < 1: logging.info( " пропало поле Generator")
    #  if rastr.tables("vetv").cols.Find("ktrRes") < 1: logging.info( " пропало поле vetv")
    rastr.Tables("node").cols.item("pn").Calc("pnRes")
    rastr.Tables("node").cols.item("qn").Calc("qnRes")
    rastr.Tables("node").cols.item("pg").Calc("pgRes")

    rastr.Tables("Generator").cols.item("P").Calc("PRes")

    rastr.Tables("vetv").cols.item("ktr").Calc("ktrRes")


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def TablF_Sub():  # ШАПКА ТАБОИЦЫ
    # dim  ndx ,ob ,idop25, osh , n, N_VetvGrp, vibor_N_VetvGrp, ndx_G_id, MIN_G_id, MAX_G_id, vibor,   ndxOO
    # dim a1, a2 , a3, a4 , a5 , a6, a7, a8, a9,a10 ,formula_1 , formula_2 ,  formula_3 , formula_4 , formula_5
    # dim a1OO ,  a2OO , a3OO , formula_ZagOO, a11, a21 , a12 , a22 , formulaMAX_I , formulaMAX_S , formula_Zag
    # dim tV13  , tN14
    tN14 = rastr.Tables("node")
    tV13 = rastr.Tables("vetv")
    rastr.CalcIdop(RG.GradusZ, float(0), "")
    logging.info("\t" + "расчетная температура (TablF_Sub): " + str(RG.GradusZ))
    GLR.XL_sheet.cell(3, 11).Value = "Контролируемые элементы"
    GLR.XL_sheet.Range(GLR.XL_sheet.cell(3, 11), GLR.XL_sheet.cell(5, 11)).Merge
    GLR.XL_sheet.cell(3, 13).Value = "Допустимый длительный ток  "  # + "в " + RG.name_list (1) + "ний период, А"
    GLR.XL_sheet.Range(GLR.XL_sheet.cell(3, 13), GLR.XL_sheet.cell(5, 13)).Merge
    GLR.XL_sheet.cell(6, 15).Value = "Мощность, МВт +JМвар/Ток ветви, А/Токовая загрузка, %"
    GLR.XL_sheet.cell(GLR.Y_list - 1, 11).Value = "ЛЭП"
    n = 0  # цикл по контр. ветвям VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_VL_L_VL_VL_VL_VL_VL_VL_VL_VL_L_VL_VL_VL_VL_VL_VL_VL_VL_L_VL_VL_VL_VL_VL_VL_VL_VL_
    for i = LBound(RG.KontrolVL) to UBound(RG.KontrolVL)  # запись всех  ветвей  ВЛ отмеченных КОНТРОЛЬ
    ndx = RG.KontrolVL(i)
    GLR.XL_sheet.cell(GLR.Y_list + i + n, GLR.X_list - 4).Value = tV13.cols.item("dname").ZS(ndx)  # имя контр вет
    if tV13.cols.item("i_dop_r").Z(ndx) = 0: GLR.XL_sheet.cell(GLR.Y_list + i + n, GLR.X_list - 2).Value = "" else:
        GLR.XL_sheet.cell(GLR.Y_list + i + n, GLR.X_list - 2).Value = tV13.cols.item("i_dop_r").Z(ndx)
    GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list + i + n, GLR.X_list - 2), GLR.XL_sheet.cell(GLR.Y_list + i + n,
                                                                                                  GLR.X_list - 1)).NumberFormat = "0"  # формат формат доп тока
    GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list + i + n, GLR.X_list - 4),
                       GLR.XL_sheet.cell(GLR.Y_list + i + n + 1,
                                          GLR.X_list - 4)).Merge  # ячейки объеденить название ВЛ
    GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list + i + n, 1),
                       GLR.XL_sheet.cell(GLR.Y_list + i + n + 1, GLR.X_list - 1)).Borders(
        9).LineStyle = 1  # низ # рамка ИМЯ ВЛ и доп парам
    GLR.XL_sheet.cell(GLR.Y_list + i + n, 1).Font.Underline = True  # подчеркнуть текст

    if tV13.cols.item("groupid").Z(
            ndx):  # выбираем максимальное и минимальной значение в группе и заносим в таблицу с учетом разных доп токов на одной вл
        N_VetvGrp = tV13.cols.item("groupid").Z(ndx)  # присваиваем номер группы
        if tV13.cols.item("i_dop_ob").Z(ndx) > 0:
            i_dop_ob = tV13.cols.item("i_dop_ob").ZS(ndx) else:
            i_dop_ob = "0"
        if tV13.cols.item("i_dop").Z(ndx) > 0:
            i_dop = tV13.cols.item("i_dop").ZS(ndx) else:
            i_dop = "0"

        if GLR.dtn_uchastki = 1:
            vibor_N_VetvGrp = "groupid=" + str(N_VetvGrp) + "+i_dop=" + str(i_dop) + "+i_dop_ob=" + str(
                i_dop_ob) + "+n_it=" + tV13.cols.item("n_it").ZS(ndx)
            dname_vetv = trim(tV13.cols.item("dname").ZS(ndx))
        elif GLR.dtn_uchastki = 0:
            vibor_N_VetvGrp = "groupid=" + str(N_VetvGrp)

        tV13.setsel(vibor_N_VetvGrp)
        ndx_G_id = tV13.FindNextSel(-1)
        MIN_G_id = ndx_G_id
        MAX_G_id = ndx_G_id

        if GLR.dtn_uchastki = 1:
            while ndx_G_id >= 0
                if tV13.cols.item("i_max").Z(ndx_G_id) < tV13.cols.item("i_max").Z(
                    MIN_G_id) and dname_vetv = trim ( tV13.cols.item("dname").ZS(ndx_G_id) ): MIN_G_id = ndx_G_id
                if tV13.cols.item("i_max").Z(ndx_G_id) > tV13.cols.item("i_max").Z(
                    MAX_G_id) and dname_vetv = trim ( tV13.cols.item("dname").ZS(ndx_G_id) ): MAX_G_id = ndx_G_id
                ndx_G_id = tV13.FindNextSel(ndx_G_id)
            wend
        elif GLR.dtn_uchastki = 0:
            while ndx_G_id >= 0
                if tV13.cols.item("i_max").Z(ndx_G_id) < tV13.cols.item("i_max").Z(MIN_G_id): MIN_G_id = ndx_G_id
                if tV13.cols.item("i_max").Z(ndx_G_id) > tV13.cols.item("i_max").Z(MAX_G_id): MAX_G_id = ndx_G_id
                ndx_G_id = tV13.FindNextSel(ndx_G_id)
            wend

        if GLR.DRVXL = 1:
            if tV13.cols.item("znak-").Z(ndx):
                GLR.XL_sheet.cell(GLR.Y_list + i + n,
                                   3).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",\"-Smax\",\"" + rastr.Tables(
                    "vetv").SelString(MIN_G_id) + "\")"
                GLR.XL_sheet.cell(GLR.Y_list + i + n + 1,
                                   3).Formula = "=-RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",\"Imax\",\"" + rastr.Tables(
                    "vetv").SelString(MIN_G_id) + "\")"
                GLR.XL_sheet.cell(GLR.Y_list + i + n,
                                   4).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",\"-Smax\",\"" + rastr.Tables(
                    "vetv").SelString(MAX_G_id) + "\")"
                GLR.XL_sheet.cell(GLR.Y_list + i + n + 1,
                                   4).Formula = "=-RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",\"Imax\",\"" + rastr.Tables(
                    "vetv").SelString(MAX_G_id) + "\")"
            else:
                GLR.XL_sheet.cell(GLR.Y_list + i + n,
                                   3).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",\"Smax\",\"" + rastr.Tables(
                    "vetv").SelString(MIN_G_id) + "\")"
                GLR.XL_sheet.cell(GLR.Y_list + i + n + 1,
                                   3).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",\"Imax\",\"" + rastr.Tables(
                    "vetv").SelString(MIN_G_id) + "\")"
                GLR.XL_sheet.cell(GLR.Y_list + i + n,
                                   4).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",\"Smax\",\"" + rastr.Tables(
                    "vetv").SelString(MAX_G_id) + "\")"
                GLR.XL_sheet.cell(GLR.Y_list + i + n + 1,
                                   4).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",\"Imax\",\"" + rastr.Tables(
                    "vetv").SelString(MAX_G_id) + "\")"

            a11 = GLR.XL_sheet.cell(GLR.Y_list + i + n, 3).Address  # S1
            a21 = GLR.XL_sheet.cell(GLR.Y_list + i + n + 1, 3).Address  # I1
            a12 = GLR.XL_sheet.cell(GLR.Y_list + i + n, 4).Address  # S2
            a22 = GLR.XL_sheet.cell(GLR.Y_list + i + n + 1, 4).Address  # I2

            formulaMAX_I = "=ЕСЛИ(ABS(" + a21 + ")>ABS(" + a22 + ");" + a21 + ";" + a22 + ")"
            formulaMAX_S = "=ЕСЛИ(ABS(" + a21 + ")>ABS(" + a22 + ");" + a11 + ";" + a12 + ")"

        With
        GLR.XL_sheet.cell(GLR.Y_list + i + n, 1)
        # .Value = tV13.cols.item("i_zag").Z(ndx)   # запись расчет загрузки значение
        .FormulaLocal = formulaMAX_S  # запись расчет загрузки формула
    End
    With

    With
    GLR.XL_sheet.cell(GLR.Y_list + i + n + 1, 1)
    # .Value = tV13.cols.item("i_zag").Z(ndx)   # запись расчет загрузки значение
    .FormulaLocal = formulaMAX_I  # запись расчет загрузки формула


End
With

else:
if GLR.DRVXL = 1:
    if tV13.cols.item("znak-").Z(ndx):
        GLR.XL_sheet.cell(GLR.Y_list + i + n,
                           1).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",\"-Smax\",\"" + rastr.Tables(
            "vetv").SelString(ndx) + "\")"
        GLR.XL_sheet.cell(GLR.Y_list + i + n + 1,
                           1).Formula = "=-RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",\"Imax\",\"" + rastr.Tables(
            "vetv").SelString(ndx) + "\")"
    else:
        GLR.XL_sheet.cell(GLR.Y_list + i + n,
                           1).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",\"Smax\",\"" + rastr.Tables(
            "vetv").SelString(ndx) + "\")"
        GLR.XL_sheet.cell(GLR.Y_list + i + n + 1,
                           1).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",\"Imax\",\"" + rastr.Tables(
            "vetv").SelString(ndx) + "\")"
a1 = GLR.XL_sheet.cell(GLR.Y_list + i + n + 1, 1).Address(False, False)  # ток
a2 = GLR.XL_sheet.cell(GLR.Y_list + i + n, 13).Address(False, True)  # доп ток
a3 = GLR.XL_sheet.cell(GLR.Y_list + i + n, 14).Address(False, True)  # доп ток
formula_Zag = "=ABS(" + a1 + "/МИН(" + a2 + ":" + a3 + ")" + "*100)"

With
GLR.XL_sheet.cell(GLR.Y_list + i + n, 2)
# .Value = tV13.cols.item("i_zag").Z(ndx)   # запись расчет загрузки значение
if GLR.XL_sheet.cell(GLR.Y_list + i + n, 13).value > 0: .
    FormulaLocal = formula_Zag  # запись расчет загрузки формула
.NumberFormat = "0"
End
With
n = n + 1

With
GLR.XL_sheet.Range(GLR.XL_sheet.cell(7, GLR.X_list - 4),
                   GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL - 1, GLR.X_list - 1))  # рамка ИМЯ ВЛ и доп парам
.Borders(7).LineStyle = 1  # лево
.Borders(8).LineStyle = 1  # верх
.Borders(9).LineStyle = 1  # низ
.Borders(10).LineStyle = 1  # право
.Borders(11).LineStyle = 1  # внутри вертикаль
End
With

With
GLR.XL_sheet.Range(GLR.XL_sheet.cell(7, 1), GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL - 1, 2))  # рамка ДРВ
.Borders(7).LineStyle = 1  # лево
.Borders(8).LineStyle = 1  # верх
.Borders(9).LineStyle = 1  # низ
.Borders(10).LineStyle = 1  # право
.Borders(11).LineStyle = 1  # внутри вертикаль
End
With

GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL, 11).Value = "АТ(Г)"
GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL, 13).Value = "Номинальный ток, А"
GLR.XL_sheet.Rows(GLR.Y_list + GLR.Y_VL).RowHeight = 30  # высота строки

n = 0  # цикл по контр. ветвям Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_
# GLR.Kol_Tr_OO = 0
for i = LBound(RG.KontrolTrans) to UBound(RG.KontrolTrans)  # запись всех  ветвей  Trans отмеченных КОНТРОЛЬ
ndx = RG.KontrolTrans(i)
GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, GLR.X_list - 4).Value = tV13.cols.item("dname").ZS(
    ndx)  # имя вет
GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, GLR.X_list - 2).Value = tV13.cols.item("i_dop_r").ZN(
    ndx)  # tV13.cols.item("i_zag").ZN(ndx)
#  if tV13.cols.item("i_dop_ob").ZN(ndx) > 0:
#      GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, GLR.X_list-2).Value = tV13.cols.item("i_dop_ob").ZN(ndx)  #  ток оборуд
#  else:
#      GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, GLR.X_list-2).Value = tV13.cols.item("i_dop").ZN(ndx) #  ток
#  # end if

if tV13.cols.item("i_dop_ob").Z(ndx) = 0 And tV13.cols.item("i_dop").Z(ndx) = 0: GLR.XL_sheet.cell(
    GLR.Y_list + GLR.Y_VL + i + 1 + n, GLR.X_list - 2).Value = ""

GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, GLR.X_list - 4),
                   GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n + 1, GLR.X_list - 4)).Merge

if GLR.DRVXL = 1: GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + n + 1,
                                     1).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",\"Smax\",\"" + rastr.Tables(
    "vetv").SelString(ndx) + "\")"
if GLR.DRVXL = 1: GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + n + 2,
                                     1).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",\"Imax\",\"" + rastr.Tables(
    "vetv").SelString(ndx) + "\")"
GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + n + 1, 1).Font.Underline = True  # подчеркнуть текст

if tV13.cols.item("KontrOO").Z(ndx):  # истина если контроль ОО АТ
    GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, GLR.X_list - 4).Value = tV13.cols.item("dname").ZS(
        ndx) + " (обмотка ВН)"
    GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n + 2,
                       GLR.X_list - 4).Value = "ток общей обмотки, А /токовая загрузка, %"
    GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n + 2, 13).Value = tV13.cols.item("IdopOO").Z(
        ndx)  # тут нада указать доп ток ОО

    vibor = "ip=" + tV13.cols.item("iq").ZS(ndx) + "+ ktr>0.2"  # выбираем обмотку СН АТ
    tV13.setsel(vibor)
    ndxOO = tV13.FindNextSel(-1)

    if ndxOO = -1:
        tV13.setsel("ip=" + tV13.cols.item("ip").ZS(ndx) + "+ ktr>0.2")
        ndxOO = tV13.FindNextSel(-1)
        ny_vn = tV13.cols.item("iq").ZN(ndx)
        ny_sn = tV13.cols.item("iq").ZN(ndxOO)
        if GLR.DRVXL = 1: GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n,
                                             3).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",""v_iq"",\"" + rastr.Tables(
            "vetv").SelString(ndx) + "\")"
        if GLR.DRVXL = 1: GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n,
                                             4).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",""pl_iq"",\"" + rastr.Tables(
            "vetv").SelString(ndx) + "\")"
        if GLR.DRVXL = 1: GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n,
                                             5).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",""ql_iq"",\"" + rastr.Tables(
            "vetv").SelString(ndx) + "\")"
        if GLR.DRVXL = 1: GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n,
                                             8).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",""node"",""delta"",\"" + "ny=" + str(
            ny_vn) + "\")"
    else:
        ny_vn = tV13.cols.item("ip").ZN(ndx)
        ny_sn = tV13.cols.item("iq").ZN(ndxOO)
        if GLR.DRVXL = 1: GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n,
                                             3).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",""v_ip"",\"" + rastr.Tables(
            "vetv").SelString(ndx) + "\")"
        if GLR.DRVXL = 1: GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n,
                                             4).Formula = "=-RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",""pl_ip"",\"" + rastr.Tables(
            "vetv").SelString(ndx) + "\")"
        if GLR.DRVXL = 1: GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n,
                                             5).Formula = "=-RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",""ql_ip"",\"" + rastr.Tables(
            "vetv").SelString(ndx) + "\")"
        if GLR.DRVXL = 1: GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n,
                                             8).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",""node"",""delta"",\"" + "ny=" + str(
            ny_vn) + "\")"

    if GLR.DRVXL = 1:
        GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 2 + n,
                           3).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",""v_iq"",\"" + rastr.Tables(
            "vetv").SelString(ndxOO) + "\")"
        GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 2 + n,
                           4).Formula = "=-RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",""pl_iq"",\"" + rastr.Tables(
            "vetv").SelString(ndxOO) + "\")"
        GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 2 + n,
                           5).Formula = "=-RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",\"vetv\",""ql_iq"",\"" + rastr.Tables(
            "vetv").SelString(ndxOO) + "\")"
        GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 2 + n,
                           8).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",""node"",""delta"",\"" + "ny=" + str(
            ny_sn) + "\")"

    U1_a = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, 3).Address  # U1
    delta1_a = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, 8).Address  # delta1
    P1_a = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, 4).Address  # P1
    Q1_a = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, 5).Address  # Q1
    U2_a = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 2 + n, 3).Address  # U2
    delta2_a = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 2 + n, 8).Address  # delta2
    P2_a = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 2 + n, 4).Address  # P2
    Q2_a = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 2 + n, 5).Address  # Q2

    ReI1_a = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, 6).Address  # ReI1
    ImI1_a = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, 7).Address  # ImI1
    ReI2_a = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 2 + n, 6).Address  # ReI2
    ImI2_a = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 2 + n, 7).Address  # ImI2

    ReU1 = U1_a + "*cos(" + delta1_a + ")"
    ImU1 = U1_a + "*sin(" + delta1_a + ")"
    ReU2 = U2_a + "*cos(" + delta2_a + ")"
    ImU2 = U2_a + "*sin(" + delta2_a + ")"

    formula_1 = "=(" + P1_a + "*" + ReU1 + "+" + Q1_a + "*" + ImU1 + ")/(КОРЕНЬ(3)*(" + ReU1 + "*" + ReU1 + "+" + ImU1 + "*" + ImU1 + "))*1000"
    formula_2 = "=(" + Q1_a + "*" + ReU1 + "-" + P1_a + "*" + ImU1 + ")/(КОРЕНЬ(3)*(" + ReU1 + "*" + ReU1 + "+" + ImU1 + "*" + ImU1 + "))*1000"
    formula_3 = "=(" + P2_a + "*" + ReU2 + "+" + Q2_a + "*" + ImU2 + ")/(КОРЕНЬ(3)*(" + ReU2 + "*" + ReU2 + "+" + ImU2 + "*" + ImU2 + "))*1000"
    formula_4 = "=(" + Q2_a + "*" + ReU2 + "-" + P2_a + "*" + ImU2 + ")/(КОРЕНЬ(3)*(" + ReU2 + "*" + ReU2 + "+" + ImU2 + "*" + ImU2 + "))*1000"

    GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, 6).FormulaLocal = formula_1  # ReI1
    GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 2 + n, 6).FormulaLocal = formula_3  # ImI1
    GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, 7).FormulaLocal = formula_2  # ReI2
    GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 2 + n, 7).FormulaLocal = formula_4  # ImI2
    formula_5 = "=корень((" + ReI1_a + "-" + ReI2_a + ")*(" + ReI1_a + "-" + ReI2_a + ")+(" + ImI1_a + "-" + ImI2_a + ")*(" + ImI1_a + "-" + ImI2_a + "))"
    GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 3 + n, 1).FormulaLocal = formula_5  #

    GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, 1),
                       GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n + 2, GLR.X_list - 1)).Borders(
        9).LineStyle = 1  # низ

    a1 = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n + 1, 1).Address(False, False)  # ток
    a2 = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, 13).Address(False, True)  # доп ток
    a3 = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, 14).Address(False, True)  # доп ток
    formula_Zag = "=ABS(" + a1 + "/МИН(" + a2 + ":" + a3 + ")" + "*100)"

    With
    GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, 2)
    # .Value = tV13.cols.item("i_zag").Z(ndx)   # запись расчет загрузки значение
    if GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, 13).value > 0: .
        FormulaLocal = formula_Zag  # запись расчет загрузки формула
    .NumberFormat = "0"
End
With

a1OO = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n + 2, 1).Address(False, False)  # ток
a2OO = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n + 2, 13).Address(False, True)  # доп ток
a3OO = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n + 2, 14).Address(False, True)  # доп ток
formula_ZagOO = "=ABS(" + a1OO + "/МИН(" + a2OO + ":" + a3OO + ")" + "*100)"

With
GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n + 2, 2)
# .Value = tV13.cols.item("i_zag").Z(ndx)   # запись расчет загрузки значение
if GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n + 2, 13).value > 0: .
    FormulaLocal = formula_ZagOO  # запись расчет загрузки формула
.NumberFormat = "0"
End
With
n = n + 1
else:
GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, 1),
                   GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n + 1, GLR.X_list - 1)).Borders(
    9).LineStyle = 1  # низ

a1 = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + n + 1 + 1, 1).Address(False, False)  # ток
a2 = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + n + 1, 13).Address(False, True)  # доп ток
a3 = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + n + 1, 14).Address(False, True)  # доп ток
formula_Zag = "=ABS(" + a1 + "/МИН(" + a2 + ":" + a3 + ")" + "*100)"

With
GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + n + 1, 2)
# .Value = tV13.cols.item("i_zag").Z(ndx)   # запись расчет загрузки значение
if GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + n + 1, 13).value > 0: .
    FormulaLocal = formula_Zag  # запись расчет загрузки формула
End
With

n = n + 1

With
GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + 1, GLR.X_list - 4),
                   GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans, GLR.X_list - 1))  # рамка
.Borders(7).LineStyle = 1  # лево
.Borders(8).LineStyle = 1  # верх
.Borders(9).LineStyle = 1  # низ
.Borders(10).LineStyle = 1  # право
.Borders(11).LineStyle = 1  # внутри вертикаль
End
With

With
GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + 1, 1),
                   GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans, 2))  # рамка
.Borders(7).LineStyle = 1  # лево
.Borders(8).LineStyle = 1  # верх
.Borders(9).LineStyle = 1  # низ
.Borders(10).LineStyle = 1  # право
.Borders(11).LineStyle = 1  # внутри вертикаль
End
With

# условное форматирование
GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list, 2),
                   GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans, 2)).FormatConditions.Add(1, 5, "=100")
With
GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list, 2),
                   GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans, 2)).FormatConditions(1).Interior
.Color = 49407
End
With

GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + 1, 11).Value = "Наименование"
GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + 1, 13).Value = "Мин. допустимое напряжение, кВ"
for i = LBound(RG.KontrolNode) to UBound(RG.KontrolNode)  # цикл по ОТМЕЧЕННЫМ узлам____цикл по ОТМЕЧЕННЫМ узлам_______цикл по ОТМЕЧЕННЫМ узлам
ndx = RG.KontrolNode(i)

GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + i + 2, GLR.X_list - 4).Value = tN14.cols.item("dname").ZS(ndx)
GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + i + 2, GLR.X_list - 2).Value = tN14.cols.item("umin").ZS(ndx)
kluch = "ny=" + tN14.cols.item("ny").ZS(ndx)
if GLR.DRVXL = 1: GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + i + 2,
                                     1).Formula = "=RTD(\"rastr.rtd\",\"\",\"V\",\"$1\",""node"",""vras"",\"" + kluch + "\")"
if tN14.cols.item("uhom").ZS(ndx) > 90:  GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + i + 2, 1).NumberFormat = "0"
_
else: GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + i + 2, 1).NumberFormat = "0.0"
# GLR.XL_sheet.cell(GLR.Y_list +  GLR.Y_VL_Trans + i + 2, 1).Value = tN14.cols.item("ny").Z(ndx)

GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + i + 2, 1),
                   GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + i + 2, GLR.X_list - 1)).Borders(
    9).LineStyle = 1  # низ

With
GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + 2, GLR.X_list - 4),
                   GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 1, GLR.X_list - 1))
.Borders(7).LineStyle = 1  # лево
.Borders(8).LineStyle = 1  # верх
.Borders(9).LineStyle = 1  # низ
.Borders(10).LineStyle = 1  # право
.Borders(11).LineStyle = 1  # внутри вертикаль
End
With
With
GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + 2, 1),
                   GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 1, 2))
.Borders(7).LineStyle = 1  # лево
.Borders(8).LineStyle = 1  # верх
.Borders(9).LineStyle = 1  # низ
.Borders(10).LineStyle = 1  # право
End
With
GLR.XL_sheet.Range(GLR.XL_sheet.cell(3, 1), GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + 1, 2)).NumberFormat = "0"


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def NodeTest():  # востонавление питания узлов
    # dim OffNodeCount, OffNodeNdx, NodeOn, OffNodeNy, VetvTestNdx, VetvOnD, Nrc , Nlc , pl , ql , OffNodes ,  TotalOnVetv , ny_konec, reserva_net
    # dim tV15  , tN15
    tN15 = rastr.Tables("node")
    tV15 = rastr.Tables("vetv")
    tN15.setsel("sta+(!staRes)+(pn!=0|qn!=0|pg!=0)")
    OffNodeNdx = tN15.FindNextSel(-1)

    ReDim
    OffNodes(tN15.Count)  # узлы  в которых нашелся резерв
    OffNodeCount = 0
    NodeOn = 0
    ReDim
    TotalOnVetv(100)  # tV15.Size глобальная переменная, записываем номера узлов
    while OffNodeNdx >= 0  # проходим по всем узлам изменившим состояние и если находим выключатель включаем узел , выкл записываем в TotalOnVetv
        reserva_net = 0  # 1 есть резерв 0 нет
        OffNodeNy = tN15.cols.item("ny").Z(OffNodeNdx)  # номер текущего узла

        tV15.setsel(
            "(ip=" + str(OffNodeNy) + "|iq=" + str(OffNodeNy) + ")+(r<0.011+x<0.011)+(sta)")  # выборка выключателей
        VetvTestNdx = tV15.FindNextSel(-1)  # номер строки текущей ветви

        while VetvTestNdx >= 0  # цикл по отфильтрованным выключателям
            # for i = LBound(RG.OTKL_masiv , 2) to Ubound(RG.OTKL_masiv , 2)#
            #     if VetvTestNdx = RG.OTKL_masiv(0, i)  And RG.OTKL_masiv(1, i) = "vetv": VetvOnD = 0
            # # next

            if OffNodeNy = tV15.cols.item("ip").Z(VetvTestNdx): ny_konec = tV15.cols.item("iq").Z(VetvTestNdx) else:
                ny_konec = tV15.cols.item("ip").Z(VetvTestNdx)

            if tN15.cols.item("sta").Z(fNDX("node", ny_konec)) = False:  # если другой конец ветви включен то
                reserva_net = 1: NodeOn = 1
                tV15.cols.item("sta").Z(VetvTestNdx) = False  # ВКЛЮЧАЕМ выключатель
                tN15.cols.item("sta").Z(OffNodeNdx) = False  # ВКЛЮЧАЕМ ТЕКУЩИЙ УЗЕЛ

                OffNodes(OffNodeCount) = OffNodeNdx  # запиываем узел в массив OffNodes
                OffNodeCount = OffNodeCount + 1
                VetvTestNdx = -1
                #  tGen.setsel ("sta+(!staRes)+Node=" + str(OffNodeNy)) # вкл ген в узле
                #  tGen.cols.item("sta").calc ("0")
            else:  # если другой конец ветви отключен то
                VetvTestNdx = tV15.FindNextSel(VetvTestNdx)

        wend
        if reserva_net = 0: RGR.NodeNetReserv = RGR.NodeNetReserv + "Отключено (нет резерва) ny=" + tN15.cols.item(
            "ny").ZS(OffNodeNdx) + " - " + tN15.cols.item("dname").ZS(OffNodeNdx) + " Pн=" + str(
            Round(tN15.cols.item("pn").Z(OffNodeNdx), 1)) + ", Qн=" + str(
            Round(tN15.cols.item("qn").Z(OffNodeNdx), 1)) + ", Pг=" + str(
            Round(tN15.cols.item("pg").Z(OffNodeNdx), 1)) + ". "

        OffNodeNdx = tN15.FindNextSel(OffNodeNdx)
    wend

    if NodeOn > 0:  # если был включен узел хотя бы один узел
        rastr.rgm("p")
        for i = 0 to OffNodeCount-1
        if not isempty(OffNodes(i)):
            OffNodeNdx = OffNodes(i)
            if tN15.cols.item("sta").Z(OffNodeNdx):  # "Не удалось восстановить питание узла "
                RGR.NodeNetReserv = RGR.NodeNetReserv + "Отключено (нет резерва) ny=" + tN15.cols.item("ny").ZS(
                    OffNodeNdx) + " - " + tN15.cols.item("dname").ZS(OffNodeNdx) + " Pн=" + str(
                    Round(tN15.cols.item("pn").Z(OffNodeNdx), 1)) + ", Qн=" + str(
                    Round(tN15.cols.item("qn").Z(OffNodeNdx), 1)) + ", Pг=" + str(
                    Round(tN15.cols.item("pg").Z(OffNodeNdx), 1)) + ". "
            else:  # "Восстановлено питание узла "
                RGR.NodeRezerv = RGR.NodeRezerv + "Востановлено питание ny=" + tN15.cols.item("ny").ZS(OffNodeNdx) + " "
                if trim(tN15.cols.item("dname").ZS(OffNodeNdx)) != "":
                    RGR.NodeRezerv = RGR.NodeRezerv + tN15.cols.item("dname").ZS(OffNodeNdx) + ". " else:
                    RGR.NodeRezerv = RGR.NodeRezerv + tN15.cols.item("name").ZS(OffNodeNdx) + ". "


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
class overload_class:  # для хранения параметров текущего расчета (нр или сочетания)
    # dim shapka , val , dictVal ,  Kontrol_Key , ndxN ,ndxV , DN1, DN2 , Ir , Idz , Ia ,Iaz , Itxt , Ur , Ud  , Ud_av,UmaxZ, P , Q , Sk , S , PS ,Kontrol_name , OO , TABL_ZAG_XL
    # dim stolb_DopNz, stolb_DopNz2, stolb_DopNz3, varN , stolb_R2, stolb_R3, stolb_Urz, stolb_Izz , stolb_R2z, stolb_R3z , stolb_DopNzz , stolb_DopNz2z, stolb_DopNz3z, flagU, flagI
    # dim cFName , cGog   , cSez   , cDN1   , cDN2   , cDN3   , cNKomb , cGrad  , cK, cO    , cR1   , cR2   , cKeyO , cKeyR1, cKeyR2,cKeyK, cTip  , cKolOtkl
    # dim cfactZag , cNodeNetReserv, cNodeRezerv , cAutoShunt , cRis , cN_KO , cIr , cDDTN , cDDTNp , cADTN , cADTNp , cTXTz , cUr , cUdop , cUdop_av , cUnr , cP , cQ , cSk , cS , cPLep , cPAv , cPAn
    private
    tV17, tN17, n

    def init_c():  #
        cFName = 0
        cGog = 1
        cSez = 2
        cDN1 = 3
        cDN2 = 4
        cDN3 = 5
        cNKomb = 6
        cGrad = 7
        cK = 8
        cO = 9
        cR1 = 10
        cR2 = 11
        cKeyO = 12
        cKeyR1 = 13
        cKeyR2 = 14
        cKeyK = 15
        cTip = 16
        cKolOtkl = 17
        cfactZag = 18
        cNodeNetReserv = 19
        cNodeRezerv = 20
        cAutoShunt = 21
        cRis = 22
        cN_KO = 23
        cIr = 24
        cDDTN = 25
        cDDTNp = 26
        cADTN = 27
        cADTNp = 28
        cTXTz = 29
        cUr = 30
        cUdop = 31
        cUdop_av = 32
        cUnr = 33
        cP = 34
        cQ = 35
        cSk = 36
        cS = 37
        cPLep = 38
        cPAv = 39
        cPAn = 40

    def init_p():  # можно удалить "Откл1 таб." , "Откл2 таб." , объединить "NodeNetReserv" , "NodeRezerv"

        Sheets_add(GLR.Book_XL, TABL_ZAG_XL, "перегрузки")
        init_c()

        redim
        shapka(cPAn)
        shapka(cFName) = "Имя файла"
        shapka(cGog) = "Год"
        shapka(cSez) = "Сезон макс/мин"
        shapka(cDN1) = "Доп. имя"
        shapka(cDN2) = "Доп. имя2"
        shapka(cDN3) = "Доп. имя3"
        shapka(cNKomb) = "N комб"
        shapka(cGrad) = "темп"
        shapka(cK) = "Контролируемые элементы"
        shapka(cO) = "Отключение"
        shapka(cR1) = "Ремонт 1"
        shapka(cR2) = "Ремонт 2"
        shapka(cKeyO) = "Ключ откл1"
        shapka(cKeyR1) = "Ключ рем.1"
        shapka(cKeyR2) = "Ключ рем.2"
        shapka(cKeyK) = "Ключ контроль"
        shapka(cTip) = "tip_comb"
        shapka(cKolOtkl) = "кол. откл. эл."
        shapka(cfactZag) = "Действие по факту загрузки"
        shapka(cNodeNetReserv) = "NodeNetReserv"
        shapka(cNodeRezerv) = "NodeRezerv"
        shapka(cAutoShunt) = "AutoShunt_list"
        shapka(cRis) = "N рисунка"
        shapka(cN_KO) = "N в таблице КО"
        shapka(cIr) = "Iрасч.,A"
        shapka(cDDTN) = "Iддтн,A"
        shapka(cDDTNp) = "Iзагр. ддтн,%"
        shapka(cADTN) = "Iадтн,A"
        shapka(cADTNp) = "Iзагр. адтн,%"
        shapka(cTXTz) = "тхт загрузка"
        shapka(cUr) = "U расч.,кВ"
        shapka(cUdop) = "U доп.,кВ"
        shapka(cUdop_av) = "U ав.доп.,кВ"
        shapka(cUnr) = "U наиб. раб.,кВ"
        shapka(cP) = "P, МВт"
        shapka(cQ) = "Q, Мвар"
        shapka(cSk) = "S, МВт+jМвар"
        shapka(cS) = "S, МВА"
        shapka(cPLep) = "P, МВт для ЛЭП S, МВА для Т"
        shapka(cPAv) = "ПА есть узел"
        shapka(cPAn) = "ПА есть ветвь"

        Print_XL(TABL_ZAG_XL, 1, 1, shapka, 1, "гор", "", "", "")
        dictVal = CreateObject("Scripting.Dictionary")  # для хранения

        nach_znach()
        n = 1

    def add():  #
        if trim(RG.DopName(0)) = "": DN1 = "-" else:
            DN1 = RG.DopName(0)
        if Ubound(RG.DopName) > 0:   if
        trim(RG.DopName(1)) = "": DN2 = "-" else: DN2 = RG.DopName(1)
        if Ubound(RG.DopName) > 1:   if
        trim(RG.DopName(2)) = "": DN3 = "-" else: DN3 = RG.DopName(2)

        if ndxV > -1:
            tV17 = rastr.Tables("vetv")
            Ir = round(Ir, 1)
            Idz = round(Idz)  # "Iзагр. ддтн,%"
            if OO = 0:
                Kontrol_name = tV17.cols.item("dname").ZS(ndxV)
                Id = round(tV17.cols.item("i_dop_r").Z(ndxV))  # "Iддтн,A"
                Ia = round(fIadtn(ndxV))  # "Iадтн,A"
                Iaz = round(Ir / fIadtn(ndxV) * 100)  # "Iзагр. адтн,%"
                Itxt = f_name_zagruzka_ris("", tV17.cols.item("dname").ZS(ndxV), Idz,
                                           round(tV17.cols.item("i_max").ZS(ndxV)),
                                           round(tV17.cols.item("i_dop_r").ZS(ndxV)), round(fIadtn(ndxV)), "")
                P = abs(round(tV17.cols.item("pl_ip").ZN(ndxV)))  # "P, МВт" #  для контралируемого элемента
                Q = abs(round(tV17.cols.item("ql_ip").ZN(ndxV)))
                Sk = tV17.cols.item("S_max").ZN(ndxV)  # "S, МВт+jМвар"
                S = round(sqr(tV17.cols.item("pl_ip").ZN(ndxV) * tV17.cols.item("pl_ip").ZN(ndxV) + tV17.cols.item(
                    "ql_ip").ZN(ndxV) * tV17.cols.item("ql_ip").ZN(ndxV)))  # "S, МВА"
                if tV17.cols.item("tip").ZN(ndxV) = 1: PS = S  # если транс
                if not tV17.cols.item("tip").ZN(ndxV) = 1: PS = P  # если ВЛ или выкл
            else:
                Kontrol_name = tV17.cols.item("dname").ZS(ndxV) + " ОО"
                Id = round(tV17.cols.item("IdopOO").Z(ndxV))
                Ia = round(fIadtn_OO(ndxV))
                Iaz = round(Ir / fIadtn_OO(ndxV) * 100)
                if ndxN > -1:
            tN17 = rastr.Tables("node")
            Kontrol_name = tN17.cols.item("dname").ZS(ndxN)
            Ur = tN17.cols.item("vras").Z(ndxN)
            Ud = tN17.cols.item("umin").Z(ndxN)
            Ud_av = tN17.cols.item("umin_av").Z(ndxN)
            UmaxZ = tN17.cols.item("umax").Z(ndxN)  #

        if ndxN = -1 and ndxV = -1: Kontrol_name = "Режим не моделируется": Idz = -1
        #  если меняется val то нада менять и сводную!!!
        redim
        val(cPAn)
        val(cFName) = RG.Name_Base
        val(cGog) = RG.god
        val(cSez) = RG.SezonName
        val(cDN1) = DN1
        val(cDN2) = DN2
        val(cDN3) = DN3
        val(cNKomb) = GLR.N_rezh
        val(cGrad) = RG.GradusZ
        val(cK) = Kontrol_name
        val(cO) = RGR.otkl_name
        val(cR1) = RGR.remont_name1
        val(cR2) = RGR.remont_name2
        val(cKeyO) = RGR.otkl_key
        val(cKeyR1) = RGR.remont_key1
        val(cKeyR2) = RGR.remont_key2
        val(cKeyK) = Kontrol_Key
        val(cTip) = Comb.tip_comb
        val(cKolOtkl) = GLR.kol_otkl
        val(cfactZag) = RGR.autoTXT_fPA
        val(cNodeNetReserv) = RGR.NodeNetReserv
        val(cNodeRezerv) = RGR.NodeRezerv
        val(cAutoShunt) = GS.AutoShunt_list
        val(cRis) = GLR.name_ris1 + str(GLR.number_pict)
        val(cN_KO) = RG.TabRgmCount
        val(cIr) = Ir
        val(cDDTN) = Id
        val(cDDTNp) = Idz
        val(cADTN) = Ia
        val(cADTNp) = Iaz
        val(cTXTz) = Itxt
        val(cUr) = Ur
        val(cUdop) = Ud
        val(cUdop_av) = Ud_av
        val(cUnr) = UmaxZ
        val(cP) = P
        val(cQ) = Q
        val(cSk) = Sk
        val(cS) = S
        val(cPLep) = PS
        val(cPAv) = RGR.autoNDX_listN_info
        val(cPAn) = RGR.autoNDX_listV_info

        dictVal.Add(n, val)  # ключ  и значение
        nach_znach()
        n = n + 1

    def print_end_p(
            zadanie_tip):  # вывод dictVal  и общее форматирование(zadanie_tip 0 обычная работа, 1 анализ перегрузок без принта)
        # dim test_iii , dict_period_tek , dict_OR_tek, arr_Kontr_el_n , otklonenie_max_test
        # dim max_count, dict_Otkl_name , dictK , i_val, i_key, i_key2 , i_Kont ,i_Kont2, dict_otkl_count ,  temp9 ,val1 , i , ii , arr_dictK_Item , dict_temp ,   wordPZ , cursor, cursorRis , wordA , wordARis

        dict_otkl_count = CreateObject("Scripting.Dictionary")  # для записи количество упоменаний -  отключение
        dict_Otkl_name = CreateObject("Scripting.Dictionary")  # для записи имя откл чтоб не повторялась
        dict_temp = createobject("Scripting.dictionary")  # (откл , ремонты)
        dictK = CreateObject(
            "Scripting.Dictionary")  # dictK (имя контр эл, array(1,2 ,3) - номера строк или dictVal.Item)
        dictK2 = CreateObject(
            "Scripting.Dictionary")  # dictK (имя контр эл, array(0,1,2,3)-объекты класса для каждого кол откл соответствующего элемента))
        #  формируем dictK
        if dictVal.count > 0:
            for i = 0 to dictVal.count-1
            i_val = dictVal.Items()(i)
            if not dictK.exists(i_val(cK)):
                dictK.add(i_val(cK), str(dictVal.Keys()(i)))
            else:
                dictK.Item(i_val(cK)) = dictK.Item(i_val(cK)) + "," + str(dictVal.Keys()(i))

        # print_dic dictK
        for each i_key in dictK.Keys
            if mid(dictK.Item(i_key), 1, 1) = ",": dictK.Item(i_key) = mid(dictK.Item(i_key), 2,
                                                                           len(dictK.Item(i_key)) - 1)
            arr_dictK_Item = split(dictK.Item(i_key), ",")
            for ii = 0 to ubound ( arr_dictK_Item)
            arr_dictK_Item(ii) = float(arr_dictK_Item(ii))

        dictK.Item(i_key) = arr_dictK_Item

    #  СОРТИРОВКА
    for Each i_Kont in dictK.Keys  # для каждого контролируемого элемента отдельно
        if ubound(dictK.Item(i_Kont)) > 0:
            dict_otkl_count.removeall
            dict_Otkl_name.RemoveAll
            max_count = 0
            for Each i in dictK.Item(
                    i_Kont)  # цикл по массиву строк dictVal, записываем сколько каких наименований отключений есть для текущего контроля
                i_val = dictVal.Item(i)
                if i_val(cTip) = "-1" and i_val(cK) = i_Kont and i_Kont != "-" and i_Kont != "":
                    if dict_otkl_count.exists(i_val(cO)):
                        dict_otkl_count.Item(i_val(cO)) = dict_otkl_count.Item(i_val(cO)) + 1  else:
                        dict_otkl_count.add(i_val(cO), 1)
                    if dict_otkl_count.exists(i_val(cR1)):
                        dict_otkl_count.Item(i_val(cR1)) = dict_otkl_count.Item(i_val(cR1)) + 1  else:
                        dict_otkl_count.add(i_val(cR1), 1)
                    if dict_otkl_count.exists(i_val(cR2)):
                        dict_otkl_count.Item(i_val(cR2)) = dict_otkl_count.Item(i_val(cR2)) + 1  else:
                        dict_otkl_count.add(i_val(cR2), 1)

            for i_Key = 0 to dict_otkl_count.count-1  # находим максимальное значение повторений наименования отключения
            if max_count < dict_otkl_count.Items()(i_Key): max_count = dict_otkl_count.Items()(i_Key)

        if max_count > 1:
            for i = max_count to 2 step -1  # от максимального количества отключений до 2

            for i_Key = 0 to dict_otkl_count.count-1  # находим ключ или имя
            if i = dict_otkl_count.Items()(i_Key):  # нашли

                name_tekk = dict_otkl_count.Keys()(i_Key)  # имя откл которое будем перестовлять

                for Each i_key2 in dictK.Item(i_Kont)  # сорт
                    val1 = dictVal.Item(i_key2)
                    if val1(cKolOtkl) > 1 and val1(cTip) = "-1" and val1(cK) = i_Kont:
                        if val1(cR1) = name_tekk and val1(cR1) != "" and val1(cR1) != "-" and val1(cO) != "" and val1(cO) != "-":
                            if not dict_Otkl_name.exists(val1(cO)):  # можно и нада менять сК1 и cO местами
                                # logging.info( i_Kont +" меняем R1 " + name_tekk +" и O " + val1(cO) +" N комб " + str( val1(cK-2)))
                                val1(cR1) = val1(cO)
                                val1(cO) = name_tekk
                                temp9 = val1(cKeyR1)
                                val1(cKeyR1) = val1(cKeyO)
                                val1(cKeyO) = temp9
                                dictVal.Item(i_key2) = val1
                                if val1(cKolOtkl) > 2:
                            if val1(cR2) = name_tekk and val1(cR2) != "" and val1(cR2) != "-" and val1(cO) != "" and val1(cO) != "-":
                                if not dict_Otkl_name.exists(val1(cO)):  # можно и нада менять сК1 и cO местами
                                    # logging.info( i_Kont +" меняем R2 " + name_tekk + " и O " + val1(cO) +" N комб " + str( val1(cK-2)))
                                    val1(cR2) = val1(cO)
                                    val1(cO) = name_tekk
                                    temp9 = val1(cKeyR2)
                                    val1(cKeyR2) = val1(cKeyO)
                                    val1(cKeyO) = temp9
                                    dictVal.Item(i_key2) = val1

                if not dict_Otkl_name.exists(name_tekk): dict_Otkl_name.add(name_tekk, 1)
                for i = max_count to 2 step -1  # от максимального количества отключений до 2
        for i_Key = 0 to dict_otkl_count.count-1  # находим ключ или имя
        if i = dict_otkl_count.Items()(i_Key):  # нашли

            name_tekk = dict_otkl_count.Keys()(i_Key)  # имя откл которое будем перестовлять
            for Each i_key2 in dictK.Item(i_Kont)  # сорт
                val1 = dictVal.Item(i_key2)
                if val1(cKolOtkl) > 2 and val1(
                        cK) = i_Kont and val1(cR2) = name_tekk and val1(cR2) != "" and val1(cR2) != "-" and val1(cR1) != "" and val1(cR1) != "-":
                    if not dict_Otkl_name.exists(val1(cR1)):  # можно и нада менять сК1 и cO местами
                        # logging.info( i_Kont +" меняем R2 " + name_tekk +" и R1 " + val1(cR1)+" N комб " + str( val1(cK-2)))
                        val1(cR2) = val1(cR1)
                        val1(cR1) = name_tekk
                        temp9 = val1(cKeyR2)
                        val1(cKeyR2) = val1(cKeyR1)
                        val1(cKeyR1) = temp9
                        dictVal.Item(i_key2) = val1

            if not dict_Otkl_name.exists(name_tekk): dict_Otkl_name.add(name_tekk, 1)
            #  принт хл


if zadanie_tip = 0:
    for Each n in dictVal.Keys  # организуем цикл по элементам  масива Keys
        Print_XL(TABL_ZAG_XL, 1, n + 1, dictVal.Item(n), 1, "гор", "", "", "")

    if TABL_ZAG_XL.UsedRange.rows.count > 1:
        TABL_ZAG_XL.ListObjects.Add(1, TABL_ZAG_XL.Range(TABL_ZAG_XL.UsedRange.address))  # используемы диапозон листа
        TABL_ZAG_XL.ListObjects(1).name = "TABL_ZAG"
        TABL_ZAG_XL.Columns("A:AA").AutoFit
        svod_all_peregruz()  # СДЕЛАТЬ СВОДНЫЕ ИЗ ПРОТАКОЛА ПЕРЕГРУЗОК            # !!!!!! ПЗ в ВОРД!!!!!!!!
if GLR.PZ_word:
    wordA = CreateObject("word.Application")
    wordPZ = wordA.Documents.Add()
    # wordA.ScreenUpdating = False
    wordA.Visible = True
    # wordA.ScreenUpdating = True
    cursor = wordA.Selection
    cursor.EndKey(6)  # перейти в конец текста
    cursor.Font.Size = 12  # шрифт
    cursor.Font.Name = "Times  Roman"
    wordARis = CreateObject("word.Application")
    wordPZRis = wordARis.Documents.Add()
    wordARis.Visible = True
    cursorRis = wordARis.Selection
    cursorRis.Font.Size = 12  # шрифт
    cursorRis.Font.Name = "Times  Roman"
    redim
    arr_Kontr_el_n(3)  # # dim
    for Each i_Kont in dictK.Keys  # обработка dictVal для каждого контролируемого элемента отдельно
        for iii = 0 to 3  # от нр до н-1
        Kontr_el_n = Kontrol_el
        Kontr_el_n.init_KE()

        for Each iiii in dictK.Item(
                i_Kont)  # цикл по строкам с этим контр элементом, iiii -ключ  номер строки в dictVal
            if iiii > 0 and iii = dictVal.item(iiii) (cKolOtkl):
                Kontr_el_n.NO_ON = True
                Kontr_el_n.kol_otkl_comb = iii
                Kontr_el_n.name_el = i_Kont

                if instr(dictVal.item(iiii)(cKeyK), "ny") > 0:
                    Kontr_el_n.TipKontr = "node"
                elif instr(dictVal.item(iiii)(cKeyK), "ip") > 0:
                    Kontr_el_n.TipKontr = "vetv"
                else:
                    Kontr_el_n.TipKontr = "not_mod"

                if not Kontr_el_n.dict_period.exists(LCase(dictVal.item(iiii)(cSez))):
                    Kontr_el_n.dict_period.add(LCase(dictVal.item(iiii)(cSez)), str(dictVal.item(iiii)(cGog)))
                else:
                    Kontr_el_n.dict_period.Item(LCase(dictVal.item(iiii)(cSez))) = Kontr_el_n.dict_period.Item(
                        LCase(dictVal.item(iiii)(cSez))) + "," + str(dictVal.item(iiii)(cGog))

                if iii > 0:
                    if not Kontr_el_n.dict_OR.exists(dictVal.item(iiii)(cO)):  # (o, r1|r2:r1|r3)
                        if iii > 2:
                            Kontr_el_n.dict_OR.add(dictVal.item(iiii)(cO),
                                                   trim(dictVal.item(iiii)(cR1)) + ", " + trim(dictVal.item(iiii)(cR2)))
                        else:
                            Kontr_el_n.dict_OR.add(dictVal.item(iiii)(cO), trim(dictVal.item(iiii)(cR1)))

                    else:  # было
                        if iii > 2:
                            Kontr_el_n.dict_OR.item(dictVal.item(iiii)(cO)) = Kontr_el_n.dict_OR.item(
                                dictVal.item(iiii)(cO)) + ";" + trim(dictVal.item(iiii)(cR1)) + " и " + trim(
                                dictVal.item(iiii)(cR2))
                        else:
                            Kontr_el_n.dict_OR.item(dictVal.item(iiii)(cO)) = Kontr_el_n.dict_OR.item(
                                dictVal.item(iiii)(cO)) + ";" + trim(dictVal.item(iiii)(cR1))
                otklonenie_max_test = False  # если истина макс и добавляем

                if Kontr_el_n.TipKontr  = "vetv":
                    if float(dictVal.item(iiii)(cADTNp)) > 100:  # cравниваем по ADTN
                        if Kontr_el_n.Azag_max < float(dictVal.item(iiii)(cIr)) / float(
                            dictVal.item(iiii)(cADTN)):   otklonenie_max_test = True
                    else:  # cравниваем по DDTN
                        if Kontr_el_n.Dzag_max < float(dictVal.item(iiii)(cIr)) / float(
                            dictVal.item(iiii)(cDDTN)):   otklonenie_max_test = True

                elif Kontr_el_n.TipKontr  = "node":
                    if float(dictVal.item(iiii)(cUr)) < Kontr_el_n.U_min_r or Kontr_el_n.U_min_r =0:
                        otklonenie_max_test = True
                        Kontr_el_n.U_min_r = float(dictVal.item(iiii)(cUr))
                        if otklonenie_max_test:
                    Kontr_el_n.strN = iiii
                    Kontr_el_n.period_max = LCase(dictVal.item(iiii)(cSez) + " " + dictVal.item(iiii)(cGog) + " г")

                    if dictVal.item(iiii)(cO) != "-":
                        Kontr_el_n.otklMaxName = " при отключении " + dictVal.item(iiii)(cO)
                        if dictVal.item(iiii)(
                            cR1) != "-": Kontr_el_n.otklMaxName = Kontr_el_n.otklMaxName + " в схеме ремонта " + dictVal.item(
                            iiii)(cR1)
                        if dictVal.item(iiii)(
                            cR2) != "-": Kontr_el_n.otklMaxName = Kontr_el_n.otklMaxName + ", " + dictVal.item(iiii)(
                            cR2)
                    elif dictVal.item(iiii)(cO) = "-" and dictVal.item(iiii) (cR1) != "-":
                        Kontr_el_n.otklMaxName = " в схеме ремонта " + dictVal.item(iiii)(cR1)
                        if dictVal.item(iiii)(
                            cR2) != "-": Kontr_el_n.otklMaxName = Kontr_el_n.otklMaxName + ", " + dictVal.item(iiii)(
                            cR1)

                    if Kontr_el_n.TipKontr  = "vetv":
                        Kontr_el_n.Azag_max = float(dictVal.item(iiii)(cIr)) / float(dictVal.item(iiii)(cADTN))
                        Kontr_el_n.Dzag_max = float(dictVal.item(iiii)(cIr)) / float(dictVal.item(iiii)(cDDTN))

                        if dictVal.item(iiii)(cADTNp) = dictVal.item(iiii) (cDDTNp):
                            Kontr_el_n.zag_maxT = str(dictVal.item(iiii)(cDDTNp)) + " % (" + str(
                                round(dictVal.item(iiii)(cIr))) + " А) от Iддтн = Iадтн = " + str(
                                dictVal.item(iiii)(cDDTN)) + " А"  # имя контр
                        else:
                            Kontr_el_n.zag_maxT = str(dictVal.item(iiii)(cDDTNp)) + " % (" + str(
                                round(dictVal.item(iiii)(cIr))) + " А) от Iддтн = " + str(
                                dictVal.item(iiii)(cDDTN)) + " А и " + str(
                                dictVal.item(iiii)(cADTNp)) + " % от Iадтн = " + str(dictVal.item(iiii)(cADTN)) + " А"

                    elif Kontr_el_n.TipKontr  = "node":
                        Kontr_el_n.zag_maxT = str(round(dictVal.item(iiii)(cUr), 1)) + " кВ при Uмдн = " + str(
                            round(dictVal.item(iiii)(cUdop), 1)) + " кВ и Uадн = " + str(
                            round(dictVal.item(iiii)(cUdop_av), 1)) + " кВ"  # имя контр

                    Kontr_el_n.nStrDictValMax = iiii
                    Kontr_el_n.File_RG2 = dictVal.item(iiii)(cFName)

        arr_Kontr_el_n(iii) = Kontr_el_n

    dictK2.add(i_Kont, arr_Kontr_el_n)

#  принт в ворд
for iii = 0 to 3  # от нр до н-1
cursor.ParagraphFormat.SpaceAfter = 0  # отступ после параграфа
cursor.ParagraphFormat.Alignment = 3  # выравнять по ширине
cursor.ParagraphFormat.CharacterUnitLeftIndent = 0  # отступ слева
cursor.ParagraphFormat.CharacterUnitFirstLineIndent = 2.36  # отступ слева 1 строки
cursor.Font.Bold = 9999998
# cursor.Style = ActiveDocument.Styles("Заголовок 1")
if iii = 0: cursor.TypeText("Анализ режимов работы электрической сети 110 кВ и выше в нормальной схеме")
if iii = 1: cursor.TypeText(
    "Анализ режимов работы электрической сети 110 кВ и выше при нормативных возмущениях в нормальной схеме")
if iii = 2: cursor.TypeText(
    "Анализ режимов работы электрической сети 110 кВ и выше при нормативных возмущениях в ремонтной схеме")
if iii = 3: cursor.TypeText(
    "Анализ режимов работы электрической сети 110 кВ и выше при нормативных возмущениях в двойной ремонтной схеме")
cursor.Font.Bold = 9999998
# cursor.Style = ActiveDocument.Styles("Обычный")

cursor.TypeParagraph

test_iii = False
for Each i_Kont2 in dictK2.Keys
    if dictK2.item(i_Kont2)(iii).NO_ON = True and dictK2.item(i_Kont2)(iii).kol_otkl_comb = iii: test_iii = True: exit
    for

if test_iii:
    cursor.TypeText("Выявлены перегрузки следующих элементов сети." + "\n")
    for Each i_Kont in dictK2.Keys
        if dictK2.item(i_Kont)(iii).NO_ON:
            cursor.Font.Bold = 9999998
            cursor.TypeText(i_Kont + "\n")
            cursor.Font.Bold = 9999998
            if i_Kont = "Режим не моделируется":
                cursor.TypeText("Режим не моделируется в ")
            else:
                if dictK2.item(i_Kont)(iii).TipKontr = "vetv":
                    cursor.TypeText("Превышение допустимой загрузки ")
                elif dictK2.item(i_Kont)(iii).TipKontr = "node":
                    cursor.TypeText("Недопустимое отклонение напряжения ")

                cursor.TypeText(i_Kont)
                cursor.TypeText(" выявлено в ")

            dict_period_tek = dictK2.item(i_Kont)(iii).dict_period

            if dict_period_tek.count > 1:
                cursor.TypeText("следующих режимно-балансовых условиях: " + "\n")

                for nnn  = 0 to dict_period_tek.count-1
                cursor.TypeText(
                    "- " + dict_period_tek.Keys()(nnn) + " (" + god_str_diapozon(dict_period_tek.items()(nnn)) + ")")
                if nnn = dict_period_tek.count-1:
                    cursor.TypeText("." + "\n")  # последнее
                else:
                    cursor.TypeText(";" + "\n")

        elif dict_period_tek.count = 1:
            cursor.TypeText(dict_period_tek.keys()(0) + " " + god_str_diapozon(
                dict_period_tek.items()(0)) + "." + "\n")  # + " " + dict_period_tek.Items()(0)

        #  во всех рассматриваемых электроэнергетических режимах?
        if iii > 0:
            dict_OR_tek = dictK2.item(i_Kont)(iii).dict_OR  # (o, r1|r2:r1|r3)

            cursor.TypeText(
                "Отклонение параметров режима от допустимых значений выявлено в следующих схемно-режимных ситуациях: " + "\n")

            for nnn = 0 to dict_OR_tek.count-1
            cursor.TypeText("- отключение " + trim(dict_OR_tek.Keys()(nnn)))

            if iii > 1:
                temp9 = split(dict_OR_tek.items()(nnn), ";")
                for yyy= 0 to ubound(temp9)
                if not dict_temp.exists(temp9(yyy)):  dict_temp.add(temp9(yyy), 0)

            if dict_temp.count > 1:
                if iii = 2: cursor.TypeText(" в схеме ремонта одного из следующих элементов сети: ")
                if iii = 3: cursor.TypeText(" в схеме ремонтов одного из сочетаний следующих элементов сети: ")

                for yyy= 0 to dict_temp.count-1

                if yyy > 0:  cursor.TypeText(", ")
                cursor.TypeText(dict_temp.keys()(yyy))

            dict_temp.removeall()
        else:
            cursor.TypeText(
                " в схеме ремонта " + dict_temp.keys()(0)) if not nnn = dict_OR_tek.count - 1: cursor.TypeText(
                ";" + "\n") else: cursor.TypeText("." + "\n")

if not i_Kont = "Режим не моделируется":
    # МОДЕЛИРУЕМ РЕЖИМ
    if GLR.ris_PZ_word:

    if dictK2.item(i_Kont)(iii).TipKontr = "vetv":
        cursor.TypeText("Максимальная токовая загрузка наблюдается в ")
    elif dictK2.item(i_Kont)(iii).TipKontr = "node":
        cursor.TypeText("Максимальное снижение напряжения наблюдается в ")

    cursor.TypeText(dictK2.item(i_Kont)(iii).period_max + dictK2.item(i_Kont)(iii).otklMaxName)
    cursor.TypeText(" и составляет " + str(dictK2.item(i_Kont)(iii).zag_maxT) + " (рисунок " + "?" + "). " + "\n")
    #  имя рисунка нр
    # RGR.name_ris_info = array ("рис",          "номер",        "сезон год",                                                                                                 "доп имя"  ,"имя кадр",     "нр/откл+действие" ,     "загрузка" )
    # RGR.name_ris      = array ( GLR.name_ris1, Npic     , str(Npic)+ dictVal.item(Kontr_el_n.strN) (cSez) + " " + str(Npic)+ dictVal.item(Kontr_el_n.strN) (cGog)+" г.", RG.txt_dop ,  "",   RGR.raschot_name + ". "  , name_zagruzka_ris     )
    cursor.TypeText("Перегрузка оборудования ликвидируется действием АОПО " + i_Kont + + " (рисунок " + str(
        "?") + ".1). " + "\n")  # ????????

else:
cursor.TypeText("Схемно-режимных ситуации, характеризующиеся выходом параметров режима из области допустимых значений ")
if iii = 0: cursor.TypeText("в нормальной схеме электрической сети ")
if iii = 1: cursor.TypeText("при нормативных возмущениях в нормальной схеме электрической сети ")
if iii = 2: cursor.TypeText("при нормативных возмущениях в ремонтных схемах электрической сети ")
if iii = 3: cursor.TypeText("при нормативных возмущениях в двойных ремонтных схемах электрической сети ")
cursor.TypeText("не выявлено." + "\n")


def god_str_diapozon(str):  # функция из "2021,2023,2025" делает "2021-2025"
    # dim strm ,  i  , god_max , god_min
    god_max = 0
    god_min = 0
    god_str = replace(str, " ", "")
    strm = split(god_str, ",")
    for i = 0 to ubound (strm)

    if float(strm(i)) > god_max: god_max = float(strm(i))
    if float(strm(i)) < god_min or god_min = 0: god_min = float(strm(i))


if god_max = god_min: god_str_diapozon = str(god_max) + " г" else:
    god_str_diapozon = str(god_min) + "–" + str(god_max) + " гг"


# End def return

def nach_znach():
    DN2 = ""  # "Доп. имя2"
    DN3 = ""  # "Доп. имя2"
    Kontrol_name = ""  # "Контролируемые элементы"
    Ir = ""  # "Iрасч.,A"
    Id = ""  # "Iддтн,A"
    Idz = ""  # "Iзагр. ддтн,%"
    Ia = ""  # "Iадтн,A"
    Iaz = ""  # "Iзагр. адтн,%"
    Itxt = ""  # "тхт загрузка"
    Ur = ""  # "U расч.,кВ"
    Ud = ""  # "U доп.,кВ"
    Kontrol_Key = ""  # "Ключ контроль"
    P = ""  # "P, МВт"
    Q = ""  # "Q, Мвар"
    Sk = ""  # "S, МВт+jМвар"
    S = ""  # "S, МВА"
    PS = ""  # "P, МВт для ЛЭП S, МВА для Т"
    ndxN = -1
    ndxV = -1
    OO = 0  # если 1 то общая обмотка

    # +++++++++++++++++++++++свод+++++++++


def svod_all_peregruz():  # СДЕЛАТЬ СВОДНЫЕ ИЗ ПРОТАКОЛА ПЕРЕГРУЗОК
    # dim tabl_peregruz

    tabl_peregruz = TABL_ZAG_XL.ListObjects("TABL_ZAG")
    stolb_DopNz = tabl_peregruz.Range.Columns(4)  # доп имя1
    stolb_DopNz2 = tabl_peregruz.Range.Columns(5)  # доп имя2
    stolb_DopNz3 = tabl_peregruz.Range.Columns(6)  # доп имя3
    stolb_K = tabl_peregruz.Range.Columns(10)  # присвоить столбец диапозон ячейки таблицы
    stolb_R2 = tabl_peregruz.Range.Columns(11)
    stolb_R3 = tabl_peregruz.Range.Columns(12)
    stolb_Iz = tabl_peregruz.Range.Columns(25)  # Iзагр. ддтн,%
    stolb_Ur = tabl_peregruz.Range.Columns(31)

    stolb_Kz = stolb_K.Range(TABL_ZAG_XL.cell(2, 1),
                             TABL_ZAG_XL.cell(stolb_K.Rows.count, 1))  # убираем ячейку заголовка
    stolb_R2z = stolb_R2.Range(TABL_ZAG_XL.cell(2, 1), TABL_ZAG_XL.cell(stolb_R2.Rows.count, 1))
    stolb_R3z = stolb_R3.Range(TABL_ZAG_XL.cell(2, 1), TABL_ZAG_XL.cell(stolb_R3.Rows.count, 1))
    stolb_Urz = stolb_Ur.Range(TABL_ZAG_XL.cell(2, 1), TABL_ZAG_XL.cell(stolb_Ur.Rows.count, 1))
    stolb_Izz = stolb_Iz.Range(TABL_ZAG_XL.cell(2, 1), TABL_ZAG_XL.cell(stolb_Iz.Rows.count, 1))
    stolb_DopNzz = stolb_DopNz.Range(TABL_ZAG_XL.cell(2, 1), TABL_ZAG_XL.cell(stolb_DopNz.Rows.count, 1))
    stolb_DopNz2z = stolb_DopNz2.Range(TABL_ZAG_XL.cell(2, 1), TABL_ZAG_XL.cell(stolb_DopNz2.Rows.count, 1))
    stolb_DopNz3z = stolb_DopNz3.Range(TABL_ZAG_XL.cell(2, 1), TABL_ZAG_XL.cell(stolb_DopNz3.Rows.count, 1))
    flagU = False
    flagI = False
    for Each varN in stolb_Urz  #
        if varN.value != "": flagU = True: exit
        for

    for Each varN in stolb_Izz  #
        if varN.value = "": flagI = True: exit
        for

    if TABL_ZAG_XL.FilterMode: TABL_ZAG_XL.ShowAllData  # убрать фильты если они есть
    # zam_range (stolb_R2z, "", "-")     #  заменит в дивапазоне значение 1 на 2
    SVOD(tabl_peregruz, SvodListK, "свод_К", "свод_К")
    SVOD(tabl_peregruz, SvodListO1, "свод_O", "свод_O")
    # SVODn2(tabl_peregruz, SvodListN2, "зад_n2", "зад_n2")

    if flagU:
        SVOD(tabl_peregruz, SvodListK, "свод_К_U", "свод_К_U")
        SVOD(tabl_peregruz, SvodListK, "свод_O_U", "свод_O_U")


def SVOD(tabl_peregruz, List, name_list, name_PT):  # диапозон источник ,сылка на лист, имя листа, имя таблицы
    # dim rng,  PTCache, PT , PositionK , pvtitem
    PositionK = 2
    GLR.Book_XL.Worksheets(1).Activate
    Sheets_add(GLR.Book_XL, List, name_list)

    PTCache = GLR.Book_XL.PivotCaches.Create(1, tabl_peregruz)  # создать КЭШ
    PT = PTCache.CreatePivotTable(name_list + "!R1C1", name_PT)  # создать сводную таблицу

    arr_str = "Контролируемые элементы;Отключение"  # строки
    arr_stb = "Год;Сезон макс/мин"  # столбцы
    arr_flt = Array("Действие по факту загрузки", "Имя файла", "кол. откл. эл.")  # фильтр

    for Each varN in stolb_R2z  #
        if varN.value != "" and varN.value != "-": arr_str = arr_str + ";Ремонт 1":  PositionK = 3: exit
        for

    for Each varN in stolb_R3z  #
        if varN.value != "" and varN.value != "-": arr_str = arr_str + ";Ремонт 2": exit
        for

    for Each varN in stolb_DopNzz  #
        if varN.value != "" and varN.value != "-": arr_stb = arr_stb + ";Доп. имя": exit
        for

    for Each varN in stolb_DopNz2z  #
        if varN.value != "" and varN.value != "-": arr_stb = arr_stb + ";Доп. имя2": exit
        for

    for Each varN in stolb_DopNz3z  #
        if varN.value != "" and varN.value != "-": arr_stb = arr_stb + ";Доп. имя3": exit
        for

    arr_str = split(arr_str, ";")
    arr_stb = split(arr_stb, ";")

    With
    PT
    .ManualUpdate = True  # не обновить сводную
    .AddFields
    arr_str, arr_stb, arr_flt, False  # добавить поля .AddFields (RowFields, ColumnFields, PageFields, AddToTable)
    .PivotFields(
        "Контролируемые элементы").ShowDetail = False  # показывать подкатегории            PivotField -   Представляет поле в отчете PivotTable
    if name_list = "свод_O" Or name_list = "свод_O_U": .
        PivotFields("Контролируемые элементы").Position = PositionK
    if name_list = "свод_К" Or name_list = "свод_O":  # .AddDataField (Field, Caption, def)
        .AddDataField.PivotFields("Iрасч.,A"), "Iрасч.,A ", -4136  # xlSum #  добавить поле в область значений
        # Caption имя заголовка, не равна имени поля # def формула расчета
        .AddDataField.PivotFields(
            "Iддтн,A"), "Iддтн,A ", -4136  # xlMax -4136 xlSum -4157 #  добавить поле в область значений
        .AddDataField.PivotFields("Iзагр. ддтн,%"), "Iзагр. ддтн,% ", -4136  #
        .AddDataField.PivotFields("Iадтн,A"), "Iадтн,A ", -4136  #
        .AddDataField.PivotFields("Iзагр. адтн,%"), "Iзагр. адтн,% ", -4136  #
        # .AddDataField .PivotFields("P, МВт для ЛЭП S, МВА для Т"), "P(S), МВт(МВА)", -4136 #

    if flagU And name_list != "свод_К":
        .AddDataField.PivotFields("U расч.,кВ"), "U расч.,кВ ", -4157
        .AddDataField.PivotFields("U доп.,кВ"), "U доп.,кВ ", -4157

    if name_list = "свод_O" Or name_list = "свод_O_U": .
        PivotFields("Отключение").ShowDetail = True  # группировка
    if name_list = "свод_К" Or name_list = "свод_К_U": .
        PivotFields("Контролируемые элементы").ShowDetail = True  # группировка
    .RowAxisLayout
    1  # 1 xlTabularRow показывать в табличной форме!!!!
    .DataPivotField.Orientation = 1  # xlRowField = 1 "Значения" в столбцах или строках xlColumnField
    # .DataPivotField.Position = 1 #  позиция в строках
    .RowGrand = False  # удалить строку общих итогов
    .ColumnGrand = False  # удалить столбец общих итогов
    .MergeLabels = True  # обединять одинаковые ячейки
    .HasAutoFormat = False  # не обновлять ширину при обнавлении
    .NullString = "--"  # заменять пустые ячейки
    .PreserveFormatting = False  # сохранять формат ячеек при обнавлении
    .ShowDrillIndicators = False  # показывать кнопки свертывания
    .PivotCache.MissingItemsLimit = xlMissingItemsNone  # для норм отображения уникальных значений автофильтра ???????
    .PivotFields("Контролируемые элементы").Subtotals = Array(False, False, False, False, False, False, False, False,
                                                              False, False, False,
                                                              False)  # промежуточные итоги и фильтры
    .PivotFields("Отключение").Subtotals = Array(False, False, False, False, False, False, False, False, False, False,
                                                 False, False)  # промежуточные итоги и фильтры
    .PivotFields("Ремонт 1").Subtotals = Array(False, False, False, False, False, False, False, False, False, False,
                                               False, False)  # промежуточные итоги и фильтры
    .PivotFields("Ремонт 2").Subtotals = Array(False, False, False, False, False, False, False, False, False, False,
                                               False, False)  # промежуточные итоги и фильтры
    .PivotFields("Год").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False,
                                          False)  # промежуточные итоги и фильтры
    .PivotFields("Доп. имя").Subtotals = Array(False, False, False, False, False, False, False, False, False, False,
                                               False, False)  # промежуточные итоги и фильтры
    .PivotFields("Доп. имя2").Subtotals = Array(False, False, False, False, False, False, False, False, False, False,
                                                False, False)  # промежуточные итоги и фильтры
    .PivotFields("Доп. имя3").Subtotals = Array(False, False, False, False, False, False, False, False, False, False,
                                                False, False)  # промежуточные итоги и фильтры
    .PivotFields("Год").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False,
                                          False)  # промежуточные итоги и фильтры
    .PivotFields("Сезон макс/мин").Subtotals = Array(False, False, False, False, False, False, False, False, False,
                                                     False, False, False)  # промежуточные итоги и фильтры

    if name_list = "свод_К_U" Or name_list = "свод_O_U":
        .PivotFields("U расч.,кВ").Orientation = 3  # xlPageField = 3
        .PivotFields("U расч.,кВ").CurrentPage = "(All)"

        for Each pvtitem in .PivotFields("U расч.,кВ").PivotItems
        if pvtitem.Name  = "(blank)": .
            PivotFields("U расч.,кВ").PivotItems("(blank)").Visible = False: exit
        for


if name_list = "свод_К":
    .PivotFields("Iзагр. ддтн,%").Orientation = 3
    .PivotFields("Iзагр. ддтн,%").CurrentPage = "(All)"
    if flagI:
        for Each pvtitem in .PivotFields("Iзагр. ддтн,%").PivotItems
        if pvtitem.Name  = "(blank)": .
            PivotFields("Iзагр. ддтн,%").PivotItems("(blank)").Visible = False: exit
        for

if name_list = "свод_К" Or name_list = "свод_O":
    .PivotFields("Iрасч.,A ").NumberFormat = "0"
    .PivotFields("Iзагр. ддтн,% ").NumberFormat = "0"
    .PivotFields("Iадтн,A ").NumberFormat = "0"

.ManualUpdate = False  # обновить сводную
.TableStyle2 = ""  # стиль
.ColumnRange.ColumnWidth = 10  # ширина строк
.RowRange.ColumnWidth = 10
.RowRange.Columns(1).ColumnWidth = 25
.RowRange.Columns(2).ColumnWidth = 25
.RowRange.Columns(3).ColumnWidth = 25
.RowRange.Columns(4).ColumnWidth = 17

.DataBodyRange.HorizontalAlignment = -4108  # xlCenter = -4108
.DataBodyRange.NumberFormat = "#,##0"

With.TableRange1  # формат
.WrapText = True  # перенос текста в ячейке
.Borders(7).LineStyle = 1  # лево
.Borders(8).LineStyle = 1  # верх
.Borders(9).LineStyle = 1  # низ
.Borders(10).LineStyle = 1  # право
.Borders(11).LineStyle = 1  # внутри вертикаль
.Borders(12).LineStyle = 1  #
End
With
End
With
# -------------------УСЛОВНОЕ Форматирование------------------------------
if name_list = "свод_К" Or name_list = "свод_O":
    uslovnoeFc(PT.DataBodyRange.Rows(3).cell(1))
    uslovnoeFc(PT.DataBodyRange.Rows(5).cell(1))
    for Each rng in PT.DataBodyRange.Rows
        if rng.Rows.count > 1:
            if List.cell(rng.row, rng.Column - 1) = "Iзагр. ддтн,% ": uslovnoeFz(
                rng)  # выделить максимальное значение жирным
            if List.cell(rng.row, rng.Column - 1) = "Iзагр. адтн,% ": uslovnoeFz(rng)


def SVODn2(tabl_peregruz, List, name_list, name_PT):  # диапозон источник ,сылка на лист, имя листа, имя таблицы
    # dim rng, PTCache, PT

    Sheets_add(GLR.Book_XL, List, name_list)
    PTCache = GLR.Book_XL.PivotCaches.Create(1, tabl_peregruz)  # создать КЭШ
    PT = PTCache.CreatePivotTable(name_list + "!R10C4", name_PT)  # создать сводную таблицу

    With
    PT
    .ManualUpdate = True  # не обновить сводную .AddFields (RowFields, ColumnFields, PageFields, AddToTable)
    .AddFields
    Array("Отключение", "Ключ откл1", "Откл1 таб.")
    _
    , Array("Ремонт 1", "Ремонт 2", "Ключ откл2", "Откл2 таб.")
    _
    , Array("Имя файла", "Доп. имя", "Год", "Сезон макс/мин", "Действие по факту загрузки"), False
    # RowFields:=строки,ColumnFields:=столбцы, PageFields:= области офильтров

.PivotFields("Контролируемые элементы").ShowDetail = False  # показывать подкатегории
.AddDataField.PivotFields("Имя файла"), "Имя файла ", -4112  # lCount = -4112 добавить поле в область значений
# Caption имя заголовка, не равна имени поля # def формула расчета

.PivotFields("Отключение").ShowDetail = True  # группировка
.RowAxisLayout
1  # показывать в табличной форме!!!!

# .DataPivotField.Position = 1 #  позиция в строках
.RowGrand = False  # удалить строку общих итогов
.ColumnGrand = False  # удалить столбец общих итогов
.MergeLabels = False  # обединять одинаковые ячейки
.HasAutoFormat = False  # не обновлять ширину при обнавлении
.PreserveFormatting = False  # сохранять формат ячеек при обнавлении
.ShowDrillIndicators = False  # показывать кнопки свертывания
.PivotCache.MissingItemsLimit = xlMissingItemsNone  # для норм отображения уникальных значений автофильтра
.PivotFields("Отключение").Subtotals = Array(False, False, False, False, False, False, False, False, False, False,
                                             False, False)  # промежуточные итоги и фильтры
.PivotFields("Ключ откл1").Subtotals = Array(False, False, False, False, False, False, False, False, False, False,
                                             False, False)  # промежуточные итоги и фильтры
.PivotFields("Откл1 таб.").Subtotals = Array(False, False, False, False, False, False, False, False, False, False,
                                             False, False)  # промежуточные итоги и фильтры
.PivotFields("Ремонт 1").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False,
                                           False)  # промежуточные итоги и фильтры
.PivotFields("Ремонт 2").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False,
                                           False)  # промежуточные итоги и фильтры
.PivotFields("Ключ откл2").Subtotals = Array(False, False, False, False, False, False, False, False, False, False,
                                             False, False)  # промежуточные итоги и фильтры
.PivotFields("Откл2 таб.").Subtotals = Array(False, False, False, False, False, False, False, False, False, False,
                                             False, False)  # промежуточные итоги и фильтры

.ManualUpdate = False  # обновить сводную

.TableStyle2 = ""  # стиль
.ColumnRange.ColumnWidth = 15  # ширина строк
.RowRange.ColumnWidth = 10
.RowRange.Columns(1).ColumnWidth = 50
.RowRange.Columns(1).ColumnWidth = 25

.DataBodyRange.HorizontalAlignment = -4108  # xlCenter = -4108
.DataBodyRange.NumberFormat = "0"

With.TableRange1  # формат
.WrapText = True  # перенос текста в ячейке
.Borders(7).LineStyle = 1  # лево
.Borders(8).LineStyle = 1  # верх
.Borders(9).LineStyle = 1  # низ
.Borders(10).LineStyle = 1  # право
.Borders(11).LineStyle = 1  # внутри вертикаль
.Borders(12).LineStyle = 1  #
End
With
End
With


def zam_range(diapazon, iz, na):  # заменит в дивапазоне значение 1 на 2
    # dim rng
    for Each rng in diapazon  # цикл по видимым ячейкам столбца
        if rng.value = iz or isempty (rng.value): rng.value = na


def uslovnoeFc(dpz):  # выделить цветом в зависимости от загрузки
    dpz.FormatConditions.AddColorScale
    2  # ColorScaleType:=2
    dpz.FormatConditions(dpz.FormatConditions.count).SetFirstPriority
    dpz.FormatConditions(1).ColorScaleCriteria(1).Type = 0  # xlConditionValueNumber = 0
    dpz.FormatConditions(1).ColorScaleCriteria(1).Value = 100

    With
    dpz.FormatConditions(1).ColorScaleCriteria(1).FormatColor
    .ThemeColor = 1  # xlThemeColorDark1 = 1
    .TintAndShade = 0


End
With
dpz.FormatConditions(1).ColorScaleCriteria(2).Type = 2  # xlConditionValueHighestValue = 2
With
dpz.FormatConditions(1).ColorScaleCriteria(2).FormatColor
.ThemeColor = 6  # xlThemeColorAccent2 = 6
.TintAndShade = -0.249977111117893
End
With
dpz.FormatConditions(1).ScopeType = 2  # xlDataFieldScope = 2 применить ко всем значениям поля


def uslovnoeFz(dpzn):  # выделить максимальное значение подчеркиванеим

    dpzn.FormatConditions.AddTop10
    dpzn.FormatConditions(dpzn.FormatConditions.count).SetFirstPriority
    With
    dpzn.FormatConditions(1)
    .TopBottom = 1  # xlTop10Top = 1
    .Rank = 1
    .Percent = True


End
With
With
dpzn.FormatConditions(1).Font
.Bold = True
.Italic = False
.Underline = 2  # xlUnderlineStyleSingle = 2
.TintAndShade = 0
End
With
dpzn.FormatConditions(1).StopIfTrue = False


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
class Kontrol_el:  # сводные данные для каждого контралируемого элемента, для соответствующего кол отключений
    # dim strN, kol_otkl_comb, name_el , dict_OR , dict_period , period_max , Azag_max , Dzag_max,zag_maxT, picte1 , picte2 , action, nStrDictValMax,   otklMaxName,  NO_ON  #  (истина есть, ложь нет)
    # dim File_RG2  , U_min_r, TipKontr #  node, vetv, not_mod
    def init_KE():
        NO_ON = False
        Azag_max = 0
        Dzag_max = 0
        U_min_r = 0
        period = ""
        period_max = ""
        dict_OR = createobject("Scripting.dictionary")  # (откл , ремонты)
        dict_period = createobject("Scripting.dictionary")  # (откл , ремонты)


class Komb_All_List_class:  # Komb_List. вывод комбинаций
    # dim dict , XL_comb
    Private
    n, val

    def init():  #
        Sheets_add(GLR.Book_XL, XL_comb, "Перечень отключений")
        Print_XL(XL_comb, 1, 1,
                 array("Имя файла", "N комб", "N в таблице КО", "Макс. загрузка  по I,%", "Мин. запас  по U,%",
                       "Откл. элемент"), 1, "гор", "", "", "")
        n = 1

    def add(max_zagruzka_dorgm, min_zapas_U):  # либо в ХЛ или в колекцию
        val = array(rg.Name_Base, GLR.N_rezh, rg.TabRgmCount, round(max_zagruzka_dorgm, 1), round(min_zapas_U, 1),
                    rgr.raschot_name)
        if GLR.vivod_komb = 1:  # 1 сразу принт 2 по завершению
            Print_XL(XL_comb, 1, n + 1, val, 1, "гор", "", "",
                     "")  # печать массива: лист ХL ,по X , по Y , массив , кол изм массива 1 или 2 , "гор" "верт" , "" или "vetv" ,"" или "name" , "" или "орвыаи " - произвольный текст
        elif GLR.vivod_komb = 2:
            dict.Add(n, val)  # ключ  и значение

        n = n + 1

    def print_end_KL():  # вывод dict
        if GLR.vivod_komb = 2:
            for Each n in dict.Keys  # организуем цикл по элементам  масива Keys
                Print_XL(XL_comb, 1, n + 1, dict.Item(n), 1, "гор", "", "", "")

        XL_comb.Columns("D:E").NumberFormat = "0"  #
        XL_comb.ListObjects.Add(1, XL_comb.Range(XL_comb.UsedRange.address))  # таблица
        XL_comb.Columns("A:AA").AutoFit
        XL_comb.UsedRange.columns(4).FormatConditions.Add(1, 5, "=100")  # 1больше
        XL_comb.UsedRange.columns(4).FormatConditions(
            XL_comb.UsedRange.columns(4).FormatConditions.Count).SetFirstPriority()
        XL_comb.UsedRange.columns(4).FormatConditions(1).Interior.Color = 49407
        XL_comb.UsedRange.columns(5).FormatConditions.Add(1, 6, "=0")  # 6меньше
        XL_comb.UsedRange.columns(5).FormatConditions(
            XL_comb.UsedRange.columns(5).FormatConditions.Count).SetFirstPriority()
        XL_comb.UsedRange.columns(5).FormatConditions(1).Interior.Color = 49407


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fIadtn(ndxTab1):  # возвращает Iadtn
    # dim tV18
    tV18 = rastr.Tables("vetv")
    if tV18.cols.item("i_dop_r_av").Z(ndxTab1) = 0:
        fIadtn = round(tV18.cols.item("i_dop_r").Z(ndxTab1))
    else:
        fIadtn = round(tV18.cols.item("i_dop_r_av").Z(ndxTab1))


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fIadtn_OO(ndxTab2):  # возвращает Iadtn
    # dim tV19
    tV19 = rastr.Tables("vetv")
    if RG.name_list(1) = "зим":  # zima
        if tV19.cols.item("i_adtn_zim_OO").Z(ndxTab2) > 0:
            fIadtn_OO = tV19.cols.item("i_adtn_zim_OO").Z(ndxTab2) else:
            fIadtn_OO = tV19.cols.item("IdopOO").Z(ndxTab2)
    else:  # if RG.name_list (1) = "лет": # leto
        if tV19.cols.item("i_adtn_let_OO").Z(ndxTab2) > 0:
            fIadtn_OO = tV19.cols.item("i_adtn_let_OO").Z(ndxTab2) else:
            fIadtn_OO = tV19.cols.item("IdopOO").Z(ndxTab2)


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def f_I_max_grouid(i):  # функция возвращает ndx ветви c максимальной токовой загрузкой по заданному участку ndx ветви
    # dim i_dop_ob , i_dop , i_dop_ob_av , i_dop_av ,vibor_N_VetvGrp , ii
    # dim tV20
    tV20 = rastr.Tables("vetv")
    f_I_max_grouid = i
    if tV20.cols.item("groupid").Z(i) > 0:  # если задан groupid
        if tV20.cols.item("i_dop_ob").Z(i):
            i_dop_ob = tV20.cols.item("i_dop_ob").ZS(i)     else:
            i_dop_ob = "0"
        if tV20.cols.item("i_dop").Z(i):
            i_dop = tV20.cols.item("i_dop").ZS(i)        else:
            i_dop = "0"
        if tV20.cols.item("i_dop_ob_av").Z(i):
            i_dop_ob_av = tV20.cols.item("i_dop_ob_av").ZS(i)  else:
            i_dop_ob_av = "0"
        if tV20.cols.item("i_dop_av").Z(i):
            i_dop_av = tV20.cols.item("i_dop_av").ZS(i)     else:
            i_dop_av = "0"
        # формируем выборку
        if GLR.dtn_uchastki = 1:   vibor_N_VetvGrp = "groupid=" + str(
            tV20.cols.item("groupid").Z(i)) + "+i_dop=" + i_dop + "+i_dop_ob=" + i_dop_ob + "+n_it=" + tV20.cols.item(
            "n_it").ZS(i)
        _
        +  "+i_dop_av=" + i_dop_av + "+i_dop_ob_av=" + i_dop_ob_av + "+n_it_av=" + tV20.cols.item("n_it_av").ZS(i)
    if GLR.dtn_uchastki = 0:   vibor_N_VetvGrp = "groupid=" + str(tV20.cols.item("groupid").Z(i))

    tV20.setsel(vibor_N_VetvGrp)
    ii = tV20.FindNextSel(-1)
    while ii >= 0
        if tV20.cols.item("i_max").Z(ii) > tV20.cols.item("i_max").Z(f_I_max_grouid) and Replace(
            tV20.cols.item("dname").ZS(ii), " ",
            "") = Replace(tV20.cols.item("dname").ZS(ii), " ", ""): f_I_max_grouid = ii
        ii = tV20.FindNextSel(ii)
    wend


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def DoRgm():
    # dim itr_dop , ADD_REZHIM_tab ,  max_zagruzka_dorgm , name_zagruzka_ris , ndxTab  ,  max_zag_ddtn , max_zag_adtn ,I_maxx , ndx_OO , max_zagruzka_ndxTab_OO
    # dim ndxLost , pl , ql , ndxNode ,  n , TOK , I_OO ,  vibor_tab , ndxOO , formula_1 ,  V , Vn , Rto, flag_izm , YV , a , b  , umin_t
    # dim autoNDX , autoNDX_listV , autoNDX_listN ,  min_zapas_U , zapas_U , ndx_gr_Imax, tV21  , tN21
    tN21 = rastr.Tables("node")
    tV21 = rastr.Tables("vetv")
    max_zagruzka_dorgm = 0
    min_zapas_U = 0
    name_zagruzka_ris = ""  # перегружаемые элементы для рисунков
    sel0()

    autoNDX_listV = ""  # ndx через \ ветвей с перегрузками и автоматикой
    autoNDX_listN = ""  # ndx через \ ветвей с перегрузками и автоматикой

    if RGR.FLAG_ris_tabl_add_PA = 1: ADD_REZHIM_tab = 1    else:
        ADD_REZHIM_tab = 0  # 0  добавляеть сталбец в таблицу, если меняет на 1 то добавлять

    GS.kod_rgm = rastr.rgm("")
    if GS.kod_rgm = 1:
        GS.kod_rgm = rastr.rgm("")
    if GS.kod_rgm = 1:
        GS.kod_rgm = rastr.rgm("p")
    if GS.kod_rgm = 1:
        GS.kod_rgm = rastr.rgm("p")
    if GS.kod_rgm = 1:
        GS.kod_rgm = rastr.rgm("p")

    if GS.kod_rgm = 1:  # если режим не моделируется kod_rgm = 1 , существует 0
        if GLR.protokol_XL: GLR.overload.ndxN = -1: GLR.overload.ndxV = -1: GLR.overload.add()
        max_zagruzka_dorgm = -1  # проверка текущего расчета на максимальную загрузу
        min_zapas_U = -1
    else:  # если режим существует
        max_zagruzka_dorgm = 0  # проверка текущего расчета на максимальную загрузу

        if tV21.cols.Find("Kontrol_all") < 1: logging.info("пропало поле Kontrol_all")
        if RG.temp_a_v_gost and GLR.kol_otkl = 2:    tV21.setsel("Kontrol_all+tip=0+i_zag_av>0.1")    else:
            tV21.setsel("Kontrol_all+tip=0+i_zag>0.1")

        if tV21.count > 0 or GLR.max_tok_save:  # 1
            tV21.setsel("")
            for i = LBound(RG.KontrolVL) to UBound(RG.KontrolVL)  # цикл по КОНТРОЛ ВЛ
            ndxTab = RG.KontrolVL(i)
            if tV21.cols.item("groupid").Z(ndxTab) > 0:
                ndx_gr_Imax = f_I_max_grouid(ndxTab)  else:
                ndx_gr_Imax = ndxTab  # находим участок с макс током

            I_maxx = ABS(tV21.cols.item("Imax").ZN(ndx_gr_Imax))  # макс ток по контрол
            max_zag_ddtn = tV21.cols.item("i_zag").ZN(ndx_gr_Imax)
            if tV21.cols.item("i_zag_av").ZN(ndx_gr_Imax) > 0:
                max_zag_adtn = tV21.cols.item("i_zag_av").ZN(ndx_gr_Imax) else:
                max_zag_adtn = max_zag_ddtn

            if GLR.max_tok_save: GLR.max_tok1.test_max_tok(ndx_gr_Imax, RG.NAME_RG2_plus2, round(I_maxx), RG.god,
                                                           RGR.raschot_name, GLR.kol_otkl)

            if max_zagruzka_dorgm < max_zag_ddtn: max_zagruzka_dorgm = max_zag_ddtn

            if max_zag_ddtn > 100:  # есть перегрузка ддтн
                if not (RG.temp_a_v_gost and GLR.kol_otkl = 2 and max_zag_adtn < 100):  # учет по гост в н-2 адтн

                    if tV21.cols.item("autosta").Z(
                            ndxTab) = 0 and StrComp(Replace(tV21.cols.item("automatika").Z(ndxTab), " ", ""), "") != 0:  # АВТОМАТИКА
                        autoNDX_listV = autoNDX_listV + str(ndxTab) + "\"    #
                        RGR.autoNDX_listV_info = RGR.autoNDX_listV_info + " " + tV21.cols.item("dname").ZS(
                            ndxTab) + " " + tV21.cols.item("automatika").ZS(ndxTab) + " \ "  #

                    if GLR.protokol_XL: GLR.overload.Kontrol_Key = rastr.Tables("vetv").SelString(
                        ndxTab): GLR.overload.ndxV = ndx_gr_Imax: GLR.overload.Ir = I_maxx
                    if GLR.protokol_XL: GLR.overload.Idz = max_zag_ddtn: GLR.overload.add()
                    if max_zag_ddtn > GLR.zagruz_add_tab: ADD_REZHIM_tab = 1
                    if max_zag_ddtn > 100 and GLR.picture_add: name_zagruzka_ris = f_name_zagruzka_ris(
                        name_zagruzka_ris, tV21.cols.item("dname").ZS(ndxTab), max_zag_ddtn,
                        tV21.cols.item("i_max").ZS(ndx_gr_Imax), tV21.cols.item("i_dop_r").ZS(ndxTab), fIadtn(ndxTab),
                        "")

    if RG.temp_a_v_gost and GLR.kol_otkl = 2:    tV21.setsel("Kontrol_all+tip=1+i_zag_av>0.1")   else:
        tV21.setsel("Kontrol_all+tip=1+i_zag>0.1")  # вл

    if tV21.count > 0 or GLR.max_tok_save:
        tV21.setsel("")  #
        for i = LBound(RG.KontrolTrans) to UBound(RG.KontrolTrans)  # цикл по КОНТРОЛ ТРАНС
        ndxTab = RG.KontrolTrans(i)
        max_zagruzka_ndxTab_OO = 0

        itr_dop = tV21.cols.item("i_dop_r").ZN(ndxTab)

        I_maxx = ABS(tV21.cols.item("Imax").ZN(ndxTab))  # макс ток по контрол
        if GLR.max_tok_save: GLR.max_tok1.test_max_tok(ndxTab, RG.NAME_RG2_plus2, round(I_maxx), RG.god,
                                                       RGR.raschot_name, GLR.kol_otkl)

        max_zag_ddtn = tV21.cols.item("i_zag").ZN(ndxTab)
        if tV21.cols.item("i_zag_av").ZN(ndxTab) > 0:
            max_zag_adtn = tV21.cols.item("i_zag_av").ZN(ndxTab) else:
            max_zag_adtn = max_zag_ddtn

        if tV21.cols.item("KontrOO").Z(ndxTab):
            I_OO = fIOO(ndxTab)
            # if GLR.max_tok_save: GLR.max_tok1.REC1  ("ТРАНС_OO" , i , round (I_maxx ) , RG.NAME_RG2_plus2 , RGR.raschot_name) # запись в макс токов
            max_zagruzka_ndxTab_OO = I_OO / tV21.cols.item("IdopOO").Z(ndxTab) * 100  # макс загрузка текущей OO

        if (max_zag_ddtn > 100 or max_zagruzka_ndxTab_OO > 100) and tV21.cols.item("autosta").Z(
                ndxTab) = 0 and StrComp(Replace(tV21.cols.item("automatika").Z(ndxTab), " ", ""), "") != 0:  # АВТОМАТИКА
            autoNDX_listV = autoNDX_listV + str(ndxTab) + "\"
            RGR.autoNDX_listV_info = RGR.autoNDX_listV_info + " " + tV21.cols.item("dname").ZS(
                ndxTab) + " " + tV21.cols.item("automatika").ZS(ndxTab) + " \ "  #

        if max_zagruzka_dorgm < max_zag_ddtn: max_zagruzka_dorgm = max_zag_ddtn
        if max_zagruzka_dorgm < max_zagruzka_ndxTab_OO: max_zagruzka_dorgm = max_zagruzka_ndxTab_OO

        if max_zag_ddtn > GLR.zagruz_add_tab Or max_zagruzka_ndxTab_OO > GLR.zagruz_add_tab: ADD_REZHIM_tab = 1

        if max_zag_ddtn > 100:  # последовательная обмотка АТ  или тр
            if not (RG.temp_a_v_gost and GLR.kol_otkl = 2 and max_zag_adtn < 100):  # учет по гост в н-2 адтн
                if GLR.protokol_XL: GLR.overload.Kontrol_Key = rastr.Tables("vetv").SelString(
                    ndxTab): GLR.overload.ndxV = ndxTab: GLR.overload.Ir = I_maxx
                if GLR.protokol_XL: GLR.overload.Idz = max_zag_ddtn: GLR.overload.add()
                if max_zag_ddtn > GLR.zagruz_add_tab: ADD_REZHIM_tab = 1
                if max_zag_ddtn > 100 and GLR.picture_add: name_zagruzka_ris = f_name_zagruzka_ris(name_zagruzka_ris,
                                                                                                   tV21.cols.item(
                                                                                                       "dname").ZS(
                                                                                                       ndxTab),
                                                                                                   max_zag_ddtn,
                                                                                                   tV21.cols.item(
                                                                                                       "i_max").ZS(
                                                                                                       ndxTab),
                                                                                                   tV21.cols.item(
                                                                                                       "i_dop_r").ZS(
                                                                                                       ndxTab),
                                                                                                   fIadtn(ndxTab), "")
        if max_zagruzka_ndxTab_OO > 100:  # общая обмотка АТ
            if not (RG.temp_a_v_gost and GLR.kol_otkl = 2 and max_zag_adtn < 100):  # учет по гост в н-2 адтн
                if GLR.protokol_XL: GLR.overload.Kontrol_Key = rastr.Tables("vetv").SelString(
                    ndxTab) + "ОО": GLR.overload.ndxV = ndxTab: GLR.overload.Ir = I_OO: GLR.overload.OO = 1
                if GLR.protokol_XL: GLR.overload.Idz = max_zagruzka_ndxTab_OO: GLR.overload.add()
                if max_zagruzka_ndxTab_OO > GLR.zagruz_add_tab: ADD_REZHIM_tab = 1
                if max_zagruzka_ndxTab_OO > 100 and GLR.picture_add: name_zagruzka_ris = f_name_zagruzka_ris(
                    name_zagruzka_ris, tV21.cols.item("dname").ZS(ndxTab), max_zagruzka_ndxTab_OO, I_OO,
                    tV21.cols.item("IdopOO").Z(ndxTab), fIadtn_OO(ndxTab), " OO ")

                #   цикл по КОНТРОЛ НАПРЯЖЕНИЕ


vibor_tab = "Kontrol+(otv_min<0|((umax-vras)<0+umax>0))+!sta": umin_t = "umin"

if GLR.gost58670:
    if RG.temp_a_v_gost and GLR.kol_otkl = 2: vibor_tab = "Kontrol+(otv_min_av<0|((umax-vras)<0+umax>0))+!sta": umin_t = "umin_av"
    if not RG.temp_a_v_gost and GLR.kol_otkl = 3: vibor_tab = "Kontrol+(otv_min_av<0|((umax-vras)<0+umax>0))+!sta": umin_t = "umin_av"

tN21.setsel(vibor_tab)  # otv_min мое поле
if tN21.count > 0:
    ADD_REZHIM_tab = 1
    ndxTab = tN21.FindNextSel(-1)
    while ndxTab >= 0  #
        if tN21.cols.item("vras").Z(ndxTab) < tN21.cols.item(umin_t).Z(ndxTab):  # напряжение ниже допустимого
            if GLR.picture_add:
                if not name_zagruzka_ris = "": name_zagruzka_ris = name_zagruzka_ris + ". "
                if umin_t = "umin": name_zagruzka_ris = name_zagruzka_ris + "Напряжения на " + tN21.cols.item(
                    "dname").ZS(ndxTab) + " составляет " + Round(tN21.cols.item("vras").ZS(ndxTab),
                                                                 0) + " кВ, при минимально допустимом " + tN21.cols.item(
                    umin_t).ZS(ndxTab) + " кВ"
                if umin_t = "umin_av": name_zagruzka_ris = name_zagruzka_ris + "Напряжения на " + tN21.cols.item(
                    "dname").ZS(ndxTab) + " составляет " + Round(tN21.cols.item("vras").ZS(ndxTab),
                                                                 0) + " кВ, при аварийно допустимом " + tN21.cols.item(
                    umin_t).ZS(ndxTab) + " кВ"

            if tN21.cols.item("autosta").Z(
                    ndxTab) = 0 and StrComp(Replace(tN21.cols.item("automatika").Z(ndxTab), " ", ""), "") != 0:  # АВТОМАТИКА# StrComp сравнение строк
                autoNDX_listN = autoNDX_listN + str(ndxTab) + "\"
                RGR.autoNDX_listN_info = RGR.autoNDX_listN_info + " " + tN21.cols.item("dname").ZS(
                    ndxTab) + " " + tN21.cols.item("automatika").ZS(ndxTab) + " \ "

            if GLR.protokol_XL: GLR.overload.Kontrol_Key = rastr.Tables("node").SelString(
                ndxTab): GLR.overload.ndxN = ndxTab: GLR.overload.add()

        elif tN21.cols.item("vras").Z(ndxTab) > tN21.cols.item("umax").Z(ndxTab) and tN21.cols.item("umax").Z(
                ndxTab) > 0:  # напряжение выше допустимого
            if GLR.picture_add:
                if not name_zagruzka_ris = "": name_zagruzka_ris = name_zagruzka_ris + ". "
                name_zagruzka_ris = name_zagruzka_ris + "Напряжения на " + tN21.cols.item("dname").ZS(
                    ndxTab) + " составляет " + str(
                    Round(tN21.cols.item("vras").Z(ndxTab), 0)) + " кВ, при допустимом " + tN21.cols.item("umax").ZS(
                    ndxTab) + " кВ"

            if tN21.cols.item("autosta").Z(
                    ndxTab) = 0 and StrComp(Replace(tN21.cols.item("automatika").Z(ndxTab), " ", ""), "") != 0:  # АВТОМАТИКА# StrComp сравнение строк
                autoNDX_listN = autoNDX_listN + str(ndxTab) + "\"
                RGR.autoNDX_listN_info = RGR.autoNDX_listN_info + " " + tN21.cols.item("dname").ZS(
                    ndxTab) + " " + tN21.cols.item("automatika").ZS(ndxTab) + " \ "

            if GLR.protokol_XL: GLR.overload.Kontrol_Key = rastr.Tables("node").SelString(
                ndxTab): GLR.overload.ndxN = ndxTab: GLR.overload.add()

        ndxTab = tN21.FindNextSel(ndxTab)
    wend

if GLR.FLAG_automatika:  # АВТОМАТИКА  АВТОМАТИКА  АВТОМАТИКА  АВТОМАТИКА
    autoNDX_listV_masiv = split(autoNDX_listV, "\")
    autoNDX_listN_masiv = split(autoNDX_listN, "\")

    if autoNDX_listV = "": NN_V = 0 else: NN_V = ubound(autoNDX_listV_masiv)
    if autoNDX_listN = "": NN_N = 0 else: NN_N = ubound(autoNDX_listN_masiv)

    NN_PA = NN_V + NN_N
    if NN_PA > 0:
        RGR.add_rgm_PA = True
    redim
    autoNDX(1, NN_PA - 1)

    for i = 0  to (NN_PA - 1)
    if i < ubound (autoNDX_listV_masiv):
        autoNDX(0, i) = "vetv"
    autoNDX(1, i) = float(autoNDX_listV_masiv(i))
    else:
    autoNDX(0, i) = "node"
    autoNDX(1, i) = float(autoNDX_listN_masiv(i - NN_V))

    if ubound(autoNDX, 2) > 0:  # сортировка ПА  по N
        for
    i = 1
    to
    ubound(autoNDX, 2)  # + 1#  цикл сорт
    key1 = autoNDX(0, i)
    key2 = autoNDX(1, i)
    j = i - 1
    do
    while j >= 0
        if fParam(autoNDX(0, j), "autoN", autoNDX(1, j)) < fParam(key1, "autoN", key2): Exit
        Do
        autoNDX(0, j + 1) = autoNDX(0, j)
        autoNDX(1, j + 1) = autoNDX(1, j)
        j = j - 1
    loop
    autoNDX(0, j + 1) = key1
    autoNDX(1, j + 1) = key2

if autoNDX(0, 0) = "node":
    RGR.AutoZad = tN21.cols.item("automatika").ZS(autoNDX(1, 0))
    tN21.cols.item("autosta").Z(autoNDX(1, 0)) = 1
elif autoNDX(0, 0) = "vetv":
    RGR.AutoZad = tV21.cols.item("automatika").ZS(autoNDX(1, 0))
    tV21.cols.item("autosta").Z(autoNDX(1, 0)) = 1

RGR.AutoKontrol(0) = autoNDX(0, 0)
RGR.AutoKontrol(1) = autoNDX(1, 0)
if max_zagruzka_dorgm = 0: max_zagruzka_dorgm = rastr.Calc("max", "vetv", "i_zag*1000", "Kontrol_all")
if abs(max_zagruzka_dorgm) > 1000000: max_zagruzka_dorgm = 0
if min_zapas_U != -1: min_zapas_U = rastr.Calc("max", "node", "otv_min",
                                               "Kontrol")  # (U-Umin)/Umin*100 те >0 есть запас, 0 < нет запаса
if abs(min_zapas_U) > 1000000: min_zapas_U = 0
if GLR.vivod_komb > 0: Komb_List.add(max_zagruzka_dorgm, min_zapas_U)

if RGR.raschot_name = "Нормальный режим": ADD_REZHIM_tab = 1
#  Вывод РИСУНКОВ
RGR.add_risunok = False  # инициализация добавления РИСУНКА
if GLR.picture_add:
    if GLR.risunok_nr  And RGR.raschot_name = "Нормальный режим":       RGR.add_risunok = True  # все нр
    if GLR.risunok_par And RGR.raschot_name != "Нормальный режим":       RGR.add_risunok = True  # все пар
    if GLR.risunok_zag and (max_zagruzka_dorgm > 100 or min_zapas_U < 0):       RGR.add_risunok = True  #
    if GLR.risunok_zag and RGR.FLAG_ris_tabl_add_PA = 1:       RGR.add_risunok = True  #

if (ADD_REZHIM_tab
= 1 Or GLR.zagruz_add_tab = 0) and GLR.Tabl_otlk_kontrol > 0: add_tabl_KontrOtkl()  # ЗАПОЛНЕНИЕ ТАБЛИЦЫ контроль - отключение

if RGR.add_risunok and GLR.picture_add:
    if name_zagruzka_ris != "": name_zagruzka_ris = name_zagruzka_ris + ". "
    RGR.name_ris_info = array("рис", "номер", "сезон год", "доп имя", "имя кадр", "нр/откл+действие", "загрузка")
    RGR.name_ris = array(GLR.name_ris1, GLR.number_pict, RG.SezonName + " " + RG.god + " г. ", RG.txt_dop, "",
                         RGR.raschot_name + ". ", name_zagruzka_ris)
    RisunokPrint()  # вставить рисунок в ворд/ сохранить rg2

RGR.autoTXT_fPA = ""
GS.AutoShunt_list = ""


# end def #
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def kontrol_groupid():  # отметить контрол ветви если каки то другие участки эотй группы отмечены groupid
    # dim ndx , tVT42 , tV22
    tV22 = rastr.Tables("vetv")
    tVT42 = rastr.Tables("vetv")
    tVT42.setsel("Kontrol")
    ndx = tVT42.FindNextSel(-1)
    while ndx >= 0  #
        if tVT42.count > 0:
            if tVT42.cols.item("groupid").Z(ndx) > 0:
                tV22.setsel("groupid=" + tVT42.cols.item("groupid").ZS(ndx))
                tV22.cols.item("Kontrol").calc(1)
                ndx = tVT42.FindNextSel(ndx)
    wend


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def add_tabl_KontrOtkl():
    # dim YV, ndx
    # dim tV23  , tN23
    tV23 = rastr.Tables("vetv")
    tN23 = rastr.Tables("node")
    GLR.XL_sheet.cell(GLR.Y_list - 3, GLR.X_list).Value = RGR.raschot_name
    GLR.XL_sheet.Columns(GLR.X_list).ColumnWidth = 14  # ширина столбца
    GLR.XL_sheet.Columns(GLR.X_list + 1).ColumnWidth = 5  # ширина столбца
    GLR.XL_sheet.cell(GLR.Y_list - 3 + 1, GLR.X_list).Value = RG.TabRgmCount
    GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list - 3 + 1, GLR.X_list),
                       GLR.XL_sheet.cell(GLR.Y_list - 3 + 1, GLR.X_list + 1)).Merge
    GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list - 3, GLR.X_list),
                       GLR.XL_sheet.cell(GLR.Y_list - 3 - 1, GLR.X_list + 1)).Merge

    RG.TabRgmCount = RG.TabRgmCount + 1
    GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 3, GLR.X_list).Value = RGR.NodeNetReserv
    GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 4, GLR.X_list).Value = RGR.NodeRezerv
    GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 5, GLR.X_list).Value = GS.AutoShunt_list

    if autoNDX_listV != "": GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 6,
                                               GLR.X_list).Value = "ПА есть вет:" + RGR.autoNDX_listV_info
    if autoNDX_listN != "": GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 7,
                                               GLR.X_list).Value = "ПА есть узел:" + RGR.autoNDX_listN_info

    if GS.kod_rgm =1:  # записывает расчетные значения в таблицу
        GLR.XL_sheet.cell(GLR.Y_list, GLR.X_list).Value = "Режим не моделируется!"
        With
        GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list, GLR.X_list),
                           GLR.XL_sheet.cell(GLR.Y_list + 1, GLR.X_list + 1))
        .Merge
        .WrapText = True  # перенос текста в ячейке
    End
    With
    GLR.X_list = GLR.X_list + 2

else:
if GLR.Tabl_otlk_kontrol = 1:
    GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list, 1), GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 1, 2)).Copy
    GLR.XL_sheet.cell(GLR.Y_list, GLR.X_list).PasteSpecial()
    GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list, GLR.X_list),
                       GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 1, GLR.X_list)).Clearcontents
    n = 0
    for i = LBound(RG.KontrolVL) To UBound(RG.KontrolVL)  # цикл записи расчетных значений-ветви VL_ VL_ VL_ VL_ VL_VL_ VL_ VL_ VL_ VL_
    ndx = RG.KontrolVL(i)
    if tV23.cols.item("sta").Z(ndx):
    else:
        if tV23.cols.item("groupid").Z(ndx):
            N_VetvGrp = tV23.cols.item("groupid").Z(ndx)  # присваиваем номер группы

            if tV23.cols.item("i_dop_ob").Z(ndx):
                i_dop_ob = tV23.cols.item("i_dop_ob").ZS(ndx) else:
                i_dop_ob = "0"
            if tV23.cols.item("i_dop").Z(ndx):
                i_dop = tV23.cols.item("i_dop").ZS(ndx) else:
                i_dop = "0"

            if GLR.dtn_uchastki = 1:
                vibor_N_VetvGrp = "groupid=" + str(N_VetvGrp) + "+i_dop=" + i_dop + "+i_dop_ob=" + str(
                    i_dop_ob) + "+n_it=" + tV23.cols.item("n_it").ZS(ndx)
                dname_vetv = Trim(tV23.cols.item("dname").ZS(ndx))
            elif GLR.dtn_uchastki = 0:
                vibor_N_VetvGrp = "groupid=" + str(N_VetvGrp)

            tV23.setsel(vibor_N_VetvGrp)
            ndx_G_id = tV23.FindNextSel(-1)
            MAX_G_id = ndx_G_id  # номер строки с макс током в текущей группе

            While
            ndx_G_id >= 0
            if tV23.cols.item("i_max").Z(ndx_G_id) > tV23.cols.item("i_max").Z(
                MAX_G_id) And dname_vetv = Trim(tV23.cols.item("dname").ZS(ndx_G_id)): MAX_G_id = ndx_G_id
            ndx_G_id = tV23.FindNextSel(ndx_G_id)
        wend

        if tV23.cols.item("znak-").Z(ndx):
            GLR.XL_sheet.cell(GLR.Y_list + i + n, GLR.X_list).Value = tV23.cols.item("-Smax").ZS(
                MAX_G_id)  # запись мощность-
            TOK = -tV23.cols.item("Imax").Z(MAX_G_id)
            GLR.XL_sheet.cell(GLR.Y_list + i + 1 + n, GLR.X_list).Value = TOK  # запись ток-
        else:
            GLR.XL_sheet.cell(GLR.Y_list + i + n, GLR.X_list).Value = tV23.cols.item("Smax").ZS(
                MAX_G_id)  # запись мощность
            GLR.XL_sheet.cell(GLR.Y_list + i + 1 + n, GLR.X_list).Value = tV23.cols.item("Imax").Z(
                MAX_G_id)  # запись ток

    else:

    if tV23.cols.item("znak-").Z(ndx):
        GLR.XL_sheet.cell(GLR.Y_list + i + n, GLR.X_list).Value = tV23.cols.item("-Smax").ZS(ndx)  # запись мощность-
        TOK = -tV23.cols.item("Imax").Z(ndx)
        GLR.XL_sheet.cell(GLR.Y_list + i + 1 + n, GLR.X_list).Value = TOK  # запись ток-
    else:
        GLR.XL_sheet.cell(GLR.Y_list + i + n, GLR.X_list).Value = tV23.cols.item("Smax").ZS(ndx)  # запись мощность
        GLR.XL_sheet.cell(GLR.Y_list + i + 1 + n, GLR.X_list).Value = tV23.cols.item("Imax").Z(ndx)  # запись ток
n = n + 1

n = 0

# Kol_Tr_OO = 0
for i = LBound(RG.KontrolTrans) To UBound(RG.KontrolTrans)  # цикл записи расчетных значений-ветви Trans_Trans_Trans_Trans_Trans_Trans_Trans_Trans_
ndx = RG.KontrolTrans(i)

GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n, GLR.X_list).Value = tV23.cols.item("Smax").ZS(
    ndx)  # запись мощность
GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n + 1, GLR.X_list).Value = tV23.cols.item("Imax").Z(
    ndx)  # запись ток                if tV23.cols.item("KontrOO").Z(ndx) = True:   #  истина если контроль ОО АТ
GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL + i + 1 + n + 2, GLR.X_list).Value = fIOO(ndx)  # расчетное
n = n + 1

n = n + 1

# условное форматирование
GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list, GLR.X_list + 1),
                   GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans, GLR.X_list + 1)).FormatConditions.Add(1, 5, "=100")
With
GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list, GLR.X_list + 1),
                   GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans, GLR.X_list + 1)).FormatConditions(1).Interior
.Color = 49407
End
With

for i = LBound(RG.KontrolNode) To UBound(RG.KontrolNode)  # ЦИКЛ по УЗЛАМ ЦИКЛ по УЗЛАМ ЦИКЛ по УЗЛАМ ЦИКЛ по УЗЛАМ ЦИКЛ по УЗЛАМ ЦИКЛ по УЗЛАМ ЦИКЛ по УЗЛАМ ЦИКЛ по УЗЛАМ ЦИКЛ по УЗЛАМ ЦИКЛ по УЗЛАМ
ndx = RG.KontrolNode(i)

V = tN23.cols.item("vras").Z(ndx)  # vras
Vn = tN23.cols.item("uhom").Z(ndx)

if Vn > 90: GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + i + 2, GLR.X_list).Value = Round(V,
                                                                                              0)  # Round - точность после запятой
if Vn < 90: GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + i + 2, GLR.X_list).Value = Round(V,
                                                                                              1)  # Round - точность после запятой

# условное форматирование
if GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + i + 2,
                      13).Value > 0 And GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + i + 2, GLR.X_list).Value > 0:
    address_dop_U = GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + i + 2, 13).Address(False, False)
    GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + i + 2, GLR.X_list).FormatConditions.Add(1, 6, "=" + address_dop_U,
                                                                                             "=1")
    GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans + i + 2, GLR.X_list).FormatConditions(1).Interior.Color = 49407

elif GLR.Tabl_otlk_kontrol = 2:  # RTD RTD RTD RTD RTD RTD RTD RTD RTD RTD RTD RTD RTD RTD RTD RTD RTD RTD RTD RTD RTD RTD
CreateObject("Rtdserver.AstraServer.1").ServeRTD(rastr, "$1")

for i = 1 To 500
flag_izm = 0  # RTD сработало
YV = GLR.Y_list + GLR.Y_VL_Trans_V

for p = GLR.Y_list To YV
a = GLR.XL_sheet.cell(p, 1).Value
b = GLR.XL_sheet.cell(p, GLR.X_list - 2).Value
# logging.info ( a): logging.info ( b)
if a = b:  # and a = ""    and VarType(a) = 8
else:
    flag_izm = 1
    Exit
    for

if flag_izm = 1: Exit
for
    if i = 500: MsgBox("не обновляется!!!")

GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list, 1), GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 1, 2)).Copy
GLR.XL_sheet.cell(GLR.Y_list, GLR.X_list).PasteSpecial()
GLR.XL_sheet.Range(GLR.XL_sheet.cell(GLR.Y_list, 1), GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 1, 1)).Copy
GLR.XL_sheet.cell(GLR.Y_list, GLR.X_list).PasteSpecial(-4163, -4142)

if RGR.add_risunok: GLR.XL_sheet.cell(GLR.Y_list + GLR.Y_VL_Trans_V + 2, GLR.X_list).Value = GLR.name_ris1 + str(
    GLR.number_pict)
GLR.X_list = GLR.X_list + 2


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def f_txt_name_good(txt):  # привести текст к доп умени файла
    f_txt_name_good = trim(
        txt)  # LTrim(), RTrim(), Trim() - возможность убрать пробелы соответственно слева, справа или и слева, и справа.
    f_txt_name_good = replace(f_txt_name_good, "<",
                              "_")  # Replace()— возможность заменить в строке одну последовательность символов на другую.
    f_txt_name_good = replace(f_txt_name_good, ">", "_")
    f_txt_name_good = replace(f_txt_name_good, ":", "_")
    f_txt_name_good = replace(f_txt_name_good, Chr(34), "_")
    f_txt_name_good = replace(f_txt_name_good, "/", "_")
    f_txt_name_good = replace(f_txt_name_good, "\","
    _
    ")
    f_txt_name_good = replace(f_txt_name_good, "|", "_")
    f_txt_name_good = replace(f_txt_name_good, "?", "_")
    f_txt_name_good = replace(f_txt_name_good, "*", "_")
    # f_txt_name_good = replace ( f_txt_name_good,".","_")
    f_txt_name_good = Left(f_txt_name_good, 250)
    # End def return
    # *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************


def rg2_ris(tip_zad_ris):  # 4 замена рисунков в ворд из rg2 # 5  замена рисунков в ворд из ворд
    # dim Folder_rg2_ris, text_color, File_ris_in, File_ris_iz, kluch_ris
    # ---4  замена рисунков в ворд из rg2 ///  можно отметить graf_load = 1 и КОНТРОЛЬ для загрузки, нр Zad_RG2
    # ----5  замена рисунков в ворд в ворд
    Folder_rg2_ris = IzFolder + "\temp\ris_rg2 !"  # 4
    text_color = 1  # (выделить красным)              # 4,5
    File_ris_in = IzFolder + "\temp\Рисунок 2.docx"  # 4,5
    File_ris_iz = IzFolder + "\temp\Рисунок 3.docx"  # 5
    kluch_ris = 0  # 0 по номеру рисунка 1 по описанию#   5


if tip_zad_ris = 4:  # 4 - замена рисунков в ворд из rg2
    GLR.risunok_word = True  # ФОРМИРОВАНИЕ РИСУНКОВ  ПРОВЕРИТЬ ?????????????????
    GLR.risunok_nr = True  # 1 рисунки нормальных режимов
    GLR.risunok_par = False  #

    GLR.word_App = CreateObject("word.Application")
    GLR.word_App.Visible = True
    GLR.word_ris_in = GLR.word_App.Documents.Open(GLR.File_ris_in)  # открыть сущ док
    for Each objFile in objFSO.GetFolder(GLR.Folder_rg2_ris).Files  # цикл по файлам в  указанной папке

        RG = CurrentFile
        RG.file_path = objFile.Path
        RG.initRG(0, array("", "", "", ""))  # разбирает file_path и тд
        if not RG.Name_st = "не подходит" and objFile.type = "Файл режима rg2":
            rastr.Load(1, objFile.Path, RG.shablon)  # загрузить режим
            rastr.Tables("log_otkl").Size = 0
            if GLR.graf_load = 1: GrfLoad()
            logging.info(RG.Name_Base)
            GS.kod_rgm = rastr.rgm("")
            if GS.kod_rgm = 1:
                GS.kod_rgm = rastr.rgm("")
            if GS.kod_rgm = 1:
                GS.kod_rgm = rastr.rgm("p")
            if GS.kod_rgm = 1:
                GS.kod_rgm = rastr.rgm("p")
            if GS.kod_rgm = 1:
                GS.kod_rgm = rastr.rgm("p")

            VIBOR_KONTROL_OTKL()  # процедура для отметки узлов и ветвей КОНТРОЛЬ и ОТКЛ

            if GLR.AutoShunt:
                AutoShunt_class_rec("")  # процедура формирует Umin , Umax, AutoBsh , nBsh
                AutoShunt_class_kor()  # процедура меняет Bsh  и записывает GS.AutoShunt_list
                GS.AutoShunt_list = ""

            rastr.CalcIdop(RG.GradusZ, float(0), ""): logging.info(
                "\t" + "расчетная температура(mainR): " + str(RG.GradusZ))
            RG.Kontrol_init()  # формирует RG.KontrolVL , RG.KontrolTrans , RG.KontrolNode

            RGR = raschot_tek_comb  # НР
            RGR.init_new()  #
            DoRgm()  # формирует name_wmf
            #  найти номер рисунка из RG2
            NN_ris_rg2_1 = split(RG.Name_Base, ")")
            NN_ris_rg2_2 = split(NN_ris_rg2_1(0), " ")
            NN_ris_rg2_3 = split(NN_ris_rg2_2(ubound(NN_ris_rg2_2)), ".")
            NN_ris_rg2 = NN_ris_rg2_3(ubound(NN_ris_rg2_3))
            logging.info("\t" + "заменить рисунок " + str(NN_ris_rg2))
            for Each control in GLR.word_ris_in.ContentControls  # определить номер группы
                if control.Range.Text = NN_ris_rg2:
                    kluch_Tag = control.Tag
                    exit
                    for

            for Each control in GLR.word_ris_in.ContentControls  # поменять рисунок
                if control.Tag = kluch_Tag and control.Title = "рисунок":
                    control.Range.InlineShapes(1).Delete
                    # objPic = GLR.word_ris_in.InlineShapes.AddPicture (name_wmf , , ,  control.Range)
                    # objPic.LockAspectRatio = 0
                    # objPic.Width = 1000 # высота
                    # objPic.Height = 500 #  ширина

            for Each control in GLR.word_ris_in.ContentControls  # поменять загрузку  и выделить цветом
                if control.Tag = kluch_Tag and control.Title = "загрузка":
                    control.Range.Delete
                    control.Range.Text = RGR.name_ris(6)
                    if GLR.text_color = 1: control.Range.Font.Color = vbRed

            if GLR.text_color = 1:
                for Each control in GLR.word_ris_in.ContentControls  # выделить цветом "рисунок"
                    if control.Tag = kluch_Tag and control.Title = "рис":
                        control.Range.Font.Color = vbRed

            #  control.Delete (истина лож )True удаляет все False - оставляет содержимое
            if GLR.filtr_file = 1: exit
            for
                if GLR.filtr_file > 1: GLR.filtr_file = GLR.filtr_file - 1

elif tip_zad_ris = 5:  # замена рисунков в ворд из ворд

    GLR.word_App = CreateObject("word.Application")
    GLR.word_App.Visible = True
    GLR.word_ris_in = GLR.word_App.Documents.Open(GLR.File_ris_in)  # открыть сущ док
    GLR.word_ris_iz = GLR.word_App.Documents.Open(GLR.File_ris_iz)  # открыть сущ док
    GLR.Dict_iz = CreateObject("Scripting.Dictionary")  # для хранения kluch_doc
    GLR.Dict_in = CreateObject("Scripting.Dictionary")  # для хранения kluch_doc
    Dict_unik_Tag_doc_sub(GLR.Dict_in, GLR.word_ris_in)
    Dict_unik_Tag_doc_sub(GLR.Dict_iz, GLR.word_ris_iz)
    # print_dic (GLR.Dict_iz )

    GLR.name_ris_zamena = array("рис", "номер", "сезон год", "доп имя", "нр/откл+действие", "загрузка")

    for Each kluch_Tag_iz in GLR.Dict_iz.Keys  # цикл по ключам GLR.Dict_iz

        for Each control in GLR.word_ris_iz.ContentControls  # записываем все значения для тек Tag GLR.word_ris_iz
            if control.Tag = kluch_Tag_iz and control.Title = "номер": GLR.name_ris_zamena(1) = control.Range.Text
            if control.Tag = kluch_Tag_iz and control.Title = "сезон год": GLR.name_ris_zamena(2) = control.Range.Text
            if control.Tag = kluch_Tag_iz and control.Title = "доп имя": GLR.name_ris_zamena(3) = control.Range.Text
            if control.Tag = kluch_Tag_iz and control.Title = "нр/откл+действие": GLR.name_ris_zamena(
                4) = control.Range.Text
            if control.Tag = kluch_Tag_iz and control.Title = "загрузка": GLR.name_ris_zamena(5) = control.Range.Text

            if control.Tag = kluch_Tag_iz and control.Title = "рисунок": control.Range.copy

        GLR.naiden = 0
        GLR.naiden_ris = 0
        for Each kluch_Tag_in in GLR.Dict_in.Keys  # цикл по ключам GLR.Dict_in
            #  по ключам нр/откл+действие" "сезон год" "доп имя" или по номеру рисунка
            if (fword_Text(GLR.word_ris_in, kluch_Tag_in, "нр/откл+действие") = GLR.name_ris_zamena (4) and _
            fword_Text ( GLR.word_ris_in, kluch_Tag_in, "сезон год")        = GLR.name_ris_zamena (2) and _
            fword_Text ( GLR.word_ris_in, kluch_Tag_in, "доп имя")          = GLR.name_ris_zamena (3) and GLR.kluch_ris = 1 ) OR _
            (fword_Text(GLR.word_ris_in, kluch_Tag_in, "номер") = GLR.name_ris_zamena (1) and GLR.kluch_ris = 0 ):
            # если истина то копируем
            for Each control in GLR.word_ris_in.ContentControls  # записываем новые значения
                if control.Tag = kluch_Tag_in:

                    if control.Title = "рисунок":
                        control.Range.InlineShapes(1).Delete
                        control.Range.paste
                        GLR.naiden_ris = 1

                    if GLR.text_color = 1 and control.Title = "рис": control.Range.Font.Color = vbRed

                    if control.Title = "загрузка":
                        control.Range.Text = GLR.name_ris_zamena(5)
                        if GLR.text_color = 1: control.Range.Font.Color = vbRed
                        exit
                        for

            logging.info("\t" + " найден Tag  iz: " + kluch_Tag_iz + " соответствует in: " + kluch_Tag_in)
            if GLR.naiden_ris = 0: logging.info("\t" + " не найден рисунок Tag  in: " + kluch_Tag_in)
            GLR.naiden = 1

if GLR.naiden = 0: logging.info("\t" + " не найден Tag  iz: " + kluch_Tag_iz)


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def RisunokPrint():  # вставить рисунок в ворд/ сохранить rg2
    # dim RISUNKI , RIS#  для шапес
    # dim kursor #  двигаем курсор
    # dim txt_temp #  для временного хранения
    # dim name_wmf , ContentControls_add , arr_graf_shot, shot_name , test_wmf

    if GLR.risunok_rg2:
        rastr.Save(Left(
            GLR.Folder_rg2 + "\"  + f_txt_name_good ( RG.Name_Base    + "(" + GLR.name_ris1 + str (GLR.number_pict) + ")
        " + RGR.name_ris(5)) , 250) + ".rg2
        ", RG.shablon )
        logging.info("сохранен файл: " + Left(
            GLR.Folder_rg2 + "\"  + f_txt_name_good ( RG.Name_Base    + "(" + GLR.name_ris1 + str (GLR.number_pict) + ")
        "+ RGR.name_ris(5)),250) + ".rg2
        ")

        if GLR.risunok_word:  # в вставить рисунок в ворд
            test_wmf = False
        ContentControls_add = False  # 1 добавить ContentControls

        kursor = GS.word.Selection  # для работы в области курсора https://vremya-ne-zhdet.ru/vba-excel/redaktirovaniye-dokumentov-word/
        RISUNKI = GLR.word2.Shapes
        # объект шапес, в нем храним все картинки в доке, при добавлении
        kursor.Font.Size = 12
        # шрифт
        kursor.Font.Name = "Times  Roman"

        for i = 0 to  ubound (GLR.graf_shot)
        arr_graf_shot = split (GLR.graf_shot (i), "/")
        if ubound (arr_graf_shot) > 0: shot_name = arr_graf_shot(1) else:
            shot_name = ""

        name_wmf = left(GLR.Folder_wmf + "\" + GLR.name_ris1  + str (GLR.number_pict)    + "(" + RG.Name_Base  +")
        " +  shot_name , 250) + ".wmf
        "

        for i_i =1 to 10
        # logging.info("\t" + "итерация  " + str(i_i) )
        rastr.SendCommandMain (23, arr_graf_shot (0), name_wmf, 100503 )  # (COMM_OPEN_GRAPH=23,номер кадра(значение от 10-ctrl+0 до 19ctrl+9),файл для сохранения,?)графика должна быть открыта
        if objFSO.fileExists(name_wmf):  # файл есть
            file_wmf = objFSO.getfile(name_wmf)
            if number_pict = number_pict_first:
        GLR.file_wmf_size(i) = file_wmf.size
        exit
        for
        else:
            if
        file_wmf.size > GLR.file_wmf_size(i) * 0.8 and file_wmf.size < GLR.file_wmf_size(i) * 1.2:  # подобный размер
        test_wmf = True
        exit
        for
        else:  # не подобный размер
            file_wmf.delete
        logging.info("\t" + "удален бракованный рисунок " + name_wmf)

        #  выделить графику по узлам несколько узлов
        #  sel0 ()
        #  for each elem in GLR.graf_shot
        #      grup_cor ( "node","sel","ny=" + str (elem),"1" )#  выделяем узлы дли позиционированмя окна графики
        #  # next
        #  rastr.SendCommandMain (23 , "sel" , "" , 10 ) # (?,выборка узлов , выборка ветвей,сохранить кадр)  позиционировать экран на  выделенных узлах(графика должна быть открыта)
        #  #  от  10 - графика кадр ctrl+0 (графика кадр ndx 1(ndx c 0))   - до  19 - графика кадр ctrl+9  (графика кадр ndx 10)
        #  sel0 ()
        #  rastr.SendCommandMain (23 , "10", name_wmf , 100503)

        kursor.EndKey(6)
        # перейти в конец текста

        if objFSO.fileExists(name_wmf):
            RISUNKI.AddPicture(name_wmf)
        # вставить рисунок_ имя,связать файл_False независимый,сохранить связь, слева, сверху,ширина, высота
        RIS = RISUNKI(1)
        RIS.ConvertToInlineShape()  # конвертируем шапес в шапе
        else:
        logging.info("файл name_wmf не найден: " + name_wmf)
        logging.info("попробуйте закрыть, а потом открыть окно графики")

        kursor.MoveRight(1, 1, 1)  # курсор вправо с зажатым шифт

        if ContentControls_add: kursor.Range.ContentControls.Add(2)  # элемент управления - 2 рисунок
        if ContentControls_add: kursor.ParentContentControl.Title = "рисунок"
        if ContentControls_add: kursor.ParentContentControl.Tag = str(GLR.number_pict)

        kursor.EndKey(6)
        # перейти в конец текста
        kursor.MoveRight(1, 1)
        kursor.TypeParagraph()
        kursor.TypeText("\t")

        for iii = 0 to 6
        if iii = 2: kursor.TypeText(" - ")
        if ContentControls_add:
            addTextContent(kursor, RGR.name_ris_info(iii), RGR.name_ris(iii), str(GLR.number_pict))
        else:
            if
        iii = 4: kursor.TypeText(shot_name) else: kursor.TypeText(str(RGR.name_ris(iii)))
        kursor.MoveRight(1, 1)

        kursor.InsertBreak(
            0)  # равзрыв:7 страницы с новой строки, 0-в той же строке,1 и  8 колонки,2-5 раздела со след стр,6 и 9-11 перенос на новую стр
        if ubound(GLR.graf_shot) > 0 and not (i= ubound (GLR.graf_shot)):
        GLR.number_pict = GLR.number_pict + 1 + GLR.nomer_ris_shag
        RGR.name_ris(1) = GLR.number_pict

        # RISUNKI(1).WrapFormat.Type = 4#  положение рис относительно текста 1,2,5,6 наклвдывается,0 и4 обтекать
        # !!!!!!!!!!!!!!!! номер рисунка соответствует количеству всех рисунков в файле
        # RISUNKI(1).PictureFormat.cropleft = 5000#  удалить слева 5000 пикселей
        # word1.SaveAs (papka+"Рисунки.docx" )#   сохранить как, имя нового дока
        # word2.Save  word2.Close - закрыть word.Quit -  сохранить
        # logging.info ( "\t" + RIS.name)
        # RIS.Rotation = 45 #  вращать

        GLR.number_pict = GLR.number_pict + 1 + GLR.nomer_ris_shag

        # *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************


def addTextContent(kursor, txt_info, txt, kluch):
    kursor.Range.ContentControls.Add(0)  # элемент управления - 0 текст
    kursor.ParentContentControl.Title = txt_info
    kursor.ParentContentControl.Tag = Left(kluch, 64)
    kursor.TypeText(str(txt))
    kursor.MoveRight(1, 1)


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def grup_kor(tabl, param, viborka,
             formula):  # групповая коррекция(таблица, параметр корр, выборка, формула для расчета параметра)
    # dim tGen , tNG , tArea1
    # dim tV24  , tN24
    tN24 = rastr.Tables("node")
    tV24 = rastr.Tables("vetv")
    tArea1 = rastr.Tables("area")
    if tabl = "node":
        tN24.setsel(viborka)
        tN24.cols.item(param).Calc(formula)
    elif tabl = "vetv":
        tV24.setsel(viborka)
        tV24.cols.item(param).Calc(formula)
    elif tabl = "area":
        tArea1.setsel(viborka)
        tArea1.cols.item(param).Calc(formula)
    elif tabl = "ngroup":
        tNG = rastr.Tables("ngroup")  # нагрузочная группа
        tNG.setsel(viborka)
        tNG.cols.item(param).Calc(formula)
    elif tabl = "Generator":
        tGen = rastr.Tables("Generator")
        tGen.setsel(viborka)
        tGen.cols.item(param).Calc(formula)


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def r_print_masiv(maviv):
    for i = 0 to ubound ( maviv )
    logging.info("\t" + str(maviv(i)))


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def DName_sub(tb, ndx_tb):  # заполняет диспетчерские наименования
    rastr.Tables(tb).cols.item("dname").ZS(ndx_tb) = trim(rastr.Tables(tb).cols.item("dname").ZS(ndx_tb))
    if len(rastr.Tables(tb).cols.item("dname").ZS(ndx_tb)) < 2: rastr.Tables(tb).cols.item("dname").ZS(
        ndx_tb) = rastr.Tables(tb).cols.item("name").ZS(ndx_tb)


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def auto_run(zadanie, KontrolEl, tip):  # выполнение задания в соответствии с таблицей AutoZad
    # zadanie - номер или имя соответствующее полю remont_add, otkl_add, automatika в таблице node,vetv
    #  tip     0  - действие по факту отключения адд откл,         1 - действие при ремонте ,       2 действие по факту перегрузки
    #  KontrolEl = array("node",121) (табл,ndx) для контроля перегрузки, если перегрузка устранена то автомитика заканчивает действие
    # dim ndx_v,ndx_n, zn, pn, pn0, TXTrez(), flag(), autoTXT , value_m, value_z
    # dim tGen , tNG, tArea2 , tArea4 , vibor_PA , max_step
    # dim tV26  , tN26 , tAZ , ndx_az , zadanie_m ,cycle_while
    if rastr.Tables.Find("AutoZad") < 0:
    # logging.info( "!!! НЕ ЗАГРУЖЕН ШАБЛОН АВТОМАТИКИ с таблицей  AutoZad (.amt) !!" )
    else:
        tAZ = rastr.Tables("AutoZad")
        if tAZ.size = 0:
        # logging.info( "!!! НЕ ЗАГРУЖЕН ФАЙЛ АВТОМАТИКИ (.amt) !!" )
        else:
            #  N        #   номер соответствующий полю remont_add, otkl_add, automatika в таблице node,vetv
            #  Nstep    #   номер ступени, для последовательности выполнения
            #  sta      #   отключение ступени
            #  action   #   0 откл|1 вкл|2 ОН|3 ОГ|4 ИЗМ       действие: отключить|включить узел, ветвь, генератор|ОН-ограничение нагрузки узла, района, территории, нагр группы, |ОГ ограничение генерации узла, генератора|ИЗМ произвольное изменение сети
            #  tabl     #    0  узел|1 ветвь|2 район|3 территроррия|4 нагр.группа|5 генератор|6 ИЗМ
            #  kluch    #      номер узла, генератора, ветви(ip,iq,np),нагрузочной группы, района...
            #  value    #       величина воздействия
            #  uslovie       #       условие выполнения
            #  name_step#      имя контролируемого элемента (имя изменяемого элемента) - заполняется макросом
            #  setpoint #      отметить если разгружать до адтн (иначе до ДДТН)
            tN26 = rastr.Tables("node")
            tV26 = rastr.Tables("vetv")
            tGen = rastr.Tables("Generator")
            tNG = rastr.Tables("ngroup")  # нагрузочная группа
            tArea2 = rastr.Tables("area")
            tArea4 = rastr.Tables("area2")
            zadanie = replace(zadanie, " ", "")
            if instr(zadanie, ",") > 0:
                zadanie_m = split(zadanie, ",")  else:
                redim
            zadanie_m(0): zadanie_m(0) = zadanie
            autoTXT = ""
            for Nzad = 0 to ubound (zadanie_m )
            zadanie = zadanie_m(Nzad)
            if isnumeric(zadanie):
                zadanie = float(zadanie)

                vibor_PA = "N=" + str(zadanie) + "+sta=0"  #
                max_step = rastr.Calc("max", "AutoZad", "Nstep", vibor_PA)
                tAZ.setsel(vibor_PA + "+!uslovie")
                count_activ_step = tAZ.count
                if count_activ_step > 0:

                    redim
                    TXTrez(max_step)
                    redim
                    flag(max_step)

                    for i = 0 to max_step
                    TXTrez(i) = ""  # для записи воздействий в цикле for
                    flag(i) = 0  # для определения начального значения в цикле"ц"

                for i = 0 to max_step
                tAZ.setsel(vibor_PA + "+!uslovie+Nstep=" + str(i))
                if tAZ.count > 0:  # есть такая ступень
                    ndx_az = tAZ.findnextsel(-1)
                    while ndx_az > -1
                        exit_cycle = 0
                        cycle_while = 0
                        if fUslovieZad(vibor_PA + "+uslovie+Nstep=" + str(
                                i)):  # функция проверяет выполнение условия (выборка: zadanie sta Nstep )
                            if tAZ.cols.item("action").Z(ndx_az) < 2:  # 0 откл|1 вкл
                                if tAZ.cols.item("tabl").Z(ndx_az) = 0:  # 0  узел
                                    ndx_n = fNDX("node", tAZ.cols.item("kluch").ZS(ndx_az))
                                    if ndx_n > -1:
                                        DName_sub("node", ndx_n) else:
                                        logging.info(
                                            "не найден узел (automatika,tip=" + str(tip) + "), задание:" + str(zadanie))
                                    if tAZ.cols.item("action").Z(ndx_az)  = 0:  # отключить

                                        if fNodeSta("ndx", ndx_n,
                                                    1):  # "ndx"/"ny"; 121 ; vkl_otkl= 1 отключить/ 0 включить)
                                            if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                            TXTrez(i) = TXTrez(i) + "отключение " + tN26.cols.item("dname").ZS(ndx_n)
                                        else:
                                            logging.info("\t" + "узел отключен до (automatika,tip=" + str(
                                                tip) + "), " + rastr.Tables("node").SelString(ndx_n) + tN26.cols.item(
                                                "dname").ZS(ndx_n) + ", N комб: " + str(GLR.N_rezh))

                                    elif tAZ.cols.item("action").Z(ndx_az)  = 1:  # включить
                                        if fNodeSta("ndx", ndx_n,
                                                    0):  # "ndx"/"ny"; 121 ; vkl_otkl= 1 отключить/ 0 включить)
                                            if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                            TXTrez(i) = TXTrez(i) + "включение " + tN26.cols.item("dname").ZS(ndx_n)
                                        else:
                                            logging.info("\t" + "ветвь отключена до (automatika,tip=" + str(
                                                tip) + "), " + rastr.Tables("node").SelString(
                                                ndx_n) + " {" + tN26.cols.item("dname").ZS(ndx_n) + "}, N комб: " + str(
                                                GLR.N_rezh)) elif tAZ.cols.item("tabl").Z(ndx_az) = 1:  # 1 ветвь

                                        if instr(tAZ.cols.item("kluch").Z(ndx_az), ",") > 1:  # задано ip,iq,np
                                            ndx_v = fNDX("vetv", tAZ.cols.item("kluch").Z(ndx_az))
                                            if ndx_v = -1:
                                                logging.info("не найдена ветвь (automatika,tip=" + str(
                                                    tip) + "), задание:" + str(zadanie) + "не найден v0/1 ndx_v " + str(
                                                    ndx_v))
                                            else:
                                                DName_sub("vetv", ndx_v)
                                                if tAZ.cols.item("action").Z(ndx_az)  = 0:  # отключить

                                                    if fVetv_Sta("ndx", ndx_v,
                                                                 1):  # "ndx"/"groupid"/"kluch"; "ip=1,iq=2,np=0"; vkl_otkl= 1 отключить/ 0 включить)
                                                        if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                                        TXTrez(i) = TXTrez(i) + "отключение " + tV26.cols.item(
                                                            "dname").ZS(ndx_v)
                                                    else:
                                                        logging.info("\t" + "ветвь отключена до (automatika,tip=" + str(
                                                            tip) + "), " + rastr.Tables("vetv").SelString(
                                                            ndx_v) + " {" + tV26.cols.item("dname").ZS(
                                                            ndx_v) + "}, N комб: " + str(GLR.N_rezh))

                                                elif tAZ.cols.item("action").Z(ndx_az)  = 1:  # включить

                                                    if fVetv_Sta("ndx", ndx_v,
                                                                 0):  # "ndx"/"groupid"/"kluch"; "ip=1,iq=2,np=0"; vkl_otkl= 1 отключить/ 0 включить)
                                                        if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                                        TXTrez(i) = TXTrez(i) + "включение " + tV26.cols.item(
                                                            "dname").ZS(ndx_v)
                                                    else:
                                                        logging.info("\t" + "ветвь отключена до (automatika,tip=" + str(
                                                            tip) + "), " + rastr.Tables("vetv").SelString(
                                                            ndx_v) + " {" + tV26.cols.item("dname").ZS(
                                                            ndx_v) + "}, N комб: " + str(GLR.N_rezh))
                                        else:  # задано groupid
                                            tV26.setsel("groupid=" + str(tAZ.cols.item("kluch").Z(ndx_az)))
                                            ndx_v = tV26.FindNextSel(-1)

                                            if ndx_v < 0:
                                                logging.info("\t" + "не найден v0 groupid " + str(
                                                    tAZ.cols.item("kluch").Z(
                                                        ndx_az)))  # если строка, те не найтен NDX, те ошибка в задании
                                            else:
                                                if tAZ.cols.item("action").Z(ndx_az)  = 0:  # отключить

                                                    if fVetv_Sta("groupid", tAZ.cols.item("kluch").Z(ndx_az),
                                                                 1):  # "ndx"/"groupid"/"kluch"; "ip=1,iq=2,np=0"; vkl_otkl= 1 отключить/ 0 включить)
                                                        if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                                        TXTrez(i) = TXTrez(i) + "отключение " + tV26.cols.item(
                                                            "dname").ZS(ndx_v)
                                                    else:
                                                        logging.info("\t" + "ветвь отключена до (automatika,tip=" + str(
                                                            tip) + "), " + rastr.Tables("vetv").SelString(
                                                            ndx_v) + ", N комб: " + str(GLR.N_rezh))

                                                elif tAZ.cols.item("action").Z(ndx_az)  = 1:  # включить

                                                    if fVetv_Sta("groupid", tAZ.cols.item("kluch").Z(ndx_az),
                                                                 0):  # "ndx"/"groupid"/"kluch"; "ip=1,iq=2,np=0"; vkl_otkl= 1 отключить/ 0 включить)
                                                        if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                                        TXTrez(i) = TXTrez(i) + "включение " + tV26.cols.item(
                                                            "dname").ZS(ndx_v)
                                                    else:
                                                        logging.info("\t" + "ветвь отключена до (automatika,tip=" + str(
                                                            tip) + "), " + rastr.Tables("vetv").SelString(
                                                            ndx_v) + ", N комб: " + str(GLR.N_rezh))
                                    # 0 откл|1 вкл|2 ОН|3 ОГ|4 ИЗМ ****  tabl  0  узел|1 ветвь|2 район|3 территроррия|4 нагр.группа|5 генератор|6 ИЗМ
                                if tAZ.cols.item("action").Z(
                                        ndx_az) = 2 or tAZ.cols.item("action").Z(ndx_az) = 3:  # 2 ОН|3 ОГ
                                    tAZ.cols.item("value").ZS(ndx_az) = Replace(tAZ.cols.item("value").ZS(ndx_az), " ",
                                                                                "")  # удалить все пробелы
                                    if tAZ.cols.item("value").ZS(ndx_az) = "": logging.info(
                                        "\t" + "значение не указано, N=" + tAZ.cols.item("N").ZS(
                                            ndx_az) + ",Nstep=" + tAZ.cols.item("Nstep").ZS(ndx_az)): exit

                                    def

                                        value_m = split(tAZ.cols.item("value").ZS(ndx_az), "*")
                                    if isnumeric(value_m(0)):
                                        value_z = float(value_m(0)) else:
                                        logging.info("\t" + "значение не является числом, N=" + tAZ.cols.item("N").ZS(
                                            ndx_az) + ",Nstep=" + tAZ.cols.item("Nstep").ZS(ndx_az)): exit

                                    def

                                    if tAZ.cols.item("action").Z(
                                            ndx_az) = 2 and tAZ.cols.item("tabl").Z(ndx_az) = 0:  # ограничение НАГРУЗКИ УЗЛА
                                        RG.loadRGM = True
                                        ndx_n = fNDX("node", tAZ.cols.item("kluch").ZS(ndx_az))
                                        if ndx_n = -1:
                                            logging.info(
                                                "не найден узел (automatika,tip=" + str(tip) + "), задание:" + str(
                                                    zadanie) + ", узел номер:" + tAZ.cols.item("kluch").ZS(ndx_az))
                                        else:
                                            pn = tN26.cols.item("pn").Z(ndx_n)
                                            if flag(i) = 0: pn0 = pn
                                            flag(i) = 1

                                            if pn < value_z:
                                                exit_cycle = 1
                                                tN26.cols.item("pn").Z(ndx_n) = 0
                                                tN26.cols.item("qn").Z(ndx_n) = 0
                                                if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                                TXTrez(i) = TXTrez(i) + "ограничение нагрузки в узле " + tN26.cols.item(
                                                    "dname").ZS(ndx_n) + " [" + tN26.cols.item("ny").ZS(
                                                    ndx_n) + "] на " + str(round(pn0, 0)) + " МВт"
                                        else:
                                        Kefn = 1 - (value_z / pn)
                                        tN26.cols.item("pn").Z(ndx_n) = tN26.cols.item("pn").Z(ndx_n) * Kefn
                                        tN26.cols.item("qn").Z(ndx_n) = tN26.cols.item("qn").Z(ndx_n) * Kefn
                                        pn = tN26.cols.item("pn").Z(ndx_n)
                                        if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                        TXTrez(i) = TXTrez(i) + "ограничение нагрузки на " + str(round(pn0 - pn,
                                                                                                       0)) + " МВт"  # + "  на "+ tN26.cols.item("dname").ZS(ndx_n)                                             elif tAZ.cols.item("action").Z(ndx_az) = 2 and tAZ.cols.item("tabl").Z(ndx_az) = 4:   #  ограничение нагрузки НАГРУЗОЧНОЙ ГРУППЫ
                                RG.loadRGM = True
                                ndx_ng = fNDX("ngroup", tAZ.cols.item("kluch").ZS(ndx_az))
                                if ndx_ng = -1:
                                    logging.info("не найден номер нагр. группы (automatika,tip=" + str(
                                        tip) + "), задание:" + str(zadanie) + ",  номер:" + tAZ.cols.item("kluch").ZS(
                                        ndx_az))
                                else:
                                    png = rastr.Calc("sum", "ngroup", "pn", "nga=" + tAZ.cols.item("kluch").ZS(
                                        ndx_az))  # tNG.cols.item("pn").Z(ndx_ng)

                                    if flag(i) = 0: png0 = png
                                    flag(i) = 1

                                    if png < value_z:

                                        exit_cycle = 1
                                        grup_kor("node", "pn", "nga=" + tAZ.cols.item("kluch").ZS(ndx_az), "0")
                                        grup_kor("node", "qn", "nga=" + tAZ.cols.item("kluch").ZS(ndx_az), "0")
                                        if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                        TXTrez(i) = TXTrez(i) + "ограничение нагрузки nga= " + tAZ.cols.item(
                                            "kluch").ZS(ndx_az) + " на " + str(round(png0, 0)) + " МВт"
                                    else:
                                        Kefng = 1 - (value_z / png)

                                        grup_kor("node", "pn", "nga=" + tAZ.cols.item("kluch").ZS(ndx_az),
                                                 "pn*" + str(Kefng))
                                        grup_kor("node", "qn", "nga=" + tAZ.cols.item("kluch").ZS(ndx_az),
                                                 "qn*" + str(Kefng))

                                        png = rastr.Calc("sum", "ngroup", "pn", "nga=" + tAZ.cols.item("kluch").ZS(
                                            ndx_az))  # tNG.cols.item("pn").Z(ndx_ng)

                                        if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                        TXTrez(i) = TXTrez(i) + "ограничение нагрузки nga= " + tAZ.cols.item(
                                            "kluch").ZS(ndx_az) + " на " + str(
                                            round(png0 - png, 0)) + " МВт" elif tAZ.cols.item("action").Z(
                                            ndx_az) = 2 and tAZ.cols.item("tabl").Z(
                                            ndx_az) = 2:  # ограничение НАГРУЗКИ РАЙОНА
                                    RG.loadRGM = True
                                    ndx_nr = fNDX("area", tAZ.cols.item("kluch").ZS(ndx_az))

                                    if ndx_nr = -1:
                                        logging.info(
                                            "не найден номер района (automatika,tip=" + str(tip) + "), задание:" + str(
                                                zadanie) + ",  номер:" + tAZ.cols.item("kluch").ZS(ndx_az))
                                    else:
                                        pnr = tArea2.cols.item("pn_sum").Z(ndx_nr)

                                        if flag(i) = 0: pnr0 = pnr
                                        flag(i) = 1

                                        if pnr < value_z:
                                            exit_cycle = 1
                                            grup_kor("node", "pn", "na=" + tAZ.cols.item("kluch").ZS(ndx_az), "0")
                                            grup_kor("node", "qn", "na=" + tAZ.cols.item("kluch").ZS(ndx_az), "0")

                                            if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                            TXTrez(i) = TXTrez(i) + "ограничение нагрузки " + str(
                                                round(pnr0, 0)) + " МВт"
                                        else:
                                            Kef = 1 - (value_z / pnr)
                                            grup_kor("node", "pn", "na=" + tAZ.cols.item("kluch").ZS(ndx_az),
                                                     "pn*" + str(Kef))
                                            grup_kor("node", "qn", "na=" + tAZ.cols.item("kluch").ZS(ndx_az),
                                                     "qn*" + str(Kef))

                                            pnr = tArea2.cols.item("pn_sum").Z(ndx_nr)

                                            if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                            TXTrez(i) = TXTrez(i) + "ограничение нагрузки " + str(
                                                round(pnr0 - pnr, 0)) + " МВт" elif tAZ.cols.item("action").Z(
                                                ndx_az) = 2 and tAZ.cols.item("tabl").Z(
                                                ndx_az) = 3:  # ограничение НАГРУЗКИ территории
                                        RG.loadRGM = True
                                        ndx_nr = fNDX("area2", tAZ.cols.item("kluch").ZS(ndx_az))

                                        if ndx_nr = -1:
                                            logging.info("не найден номер территории (automatika,tip=" + str(
                                                tip) + "), задание:" + str(zadanie) + ",  номер:" + tAZ.cols.item(
                                                "kluch").ZS(ndx_az))
                                        else:
                                            pnr = tArea4.cols.item("pn_sum").Z(ndx_nr)

                                            if flag(i) = 0: pnr0 = pnr
                                            flag(i) = 1

                                            if pnr < value_z:
                                                exit_cycle = 1
                                                grup_kor("node", "pn", "npa=" + tAZ.cols.item("kluch").ZS(ndx_az), "0")
                                                grup_kor("node", "qn", "npa=" + tAZ.cols.item("kluch").ZS(ndx_az), "0")

                                                if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                                TXTrez(i) = TXTrez(i) + "ограничение нагрузки " + str(
                                                    round(pnr0, 0)) + " МВт"
                                            else:
                                                Kef = 1 - (value_z / pnr)
                                                grup_kor("node", "pn", "npa=" + tAZ.cols.item("kluch").ZS(ndx_az),
                                                         "pn*" + str(Kef))
                                                grup_kor("node", "qn", "npa=" + tAZ.cols.item("kluch").ZS(ndx_az),
                                                         "qn*" + str(Kef))

                                                pnr = tArea4.cols.item("pn_sum").Z(ndx_nr)

                                                if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                                TXTrez(i) = TXTrez(i) + "ограничение нагрузки " + str(
                                                    round(pnr0 - pnr, 0)) + " МВт" elif tAZ.cols.item("action").Z(
                                                    ndx_az) = 3 and tAZ.cols.item("tabl").Z(
                                                    ndx_az) = 0:  # ограничение ГЕНЕРАЦИИ УЗЛА
                                            RG.loadRGM = True
                                            ndx_g = fNDX("node", tAZ.cols.item("kluch").ZS(ndx_az))

                                            if ndx_g = -1:
                                                logging.info("не найден номер узла (automatika,tip=" + str(
                                                    tip) + "), задание:" + str(zadanie) + ",  номер:" + tAZ.cols.item(
                                                    "kluch").ZS(ndx_az))
                                            else:
                                                pg = tN26.cols.item("pg").Z(ndx_g)
                                                if flag(i) = 0: pg0 = pg
                                                flag(i) = 1

                                                if pg < value_z:
                                                    exit_cycle = 1
                                                    tN26.cols.item("pg").Z(ndx_g) = 0
                                                    tN26.cols.item("qg").Z(ndx_g) = 0
                                                    tN26.cols.item("qmax").Z(ndx_g) = 0
                                                    tN26.cols.item("qmin").Z(ndx_g) = 0
                                                    if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                                    TXTrez(i) = TXTrez(i) + "ограничение генерации на " + str(round(pg0,
                                                                                                                    0)) + " МВт"  # + "  на "+ tN26.cols.item("dname").ZS(ndx_n)
                                                else:
                                                    tN26.cols.item("pg").Z(ndx_g) = tN26.cols.item("pg").Z(
                                                        ndx_g) - value_z
                                                    pg = tN26.cols.item("pg").Z(ndx_g)
                                                    if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                                    TXTrez(i) = TXTrez(i) + "ограничение генерации на " + str(
                                                        round(pg0 - pg,
                                                              0)) + " МВт"  # + "  на "+ tN26.cols.item("dname").ZS(ndx_n)                                             elif tAZ.cols.item("action").Z(ndx_az) = 3 and tAZ.cols.item("tabl").Z(ndx_az) = 5:   #  ограничение ГЕНЕРАЦИИ ГЕНЕРАТОРА

                                            RG.loadRGM = True
                                            ndx_gg = fNDX("Generator", tAZ.cols.item("kluch").ZS(ndx_az))

                                            if ndx_gg = -1:
                                                logging.info("не найден номер генератора (automatika,tip=" + str(
                                                    tip) + "), задание:" + str(zadanie) + ",  номер:" + tAZ.cols.item(
                                                    "kluch").ZS(ndx_az))
                                            else:
                                                pgg = tGen.cols.item("P").Z(ndx_gg)
                                                pgg_min = tGen.cols.item("Pmin").Z(ndx_gg)
                                                pgg_max = tGen.cols.item("Pmax").Z(ndx_gg)
                                                if flag(i) = 0: pgg0 = pgg
                                                flag(i) = 1

                                                if value_z > 0:  # нада снизить  ген
                                                    if not tGen.cols.item("sta").Z(ndx_gg):  # истина откл ген
                                                        if pgg <= value_z:
                                                            exit_cycle = 1
                                                            tGen.cols.item("sta").Z(ndx_gg) = 1
                                                            if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                                            TXTrez(i) = TXTrez(i) + "ограничение генерации на " + str(
                                                                round(pgg0,
                                                                      0)) + " МВт"  # + "  на "+ tN26.cols.item("dname").ZS(ndx_n)
                                                        else:
                                                            if pgg - pgg_min < value_z and pgg_min > 0:
                                                                exit_cycle = 1
                                                                tGen.cols.item("P").Z(ndx_gg) = tGen.cols.item(
                                                                    "Pmin").Z(ndx_gg)
                                                                if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                                                TXTrez(i) = TXTrez(
                                                                    i) + "ограничение генерации на " + str(
                                                                    round(pgg0 - pgg_min,
                                                                          0)) + " МВт"  # + "  на "+ tN26.cols.item("dname").ZS(ndx_n)
                                                            else:
                                                                tGen.cols.item("P").Z(ndx_gg) = tGen.cols.item("P").Z(
                                                                    ndx_gg) - value_z
                                                                pgg = tGen.cols.item("P").Z(ndx_gg)
                                                                if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                                                TXTrez(i) = TXTrez(
                                                                    i) + "ограничение генерации на " + str(
                                                                    round(pgg0 - pgg,
                                                                          0)) + " МВт"  # + "  на "+ tN26.cols.item("dname").ZS(ndx_n)
                                                elif value_z < 0:  # нада увеличить  ген
                                                    if tGen.cols.item("sta").Z(ndx_gg):  # истина откл ген
                                                        tGen.cols.item("sta").Z(ndx_gg) = 0

                                                        if pgg_max > - value_z:  #
                                                            tGen.cols.item("P").Z(ndx_gg) = - value_z
                                                            if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                                            TXTrez(i) = TXTrez(i) + "ген вкл P= " + str(
                                                                - value_z) + " МВт"

                                                        if pgg_max < - value_z:
                                                            tGen.cols.item("P").Z(ndx_gg) = pgg_max
                                                            if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                                            TXTrez(i) = TXTrez(i) + "ген вкл P= " + str(
                                                                pgg_max) + " МВт"
                                                            exit_cycle = 1

                                                    else:  # ген включен
                                                        if - value_z > pgg_max:
                                                            tGen.cols.item("P").Z(ndx_gg) = - value_z else:
                                                            tGen.cols.item("P").Z(ndx_gg) = pgg_max: exit_cycle = 1
                                                        pgg = tGen.cols.item("P").Z(ndx_gg)
                                                        if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                                        TXTrez(i) = TXTrez(i) + "увеличение генерации на " + str(
                                                            round(-pgg0 + pgg,
                                                                  0)) + " МВт"  # + "  на "+ tN26.cols.item("dname").ZS(ndx_n)                                         if tAZ.cols.item("action").Z(ndx_az) = 4:                                                #  вызов процедуры cor
                                        RG.loadRGM = True
                                        cor(tAZ.cols.item("kluch").ZS(ndx_az), tAZ.cols.item("value").ZS(ndx_az))
                                        if not TXTrez(i) = "": TXTrez(i) = TXTrez(i) + ", "
                                        TXTrez(i) = TXTrez(i) + TXTrez(i) = "изм_" + tAZ.cols.item("kluch").ZS(
                                            ndx_az) + "_" + tAZ.cols.item("value").ZS(ndx_az)

                                    if tip = 2:  # (если otkl_add или remont_add то не нада расчитывать )

                                        if tAZ.findnextsel(ndx_az) = -1:  # если  последняя строка  ступени то расчет
                                            GS.kod_rgm = rastr.rgm("")
                                        if GS.kod_rgm = 1:
                                            GS.kod_rgm = rastr.rgm("")
                                        if GS.kod_rgm = 1:
                                            GS.kod_rgm = rastr.rgm("p")
                                        if GS.kod_rgm = 1:
                                            GS.kod_rgm = rastr.rgm("p")
                                        if GS.kod_rgm = 1:
                                            GS.kod_rgm = rastr.rgm("p")

                                        if GLR.ris_tabl_add_PA:
                                            RGR.FLAG_ris_tabl_add_PA = 1  else:
                                            RGR.FLAG_ris_tabl_add_PA = 0

                                        if KontrolEl(0) = "node":
                                            if tAZ.cols.item("setpoint").ZS(ndx_az):
                                                if tN26.cols.item("vras").Z(KontrolEl(1)) > tV26.cols.item(
                                                    "umin_av").ZN(KontrolEl(1)): exit
                                                for
                                            else:
                                                if tN26.cols.item("vras").Z(KontrolEl(1)) > tV26.cols.item("umin").ZN(
                                                    KontrolEl(1)): exit
                                                for

                                        elif KontrolEl(0) = "vetv":
                                            if tV26.cols.item("groupid").Z(KontrolEl(1)) > 0: KontrolEl(
                                                1) = f_I_max_grouid(KontrolEl(1))
                                            if tAZ.cols.item("setpoint").ZS(ndx_az):
                                                if tV26.cols.item("i_zag_av").ZN(KontrolEl(1)) < 100: exit
                                                for
                                            else:
                                                if tV26.cols.item("i_zag").ZN(KontrolEl(1)) < 100: exit
                                                for if instr (tAZ.cols.item("value").ZS(ndx_az), "*") > 0 and exit_cycle = 0: cycle_while = 1 if cycle_while = 0: ndx_az = tAZ.findnextsel(
                                                    ndx_az)
                            wend

                    for i = 0 to max_step
                    if (not autoTXT = "") and (not TXTrez(i) = ""): autoTXT = autoTXT + ", "
                    autoTXT = autoTXT + trim(TXTrez(i))
                    flag(i) = 0

            else:
                logging.info("\t" + "!!! не найдена ступень " + vibor_PA + "!!!")

    if tip < 2:
        if not autoTXT = "":
            RGR.autoTXT_fact_Otkl_Remont = RGR.autoTXT_fact_Otkl_Remont + "[" + autoTXT + "]"  # переменная для записи autoTXT при откл и ремонте
            RGR.autoTXT_fact_Otkl_Remont_tek = "[" + autoTXT + "]"  # переменная для записи autoTXT при откл или ремонте

    elif tip = 2:
        if autoTXT != "": RGR.autoTXT_fPA = "Действие на " + autoTXT  # + "."   #  переменная для записи autoTXT при действии ПА


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fUslovieZad(vibor_PA_Zad):  # функция проверяет выполнение условия (выборка: zadanie sta ступень )
    # dim tAZ2 , ndx_azu
    tAZ2 = rastr.Tables("AutoZad")
    fUslovieZad = True
    tAZ2.setsel(vibor_PA_Zad)
    if tAZ2.count > 0:
        ndx_azu = tAZ2.findnextsel(-1)
        while ndx_azu > -1
            if tAZ2.cols.item("tabl").Z(ndx_azu) = 0: tabl2 = "node"
            if tAZ2.cols.item("tabl").Z(ndx_azu) = 1: tabl2 = "vetv"

            if tAZ2.cols.item("action").Z(ndx_azu) < 2:  # action   #   0 откл|1 вкл
                if tAZ2.cols.item("tabl").Z(ndx_azu) < 2:  # 0  узел
                    if fParam_kkluch(tabl2, "sta", tAZ2.cols.item("kluch").Z(
                        ndx_azu)) = tAZ2.cols.item("action").Z(ndx_azu): fUslovieZad = False: exit

                    def  # kluch
                else:
                    logging.info("\t" + "условие не распознано:" + vibor_PA_Zad)

            else:
                logging.info("\t" + "условие не распознано:" + vibor_PA_Zad)

            ndx_azu = tAZ2.findnextsel(ndx_azu)
        wend


# End def return

# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def name_stepVN(dict_N_PA, zadanie1, tip_str):
    # dim zadanie_m1 , ii
    if not zadanie1 = "":
        if instr(zadanie1, ",") > 0:
            zadanie_m1 = split(zadanie1, ",")  else:
            redim
        zadanie_m1(0): zadanie_m1(0) = zadanie1
        for ii = 0 to ubound(zadanie_m1)
        if not dict_N_PA.Exists(zadanie_m1(ii)):
            dict_N_PA.Add(zadanie_m1(ii), tip_str)
        else:
            dict_N_PA.Item(zadanie_m1(ii)) = dict_N_PA.Item(zadanie_m1(ii)) + ", " + tip_str


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fIOO(ndx_vn):  # расчет тока ОО АТ (ndx_vn ветви ВН ктр=1)
    # dim U1 , delta1 , U2, delta2 , P1 , Q1 ,P2 , Q2 , ny_vn , ny_sn #  vn или 1 высокое напряжение, sn или 2 высокое напряжение
    # dim tV27  , tN27
    tN27 = rastr.Tables("node")
    tV27 = rastr.Tables("vetv")
    if tV27.cols.item("ktr").Z(ndx_vn) = 1:
        tV27.setsel("ip=" + tV27.cols.item("iq").ZS(ndx_vn) + "+ktr>0.2+ktr<1")  # находим вторую ветвь
        ndx_sn = tV27.FindNextSel(-1)

        if ndx_sn = -1:  # если начало ветви ВН центр трансформатора
            tV27.setsel("ip=" + tV27.cols.item("ip").ZS(ndx_vn) + "+ktr>0.2+ktr<1")
            ndx_sn = tV27.FindNextSel(-1)
            if ndx_sn = -1:  # значит АТ задан одной ветвью(не звезда)
                logging.info("\t" + "не найдена обмотка СН " + tV27.cols.item("dname").ZS(ndx_vn))
            else:
            ny_vn = tV27.cols.item("iq").ZN(ndx_vn)
            ny_sn = tV27.cols.item("iq").ZN(ndx_sn)
            U1 = tV27.cols.item("v_iq").ZN(ndx_vn)
            delta1 = tN27.cols.item("delta").ZN(fNDX("node", ny_vn))
            U2 = tV27.cols.item("v_iq").ZN(ndx_sn)
            delta2 = tN27.cols.item("delta").ZN(fNDX("node", ny_sn))
            P1 = - tV27.cols.item("pl_iq").ZN(ndx_vn)
            Q1 = - tV27.cols.item("ql_iq").ZN(ndx_vn)
            P2 = tV27.cols.item("pl_iq").ZN(ndx_sn)
            Q2 = tV27.cols.item("ql_iq").ZN(ndx_sn)

    else:  # если начало ветви ВН не центр трансформатора
        ny_vn = tV27.cols.item("ip").ZN(ndx_vn)
        ny_sn = tV27.cols.item("iq").ZN(ndx_sn)
        U1 = tV27.cols.item("v_ip").ZN(ndx_vn)
        delta1 = tN27.cols.item("delta").ZN(fNDX("node", ny_vn))
        U2 = tV27.cols.item("v_iq").ZN(ndx_sn)
        delta2 = tN27.cols.item("delta").ZN(fNDX("node", ny_sn))
        P1 = tV27.cols.item("pl_ip").ZN(ndx_vn)
        Q1 = tV27.cols.item("ql_ip").ZN(ndx_vn)
        P2 = tV27.cols.item("pl_iq").ZN(ndx_sn)
        Q2 = tV27.cols.item("ql_iq").ZN(ndx_sn)

else:
if tV27.cols.item("v_ip").ZN(ndx_vn) > tV27.cols.item("v_iq").ZN(ndx_vn): ny_vn = tV27.cols.item("ip").ZN(
    ndx_vn): ny_sn = tV27.cols.item("iq").ZN(ndx_vn)
if tV27.cols.item("v_ip").ZN(ndx_vn) < tV27.cols.item("v_iq").ZN(ndx_vn): ny_vn = tV27.cols.item("iq").ZN(
    ndx_vn): ny_sn = tV27.cols.item("ip").ZN(ndx_vn)
U1 = tN27.cols.item("vras").Z(fNDX("node", ny_vn))
delta1 = tN27.cols.item("delta").ZN(fNDX("node", ny_vn))
U2 = tN27.cols.item("vras").Z(fNDX("node", ny_sn))
delta2 = tN27.cols.item("delta").ZN(fNDX("node", ny_sn))
P1 = tV27.cols.item("pl_ip").ZN(ndx_vn)
Q1 = tV27.cols.item("ql_ip").ZN(ndx_vn)
P2 = tV27.cols.item("pl_iq").ZN(ndx_vn)
Q2 = tV27.cols.item("ql_iq").ZN(ndx_vn)

fIOO = fRaschot_IOO(U1, delta1, U2, delta2, P1, Q1, P2, Q2)  # расчет тока ОО АТ,  положительное напрваление от ВН в СН


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fRaschot_IOO(U1, delta1, U2, delta2, P1, Q1, P2, Q2):  # расчет тока ОО АТ,  положительное напрваление от ВН в СН
    # dim ReI1 , ImI1 , ReI2 , ImI2 , ReIOO , ImIOO ,  ReU1 , ImU1 , ReU2 , ImU2
    ReU1 = U1 * cos(delta1 * (4 * atn(1)) / 180)  # pi=4*atn(1) = 3.14
    ImU1 = U1 * sin(delta1 * (4 * atn(1)) / 180)
    ReU2 = U2 * cos(delta2 * (4 * atn(1)) / 180)
    ImU2 = U2 * sin(delta2 * (4 * atn(1)) / 180)

    if U1 != 0 And U2 != 0:
        ReI1 = (P1 * ReU1 + Q1 * ImU1) / (Sqr(3) * (ReU1 * ReU1 + ImU1 * ImU1))
        ImI1 = (Q1 * ReU1 - P1 * ImU1) / (Sqr(3) * (ReU1 * ReU1 + ImU1 * ImU1))
        ReI2 = (P2 * ReU2 + Q2 * ImU2) / (Sqr(3) * (ReU2 * ReU2 + ImU2 * ImU2))
        ImI2 = (Q2 * ReU2 - P2 * ImU2) / (Sqr(3) * (ReU2 * ReU2 + ImU2 * ImU2))
        ReIOO = ReI1 - ReI2
        ImIOO = ImI1 - ImI2
        fRaschot_IOO = sqr(ReIOO * ReIOO + ImIOO * ImIOO) * 1000
    else:
        fRaschot_IOO = 0


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fParam_kkluch(tabl, param, kkluch):  # функция
    tTabl = rastr.Tables(tabl)

    if tabl = "node":
        tTabl.setsel("ny=" + str(kkluch))
    elif tabl = "vetv":
        kkluchM = split(kkluch, ",")
        tTabl.setsel("ip=" + str(kkluchM(0)) + "+iq=" + str(kkluchM(1)) + "+np=" + str(kkluchM(2)))
    elif tabl = "area":
        tTabl.setsel("na=" + str(kkluch))
    elif tabl = "ngroup":
        tTabl.setsel("nga=" + str(kkluch))
    elif tabl = "Generator":
        tTabl.setsel("Num=" + str(kkluch))

    index = tTabl.FindNextSel(-1)
    if index = -1:
        logging.info("\t" + "fParam_kkluch в таблице " + tabl + " не найдено " + param + " ключ " + str(kkluch))
        fParam_kkluch = "не найдено"
    else:
        fParam_kkluch = tTabl.cols.item(param).Z(index)


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fParam_kluch(tabl, param, kluch):  # функция
    tTabl = rastr.Tables(tabl)
    tTabl.setsel(kluch)
    index = tTabl.FindNextSel(-1)
    if index = -1:         logging.info(
        "fParam_kluch в таблице " + tabl + " не найдено " + param + " ключ " + str(kluch))
    fParam_kkluch = tTabl.cols.item(param).Z(index)


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fVetvKey(index):  # возвращает сортировочный ключь
    fVetvKey = rastr.Tables("vetv").cols.item("_SortKey").Z(index)


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fSort_N(index):  # возвращает сортировочный N транс
    fSort_N = rastr.Tables("vetv").cols.item("N").Z(index)


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fSort_NNod(index, param):  # возвращает сортировочный N узла
    fSort_NNod = rastr.Tables("node").cols.item(param).Z(index)


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def f_name_zagruzka_ris(name_zagruzka_ris, name, zagruzka, I_maxx, Ipr, I_adtn, dop_name):
    f_name_zagruzka_ris = str(I_maxx)

    name = trim(name)  # удалить пробелы в начале и конце строки
    if I_adtn > Ipr * 1.01:
        if name_zagruzka_ris = "":  # имя Загрузка для рисунков
            f_name_zagruzka_ris = "Загрузка " + name + dop_name + " составляет " + str(
                Round(zagruzka, 0)) + " % (" + str(I_maxx) + " А) от Iддтн = " + str(round(Ipr)) + " А и " + str(
                round(I_maxx / I_adtn * 100, 0)) + " % от Iадтн = " + str(round(I_adtn)) + " А"  # имя контр
        else:
            f_name_zagruzka_ris = name_zagruzka_ris + ", " + name + dop_name + " - " + str(
                Round(zagruzka, 0)) + " % (" + str(I_maxx) + " А) от Iддтн = " + str(round(Ipr)) + " А и " + str(
                round(I_maxx / I_adtn * 100, 0)) + " % от Iадтн = " + str(round(I_adtn)) + " А"  # имя контр

    else:
        if name_zagruzka_ris = "":  # имя Загрузка для рисунков
            f_name_zagruzka_ris = "Загрузка " + name + dop_name + " составляет " + str(
                Round(zagruzka, 0)) + " % (" + str(I_maxx) + " А) от Iддтн = Iадтн = " + str(
                round(Ipr)) + " А"  # имя контр
        else:
            f_name_zagruzka_ris = name_zagruzka_ris + ", " + name + dop_name + " - " + str(
                Round(zagruzka, 0)) + " % (" + str(I_maxx) + " А) от Iддтн = Iадтн = " + str(
                round(Ipr)) + " А"  # имя контр# End def return


# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fword_Tag(word_doc, Text,
              Title):  # ф-я  ищет  в word_doc  ContentControls по Text , Title и возвращает Tag(id gruoup)
    for Each control_i in word_doc.ContentControls  # поменять рисунок
        if control_i.Title = Title and control_i.Range.Text = Text:
            fTestKluch = control_i.Tag
            exit
            for


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def fword_Text(word_doc, Tag, Title):  # ф-я  ищет  в word_doc  ContentControls по Tag , Title и возвращает Text
    for Each control_i in word_doc.ContentControls  # поменять рисунок
        if control_i.Title = Title and control_i.Tag = Tag:
            fword_Text = control_i.Range.Text
            exit
            for


# End def return
# *******************РАСЧЕТ РЕЖИМОВ****************************************************************************************************************************
def Dict_unik_Tag_doc_sub(Dict, doc):  # наполняет Dict  уникальными значениями ключей
    Dict.RemoveAll()
    for Each control in doc.ContentControls
        Dict_kluch = control.Tag
        if not Dict.Exists(Dict_kluch):  # Exists проверяет наличие ключа
            Dict.Add(Dict_kluch, 1)


# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в расчет +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def Sheets_add(Book_Excel, set_Sheets, Name_Sheets):  # добавить лист
    Book_Excel.Sheets.Add()
    set_Sheets = Book_Excel.Worksheets(1)
    Book_Excel.Worksheets(1).Name = Name_Sheets


# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в расчет +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def AutoShunt_class_rec(vibor):  # процедура формирует Umin , Umax, AutoBsh , nBsh
    # dim ndxUzel , ASC , AutoBshStr , U_U #  ASC -  Auto Shunt class
    TabNode = rastr.Tables("node")
    TabNode.setsel(vibor)  # "Kontrol"
    if TabNode.count > 0:
        RG.Dict_AutoShunt = CreateObject("Scripting.Dictionary")  # для хранения AutoShunt_class
        ndxUzel = TabNode.FindNextSel(-1)
        while ndxUzel >= 0
            if Replace(TabNode.cols.item("AutoBsh").ZS(ndxUzel), " ", "") != "":  # есть автоматика
                ASC = AutoShunt_class
                ASC.ndx = ndxUzel
                ASC.ny = TabNode.cols.item("ny").Z(ndxUzel)
                AutoBshStr = Replace(TabNode.cols.item("AutoBsh").ZS(ndxUzel), " ", "")  # 654(110-125)*3
                AutoBshStr = Replace(AutoBshStr, ",", ".")
                if instr(AutoBshStr, "(") > 0:
                    ASC.AutoBsh = float(mid(AutoBshStr, 1, instr(AutoBshStr, "(") - 1))
                    U_U = mid(AutoBshStr, instr(AutoBshStr, "(") + 1,
                              instr(AutoBshStr, ")") - instr(AutoBshStr, "(") - 1)
                    ASC.Umin = float(mid(U_U, 1, instr(U_U, "-") - 1))
                    ASC.Umax = float(mid(U_U, instr(U_U, "-") + 1, len(U_U) - instr(U_U, "-")))
                else:
                    if instr(AutoBshStr, "*") > 0:
                        ASC.AutoBsh = float(mid(AutoBshStr, 1, instr(AutoBshStr, "*") - 1)) else:
                        ASC.AutoBsh = float(AutoBshStr)
                    if TabNode.cols.item("uhom").ZN(ndxUzel) > 300: ASC.Umin = TabNode.cols.item("uhom").ZN(
                        ndxUzel) * 0.95: ASC.Umax = TabNode.cols.item("uhom").ZN(ndxUzel) * 1.05
                    if TabNode.cols.item("uhom").ZN(ndxUzel) < 300: ASC.Umin = TabNode.cols.item("uhom").ZN(
                        ndxUzel) * 0.95: ASC.Umax = TabNode.cols.item("uhom").ZN(ndxUzel) * 1.15

                if instr(AutoBshStr, "*") > 0:
                    ASC.nBsh = float(
                        mid(AutoBshStr, instr(AutoBshStr, "*") + 1, len(AutoBshStr) - instr(AutoBshStr, "*"))) else:
                    ASC.nBsh = 1
                ASC.list_info = ""
                ASC.nyk = ftest_skrm(TabNode.cols.item("ny").ZN(
                    ndxUzel))  # если в узеле нет нагрузки и генерации, он подключен через одну ветвь - выключатель или малое сопротивлние и узел конца sta=0  то возвращает номер узла в противном случае - 0

                RG.Dict_AutoShunt.Add(ASC.ny, ASC)  # ключ  и значение

            ndxUzel = TabNode.FindNextSel(ndxUzel)
        wend


# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в расчет +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def AutoShunt_class_kor():  # процедура меняет Bsh  и записывает GS.AutoShunt_list
    # dim ASC
    GS.AutoShunt_list = ""
    if not isempty(RG.Dict_AutoShunt):
        for EACH ASC in RG.Dict_AutoShunt.Items
            ASC.kor_Bsh()
            if not ASC.list_info = "":
                GS.AutoShunt_list = GS.AutoShunt_list + "; " + ASC.list_info
                logging.info("\t" + "\t" + "AutoShunt_list: " + ASC.list_info)
                ASC.list_info = ""


# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в расчет +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
class AutoShunt_class:  # АВТО ШУНТ автоматика включения Bsh нр "-3500(120-124)*2" те БСК две по -3500, вкл при 120 кВ откл при 124 кВ
    # dim  nn , tek_Bsh, ish_Bsh , ot_for, do_for, st_for, vrs , ish_vrs
    # dim TabNode , ish_sta
    # dim Umin , Umax , AutoBsh , nBsh , list_info , ndx, ny, nyk , nyk_ndx#  если nyk > 0 то СКРМ задано в отдельном узле, если 0 то нет отдельного узла

    def kor_Bsh():  # процедура меняет Bsh и записывает list_info
        TabNode = rastr.Tables("node")
        ish_sta = TabNode.cols.item("sta").Z(ndx)

        if nyk > 0:  # эта скрм на отдельном узле через выключатель
            nyk_ndx = fNDX("node", nyk)
            vrs = TabNode.cols.item("vras").ZN(nyk_ndx)
        else:
            vrs = TabNode.cols.item("vras").ZN(ndx)

        ish_Bsh = TabNode.cols.item("bsh").ZN(ndx)
        ish_vrs = vrs
        if vrs > Umax:  # нада ВКЛЮЧИТЬ РЕАКТОРЫ или ОТКЛЮЧИТЬ БСК
            if AutoBsh < 0 and ish_sta = True: exit

            def  # если напряжение выше макс,  это БСК   и она отключена то выход

                if ish_sta = True and AutoBsh > 0: sta_node(TabNode.cols.item("ny").ZS(ndx), False): TabNode.cols.item(
                    "bsh").Z(
                    ndx) = 0  # если ШР, он отключен то вкл узел с реактором включить False; отключить True  и обнулить шунт

            if AutoBsh > 0: ot_for = 1: do_for = nBsh: st_for = 1  # РЕАКТОРЫ ВКЛЮЧИТЬ
            if AutoBsh < 0: ot_for = nBsh - 1: do_for = 0: st_for = -1  # БСК ОТКЛЮЧИТЬ

            for nn = ot_for to do_for step st_for
            if TabNode.cols.item("bsh").ZN(ndx) < AutoBsh * nn:  # если уже включено меньше чем  даст текущая ступень то
                tek_Bsh = TabNode.cols.item("bsh").ZN(ndx)
                if nn = 0 and nyk > 0: TabNode.cols.item("sta").Z(ndx) = True else:
                    TabNode.cols.item("bsh").ZN(ndx) = AutoBsh * nn  # отключить если это последняя ступеть БСК
                GS.kod_rgm = rastr.rgm("")
                if GS.kod_rgm = 1:
                    GS.kod_rgm = rastr.rgm("")
                if GS.kod_rgm = 1:
                    GS.kod_rgm = rastr.rgm("p")
                if GS.kod_rgm = 1:
                    GS.kod_rgm = rastr.rgm("p")
                if GS.kod_rgm = 1:
                    GS.kod_rgm = rastr.rgm("p")
                vrs = TabNode.cols.item("vras").ZN(ndx)
                if vrs < Umin:  TabNode.cols.item("bsh").ZN(ndx) = tek_Bsh: exit
                for
                    if vrs < Umax:  exit
                for

    elif vrs < Umin and vrs > 0:  # ОТКЛЮЧИТЬ  РЕАКТОРЫ или ВКЛЮЧИТЬ БСК
    if AutoBsh > 0 and ish_sta = True: exit

    def  # если это ШР напряжение ниже уставки  и он отключена то выход

        if ish_sta = True and AutoBsh < 0: sta_node(TabNode.cols.item("ny").ZS(ndx), False): TabNode.cols.item("bsh").Z(
            ndx) = 0  # вкл узел с бск

    if AutoBsh > 0: ot_for = nBsh - 1: do_for = 0: st_for = -1  # РЕАКТОРЫ ОТКЛЮЧИТЬ
    if AutoBsh < 0: ot_for = 1: do_for = nBsh: st_for = 1  # БСК ВКЛЮЧИТЬ

    for nn = ot_for to do_for step st_for
    if TabNode.cols.item("bsh").ZN(ndx) > AutoBsh * nn:  # если  уже включено больше чем  даст текущая ступень то
        tek_Bsh = TabNode.cols.item("bsh").ZN(ndx)
        if nn = 0 and nyk > 0: TabNode.cols.item("sta").Z(ndx) = True else:
            TabNode.cols.item("bsh").ZN(ndx) = AutoBsh * nn  # включить
        GS.kod_rgm = rastr.rgm("")
        if GS.kod_rgm = 1:
            GS.kod_rgm = rastr.rgm("")
        if GS.kod_rgm = 1:
            GS.kod_rgm = rastr.rgm("p")
        if GS.kod_rgm = 1:
            GS.kod_rgm = rastr.rgm("p")
        if GS.kod_rgm = 1:
            GS.kod_rgm = rastr.rgm("p")
        vrs = TabNode.cols.item("vras").ZN(ndx)
        if vrs > Umax:  TabNode.cols.item("bsh").ZN(ndx) = tek_Bsh: exit
        for
            if vrs > Umin:  exit
        for


if (not ish_sta = TabNode.cols.item("sta").Z(ndx)) or ( not ish_Bsh = TabNode.cols.item("bsh").ZN(ndx)):
    if nyk > 0:
        vrs = TabNode.cols.item("vras").ZN(nyk_ndx) else:
        vrs = TabNode.cols.item("vras").ZN(ndx)
    list_info = list_info + TabNode.cols.item("name").ZS(ndx)
    if nyk > 0: list_info = list_info + "/" + TabNode.cols.item("name").ZN(nyk_ndx) + "/"
    list_info = list_info + " (" + str(ny) + ",U=" + str(round(ish_vrs)) + "/" + str(Umin) + "-" + str(Umax) + "/" + ")"

    if (not ish_Bsh = TabNode.cols.item("bsh").ZN(ndx)): list_info = list_info + ", изм с bsh=" + str(
        ish_Bsh) + " на " + str(TabNode.cols.item("bsh").ZN(ndx))

    if (not ish_sta = TabNode.cols.item("sta").Z(ndx)):  # изм состояния узла
        if TabNode.cols.item("sta").Z(ndx):
            list_info = list_info + ", узел отключен"
        else:
            list_info = list_info + ", узел включен"
            list_info = list_info + " (Uрез=" + str(round(vrs)) + "). "  # + vbCrLf

vrs = 0


# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в расчет +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def ftest_skrm(
        nyT):  # если в узеле нет нагрузки и генерации, он подключен через одну ветвь - выключатель или малое сопротивлние и узел конца sta=0  то возвращает номер узла в противном случае - 0
    # dim tNodeT, tVetvT, ndx_node1, ndx_vetvT, ndx_node2
    ftest_skrm = 0
    tNodeT = rastr.Tables("node")
    tVetvT = rastr.Tables("vetv")
    ndx_node1 = fNDX("node", nyT)
    tVetvT.setsel("ip=" + str(nyT) + "|iq=" + str(nyT))  # выбор примыкающих ветвей
    ndx_vetvT = tVetvT.FindNextSel(-1)
    if tVetvT.count = 1 and tNodeT.cols.item("pn").Z(ndx_node1)+tNodeT.cols.item("qn").Z(ndx_node1)+tNodeT.cols.item("pg").Z(ndx_node1) = 0:
        if tVetvT.cols.item("r").Z(ndx_vetvT) + tVetvT.cols.item("x").Z(ndx_vetvT) < 0.2:
            if nyT = tVetvT.cols.item("ip").Z(ndx_vetvT): ftest_skrm = tVetvT.cols.item("iq").Z(ndx_vetvT)
            if nyT = tVetvT.cols.item("iq").Z(ndx_vetvT): ftest_skrm = tVetvT.cols.item("ip").Z(ndx_vetvT)
            ndx_node2 = fNDX("node", ftest_skrm)
            if tNodeT.cols.item("sta").Z(ndx_node2): ftest_skrm = 0  # End def return


# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в расчет +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def fNDX(tabl, kkluch):  # возвращает номер строки ("таблица", "краткий ключ")
    # dim tTabl , kkluchM
    if not tabl = "" and not kkluch = "":
        tTabl = rastr.Tables(tabl)
        if tabl = "node":
            tTabl.setsel("ny=" + str(kkluch))
        elif tabl = "vetv":
            kkluchM = split(str(kkluch), ",")
            tTabl.setsel("ip=" + str(kkluchM(0)) + "+iq=" + str(kkluchM(1)) + "+np=" + str(kkluchM(2)))
        elif tabl = "area":
            tTabl.setsel("na=" + str(kkluch))
        elif tabl = "area2":
            tTabl.setsel("npa=" + str(kkluch))
        elif tabl = "ngroup":
            tTabl.setsel("nga=" + str(kkluch))
        elif tabl = "Generator":
            tTabl.setsel("Num=" + str(kkluch))
        elif tabl = "sechen":
            tTabl.setsel("ns=" + str(kkluch))

        fNDX = tTabl.FindNextSel(-1)
        if fNDX == -1:
            logging.info("\t" + "fNDX в таблице " + tabl + " не найдено " + str(kkluch))  #: exit  def
    else:
        fNDX = -1
        logging.info("\t" + "fNDX не правельное  задание: " + tabl + " не найдено " + str(kkluch))


# End def return
# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в расчет +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def fNDX_p(tabl, kluch):  # возвращает номер строки (таблица, полный ключ типа ip=251+iq=256+np=0)
    tTabl = rastr.Tables(tabl)
    tTabl.setsel(kluch)
    fNDX_p = tTabl.FindNextSel(-1)
    if fNDX_p = -1: logging.info("\t" + "fNDX_p в таблице " + tabl + " не найдено " + str(kluch))


# End def return
# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в расчет +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def fParam_str(arr_flag, arr_z):  # склеить стоку параметров
    # dim iii
    fParam_str = ""
    for iii = 0 to ubound ( arr_flag )
    if arr_flag(iii):
        fParam_str = fParam_str + arr_z(iii) + ","


fParam_str = Left(fParam_str, Len(fParam_str) - 1)


# End def return
# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в расчет +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def Folder_add_sub(add_Folder):  # создать папку
    if Not objFSO.FolderExists (add_Folder):  # если каталог add_Folder не существует создаем его
        objFSO.CreateFolder(add_Folder)


# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в расчет +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def fParam(tabl, param, index):  # возвращает нужный параметр из нужной таблице
    return rastr.Tables(tabl).cols.item(param).Z(index)


# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def add_AutoBsh(viborka):  # процедура записывает из поля bsh в поле AutoBsh
    # dim Nod , ndxUzel , U
    Nod = rastr.tables("node")
    Nod.setsel("bsh>0|bsh<0")
    ndxUzel = Nod.FindNextSel(-1)
    while ndxUzel >= 0
        if Nod.cols.item("AutoBsh").Z(ndxUzel) = "":
            if Nod.cols.item("uhom").ZN(ndxUzel) > 300: U = str(
                round(Nod.cols.item("uhom").ZN(ndxUzel) * 0.98, 1)) + "-" + str(
                round(Nod.cols.item("uhom").ZN(ndxUzel) * 1.05, 1))
            if Nod.cols.item("uhom").ZN(ndxUzel) < 300: U = str(
                round(Nod.cols.item("uhom").ZN(ndxUzel) * 0.95, 1)) + "-" + str(
                round(Nod.cols.item("uhom").ZN(ndxUzel) * 1.15, 1))
            Nod.cols.item("AutoBsh").Z(ndxUzel) = str(Nod.cols.item("bsh").Z(ndxUzel) * 1000000) + "(" + U + ")"
            logging.info(Nod.cols.item("name").ZS(ndxUzel) + ": " + Nod.cols.item("AutoBsh").ZS(ndxUzel))

        ndxUzel = Nod.FindNextSel(ndxUzel)
    wend


# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def vetv_vikl_add(viborka):  # для ветвей добавить выкл в начале и в конце
    # dim tV28  , tN28
    tN28 = rastr.Tables("node")
    tV28 = rastr.Tables("vetv")

    tV28.setsel(viborka)  # ВЕТВИ
    ndx_tV = tV28.FindNextSel(-1)
    while ndx_tV >= 0  #
        ny_n = tV28.cols.item("ip").Z(ndx_tV)
        ny_k = tV28.cols.item("iq").Z(ndx_tV)
        uhom = tN28.cols.item("uhom").Z(fNDX("node", str(ny_n)))
        na_n = tN28.cols.item("na").Z(fNDX("node", str(ny_n)))
        na_k = tN28.cols.item("na").Z(fNDX("node", str(ny_k)))
        npa_n = tN28.cols.item("npa").Z(fNDX("node", str(ny_n)))
        npa_k = tN28.cols.item("npa").Z(fNDX("node", str(ny_k)))
        # имя начала конца
        if tV28.cols.item("dname").Z(ndx_tV) != "":
            name_VL = tV28.cols.item("dname").Z(ndx_tV) else:
            name_VL = tV28.cols.item("name").Z(ndx_tV)
        name_VL_n = "В-" + str(tN28.cols.item("uhom").ZS(fNDX("node", str(ny_n)))) + " кВ на " + tN28.cols.item(
            "name").ZS(fNDX("node", str(ny_n))) + " в цепи " + name_VL
        name_VL_k = "В-" + str(tN28.cols.item("uhom").ZS(fNDX("node", str(ny_k)))) + " кВ на " + tN28.cols.item(
            "name").ZS(fNDX("node", str(ny_k))) + " в цепи " + name_VL
        # добавить узлы
        ny_new_n = fNode_add(name_VL_n, na_n, npa_n, uhom, 0)  # (name , na , npa ) #  добавить узел и вернуть номер
        ny_new_k = fNode_add(name_VL_k, na_n, npa_k, uhom, 0)  # (name , na , npa ) #  добавить узел и вернуть номер
        # поменять номера начала и конца
        tV28.cols.item("ip").Z(ndx_tV) = ny_new_n
        tV28.cols.item("iq").Z(ndx_tV) = ny_new_k
        # добавить ВЛ сначала и конца
        ndx_vetv_n = fVetv_add_ndx(name_VL_n, ny_n, ny_new_n, 0, 0.0, 0,
                                   0)  # (dname , ip , iq , np ) #  добавить ветвь и вернуть ndx
        ndx_vetv_k = fVetv_add_ndx(name_VL_k, ny_k, ny_new_k, 0, 0.0, 0,
                                   0)  # (dname , ip , iq , np ) #  добавить ветвь и вернуть ndx
        ndx_tV = tV28.FindNextSel(ndx_tV)
    wend
    logging.info("\t" + "vetv_vikl_add / " + viborka)


def node_ku_add(viborka):  # к узлам присоединить новый узел и перенести ШР БСК УШР
    # dim tN29
    tN29 = rastr.Tables("node")
    tN29.setsel(viborka)  # УЗЛЫ
    ndx_tN = tN29.FindNextSel(-1)
    while ndx_tN >= 0  #
        ny = tN29.cols.item("ny").Z(ndx_tN)
        uhom = tN29.cols.item("uhom").Z(ndx_tN)
        na = tN29.cols.item("na").Z(ndx_tN)
        npa = tN29.cols.item("npa").Z(ndx_tN)
        # имя
        name = tN29.cols.item("name").ZS(ndx_tN) + " "
        if tN29.cols.item("bsh").Z(ndx_tN) > 0:
            name = name + "ШР"
            if tN29.cols.item("qmin").Z(ndx_tN) < 0: name = name + ", УШР"
        elif tN29.cols.item("bsh").Z(ndx_tN) < 0:
            name = name + "БСК"
            if tN29.cols.item("qmin").Z(ndx_tN) < 0: name = name + ", УШР"
        else:
            if tN29.cols.item("qmin").Z(ndx_tN) < 0: name = name + " УШР"

        # добавить узел
        ny_new = fNode_add(name, na, npa, uhom, 0)  # (name , na , npa , uhom) #  добавить узел и вернуть номер

        tN29.cols.item("bsh").Z(fNDX("node", str(ny_new))) = tN29.cols.item("bsh").Z(ndx_tN)
        tN29.cols.item("vzd").Z(fNDX("node", str(ny_new))) = tN29.cols.item("vzd").Z(ndx_tN)
        tN29.cols.item("qmin").Z(fNDX("node", str(ny_new))) = tN29.cols.item("qmin").Z(ndx_tN)
        tN29.cols.item("bsh").Z(ndx_tN) = 0
        tN29.cols.item("vzd").Z(ndx_tN) = 0
        tN29.cols.item("qmin").Z(ndx_tN) = 0
        tN29.cols.item("qg").Z(ndx_tN) = 0

        # добавить ВЛ
        ndx_vetv = fVetv_add_ndx(name, ny_new, ny, 0, 0.01, 0,
                                 0)  # (dname , ip , iq , np ) #  добавить ветвь и вернуть ndx
        ndx_tN = tN29.FindNextSel(ndx_tN)
    wend
    logging.info("\t" + "node_ku_add / " + viborka)


# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def sel_ssh2_add():  # к отмеченным узлам присоединить новый узел через выключатель и перенести верви с np=2,4,6
    rastr.RenumWP = True
    # dim tV29  , tN30
    tN30 = rastr.Tables("node")
    tV29 = rastr.Tables("vetv")

    tN30.setsel("sel")  # УЗЛЫ
    ndx_tN = tN30.FindNextSel(-1)
    while ndx_tN >= 0  #
        ny = tN30.cols.item("ny").Z(ndx_tN)
        uhom = tN30.cols.item("uhom").Z(ndx_tN)
        na = tN30.cols.item("na").Z(ndx_tN)
        npa = tN30.cols.item("npa").Z(ndx_tN)
        name = "2сш " + str(uhom) + " кВ " + tN30.cols.item("name").ZS(ndx_tN)
        ny_new = fNode_add(name, na, npa, uhom, 0)  # (name , na , npa , uhom) #  добавить узел и вернуть номер
        ndx_vetv = fVetv_add_ndx("ШСВ " + str(uhom) + " кВ " + tN30.cols.item("name").ZS(ndx_tN), ny_new, ny, 0, 0, 0,
                                 0)  # (dname , ip , iq , np ) #  добавить ветвь и вернуть ndx
        tV29.setsel("(ip=" + str(ny) + "|iq=" + str(ny) + ")+(np=2|np=4|np=6)")  # ВЕТВИ
        ndx_tV = tV29.FindNextSel(-1)
        while ndx_tV >= 0  #
            if ny = tV29.cols.item("ip").Z(ndx_tV):  cor(tV29.SelString(ndx_tV), "ip=" + str(ny_new))
            if ny = tV29.cols.item("iq").Z(ndx_tV): cor(tV29.SelString(ndx_tV), "iq=" + str(ny_new))
            ndx_tV = tV29.FindNextSel(ndx_tV)
        wend
        tN30.cols.item("name").Z(ndx_tN) = "1сш " + str(uhom) + " кВ " + tN30.cols.item("name").ZS(ndx_tN)
        ndx_tN = tN30.FindNextSel(ndx_tN)
    wend
    logging.info("\t" + "sel_ssh2_add ()")
    rastr.RenumWP = False


# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def fVetv_add_ndx(dname, ip, iq, np, r, x, b):  # добавить ветвь и вернуть ndx
    tV_fVetv_add = rastr.tables("vetv")
    tV_fVetv_add.AddRow
    fVetv_add_ndx = tV_fVetv_add.size - 1

    tV_fVetv_add.cols.item("dname").Z(fVetv_add_ndx) = dname
    tV_fVetv_add.cols.item("ip").Z(fVetv_add_ndx) = ip
    tV_fVetv_add.cols.item("iq").Z(fVetv_add_ndx) = iq
    tV_fVetv_add.cols.item("np").Z(fVetv_add_ndx) = np
    tV_fVetv_add.cols.item("r").Z(fVetv_add_ndx) = r
    tV_fVetv_add.cols.item("x").Z(fVetv_add_ndx) = x
    tV_fVetv_add.cols.item("b").Z(fVetv_add_ndx) = b

    logging.info("\t" + "добавлен узел ip=" + str(ip) + ", iq=" + str(ip) + ", np=" + str(np) + ", dname=" + dname)


# End def return
# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

def fNode_add(name, na, npa, uhom,
              ny_zad):  # добавить узел и вернуть номер (name , na , npa , uhom , ny_zad или 0 - max 1)
    tN_fNode_add = rastr.tables("node")
    tN_fNode_add.AddRow
    if ny_zad = 0: fNode_add = rastr.Calc("max", "node", "ny", "ny>0") + 1 else:
        fNode_add = ny_zad
    tN_fNode_add.cols.item("ny").Z(tN_fNode_add.size - 1) = fNode_add
    tN_fNode_add.cols.item("name").Z(tN_fNode_add.size - 1) = name
    tN_fNode_add.cols.item("na").Z(tN_fNode_add.size - 1) = na
    tN_fNode_add.cols.item("npa").Z(tN_fNode_add.size - 1) = npa
    tN_fNode_add.cols.item("uhom").Z(tN_fNode_add.size - 1) = uhom
    logging.info("\t" + "добавлен узел ny=" + str(fNode_add) + ", name=" + name + ", uhom=" + str(uhom))


# End def return


# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def Tc_0_Sub():  # обнулить расчетную температуру в ветвях районах и тд
    rastr.Tables("vetv").cols.item("Tc").Calc("0")
    rastr.Tables("area").cols.item("Tc").Calc("0")
    rastr.Tables("area2").cols.item("Tc").Calc("0")
    rastr.Tables("darea").cols.item("Tc").Calc("0")
    logging.info("\t" + "обнулена температура в ветвях, районах, территориях, объединениях")


# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def groupid_sel_sub():  # отметить groupid отмеченных узлов
    # dim tN32
    tN32 = rastr.Tables("node")
    N_groupid = rastr.Calc("max", "vetv", "groupid", "ip>0") + 1

    tN32.setsel("")
    if rastr.tables("node").cols.Find("value") < 1: rastr.tables("node").Cols.Add
    "value", 1  # добавить столбцы
    tN32.cols.item("value").calc("sel")

    tN32.setsel("sel")
    ndx_tNode = tN32.FindNextSel(-1)

    while ndx_tNode >= 0  #
        ny = tN32.cols.item("ny").ZS(ndx_tNode)
        if tN32.cols.item("value").Z(ndx_tNode) = 1:
            ny_next_str = str(ny)
            while ny_next_str != ""  #
                ny_next_arr = split(ny_next_str, ",")
                ny_next_str = ""

                for each ny_next in ny_next_arr
                    if ny_next > < "":
                        ny_next_str = ny_next_str + str(fNextNy(ny_next, N_groupid))
                        tN32.cols.item("value").Z(fNDX("node", float(ny_next))) = 0

            wend
            N_groupid = N_groupid + 1

    # rastr.tables("vetv").setsel ("groupid=" + N_groupid)
    # rastr.tables("vetv").cols.item("value").calc (0)
    ndx_tNode = tN32.FindNextSel(ndx_tNode)


wend

# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def fNextNy(ny, id):  # для  groupid_sel_sub #
    # dim ny_str, ndx  ,   tV39
    fNextNy = ""
    tV39 = rastr.Tables("vetv")

    ny_str = ""
    tV39.setsel("(ip=" + str(ny) + "|iq=" + str(ny) + ")")  # +sta=0")
    tV39.cols.item("groupid").calc(str(id))

    tV39.setsel("(ip=" + str(ny) + "|iq=" + str(ny) + ")+ip.value=1+iq.value=1")  # +sta=0"  )

    if tV39.count > 0:

        ndx = tV39.FindNextSel(-1)
        while ndx >= 0  #
            if tV39.cols.item("ip").ZS(ndx) = ny: ny_str = ny_str + "," + tV39.cols.item("iq").ZS(ndx)
            if tV39.cols.item("iq").ZS(ndx) = ny: ny_str = ny_str + "," + tV39.cols.item("ip").ZS(ndx)

            ndx = tV39.FindNextSel(ndx)
        wend

    fNextNy = ny_str


# End def return
# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def Qgen_node_in_gen_sub():  # посчитать Q ГЕН по  Q в узле
    # dim ndx , tN33
    tN33 = rastr.Tables("node")

    tG = rastr.tables("Generator")


tN33.setsel("qmin>0|qmax>0")
ndx = tN33.FindNextSel(-1)
while ndx >= 0  #
    qmin = tN33.cols.item("qmin").ZN(ndx)
    qmax = tN33.cols.item("qmax").ZN(ndx)
    pg = tN33.cols.item("pg").ZN(ndx)
    if pg != 0:
        tG.setsel("Node=" + tN33.cols.item("ny").ZS(ndx))
        ndx_g = tG.FindNextSel(-1)
        while ndx_g >= 0  #
            if tG.cols.item("NumPQ").ZN(ndx_g) = 0:
                if tG.cols.item("Qmax").ZN(ndx_g) = 0 and qmax != 0:
                    tG.cols.item("Qmax").ZN(ndx_g) = qmax * tG.cols.item("P").ZN(
                        ndx_g) / pg  #: logging.info ("узел " + tN33.cols.item("ny").ZN(ndx)  + " ген " + tG.cols.item("Num").ZN(ndx_g)   +  " qmax " +  tG.cols.item("Qmax").ZN(ndx_g))

                if tG.cols.item("Qmin").ZN(ndx_g) = 0 and qmin != 0:
                    tG.cols.item("Qmin").ZN(ndx_g) = qmin * tG.cols.item("P").ZN(
                        ndx_g) / pg  #: logging.info ("узел " + tN33.cols.item("ny").ZN(ndx)  + " ген " + tG.cols.item("Num").ZN(ndx_g)   +  " qmin " +  tG.cols.item("Qmin").ZN(ndx_g))                ndx_g = tG.FindNextSel(ndx_g)
        wend

    ndx = tN33.FindNextSel(ndx)
wend

# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def Delet_node_VL_sub():  # ТКЗ#  #  удалить промежкточные точки на ЛЭП при отсутствии магнитной связи
    # dim tV30  , tN34
    tN34 = rastr.Tables("node")
    tV30 = rastr.Tables("vetv")
    sel0()
    count = 0

    tN34 = rastr.tables("node")


tN34.setsel("")
ndx_tN = tN34.FindNextSel(-1)

while ndx_tN >= 0  # по узлам
    if Left(tN34.cols.item("name").ZS(ndx_tN), 2) = "ПТ":

        ny = tN34.cols.item("ny").Z(ndx_tN)

        tV30.setsel("(ip=" + str(ny) + "|iq=" + str(ny))

        if tV30.count = 2:
            tV30.setsel("(ip=" + str(ny) + "|iq=" + str(ny) + ")+MiGr<1+tip=0")
            if tV30.count = 2:
                tN34.cols.item("sel").Z(ndx_tN) = 1  # отметить узлы для удаления

                Rekv = 0: Rekv0 = 0
                Xekv = 0: Xekv0 = 0
                Bekv = 0: Bekv0 = 0
                NPekv = 0
                ip = 0: iq = 0
                vetv_tek = 0
                ndx_tV = tV30.FindNextSel(-1)
                while ndx_tV >= 0  # по ветвям
                    if vetv_tek = 0:

                        if tV30.cols.item("ip").Z(ndx_tV) = ny: ip = tV30.cols.item("iq").Z(ndx_tV) else:
                            ip = tV30.cols.item("ip").Z(ndx_tV)
                    elif vetv_tek = 1:
                        if tV30.cols.item("ip").Z(ndx_tV) = ny: iq = tV30.cols.item("iq").Z(ndx_tV) else:
                            iq = tV30.cols.item("ip").Z(ndx_tV)

                    Rekv = Rekv + tV30.cols.item("r").Z(ndx_tV)
                    Rekv0 = Rekv0 + tV30.cols.item("r0").Z(ndx_tV)
                    Xekv = Xekv + tV30.cols.item("x").Z(ndx_tV)
                    Xekv0 = Xekv0 + tV30.cols.item("x0").Z(ndx_tV)
                    Bekv = Bekv + tV30.cols.item("b").Z(ndx_tV)
                    Bekv0 = Bekv0 + tV30.cols.item("b0").Z(ndx_tV)
                    np = tV30.cols.item("np").Z(ndx_tV)

                    tV30.cols.item("sel").Z(ndx_tV) = 1  # отметить ветви для удаления
                    vetv_tek = vetv_tek + 1
                    ndx_tV = tV30.FindNextSel(ndx_tV)
                wend

                # добавить ветвь
                ndx_tV_new = fVetv_add_ndx("", ip, iq, np, Rekv, Xekv, Bekv)  # dname , ip , iq , np , r , x , b
                tV30.cols.item("r0").Z(ndx_tV_new) = Rekv0
                tV30.cols.item("x0").Z(ndx_tV_new) = Xekv0
                tV30.cols.item("b0").Z(ndx_tV_new) = Bekv0

                count = count + 1
                logging.info(str(ny) + tN34.cols.item("name").ZS(ndx_tN))
    ndx_tN = tN34.FindNextSel(ndx_tN)
wend
Del_sel()
logging.info("эквивалентировано узлов: " + str(count))






# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def np_zad_sub():
    # dim ndx
    # dim tV33
    tV33 = rastr.Tables("vetv")
    c_ip = tV33.cols.item("ip")
    c_iq = tV33.cols.item("iq")
    c_np = tV33.cols.item("np")
    nv = tV33.Size - 1
    logging.info("\t" + "задать номер паралельности у ветвей")
    for i = 0 to nv
    ip = c_ip.Z(i)
    iq = c_iq.Z(i)
    np = c_np.Z(i)
    if np = 0:
        tV33.setsel("ip=" + str(ip) + "+ iq=" + str(iq) + " + np=0")
        ndx = tV33.FindNextSel(-1)
        if tV33.count > 1:
            np_i = 1
            while ndx >= 0  #
                c_np.Z(ndx) = np_i
                np_i = np_i + 1
                logging.info("\t" + "\t" + "задан np ветви " + str(ip) + "," + str(iq) + "," + c_np.ZS(
                    ndx) + "-" + tV33.cols.item("name").ZS(ndx))
                ndx = tV33.FindNextSel(ndx)
            wend

        # +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


def rastr_xl_tab(tab, export_xl_on, book_str, sheet_name, tip_export_xl):
    #  (таблица растр "graphik2", True в растр False наоборот,  книга "D:\ген.xlsx", лист "rastr", 1 загрузить 0 присоединить 2 обновить)
    # dim BD_table, str_param , book , sheet ,  Excel_tab , table_kor , CSVfilename , CSVtext , tab_col_all , tab_col_all_arr
    table_kor = rastr.tables(tab)
    Excel_tab = CreateObject("Excel.Application")
    # Excel_tab.Visible = True
    if export_xl_on:  # 1  export из CS.excel в растр
        logging.info("\t" + "экспорт из книги: " + book_str + ", листа:" + sheet_name + " в rastrwin")
        if not objFSO.FileExists(book_str): msgbox(book_str + " - не найден файл"): exit

        def

            book = Excel_tab.Workbooks.Open(book_str)
        if not SheetExists(book, sheet_name): msgbox(sheet_name + " - не найден лист"): exit

        def

            sheet = book.Sheets(sheet_name)

        BD_table = sheet.UsedRange.Value  # номирация начинается с 1
        str_param = fArrCSV(BD_table, ",", 1,
                            1)  # массив, разделитель , номер первой строки  и последней или 0- последняя
        logging.info("\t" + "строка параметров: " + str_param)
        CSVtext = fArrCSV(BD_table, ";", 2, 0)  #
        CSVfilename = book.Path + "\" +  sheet_name + ".csv
        "
        logging.info("\t" + "файл CSV: " + CSVfilename)
        SaveTXTfile_sub(CSVfilename, CSVtext)
        ImportCSV(CSVfilename, tab, str_param, tip_export_xl)
        book.close
        Excel_tab.quit
    else:  # из  растр в xl
        logging.info("\t" + "экспорт из rastrwin в новую книгу excel")
        Excel_tab.Visible = True
        book = Excel_tab.Workbooks.Add
        sheet = book.Worksheets(1)
        book.Worksheets(1).Name = RG.Name_Base + " (" + tab + ")"
        tab_col_all = all_cols(tab)
        BD_table = table_kor.writesafearray(tab_col_all, "000")
        tab_col_all_arr = split(tab_col_all, ",")
        # печать массива: лист ХL ,по X , по Y , массив , кол изм массива 1 или 2 , "гор" "верт" , "" или "vetv" ,"" или "name" , "" или "орвыаи " - произвольный текст
        Print_XL(sheet, 1, 1, tab_col_all_arr, 1, "гор", "", "", "")
        Print_XL(sheet, 1, 2, BD_table, 2, "верт", "", "", "")
        book.SaveAs(CS.KIzFolder + "\" + RG.Name_Base + "(" + tab + ").xlsm
        " , 52)

        # +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


def fArrCSV(arr, ColumnsSeparator, Nstr, Kstr):  # массив, разделитель , номер первой строки  и последней
    # dim buffer, i_z
    buffer = ""  #
    if Nstr = 0: Nstr = LBound(arr, 1)
    if Kstr = 0: Kstr = UBound(arr, 1)

    for i = Nstr To Kstr  # по строкам
    txt = ""
    for j = LBound(arr, 2) To UBound(arr, 2)  # по столцам
    i_z = str(arr(i, j))
    if i_z = "": i_z = "-"

    txt = txt + Replace(str(Replace(i_z, ";", "_")), ColumnsSeparator, "") + str(ColumnsSeparator)


fArrCSV = fArrCSV + txt
if Nstr < Kstr: fArrCSV = fArrCSV + vbNewLine
if Len(fArrCSV) > 50000: buffer = buffer + fArrCSV: fArrCSV = ""

fArrCSV = buffer + fArrCSV


# End def return
# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def SaveTXTfile_sub(filename, txt):
    ts = objFSO.CreateTextFile(filename, True)
    ts.Write(txt)
    ts.Close


# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
class uzel_mdp_class:  # храним параметры узла и колекцию  генераторов
    Private
    tN37, tG, ndx, deltaPG, pgGen, StaGen, kod

    # dim ny , DGen,rezerv_P_UP,rezerv_P_DOWN , tip #  tip  = 1 узел без ген, 2 с генераторами    0 не использовать узел
    # dim ndxNode , pg_max ,pg_min , dr_p , name, pgNode , up_pgen, txt #  True нада увеличить генерацияю в узле , False уменьшить
    def init():  # ИНИТ - записываем общие параметры pg_max pg_min tip UGMC.init конкретного узла
        DGen = CreateObject("Scripting.Dictionary")  # для хранения
        tN37 = rastr.tables("node")
        tG = rastr.tables("Generator")
        dr_p = tN37.cols.item("dr_p").Z(ndxNode)
        name = tN37.cols.item("name").ZS(ndxNode)
        txt = "\t" + "\t" + "узел " + str(ny) + ": " + name
        tG.setsel("Node=" + str(ny))
        if tG.count = 0:
            tip = 1  # узел без генераторов
            pg_max = tN37.cols.item("pg_max").Z(ndxNode)
            pg_min = tN37.cols.item("pg_min").Z(ndxNode)
        else:
            tip = 0  # не использовать узел если далее не найдем ген для коррекции (должны быть sel)
            tG.setsel("Node=" + str(ny) + "+sel")  # тк все генераторы дб отмечены, если не отмечен то не используем его
            ndx = tG.FindNextSel(-1)
            while ndx >= 0  # ЦИКЛ ген
                UGMC = uzel_gen_mdp_class
                UGMC.Num = tG.cols.item("Num").Z(ndx)
                UGMC.ndxGen = ndx
                UGMC.init()
                DGen.Add(UGMC.Num, UGMC)  # ключ  и значение
                tip = 2  # c ген
            ndx = tG.FindNextSel(ndx)
            wend

    def init_rezerv_P():  # РЕЗЕРВ конкретного узла
        rezerv_P_UP = 0:rezerv_P_DOWN = 0
        if tip  = 1:  # если нет генераторов в узле
            if pg_max > 0:
                rezerv_P_UP = pg_max - tN37.cols.item("pg").Z(ndxNode) else:
                logging.info("в узле " + str(ny) + " " + name + "не задано поле pg_max")
            rezerv_P_DOWN = tN37.cols.item("pg").Z(ndxNode)
        elif tip  = 2:  # если есть генераторы в узле
            for EACH UGMC in DGen.Items
                if UGMC.tip = 1:
                    UGMC.init_rezerv_P()
                    rezerv_P_UP = rezerv_P_UP + UGMC.rezerv_P_UP
                    rezerv_P_DOWN = rezerv_P_DOWN + UGMC.rezerv_P_DOWN

    def korr():  # КОРРР конкретного узла
        if tN37.cols.item("sta").Z(ndxNode): exit

        def

        # ny = fParam ("node","ny", ndxNode)
        pgNode = tN37.cols.item("pg").Z(ndxNode)
        if up_pgen:    deltaPG = KSC.KefPG_UP * rezerv_P_UP  # на сколько нужно измнеить генерацию в узле
        if not up_pgen:    deltaPG = pgNode * KSC.KefPG_Down
        if KSC.dr_p_keff =1: if
        abs(dr_p) > 0.5: deltaPG = deltaPG * abs(1 / dr_p) else: deltaPG = deltaPG * abs(1 / 0.5)
        if deltaPG = 0: exit

        def

            if KSC.nebalans_izm_ps > 0:
                if KSC.nebalans_izm_ps > deltaPG:  KSC.nebalans_izm_ps = KSC.nebalans_izm_ps - deltaPG: deltaPG = 0
                if KSC.nebalans_izm_ps < deltaPG:  deltaPG = deltaPG - KSC.nebalans_izm_ps: KSC.nebalans_izm_ps = 0

        KSC.izm_ps = abs(KSC.izm_ps)

        if tip = 1 and KSC.tipIzm = 1:  # если нет генераторов в УЗЛЕ и учитывем узлы без генераторов
            if up_pgen:  # увеличиваем генерацию узла, KSC.KefPG_UP
                if pg_max > 0 and pg_max > pgNode:
                    if pg_min > pgNode + deltaPG:  # (от 0 до pg_min)
                        if pg_min > 0 and KSC.net_Pmin = 0:  # если есть Рмин и учитываем Рмин то
                            if KSC.izm_ps > pg_min:
                                tN37.cols.item("pg").Z(ndxNode) = pg_min
                                KSC.nebalans_izm_ps = KSC.nebalans_izm_ps + (pg_min - deltaPG)
                                KSC.izm_ps = KSC.izm_ps - pg_min

                        else:  # нет Рмин и не учитываем Рмин то
                            tN37.cols.item("pg").Z(ndxNode) = pgNode + deltaPG
                            KSC.izm_ps = KSC.izm_ps - deltaPG

                    elif pg_max > pgNode + deltaPG and (pg_min < pgNode + deltaPG or pg_min
                    = pgNode + deltaPG):  # (от pg_min (включительно) до pg_max)v
                        tN37.cols.item("pg").Z(ndxNode) = pgNode + deltaPG
                        KSC.izm_ps = KSC.izm_ps - deltaPG
                    elif pg_max < pgNode + deltaPG or pg_max = pgNode + deltaPG:  # (больше или равно pg_max)
                        tN37.cols.item("pg").Z(ndxNode) = pg_max
                        KSC.izm_ps = KSC.izm_ps - (pg_max - pgNode) else:  # снижаем генерацию узла,KefPG_Down
                    if pg_min < pgNode - deltaPG or pg_min = pgNode - deltaPG:  # (от pg_min (включительно) до pgNode)
                        tN37.cols.item("pg").Z(ndxNode) = pgNode - deltaPG
                        KSC.izm_ps = KSC.izm_ps - deltaPG
                    elif pg_min > pgNode - deltaPG and pgNode - deltaPG > 0:  # (от 0 до pg_min)
                        if pg_min > 0 and KSC.net_Pmin = 0:  # если есть Рмин и учитываем Рмин то
                            tN37.cols.item("pg").Z(ndxNode) = pg_min
                            KSC.izm_ps = KSC.izm_ps - (pgNode - pg_min)
                            deltaPG = deltaPG - (pgNode - pg_min)
                            if KSC.izm_ps > pg_min:
                                tN37.cols.item("sta").Z(ndxNode) = True
                                KSC.nebalans_izm_ps = KSC.nebalans_izm_ps + (pg_min - deltaPG)
                                KSC.izm_ps = KSC.izm_ps - pg_min

                        else:  # если  Рмин не учитываем
                            tN37.cols.item("pg").Z(ndxNode) = pgNode - deltaPG
                            KSC.izm_ps = KSC.izm_ps - deltaPG

                    elif pgNode - deltaPG < 0 or pgNode = deltaPG:  # (меньше или равно 0)
                        tN37.cols.item("pg").Z(ndxNode) = 0
                        KSC.izm_ps = KSC.izm_ps - pgNode
                        kod = rastr.rgm(""): if kod != 0: kod = rastr.rgm(""): if kod != 0: kod = rastr.rgm(
                            ""): if kod != 0: kod = rastr.rgm("p"):  if kod != 0: kod = rastr.rgm("p")
                if kod != 0 OR fKontrSech (ns)  = False:  # fKontrSech возвращает истина если мощность в сечениях отмеченных контроль (sta) не превышена   или ложь (исключение)
                    tN37.cols.item("pg").Z(ndxNode) = pgNode
                    logging.info("\t" + "Разваливается узел ny = " + str(ny) + " генерацию вернули назад")
                    kod = rastr.rgm(""): if kod != 0: kod = rastr.rgm(""): if kod != 0: kod = rastr.rgm(
                        ""): if kod != 0: kod = rastr.rgm("p"):  if kod != 0: kod = rastr.rgm("p")
                    if kod != 0: logging.info("\t" + "Аварийное завершение расчета режима, подвел узел ny = " + str(
                        ny)): KSC.inini = KSC.Ncikl
                    if kod != 0: exit

                    def

                # if KSC.print_gen = 1: logging.info (txt + ": увеличение на " + str(pg_max -  pgNode))
            elif tip = 2:  # если есть ГЕНЕРАТОРЫ в узле
                for EACH UGMC in DGen.Items
                    pgGen = tG.cols.item("P").Z(UGMC.ndxGen)  # запомнить начальное состояние
                    StaGen = tG.cols.item("sta").Z(UGMC.ndxGen)  # запомнить начальное состояние
                    if StaGen: pgGen = 0

                    if up_pgen and deltaPG != 0:  # если УВЕЛИЧИТЬ ГЕНЕРАЦИЮ
                        if UGMC.Pmax > 0 and pgGen < UGMC.Pmax:  # если задан  UGMC.Pmax
                            if pgGen + deltaPG < UGMC.Pmin:  # (от 0 до Pmin ) сравнивеем резерв снижения Р в узле и величину Р на которую нада изменить мощнсть узла
                                if UGMC.Pmin > 0 and KSC.net_Pmin = 0:  # если есть Pmin и учитываем ее
                                    if KSC.izm_ps > UGMC.Pmin:
                                        if StaGen: tG.cols.item("sta").Z(UGMC.ndxGen) = False
                                        tG.cols.item("P").Z(UGMC.ndxGen) = UGMC.Pmin
                                        KSC.nebalans_izm_ps = KSC.nebalans_izm_ps + (UGMC.Pmin - deltaPG)
                                        KSC.izm_ps = KSC.izm_ps - UGMC.Pmin
                                        deltaPG = 0

                                else:  # если нет Pmin или не  учитываем ее
                                    if StaGen: tG.cols.item("sta").Z(UGMC.ndxGen) = False
                                    tG.cols.item("P").Z(UGMC.ndxGen) = pgGen + deltaPG  #
                                    KSC.izm_ps = KSC.izm_ps - deltaPG
                                    deltaPG = 0  # если этого ген достаточно для снижения и другие ген трогать не нада

                            elif pgGen + deltaPG < UGMC.Pmax and (pgGen + deltaPG > UGMC.Pmin or pgGen + deltaPG
                            = UGMC.Pmin):  # (от Pmin (включительно) до  Pmax)
                                if StaGen: tG.cols.item("sta").Z(UGMC.ndxGen) = False
                                tG.cols.item("P").Z(UGMC.ndxGen) = pgGen + deltaPG
                                KSC.izm_ps = KSC.izm_ps - deltaPG
                                deltaPG = 0  # если этого ген достаточно для снижения и другие ген трогать не нада
                            elif pgGen + deltaPG > UGMC.Pmax or pgGen + deltaPG = UGMC.Pmax:  # (равно или больше Pmax)
                                if StaGen: tG.cols.item("sta").Z(UGMC.ndxGen) = False
                                tG.cols.item("P").Z(UGMC.ndxGen) = UGMC.Pmax
                                KSC.izm_ps = KSC.izm_ps - (UGMC.Pmax - pgGen)
                                deltaPG = deltaPG - (UGMC.Pmax - pgGen)

                        else:
                            if UGMC.Pmax = 0: logging.info(
                                "\t" + "Не задан Pmax у генератора " + str(UGMC.Num) + "в узле " + str(
                                    ny) + "исключен из рассматрения")

                    elif not up_pgen and deltaPG != 0:  # если СНИЖЕНИЕ ГЕНЕРАЦИИ
                        if StaGen = False:  # если ген включен
                            if pgGen - UGMC.Pmin > deltaPG or pgGen - UGMC.Pmin = deltaPG:  # (от Pmin(включительно) до pgGen) сравнивеем резерв снижения Р в узле и величину Р на которую нада изменить мощнсть узла
                                tG.cols.item("P").Z(UGMC.ndxGen) = pgGen - deltaPG  #
                                KSC.izm_ps = KSC.izm_ps - deltaPG
                                deltaPG = 0  # если этого ген достаточно для снижения и другие ген трогать не нада
                            elif pgGen < deltaPG or pgGen = deltaPG:  # (меньше или равно 0)
                                tG.cols.item("sta").Z(UGMC.ndxGen) = True
                                KSC.izm_ps = KSC.izm_ps - (pgGen)
                                deltaPG = deltaPG - pgGen
                            elif pgGen - deltaPG > 0 and UGMC.Pmin > pgGen - deltaPG:  # (от 0 до Pmin)
                                if UGMC.Pmin > 0 and KSC.net_Pmin = 0:  # если есть Pmin и учитываем ее
                                    tG.cols.item("P").Z(UGMC.ndxGen) = UGMC.Pmin
                                    KSC.izm_ps = KSC.izm_ps - (pgGen - UGMC.Pmin)
                                    deltaPG = deltaPG - (pgGen - UGMC.Pmin)
                                    if KSC.izm_ps > UGMC.Pmin:
                                        tG.cols.item("sta").Z(UGMC.ndxGen) = True
                                        KSC.nebalans_izm_ps = KSC.nebalans_izm_ps + (UGMC.Pmin - deltaPG)
                                        KSC.izm_ps = KSC.izm_ps - UGMC.Pmin

                                else:  # если нет Pmin или не  учитываем ее
                                    tG.cols.item("P").Z(UGMC.ndxGen) = pgGen - deltaPG  #
                                    KSC.izm_ps = KSC.izm_ps - deltaPG
                                    deltaPG = 0  # если этого ген достаточно для снижения и другие ген трогать не нада

                    kod = rastr.rgm(""): if kod != 0: kod = rastr.rgm(""): if kod != 0: kod = rastr.rgm(
                        ""): if kod != 0: kod = rastr.rgm("p"):  if kod != 0: kod = rastr.rgm("p")
                    if kod != 0 OR fKontrSech (ns)  = False:  # fKontrSech возвращает истина если мощность в сечениях отмеченных контроль (sta) не превышена   или ложь (исключение)
                        logging.info("\t" + "Разваоивается узел ny = " + str(ny) + " генерацию вернули назад")
                        tG.cols.item("P").Z(UGMC.ndxGen) = pgGen
                        tG.cols.item("sta").Z(UGMC.ndxGen) = StaGen
                        kod = rastr.rgm(""): if kod != 0: kod = rastr.rgm(""): if kod != 0: kod = rastr.rgm(
                            ""): if kod != 0: kod = rastr.rgm("p"):  if kod != 0: kod = rastr.rgm("p")
                        if kod != 0: logging.info(
                            "\t" + "Аварийное завершение расчета режима, подвел узел ny = " + str(ny) + "ген N= " + str(
                                Num)): KSC.inini = KSC.Ncikl
                        if kod != 0: exit

                        def

    # +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    class uzel_gen_mdp_class:  # UGMC. храним параметры генератора
        Private
        tG

        # dim Num ,gname, rezerv_P_UP ,rezerv_P_DOWN , Pmax ,Pmin , ndxGen , tip #  tip  = 1 используем    0 не использовать узел
        def init():
            tG = rastr.tables("Generator")
            Pmax = tG.cols.item("Pmax").Z(ndxGen)
            if Pmax = 0: logging.info("у генератора " + str(Num) + " " + gname + " не задано Pmax")
            Pmin = tG.cols.item("Pmin").Z(ndxGen)
            gname = tG.cols.item("Name").ZS(ndxGen)
            tip = 1

        def init_rezerv_P():
            rezerv_P_UP = 0:rezerv_P_DOWN = 0
            if tG.cols.item("sta").Z(ndxGen):  # если генератор отключен
                rezerv_P_UP = Pmax
            else:  # если генератор включен
                rezerv_P_DOWN = tG.cols.item("P").Z(ndxGen)
                if Pmax > 0: rezerv_P_UP = Pmax - tG.cols.item("P").Z(ndxGen)

    # +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    class KorSech_class:  # KSC. изменить мощность сечения
        Private
        tN38, tG, viborka1

        # dim Ncikl, dr_p_zad, epss, tipIzm, net_Pmin, Pmin_rezerv, print_gen , ns, newp, vibor ,  Dict_node_mdp , UMC , ps, KefPG_Down, KefPG_UP , dr_p_keff , nebalans_izm_ps
        # dim  db , ndxNode    , inini , ps_save , KolNySel , rezerv_P_sum , rezerv_P_sum0 , sum_gen_sech_UP , sum_gen_sech_DOWN, izm_ps, snisimP, uvelichimP , test_izm_P
        def init_sech(net_Pmin_zad):  # (  net_Pmin_zad)
            tN38 = rastr.tables("node")
            tG = rastr.tables("Generator")
            # --------------настройки----------
            Ncikl = 20  # максимальное количество циклов
            dr_p_zad = 0.01  # величина реакции начальная
            dr_p_keff = 0  # если  1 то умножаем дополнительно на dr_p в этом случае больше загружаются генераторы которые меньше влияют на изменение мощности в сечении
            epss = 0.05  # процент , точность задания мощности сечения, но не превышает заданную
            tipIzm = 1  # 1  ген узлов и ген ген, 2 ген ген
            net_Pmin = net_Pmin_zad  # 1 не учитывать Pmin - задается
            Pmin_rezerv = 0.1  # мин резерв в узле , МВт
            print_gen = 1  # вывод на печать измы ген
            # ----------------------------------------------------

        def korr():
            nebalans_izm_ps = 0
            viborka1 = "tip>1 +!sta + abs(dr_p) > " + str(dr_p_zad)  # tip>1 ген   !sta вкл
            db = abs(rastr.Calc("sum", "node", "dr_p", viborka1 + "+dr_p>0")) + abs(
                rastr.Calc("sum", "node", "dr_p", viborka1 + "+dr_p<0"))  # сумма реакций
            if (db < dr_p_zad):   logging.info(
                "Невозможно изменить мощность по сечению (с учетом всех узлов в модели)"): exit

            def

            if vibor = "":  # отмечаем узлы и генераторы в выборке
                sel0()  # авто отметка ген узлов
                grup_cor("node", "sel", viborka1, "1")  # авто отметка ген узлов
            else:
                if vibor != "sel":  sel0(): tN38.setsel(vibor): tN38.cols.item("sel").Calc(1)

            # отметить генераторы у отмеченных узлов
            tN38.setsel("sel")
            ndxNode = tN38.FindNextSel(-1)
            while ndxNode >= 0  # ЦИКЛ узел
                grup_cor("Generator", "sel", "Node=" + tN38.cols.item("ny").ZS(ndxNode), "1")  #
                ndxNode = tN38.FindNextSel(ndxNode)
            wend
            # отметить узлы у отмеченных генераторов
            tG.setsel("sel")
            ndxNode = tG.FindNextSel(-1)
            while ndxNode >= 0  # ЦИКЛ узел
                grup_cor("node", "sel", "ny=" + tG.cols.item("Node").ZS(ndxNode), "1")  #
                ndxNode = tG.FindNextSel(ndxNode)
            wend
            db = abs(rastr.Calc("sum", "node", "dr_p", viborka1 + "+dr_p>0+sel")) + abs(
                rastr.Calc("sum", "node", "dr_p", viborka1 + "+dr_p<0+sel"))  # сумма реакций
            if (db < dr_p_zad):   logging.info(
                "Невозможно изменить мощность по сечению (с учетом выбранных узлов)"): exit

            def

            # В итоге должны быть отмечены все узлы где что то делаем и только те ген которые нада корр
            tN38.setsel("sel")  # # записываем узлы и  узлы генераторы в классы
            if tN38.count > 0:
                Dict_node_mdp = CreateObject("Scripting.Dictionary")  # для хранения uzel_class.init
                ndxNode = tN38.FindNextSel(-1)
                while ndxNode >= 0
                    UMC = uzel_mdp_class
                    UMC.ny = tN38.cols.item("ny").Z(ndxNode)
                    UMC.ndxNode = ndxNode
                    UMC.init()
                    Dict_node_mdp.Add(UMC.ny, UMC)  # ключ  и значение
                    ndxNode = tN38.FindNextSel(ndxNode)
                wend

            test_izm_P = True  # истина если мощность в сечении в циклах меняется, ложь если нет

            for inini = 1 to Ncikl
            logging.info("\t" + "Итерация " + str(inini))

            ps = rastr.Calc("sum", "sechen", "psech", "ns=" + str(ns))  # текущая мощность в сечении
            ps_save = ps
            izm_ps = newp - ps

            if newp = 0: newp = 0.0001
            if abs(izm_ps / newp) * 100 < epss and izm_ps < newp and newp > 0:
                logging.info(
                    "\t" + "Заданная точность достигнута, итераций " + str(inini) + ",izm_ps=" + str(round(izm_ps)))
                exit

                def  # выход

            if abs(izm_ps / newp) * 100 < epss and izm_ps > newp and newp < 0:
                logging.info(
                    "\t" + "Заданная точность достигнута, итераций " + str(inini) + ",izm_ps=" + str(round(izm_ps)))
                exit

                def  # выход

            if inini > 1:
                if round(ps_save, 1) = round ( ps, 1): test_izm_P = False  # мощность в цикле не поменялся

            KolNySel = 0:        rezerv_P_sum = 0:        rezerv_P_sum0 = 0
            sum_gen_sech_UP = 0: sum_gen_sech_DOWN = 0  # резерв увеличения:уменьшения Р для задания сечения

            for EACH UMC in Dict_node_mdp.Items  # формируем rezerv_P_UP и rezerv_P_DOWN
                if UMC.tip > 0:  # #  tip  = 1 узел без ген, 2 с генераторами    0 не использовать узел
                    UMC.init_rezerv_P()
                    if UMC.rezerv_P_UP < Pmin_rezerv:  UMC.rezerv_P_UP = 0
                    KolNySel = KolNySel + 1
                    rezerv_P_sum = rezerv_P_sum + UMC.rezerv_P_UP
                    rezerv_P_sum0 = rezerv_P_sum0 + UMC.rezerv_P_DOWN

                    if izm_ps * UMC.dr_p > 0:
                        sum_gen_sech_UP = sum_gen_sech_UP + UMC.rezerv_P_UP
                        UMC.up_pgen = True  # нада увеличить генерацияю в узле , False уменьшить
                    elif izm_ps * UMC.dr_p < 0:
                        sum_gen_sech_DOWN = sum_gen_sech_DOWN + UMC.rezerv_P_DOWN
                        UMC.up_pgen = False  # уменьшить

            logging.info("\t" + "\t" + "Мощность сечения = " + str(round(ps, 0)) + " МВт, нужно получить: " + str(
                newp) + " МВт. Изменть на " + str(round(izm_ps)) + " МВт. Отклонение = " + str(
                round(abs(izm_ps / newp) * 100)) + " %")
            logging.info("\t" + "\t" + "Количество узлов в выборке " + str(
                KolNySel) + ". Суммарный резерв на увеличении генерации " + str(round(rezerv_P_sum)) + " МВт.")
            logging.info("\t" + "\t" + "P снижение " + str(round(sum_gen_sech_DOWN)) + " МВт. P увеличение " + str(
                round(sum_gen_sech_UP)) + " МВт.")
            if KolNySel = 0:  logging.info("\t" + "выход тк нет ген узлов"): exit

            def

                if (sum_gen_sech_DOWN + sum_gen_sech_UP) = 0:  logging.info(
                    "\t" + "выход тк (sum_gen_sech_DOWN + sum_gen_sech_UP ) = 0"): exit

            def

                calc_KefPG()

            for EACH UMC in Dict_node_mdp.Items  # корр
                if UMC.tip > 0: UMC.korr()  # tip  = 1 узел без ген, 2 с генераторами    0 не использовать узел        kod = rastr.rgm ("")

        if kod != 0: kod = rastr.rgm("")
        if kod != 0: logging.info("\t" + "---------Аварийное завершение расчета режима---------- KorSech_class.korr")
        if inini = Ncikl: logging.info("\t" + "не хватило итераций")

    def calc_KefPG():
        if (sum_gen_sech_DOWN + sum_gen_sech_UP) != 0:
            snisimP = abs(sum_gen_sech_DOWN / (sum_gen_sech_DOWN + sum_gen_sech_UP) * izm_ps) else:
            snisimP = 0  # на сколько МВт нужно снизить Р
        if sum_gen_sech_DOWN < snisimP: snisimP = sum_gen_sech_DOWN
        uvelichimP = abs(abs(izm_ps) - snisimP)  # на сколько МВт нужно увеличить Р
        if sum_gen_sech_UP < uvelichimP: uvelichimP = sum_gen_sech_UP

        if (sum_gen_sech_DOWN + sum_gen_sech_UP) < izm_ps: logging.info("\t" + "\t" + "генерации не хватает")

        if sum_gen_sech_DOWN > 0:
            KefPG_Down = 1 - (sum_gen_sech_DOWN - snisimP) / sum_gen_sech_DOWN else:
            KefPG_Down = 0  # кэфф на сколько нужно умножить рген и прибавить к рген, будем использовать для снижения генерации
        if sum_gen_sech_UP > 0:
            KefPG_UP = 1 - (sum_gen_sech_UP - uvelichimP) / sum_gen_sech_UP else:
            KefPG_UP = 0  # кэфф на сколько нужно умножить резерв Р и прибави��ь его к Рген, будем использовать для снижения увеличения генерации


# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def KorSech(ns, newp, vibor, tip,
            net_Pmin_zad):  # номер сеч, новая мощность в сеч (значение или "max" "min"), выбор корр узлов  (нр "sel"- отмеченные узлы и генераторыили "" - авто) ,  tip - pn или qn
    # net_Pmin_zad #  1 не учитывать Pmin - задается
    # если не задано Р макс то ген в узле не увеличиваем
    #  чтобы учесть ограничения по прочим сечениям нада отметить sta  - Контроль в таблице сечений
    # dim tGrline, tN39, ndxGrline
    # dim P_izm ,P_sum_vibor , ndxtN , keff_pn , dr_p_sum
    tN39 = rastr.tables("node")
    tGrline = rastr.Tables("grline")
    #  newp - инициализация
    if fNDX("sechen", ns) = -1:  logging.info(
        "\t" + "KorSech-выход тк нет сечения (возможно нужно загрузить файл сечений)"): exit

    def

        if newp = "max": newp = rastr.tables("sechen").cols.item("pmax").Z(fNDX("sechen", ns))
    if newp = "min": newp = rastr.tables("sechen").cols.item("pmin").Z(fNDX("sechen", ns))
    # реакции во всех узлах, реакции в узле положительные если увеличение P приводит к увеличению перетока в сечении
    rastr.sensiv_start("")
    tGrline.setsel("ns=" + str(ns))
    ndxGrline = tGrline.FindNextSel(-1)
    While
    ndxGrline != -1
    rastr.sensiv_back(4, 1., tGrline.cols.item("ip").Z(ndxGrline), tGrline.cols.item("iq").Z(ndxGrline), 0)
    ndxGrline = tGrline.FindNextSel(ndxGrline)


Wend
rastr.sensiv_write("")
rastr.sensiv_end()

if tip = "pg":
    KSC = KorSech_class
    KSC.ns = ns
    KSC.newp = newp
    KSC.vibor = vibor
    KSC.init_sech(net_Pmin_zad)
    KSC.korr()
elif tip = "pn":
    P_sum_vibor = 0
    dr_p_sum = 0
    for i = 0 to 5

    P_izm = newp - rastr.Calc("sum", "sechen", "psech", "ns=" + str(ns))  # текущая мощность в сечении
    if abs(P_izm) < 0.1: logging.info(
        "\t" + "заданная точность достигнута, сечение " + str(ns) + ",P_izm=" + str(round(P_izm))): exit
    for
        tN39.setsel(vibor)
    ndxtN = tN39.FindNextSel(-1)
    while ndxtN >= 0  # посчитали суммы
        P_sum_vibor = P_sum_vibor + tN39.cols.item("pn").Z(ndxtN)
        dr_p_sum = dr_p_sum + tN39.cols.item("dr_p").Z(ndxtN)
        ndxtN = tN39.FindNextSel(ndxtN)
    wend

    if P_sum_vibor > 0:
        if dr_p_sum < 0:
            keff_pn = 1 + (1 - (P_sum_vibor - P_izm) / P_sum_vibor)
        else:
            keff_pn = (P_sum_vibor - P_izm) / P_sum_vibor

        tN39.cols.item("pn").Calc("pn*" + str(keff_pn))
        tN39.cols.item("qn").Calc("qn*" + str(keff_pn))
        rastr.rgm("")
    else:
        logging.info("\t" + "KorSech P_sum_vibor=0")


# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def fKontrSech(
        ns):  # возвращает истина если мощность в сечениях отмеченных контроль (sta) не превышена   или ложь (исключение)
    # dim tS , ndxSech
    tS = rastr.tables("sechen")


fKontrSech = True  # истина

tS.setsel("sta")  #
ndxSech = tS.FindNextSel(-1)

while ndxSech >= 0
    if tS.cols.item("pmax").Z(ndxSech) AND tS.cols.item("ns").Z(ndxSech) != ns:
        if tS.cols.item("psech").Z(ndxSech) > tS.cols.item("pmax").Z(ndxSech):
            fKontrSech = False
            logging.info("\t" + "превышение в сечении  " + tS.cols.item("ns").ZS(ndxSech))
            ndxSech = tS.FindNextSel(ndxSech)
wend


# End def return
# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def otklonenie_seshen(nomer_sesh):  # возвращает величину отклонения psech от  pmax   + превышение; - недобор
    # dim ndxSech
    tS = rastr.tables("sechen")


ndxSech = fNDX("sechen", nomer_sesh)
otklonenie_seshen = tS.cols.item("psech").Z(ndxSech) - tS.cols.item("pmax").Z(ndxSech)


# End def return
# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def fZ(tabl, kkluch, param):  # возвращает заданное значение
    # dim tTabl , NDX, kkluchM
    tTabl = rastr.Tables(tabl)
    if tabl = "node":
        tTabl.setsel("ny=" + str(kkluch))
    elif tabl = "vetv":
        kkluchM = split(kkluch, ",")
        tTabl.setsel("ip=" + str(kkluchM(0)) + "+iq=" + str(kkluchM(1)) + "+np=" + str(kkluchM(2)))
    elif tabl = "area":
        tTabl.setsel("na=" + str(kkluch))
    elif tabl = "ngroup":
        tTabl.setsel("nga=" + str(kkluch))
    elif tabl = "Generator":
        tTabl.setsel("Num=" + str(kkluch))

    NDX = tTabl.FindNextSel(-1)
    if NDX == -1:
        logging.info("\t" + "fZ в таблице " + tabl + " не найдено " + str(kkluch)): exit

        def
    fZ = tTabl.cols.item("param").Z(NDX)


# End def return
# +++++++++++++++++++++++КОРРРРРРРРРРРРРРРРРРР +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def PGen_cor(sel_gen):
    # если мощность P больше Pmax то изменить мощность генератора  на Pmax, если P меньше Pmin но больше 0 - то на Pmin
    #  если P ген = 0 то отключить генератор, чтоб реактивка не выдавалась
    tG = rastr.tables("Generator")


tG.setsel(sel_gen)
ndxGen = tG.FindNextSel(-1)
while ndxGen >= 0
    if tG.cols.item("Pmax").Z(ndxGen) < tG.cols.item("P").Z(ndxGen) and tG.cols.item("Pmax").Z(ndxGen) > 0:
        logging.info(
            "\t" + "изменена генерация генератора: " + tG.cols.item("Name").ZS(ndxGen) + ", номер: " + tG.cols.item(
                "Num").ZS(ndxGen) + ", c " + tG.cols.item("P").ZS(ndxGen) + " на " + tG.cols.item("Pmax").ZS(
                ndxGen) + ", ny=" + tG.cols.item("Node").ZS(ndxGen))
        tG.cols.item("P").Z(ndxGen) = tG.cols.item("Pmax").Z(ndxGen)

    if tG.cols.item("Pmin").Z(ndxGen) > tG.cols.item("P").Z(ndxGen):
        logging.info(
            "\t" + "изменена генерация генератора: " + tG.cols.item("Name").ZS(ndxGen) + ", номер: " + tG.cols.item(
                "Num").ZS(ndxGen) + ", c " + tG.cols.item("P").ZS(ndxGen) + " на " + tG.cols.item("Pmin").ZS(
                ndxGen) + ", ny=" + tG.cols.item("Node").ZS(ndxGen))
        tG.cols.item("P").Z(ndxGen) = tG.cols.item("Pmin").Z(ndxGen)

    if tG.cols.item("P").Z(ndxGen) = 0 and tG.cols.item("sta").Z(ndxGen) = 0:
        logging.info(
            "\t" + "генератор отключен: " + tG.cols.item("Name").ZS(ndxGen) + ", номер: " + tG.cols.item("Num").ZS(
                ndxGen) + ", P=" + tG.cols.item("P").ZS(ndxGen) + ", ny=" + tG.cols.item("Node").ZS(ndxGen))
        tG.cols.item("P").Z(ndxGen) = tG.cols.item("Pmin").Z(ndxGen)

    ndxGen = tG.FindNextSel(ndxGen)

def XL_print_tabl_SUB(param_vibor):  # печать параметров узлов
    # dim kkluch_tek , ii , param_vibor_m , indx , i ,tTabl, temp
    tTabl = rastr.Tables(CS.print_tabl_name)
    # x_ny = 3 меняем  Y_ny = 3 не меняем
    param_vibor_m = split(CS.print_param, ",")
    ii = Y_ny
    if XL_print_ny.cell(3, 1).Value = "":  # истина то нада сделать принт узлов
        for Each kkluch_tek in CS.dict_tabl.Keys  # организуем цикл по элементам  масива Keys
            XL_print_ny.cell(ii, 1).Value = kkluch_tek
            XL_print_ny.cell(ii, 2).Value = CS.dict_tabl.Item(kkluch_tek)
            ii = ii + 1

    XL_print_ny.cell(1, x_ny).Value = RG.Name_Base
    ii = Y_ny
    for Each kkluch_tek in CS.dict_tabl.Keys  # организуем цикл по элементам  масива Keys
        if CS.print_tabl_name = "node": tTabl.setsel("ny=" + str(kkluch_tek))
        if CS.print_tabl_name = "Generator": tTabl.setsel("Num=" + str(kkluch_tek))
        if CS.print_tabl_name = "vetv":
            temp3 = split(kkluch_tek, ",")
            # temp3 = "ip=" + temp(0) + "+iq=" + temp(1) "+np=" + temp(2)
            tTabl.setsel("ip=" + temp3(0) + "+iq=" + temp3(1) + "+np=" + temp3(2))

        indx = tTabl.FindNextSel(-1)

        for i=0 to ubound ( param_vibor_m )
        if XL_print_ny.cell(2, x_ny + i).Value = "": XL_print_ny.cell(2, x_ny + i).Value = param_vibor_m(i)
        if indx > -1:
            XL_print_ny.cell(ii, x_ny + i).Value = tTabl.cols.item(param_vibor_m(i)).Z(indx)
        else:
            XL_print_ny.cell(ii, x_ny + i).Value = "-"

    # XL_print_ny.cell(ii , 1).Value = kkluch_tek
    ii = ii + 1


x_ny = x_ny + ubound(param_vibor_m) + 1


def XL_print_balans_Q_zap():  # БАЛАНС Q
    # dim row_name , row_qn , row_dq_sum , row_dq_line , row_dq_tran , row_shq_tran , row_skrm_potr , row_sum_port_Q, row_qg, row_skrm_gen, row_qg_min, row_qg_max, row_shq_line, row_sum_QG, row_Q_itog, row_Q_itog_gmin
    # dim row_Q_itog_gmax, ndx, tA
    row_name = 6  # Наименование
    row_qn = row_name + 1  # 7 Реактивная мощность нагрузки, Мвар
    row_dq_sum = row_qn + 1  # 8 Нагрузочные потери, Мвар
    row_dq_line = row_dq_sum + 1  # 9 в т.ч. потери в ЛЭП
    row_dq_tran = row_dq_line + 1  # 10 потери в трансформаторах
    row_shq_tran = row_dq_tran + 1  # 11 потери Х.Х. в трансформаторах
    row_skrm_potr = row_shq_tran + 1  # 12 Потребление ШР  и УШР
    row_sum_port_Q = row_skrm_potr + 1  # 13 Суммарное потребление реактивной мощности, Мвар
    row_qg = row_sum_port_Q + 1  # 14 Генерация реактивной мощности электростанциями, БСК, Мвар
    row_skrm_gen = row_qg + 1  # 15 Генерация реактивной мощности электростанциями, БСК, Мвар
    row_qg_min = row_skrm_gen + 1  # 16 Минимальная генерация реактивной мощности электростанциями, Мвар
    row_qg_max = row_qg_min + 1  # 17 Максимальная генерация реактивной мощности электростанциями, Мвар
    row_shq_line = row_qg_max + 1  # 18 Зарядная мощность ЛЭП, Мвар
    row_sum_QG = row_shq_line + 1  # 19 Суммарная генерация реактивной мощности, Мвар
    row_Q_itog = row_sum_QG + 1  # 20 Внешний переток реактивной мощности (избыток/дефицит +/-), Мвар
    row_Q_itog_gmin = row_Q_itog + 1  # 21 Внешний переток реактивной мощности при минимальной генерации реактивной мощности
    row_Q_itog_gmax = row_Q_itog_gmin + 1  # 22 Внешний переток реактивной мощности при максимальной генерации реактивной мощности
    if balans_Q_X0 = 5:
        XL_print_balans_Q.cell(row_name - 1, 4).Value = "в Мвар"
        XL_print_balans_Q.cell(row_name, 4).Value = "Наименование"
        XL_print_balans_Q.cell(row_qn, 4).Value = "Реактивная мощность нагрузки"
        XL_print_balans_Q.cell(row_qn, 4).AddComment
        XL_print_balans_Q.cell(row_qn, 4).Comment.Visible = False
        XL_print_balans_Q.cell(row_qn, 4).Comment.Text
        "Реактивная мощность нагрузки: " + Chr(10) + "Calc(sum,area,qn,vibor)"

        XL_print_balans_Q.cell(row_dq_sum, 4).Value = "Нагрузочные потери"
        XL_print_balans_Q.cell(row_dq_line, 4).Value = "в т.ч. потери в ЛЭП"
        XL_print_balans_Q.cell(row_dq_line, 4).AddComment
        XL_print_balans_Q.cell(row_dq_line, 4).Comment.Visible = False
        XL_print_balans_Q.cell(row_dq_line, 4).Comment.Text
        "потери в ЛЭП: " + Chr(10) + "Calc(sum,area,dq_line,vibor)"

        XL_print_balans_Q.cell(row_dq_tran, 4).Value = "потери в трансформаторах"
        XL_print_balans_Q.cell(row_dq_tran, 4).AddComment
        XL_print_balans_Q.cell(row_dq_tran, 4).Comment.Visible = False
        XL_print_balans_Q.cell(row_dq_tran, 4).Comment.Text
        "потери в трансформаторах: " + Chr(10) + "Calc(sum,area,dq_tran,vibor)"

        XL_print_balans_Q.cell(row_shq_tran, 4).Value = "потери Х.Х. в трансформаторах"
        XL_print_balans_Q.cell(row_shq_tran, 4).AddComment
        XL_print_balans_Q.cell(row_shq_tran, 4).Comment.Visible = False
        XL_print_balans_Q.cell(row_shq_tran, 4).Comment.Text
        "потери Х.Х. в трансформаторах: " + Chr(10) + "Calc(sum,area,shq_tran,vibor)"

        XL_print_balans_Q.Range(XL_print_balans_Q.cell(row_dq_line, 4),
                                XL_print_balans_Q.cell(row_shq_tran, 4)).HorizontalAlignment = -4152  # лево
        XL_print_balans_Q.cell(row_skrm_potr, 4).Value = "Потребление реактивной мощности СКРМ (ШР, УШР, СК, СТК)"
        XL_print_balans_Q.cell(row_skrm_potr, 4).AddComment
        XL_print_balans_Q.cell(row_skrm_potr, 4).Comment.Visible = False
        XL_print_balans_Q.cell(row_skrm_potr, 4).Comment.Text
        "Потребление реактивной мощности СКРМ (ШР, УШР, СК, СТК): " + Chr(
            10) + "Calc(sum,node,qsh,qsh>0 + vibor) - Calc(sum,node,qg,qg<0+pg<0.1+pg>-0.1 + vibor)"

        XL_print_balans_Q.cell(row_sum_port_Q, 4).Value = "Суммарное потребление реактивной мощности"
        XL_print_balans_Q.cell(row_qg, 4).Value = "Генерация реактивной мощности электростанциями"
        XL_print_balans_Q.cell(row_qg, 4).AddComment
        XL_print_balans_Q.cell(row_qg, 4).Comment.Visible = False
        XL_print_balans_Q.cell(row_qg, 4).Comment.Text
        "Генерация реактивной мощности электростанциями: " + Chr(10) + "Calc(sum,node,qg,(pg>0.1|pg<-0.1) + vibor)"

        XL_print_balans_Q.cell(row_skrm_gen, 4).Value = "Генерация реактивной мощности СКРМ (БСК, СК, СТК)"
        XL_print_balans_Q.cell(row_qg_min, 4).Value = "Минимальная генерация реактивной мощности электростанциями"
        XL_print_balans_Q.cell(row_qg_min, 4).AddComment
        XL_print_balans_Q.cell(row_qg_min, 4).Comment.Visible = False
        XL_print_balans_Q.cell(row_qg_min, 4).Comment.Text
        "Минимальная генерация реактивной мощности электростанциями: " + Chr(10) + "Calc(sum,node,qmin,pg>0.1+ vibor)"

        XL_print_balans_Q.cell(row_qg_max, 4).Value = "Максимальная генерация реактивной мощности электростанциями"
        XL_print_balans_Q.cell(row_qg_max, 4).AddComment
        XL_print_balans_Q.cell(row_qg_max, 4).Comment.Visible = False
        XL_print_balans_Q.cell(row_qg_max, 4).Comment.Text
        "Максимальная генерация реактивной мощности электростанциями: " + Chr(10) + "Calc(sum,node,qmax,pg>0.1+ vibor)"

        XL_print_balans_Q.cell(row_qg_max, 4).Interior.Color = vbRed
        XL_print_balans_Q.cell(row_qg_min, 4).Interior.Color = vbRed
        XL_print_balans_Q.cell(row_shq_line, 4).Value = "Зарядная мощность ЛЭП"
        XL_print_balans_Q.cell(row_shq_line, 4).AddComment
        XL_print_balans_Q.cell(row_shq_line, 4).Comment.Visible = False
        XL_print_balans_Q.cell(row_shq_line, 4).Comment.Text
        "Зарядная мощность ЛЭП: " + Chr(10) + "Calc(sum,area,shq_line, vibor)"

        XL_print_balans_Q.cell(row_sum_QG, 4).Value = "Суммарная генерация реактивной мощности"
        XL_print_balans_Q.cell(row_Q_itog, 4).Value = "Внешний переток реактивной мощности (избыток/дефицит +/-)"
        XL_print_balans_Q.cell(row_Q_itog_gmin,
                                4).Value = "Внешний переток реактивной мощности при минимальной генерации реактивной мощности электростанциями и КУ(избыток/дефицит +/-)"
        XL_print_balans_Q.cell(row_Q_itog_gmax,
                                4).Value = "Внешний переток реактивной мощности при максимальной генерации реактивной мощности электростанциями и КУ(избыток/дефицит +/-)"
        XL_print_balans_Q.cell(row_Q_itog_gmin, 4).Interior.Color = vbRed
        XL_print_balans_Q.cell(row_Q_itog_gmax, 4).Interior.Color = vbRed
        XL_print_balans_Q.cell(row_sum_port_Q, 4).Font.Bold = True
        XL_print_balans_Q.cell(row_sum_QG, 4).Font.Bold = True
        XL_print_balans_Q.cell(row_Q_itog, 4).Font.Bold = True

    XL_print_balans_Q.cell(6, balans_Q_X0).Value = RG.SezonName + " " + RG.god + " г" + " " + str(RG.DopName(0))
    XL_print_balans_Q.cell(6, balans_Q_X0).Orientation = 90
    tA = rastr.Tables("area")
    tA.setsel(CS.balans_Q_vibor)
    ndx = tA.FindNextSel(-1)

    XL_print_balans_Q.cell(row_qn, balans_Q_X0).Value = rastr.Calc("sum", "area", "qn",
                                                                    CS.balans_Q_vibor)  # Нагрузка Q sum("node","qnr","na="+str(na))
    address_qn = XL_print_balans_Q.cell(row_qn, balans_Q_X0).Address(False, False)  #

    XL_print_balans_Q.cell(row_dq_line, balans_Q_X0).Value = rastr.Calc("sum", "area", "dq_line",
                                                                         CS.balans_Q_vibor)  # Потери Q в ЛЭП
    address_dq_line = XL_print_balans_Q.cell(row_dq_line, balans_Q_X0).Address(False, False)  #

    XL_print_balans_Q.cell(row_dq_tran, balans_Q_X0).Value = rastr.Calc("sum", "area", "dq_tran",
                                                                         CS.balans_Q_vibor)  # Потери Q в Трансформаторах
    address_dq_tran = XL_print_balans_Q.cell(row_dq_tran, balans_Q_X0).Address(False, False)  #

    XL_print_balans_Q.cell(row_shq_tran, balans_Q_X0).Value = rastr.Calc("sum", "area", "shq_tran",
                                                                          CS.balans_Q_vibor)  # Потери Q_ХХ в Трансформаторах
    address_shq_tran = XL_print_balans_Q.cell(row_shq_tran, balans_Q_X0).Address(False, False)

    XL_print_balans_Q.cell(row_skrm_potr, balans_Q_X0).Value = rastr.Calc("sum", "node", "qsh",
                                                                           "qsh>0+" + CS.balans_Q_vibor) - rastr.Calc(
        "sum", "node", "qg", "qg<0+pg<0.1+pg>-0.1+" + CS.balans_Q_vibor)  # ШР УШР без бСК
    address_SHR = XL_print_balans_Q.cell(row_skrm_potr, balans_Q_X0).Address(False, False)

    XL_print_balans_Q.cell(row_qg, balans_Q_X0).Value = rastr.Calc("sum", "node", "qg",
                                                                    "(pg>0.1|pg<-0.1)+" + CS.balans_Q_vibor)  # Генерация Q  генераторов
    address_qg = XL_print_balans_Q.cell(row_qg, balans_Q_X0).Address(False, False)  #

    XL_print_balans_Q.cell(row_skrm_gen, balans_Q_X0).Value = -rastr.Calc("sum", "node", "qsh",
                                                                           "qsh<0+" + CS.balans_Q_vibor) + rastr.Calc(
        "sum", "node", "qg", "qg>0+pg<0.1+pg>-0.1+" + CS.balans_Q_vibor)  # Генерация БСК шунтом и СТК СК
    address_skrm_gen = XL_print_balans_Q.cell(row_skrm_gen, balans_Q_X0).Address(False, False)  #

    XL_print_balans_Q.cell(row_qg_min, balans_Q_X0).Value = rastr.Calc("sum", "node", "qmin",
                                                                        "pg>0.1+" + CS.balans_Q_vibor)  # минимальная генерация реактивной мощности в узлах выборки
    address_qg_min = XL_print_balans_Q.cell(row_qg_min, balans_Q_X0).Address(False, False)  #

    XL_print_balans_Q.cell(row_qg_max, balans_Q_X0).Value = rastr.Calc("sum", "node", "qmax",
                                                                        "pg>0.1+" + CS.balans_Q_vibor)  # максимальная генерация реактивной мощности в узлах выборки
    address_qg_max = XL_print_balans_Q.cell(row_qg_max, balans_Q_X0).Address(False, False)  #

    XL_print_balans_Q.cell(row_shq_line, balans_Q_X0).Value = - rastr.Calc("sum", "area", "shq_line",
                                                                            CS.balans_Q_vibor)  # Генерация Q в ЛЭП
    address_shq_line = XL_print_balans_Q.cell(row_shq_line, balans_Q_X0).Address(False, False)  #

    XL_print_balans_Q.cell(row_dq_sum,
                            balans_Q_X0).Value = "=" + address_dq_line + "+" + address_dq_tran + "+" + address_shq_tran
    address_poteri = XL_print_balans_Q.cell(row_dq_sum, balans_Q_X0).Address(False, False)  #

    XL_print_balans_Q.cell(row_sum_port_Q,
                            balans_Q_X0).Value = "=" + address_qn + "+" + address_poteri + "+" + address_SHR
    address_nagruz = XL_print_balans_Q.cell(row_sum_port_Q, balans_Q_X0).Address(False, False)  #

    XL_print_balans_Q.cell(row_sum_QG,
                            balans_Q_X0).Value = "=" + address_qg + "+" + address_shq_line + "+" + address_skrm_gen
    address_sum_gen = XL_print_balans_Q.cell(row_sum_QG, balans_Q_X0).Address(False, False)  #

    XL_print_balans_Q.cell(row_Q_itog, balans_Q_X0).Value = "=-" + address_nagruz + "+" + address_sum_gen
    XL_print_balans_Q.cell(row_Q_itog_gmin,
                            balans_Q_X0).Value = "=-" + address_nagruz + "+" + address_qg_min + "+" + address_shq_line
    XL_print_balans_Q.cell(row_Q_itog_gmax,
                            balans_Q_X0).Value = "=-" + address_nagruz + "+" + address_qg_max + "+" + address_shq_line

    balans_Q_X0 = balans_Q_X0 + 1


def f_gen_qg(ny, zad):  # возвращает минимальную генерацию Q  выборки CS.balans_Q_vibor таблица ГЕНЕРАТОРЫ
    # dim  tG
    tG = rastr.Tables("Generator")
    f_gen_qg = 0

    tG.setsel("Node=" + str(ny))

    ndx_tG = tG.FindNextSel(-1)
    if ndx_tG = -1: f_gen_qg = "не найден"

    while ndx_tG >= 0  #
        if not tG.cols.item("sta").Z(ndx_tG) = 1:
            if zad = "min":
                f_gen_qg = f_gen_qg + tG.cols.item("Qmin").Z(ndx_tG)
            elif zad = "max":
                f_gen_qg = f_gen_qg + tG.cols.item("Qmax").Z(ndx_tG)
                ndx_tG = tG.FindNextSel(ndx_tG)


# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в  коррррр+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def kor1(k_kluch, param_kor,
         value_param):  # коррекция одного уникальнгого занчения(краткийй ключ, параметр корр, значение) например("7","name","Юж")
    # dim kkluch_m , ptabl , viborka , ndx
    kkluch_m = split(k_kluch, ",")

    if ubound(kkluch_m) = 0:  # УЗЕЛ
        ptabl = rastr.Tables("node")
        viborka = "ny=" + str(k_kluch)
    elif ubound(kkluch_m) = 2:  # Ветвь
        ptabl = rastr.Tables("vetv")
        viborka = "ip=" + str(kkluch_m(0)) + "+iq=" + str(kkluch_m(1)) + "+np=" + str(kkluch_m(2))
    elif ubound(kkluch_m) = 1:  # генератор или что то еще
        if kkluch_m(0) = "g":  # генератор
            ptabl = rastr.Tables("Generator")
            viborka = "Num=" + str(kkluch_m(1))

        if kkluch_m(0) = "no":  # объединене
            ptabl = rastr.Tables("darea")
            viborka = "no=" + str(kkluch_m(1))

        if kkluch_m(0) = "na":  # районы
            ptabl = rastr.Tables("area")
            viborka = "na=" + str(kkluch_m(1))

        if kkluch_m(0) = "npa":  # территория
            ptabl = rastr.Tables("area2")
            viborka = "npa=" + str(kkluch_m(1))
            ptabl.setsel(viborka)
    ndx = ptabl.FindNextSel(-1)
    if ndx > -1:
        ptabl.cols.item(param_kor).Z(ndx) = value_param
        logging.info("\t" + "kor1 " + k_kluch + " | " + param_kor + " | " + value_param)  #
    else:
        logging.info("\t" + "НЕ НАЙДЕН kor1 " + k_kluch + " | " + param_kor + " | " + value_param)  #

    # +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в  коррррр+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


def cor_one(tabl, param, viborka,
            formula):  # коррекция одного уникальнгого занчения(таблица, параметр корр, выборка, формула для расчета параметра) "node","name","ny=7","Юж" УСТАРЕЛО
    # dim ndx
    ptabl = rastr.Tables(tabl)
    pparam = ptabl.cols.item(param)
    ptabl.setsel(viborka)
    ndx = ptabl.FindNextSel(-1)
    if ndx > -1:
        pparam.Z(ndx) = formula
    else:
        logging.info(
            "\t" + "не найден: " + str(tabl) + " / " + str(param) + " / " + str(viborka) + " / " + str(formula))

    logging.info(
        "\t" + "cor_one: tabl=" + tabl + "/param=" + param + tabl + "/viborka=" + viborka + tabl + "/formula=" + formula)


# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в  коррррр+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def Del(tabl, viborka):  #
    # dim tVet , tNod , it
    ptabl = rastr.Tables(tabl)
    logging.info("\t" + "Del: tabl=" + tabl + "/viborka=" + viborka)
    if viborka = "net":
        if tabl = "node":
            tVet = rastr.Tables("vetv")
            ptabl.setsel("sta")
            it = ptabl.FindNextSel(-1)
            while it >= 0
                tVet.setsel("ip=" + ptabl.cols.item("ny").ZS(it) + "|iq=" + ptabl.cols.item("ny").ZS(it))
                if tVet.count = 0:
                    logging.info(
                        "\t" + "\t" + "узел без связей Del: tabl=" + tabl + "/viborka=" + "ny=" + ptabl.cols.item(
                            "ny").ZS(it) + "/" + ptabl.cols.item("name").ZS(it))
                    Del("node", "ny=" + ptabl.cols.item("ny").ZS(it))
                    ptabl.setsel("sta")
                    it = ptabl.FindNextSel(-1)

                if it > -1: it = ptabl.FindNextSel(it)
            wend
        elif tabl = "vetv":
            tNod = rastr.Tables("node")
            ptabl.setsel("sta")
            it = ptabl.FindNextSel(-1)
            while it >= 0
                tNod.setsel("ny=" + ptabl.cols.item("ip").ZS(it) + "|ny=" + ptabl.cols.item("iq").ZS(it))
                if tNod.count < 2:
                    logging.info(
                        "\t" + "\t" + "ветв без узла начали конча Del: tabl=" + tabl + "/viborka=" + ptabl.SelString(
                            it) + "/" + ptabl.cols.item("name").ZS(it))
                    # Del("vetv",ptabl.SelString(it))

                it = ptabl.FindNextSel(it)
            wend
            ptabl.setsel("ip.ny=0|iq.ny=0")  # удалить  ветви связанные с уделенными узлами
            ptabl.DelRows

    else:
        ptabl.setsel(viborka)
        ptabl.DelRows

    logging.info("\t" + "Del, таблица:" + tabl + ",выборка:" + viborka)


# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в  коррррр+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def Del_sel():  # удалить отмеченные узлы (c ветвями) ветви и генераторы
    # dim tV34  , tN40
    tV34 = rastr.Tables("vetv")
    tN40 = rastr.tables("node")
    tG = rastr.tables("Generator")
    tN40.setsel("sel")
    tN40.DelRows
    tV34.setsel("ip.ny=0|iq.ny=0")  # удалить  ветви связанные с уделенными узлами
    tV34.DelRows
    tV34.setsel("sel")
    tV34.DelRows
    tG.setsel("sel")
    tG.DelRows
    logging.info("\t" + "Del_sel: удалены все отмеченные узлы (c ветвями), ветви и генераторы")


# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в  коррррр+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def sta_node(str_ny, on_off):  # узлы с ветвями (номера узлов через пробел) включить False; отключить True
    # dim   masiv_ny , i     , tN41
    # dim tV35
    tV35 = rastr.Tables("vetv")
    tN41 = rastr.tables("node")
    masiv_ny = split(str_ny, " ")
    for each i in masiv_ny
        if fNDX("node", float(i)) > -1:
            tN41.cols.item("sta").Z(fNDX("node", float(i))) = on_off
            tV35.setsel("ip=" + str(i) + "|iq=" + str(i))
            if on_off:
                tV35.cols.item("sta").calc(1) else:
                tV35.cols.item("sta").calc(0)
        else:
            logging.info("\t" + "sta_node: не найден узел " + str(i))

    logging.info("\t" + "sta_node: str_ny=" + str_ny + "/on_off=" + on_off)

# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в  коррррр+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def name0():  # поиск узлов и генераторов без имени
    nodee = rastr.tables("node")
    for i=0 to nodee.size-1
    if Replace(nodee.cols.item("name").ZS(i), " ", "") = "": logging.info(
        "\t" + "узел без имени ny: " + nodee.cols.item("ny").ZS(i))


genn = rastr.tables("Generator")
for i=0 to genn.size-1
if Replace(genn.cols.item("Name").ZS(i), " ", "") = "": logging.info(
    "\t" + "генератор без имени Num: " + genn.cols.item("Num").ZS(i))


# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в  коррррр+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def vzd0():  # поиск узлов где напряжение vzd задано а диапозона реактивки нет и удаляет vzd
    nodee = rastr.tables("node")
    for i=0 to nodee.size-1
    if nodee.cols.item("vzd").Z(i) > 0 and nodee.cols.item("qmin").Z(i) = 0 and nodee.cols.item("qmax").Z(i) = 0:
        logging.info(
            "\t" + "узел c qmin=qmax=0 vzd = " + nodee.cols.item("vzd").ZS(i) + " ny = " + nodee.cols.item("ny").ZS(
                i) + "(" + nodee.cols.item("name").ZS(i) + "), vzd изменено на 0")
        cor(nodee.cols.item("ny").ZS(i), "vzd=0")


# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в  коррррр+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def nyNum0():  # поиск узлов и генераторов с номером 0
    nodee = rastr.tables("node")
    nodee.setsel("ny=0")
    if nodee.count > 0:
        logging.info("\t" + "найдено нулевых узлов:" + str(nodee.count) + ", удалены")
        nodee.delRows

    genn = rastr.tables("Generator")
    genn.setsel("Num=0")
    if genn.count > 0:
        logging.info("\t" + "найдено нулевых генераторов:" + str(genn.count) + ", удалены")
        genn.delRows


# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в  коррррр+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def SEL(zadanie,
        no_off):  # отметить например "123 123,312,1 g,12" отметить узел 123 и ветвь 123,312,0  и генератор 12,  no_off = "0" снять отсетку "1" отметить
    # dim z_test , zadanie_m, z
    zadanie_m = split(zadanie)
    for Each z in zadanie_m
        z_test = split(z, ",")
        if ubound(z_test) = 0:  # УЗЕЛ
            if z != "": grup_cor("node", "sel", "ny=" + str(z), no_off)  #
        elif ubound(z_test) = 2:  # Ветвь
            grup_cor("vetv", "sel", "ip=" + z_test(0) + "+iq=" + z_test(1) + "+np=" + z_test(2), no_off)
        elif ubound(z_test) = 1:  # генератор или что то еще
            if z_test(0) = "g":  # генератор
                grup_cor("Generator", "sel", "Num=" + str(z_test(1)), no_off)

    logging.info("\t" + "SEL: zadanie=" + zadanie + "/no_off=" + no_off)


# +++++++++++++++++++++++ОБЩЕЕ ДЛЯ МАКРОСОВ хранить в  коррррр++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def fTabParamMax(tab, par):
    tTabl = rastr.Tables(tab)
    tPar = tTabl.cols.item(par)
    max_pram = 0
    tTabl.setsel(par + ">0")  #
    ndx_tTabl = tTabl.FindNextSel(-1)
    while ndx_tTabl >= 0  #
        if tPar.Z(ndx_tTabl) > max_pram:  max_pram = tPar.Z(ndx_tTabl)

        ndx_tTabl = tTabl.FindNextSel(ndx_tTabl)
    wend

    fTabParamMax = max_pram
