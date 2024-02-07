
"""Программа для автоматизации работы ПК RASTRWIN3"""

from qt.qt_choice import Ui_choice  # pyuic5 qt.qt_choice.ui -o qt.qt_choice.py
from qt.qt_set import Ui_Settings  # pyuic5 qt_set.ui -o qt_set.py
from qt.qt_cor import Ui_cor  # pyuic5 qt_cor.ui -o qt_cor.py
from qt.qt_calc_ur import Ui_calc_ur  # pyuic5 qt_calc_ur.ui -o qt_calc_ur.py
from qt.qt_calc_ur_set import Ui_calc_ur_set  # pyuic5 qt_calc_ur_set.ui -o qt_calc_ur_set.py
# Мои модули
from general_settings import *
from ini import Ini
# import report
# import loading_sections as ls
# from my_error import *


class Window:
    """ Класс с общими методами для QT. """

    @staticmethod
    def check_status(set_checkbox_element: tuple):
        """
        По состоянию CheckBox изменить состояние видимости соответствующего элемента.
        :param set_checkbox_element: Картеж картежей (checkbox, element).
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
        fileName_choose, _ = QtWidgets.QFileDialog.getOpenFileName(self,
                                                                   directory=directory,
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

        self.b_task_save.clicked.connect(self.task_save_yaml)
        self.b_task_load.clicked.connect(self.task_load_yaml)

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
        self.te_path_initial_models.setPlainText(ini.read_ini(section='save_form_folder_calc', key="path"))
        # Подсказки
        self.le_control_sel.setToolTip("Если поле не заполнено, то контролируются все ветви и узлы РМ")
        self.te_path_initial_models.setToolTip("Для расчета файлов во всех вложенных папках нужно в конце поставить *")

    def task_save_yaml(self):
        name_file_save = self.save_file(directory=self.te_path_initial_models.toPlainText(),
                                        filter_="YAML Files (*.yaml)")
        if name_file_save:
            with open(name_file_save, 'w') as f:
                yaml.dump(data=self.fill_task_calc(), stream=f, default_flow_style=False, sort_keys=False)

    def task_load_yaml(self):
        name_file_load = self.choice_file(directory=self.te_path_initial_models.toPlainText().replace('*', ''),
                                          filter_="YAML Files (*.yaml)")
        if not name_file_load:
            return
        with open(name_file_load) as f:
            task_yaml = yaml.safe_load(f)
        if not task_yaml:
            return

        # Окно запуска расчета.
        self.te_path_initial_models.setPlainText(task_yaml["calc_folder"])
        # Выборка файлов.
        self.cb_filter.setChecked(task_yaml["Filter_file"])  # QCheckBox
        self.sb_count_file.setValue(task_yaml["file_count_max"])  # QSpainBox
        self.le_condition_file_years.setText(task_yaml["calc_criterion"]["years"])  # QLineEdit text()
        self.le_condition_file_season.setCurrentText(task_yaml["calc_criterion"]["season"])  # QComboBox
        self.le_condition_file_max_min.setCurrentText(task_yaml["calc_criterion"]["max_min"])
        self.le_condition_file_add_name.setText(task_yaml["calc_criterion"]["add_name"])
        # Корректировка в txt.
        self.cb_cor_txt.setChecked(task_yaml["cor_rm"]['add'])
        self.te_cor_txt.setPlainText(task_yaml["cor_rm"]['txt'])
        # Импорт ИД для расчетов УР из моделей.
        self.cb_import_model.setChecked(task_yaml['CB_Import_Rg2'])
        self.te_path_import_rg2.setPlainText(task_yaml["Import_file"])
        self.te_import_rg2.setPlainText(task_yaml['txt_Import_Rg2'])
        # Расчет всех возможных сочетаний. Отключаемые элементы.
        self.cb_disable_comb.setChecked(task_yaml['cb_disable_comb'])
        self.cb_n1.setChecked(task_yaml['SRS']['n-1'])
        self.cb_n2_abv.setChecked(task_yaml['SRS']['n-2_abv'])
        self.cb_n2_gd.setChecked(task_yaml['SRS']['n-2_gd'])
        self.cb_n3.setChecked(task_yaml['SRS']['n-3'])

        self.cb_auto_disable.setChecked(task_yaml['cb_auto_disable'])
        self.le_auto_disable_choice.setText(task_yaml['auto_disable_choice'])

        self.cb_comb_field.setChecked(task_yaml['cb_comb_field'])
        self.le_comb_field.setText(task_yaml['comb_field'])

        self.cb_filter_comb.setChecked(task_yaml['filter_comb'])
        self.sb_filter_comb_val.setValue(task_yaml['filter_comb_val'])
        # Импорт перечня расчетных сочетаний из EXCEL
        self.cb_disable_excel.setChecked(task_yaml['cb_disable_excel'])
        self.te_XL_path.setPlainText(task_yaml['srs_XL_path'])
        self.le_XL_sheets.setText(task_yaml['srs_XL_sheets'])
        # Расчет всех возможных сочетаний. Контролируемые элементы.
        self.cb_control.setChecked(task_yaml['cb_control'])
        self.cb_control_field.setChecked(task_yaml['cb_control_field'])
        self.le_control_field.setText(task_yaml['control_field'])
        self.cb_control_sel.setChecked(task_yaml['cb_control_sel'])
        self.le_control_sel.setText(task_yaml['control_sel'])
        # Результаты в EXCEL: таблицы контролируемые - отключаемые элементы
        self.cb_save_i.setChecked(task_yaml['cb_save_i'])
        self.cb_tab_KO.setChecked(task_yaml['cb_tab_KO'])
        self.te_tab_KO_info.setPlainText(task_yaml['te_tab_KO_info'])
        # Результаты в RG2
        self.cb_results_pic.setChecked(task_yaml['results_RG2'])
        self.cb_pic_overloads.setChecked(task_yaml['pic_overloads'])
        self.te_name_pic.setPlainText(task_yaml['name_pic'])

        self.check_status(self.check_status_visibility)

    def start(self):
        """
        Запуск расчета моделей
        """
        ini.write_ini(section='save_form_folder_calc',
                      key="path",
                      value=self.te_path_initial_models.toPlainText())

        CalcModel(self.fill_task_calc(), ini.to_dict()).run_calc()

    def fill_task_calc(self) -> dict:
        """ Возвращает данные с формы QT. """
        task_calc = {
            # Окно запуска расчета.
            'calc_folder': self.te_path_initial_models.toPlainText().strip(),
            # Выборка файлов.
            'Filter_file': self.cb_filter.isChecked(),  # QCheckBox
            'file_count_max': self.sb_count_file.value(),  # QSpainBox
            'calc_criterion': {'years': self.le_condition_file_years.text(),  # QLineEdit text()
                               'season': self.le_condition_file_season.currentText(),  # QComboBox
                               'max_min': self.le_condition_file_max_min.currentText(),
                               'add_name': self.le_condition_file_add_name.text()},
            # Корректировка в txt.
            'cor_rm': {'add': self.cb_cor_txt.isChecked(),
                       'txt': self.te_cor_txt.toPlainText()},
            # Импорт ИД для расчетов УР из моделей.
            'CB_Import_Rg2': self.cb_import_model.isChecked(),
            'Import_file': self.te_path_import_rg2.toPlainText(),
            'txt_Import_Rg2': self.te_import_rg2.toPlainText(),
            # Расчет всех возможных сочетаний. Отключаемые элементы.
            'cb_disable_comb': self.cb_disable_comb.isChecked(),
            'SRS': {'n-1': self.cb_n1.isChecked(),
                    'n-2_abv': self.cb_n2_abv.isChecked(),
                    'n-2_gd': self.cb_n2_gd.isChecked(),
                    'n-3': self.cb_n3.isChecked()},

            'cb_auto_disable': self.cb_auto_disable.isChecked(),
            'auto_disable_choice': self.le_auto_disable_choice.text(),

            'cb_comb_field': self.cb_comb_field.isChecked(),
            "comb_field": self.le_comb_field.text(),

            'filter_comb': self.cb_filter_comb.isChecked(),
            'filter_comb_val': self.sb_filter_comb_val.value(),
            # Импорт перечня расчетных сочетаний из EXCEL
            'cb_disable_excel': self.cb_disable_excel.isChecked(),
            'srs_XL_path': self.te_XL_path.toPlainText(),
            'srs_XL_sheets': self.le_XL_sheets.text(),
            # Расчет всех возможных сочетаний. Контролируемые элементы.
            'cb_control': self.cb_control.isChecked(),
            'cb_control_field': self.cb_control_field.isChecked(),
            'cb_control_sel': self.cb_control_sel.isChecked(),
            'control_field': self.le_control_field.text(),
            'control_sel': self.le_control_sel.text(),

            # Результаты в EXCEL: таблицы контролируемые - отключаемые элементы
            'cb_save_i': self.cb_save_i.isChecked(),
            'cb_tab_KO': self.cb_tab_KO.isChecked(),
            'te_tab_KO_info': self.te_tab_KO_info.toPlainText(),

            # Результаты в RG2
            'results_RG2': self.cb_results_pic.isChecked(),
            'pic_overloads': self.cb_pic_overloads.isChecked(),
            'name_pic': self.te_name_pic.toPlainText(),
        }
        return task_calc


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
        if ini.exists():
            config = ini.read_ini(section='CalcSetWindow')
            try:
                self.cb_gost.setChecked(eval(config["gost"]))
                self.cb_skrm.setChecked(eval(config["skrm"]))
                self.cb_avr.setChecked(eval(config["avr"]))
                self.cb_add_disabling_repair.setChecked(eval(config["add_disabling_repair"]))
                self.cb_pa.setChecked(eval(config["pa"]))
            except LookupError:
                log.error(f'Файл {ini.name} [CalcSetWindow] не читается, перезаписан.')
                self.save_ini_ur()
        else:
            log.info(f'Создан файл {ini.name}.')
            self.save_ini_ur()

    def save_ini_ur(self):
        ini.save(info={"gost": self.cb_gost.isChecked(),
                       "skrm": self.cb_skrm.isChecked(),
                       "avr": self.cb_avr.isChecked(),
                       "add_disabling_repair": self.cb_add_disabling_repair.isChecked(),
                       "pa": self.cb_pa.isChecked()},
                 key='CalcSetWindow')


class SetWindow(QtWidgets.QMainWindow, Ui_Settings, Window):
    """
    Окно общих настроек.
    """

    def __init__(self):
        super(SetWindow, self).__init__()
        self.setupUi(self)
        self.load_ini()
        self.PB_qt_set_save.clicked.connect(lambda: self.save_ini())

    def load_ini(self):
        """Загрузить, создать или перезаписать файл .ini """
        if ini.exists():
            config = ini.read_ini(section='DEFAULT')
            try:
                self.LE_shablon_rg2.setText(config["шаблон rg2"])
                self.LE_shablon_rst.setText(config["шаблон rst"])
                self.LE_shablon_sch.setText(config["шаблон sch"])
                self.LE_shablon_trn.setText(config["шаблон trn"])
                self.LE_shablon_anc.setText(config["шаблон anc"])
                self.CB_load_trn_anc.setChecked(eval(config["load_trn_anc"]))
            except LookupError:
                log.error(f'Файл {ini.name} [DEFAULT] не читается, перезаписан.')
                self.save_ini()
        else:
            log.info(f'Создан файл {ini.name}.')
            self.save_ini()

    def save_ini(self):
        ini.save(info={"шаблон rg2": self.LE_shablon_rg2.text(),
                       "шаблон rst": self.LE_shablon_rst.text(),
                       "шаблон sch": self.LE_shablon_sch.text(),
                       "шаблон trn": self.LE_shablon_trn.text(),
                       "шаблон anc": self.LE_shablon_anc.text(),
                       "load_trn_anc": self.CB_load_trn_anc.isChecked()},
                 key='DEFAULT')


class EditWindow(QtWidgets.QMainWindow, Ui_cor, Window):
    """
    Окно корректировки моделей.
    """

    def __init__(self):
        super(EditWindow, self).__init__()  # *args, **kwargs
        self.setupUi(self)
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
        self.T_IzFolder.setPlainText(ini.read_ini(section='save_form_folder_edit', key="path"))

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
        if not name_file_save:
            raise ValueError('Не указан путь к сохраняемому файлу.')
        with open(name_file_save, 'w') as f:
            yaml.dump(data=self.fill_task_ui(),
                      stream=f, default_flow_style=False, sort_keys=False)

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

        self.CB_kontrol_rg2.setChecked(task_yaml["checking_parameters_rg2"])
        self.CB_U.setChecked(task_yaml["control_rg2_task"]['node'])
        self.CB_I.setChecked(task_yaml["control_rg2_task"]['vetv'])
        self.CB_gen.setChecked(task_yaml["control_rg2_task"]['Gen'])
        self.kontrol_rg2_Sel.setText(task_yaml["control_rg2_task"]['sel_node'])

        self.CB_printXL.setChecked(task_yaml["printXL"])
        self.CB_print_sech.setChecked(task_yaml['set_printXL']["sechen"]['add'])
        self.CB_print_area.setChecked(task_yaml['set_printXL']["area"]['add'])
        self.CB_print_area2.setChecked(task_yaml['set_printXL']["area2"]['add'])
        self.CB_print_darea.setChecked(task_yaml['set_printXL']["darea"]['add'])
        for key in task_yaml['set_printXL']:
            if key not in ["sechen", "area", "area2", "darea"]:
                self.CB_print_tab_log.setChecked(task_yaml['set_printXL'][key]['add'])
                self.print_tab_log_ar_set.setText(task_yaml['set_printXL'][key]["sel"])
                self.print_tab_log_ar_cols.setText(task_yaml['set_printXL'][key]['par'])
                self.print_tab_log_rows.setText(task_yaml['set_printXL'][key]['rows'])
                self.print_tab_log_cols.setText(task_yaml['set_printXL'][key]['columns'])
                self.print_tab_log_vals.setText(task_yaml['set_printXL'][key]['values'])
                break

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
        ini.write_ini(section='save_form_folder_edit',
                      key="path",
                      value=self.T_IzFolder.toPlainText())
        if self.print_tab_log_ar_tab.text() in ['area', 'area2', 'darea', 'sechen']:
            raise ValueError('В поле таблица на выбор нельзя задавать таблицы: area, area2, darea, sechen.')

        EditModel(self.fill_task_ui(), ini.to_dict()).run_cor()

    def fill_task_ui(self) -> dict:
        """
        Возвращает данные с формы QT в формате dict.
        """
        task_ui = {
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

            # Расчет режима и контроль параметров режима
            "checking_parameters_rg2": self.CB_kontrol_rg2.isChecked(),
            "control_rg2_task": {'node': self.CB_U.isChecked(),
                                 'vetv': self.CB_I.isChecked(),
                                 'Gen': self.CB_gen.isChecked(),
                                 'sel_node': self.kontrol_rg2_Sel.text()},
            # Выводить данные из моделей в XL
            "printXL": self.CB_printXL.isChecked(),
            "set_printXL": {
                "sechen": {'add': self.CB_print_sech.isChecked(),
                           "sel": 'ns>0',
                           'par': '',  # "ns,name,pmin,pmax,psech",
                           "rows": "ns,name",  # поля строк в сводной
                           "columns": "Год,Сезон макс/мин,Доп. имя1,Доп. имя2,Доп. имя3",  # поля столбцов в сводной
                           "values": "psech,pmax,difference_p"},
                "area": {'add': self.CB_print_area.isChecked(),
                         "sel": 'na>0',
                         'par': '',  # 'na,name,no,pg,pn,pn_sum,dp,pop,set_pop,qn_sum,pg_max,pg_min,poq,qn,qg,dev_pop',
                         "rows": "na,name,Сезон макс/мин,Доп. имя1,Доп. имя2,Доп. имя3",  # поля строк в сводной
                         "columns": "Год",  # поля столбцов в сводной
                         "values": "pop,difference_p"},
                "area2": {'add': self.CB_print_area2.isChecked(),
                          "sel": 'npa>0',
                          'par': '',  # 'npa,name,pg,pn,dp,pop,vnp,qg,qn,dq,poq,vnq,pn_sum,qn_sum,set_pop,dev_pop',
                          "rows": "npa,name,Сезон макс/мин,Доп. имя1,Доп. имя2,Доп. имя3",  # поля строк в сводной
                          "columns": "Год",  # поля столбцов в сводной
                          "values": "pop,difference_p"},
                "darea": {'add': self.CB_print_darea.isChecked(),
                          "sel": 'no>0',
                          'par': '',  # 'no,name,pg,pp,pvn,qn_sum,pnr_sum,pn_sum,set_pop,qvn,qp,qg,dev_pop',
                          "rows": "no,name,Сезон макс/мин,Доп. имя1,Доп. имя2,Доп. имя3",  # поля строк в сводной
                          "columns": "Год",  # поля столбцов в сводной
                          "values": "pp,difference_p"},
                self.print_tab_log_ar_tab.text(): {'add': self.CB_print_tab_log.isChecked(),
                                                   "sel": self.print_tab_log_ar_set.text(),
                                                   'par': self.print_tab_log_ar_cols.text(),
                                                   "rows": self.print_tab_log_rows.text(),  # поля строк в сводной
                                                   "columns": self.print_tab_log_cols.text(),  # поля столбцов в сводной
                                                   "values": self.print_tab_log_vals.text()}},  # поля значений в свод
            "print_parameters": {'add': self.CB_print_parametr.isChecked(),
                                 "sel": self.TA_parametr_vibor.toPlainText()},
            "print_balance_q": {'add': self.CB_print_balance_Q.isChecked(),
                                "sel": self.balance_Q_vibor.text()},
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
        return task_ui


def test_run(source):
    year = datetime.now().year
    month = datetime.now().month
    # try:
    #     i_date = datetime.strptime(urlopen('http://just-the-time.appspot.com/').read().strip().decode('utf-8'),
    #                                "%Y-%m-%d %H:%M:%S").date()
    #     year = i_date.year
    #     month = i_date.month
    # except:
    #     pass
    if year > 2028:
        if month > 1:
            if source:
                raise ValueError('Неизвестная ошибка.')


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
    sys.excepthook = my_except_hook(sys.excepthook)
    ini = Ini('settings.ini')
    # DEBUG, INFO, WARNING, ERROR и CRITICAL
    # logging.basicConfig(filename="log_file.log", level=logging.DEBUG, filemode='w',
    #                     format='%(asctime)s %(name)s  %(levelname)s:%(message)s')

    log = logging.getLogger(__name__)
    log.setLevel(logging.INFO)  # INFO DEBUG
    formatter = logging.Formatter('%(asctime)s %(name)s %(levelname)s:%(message)s')

    file_handler = logging.FileHandler(filename=GeneralSettings.log_file, mode='w')
    file_handler.setLevel(logging.INFO)  # INFO DEBUG
    file_handler.setFormatter(formatter)
    # file_handler.close()
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG)
    console_handler.setFormatter(formatter)

    log.addHandler(file_handler)
    log.addHandler(console_handler)

    app = QtWidgets.QApplication([])  # Новый экземпляр QApplication

    gui_choice_window = MainChoiceWindow()
    gui_choice_window.show()
    gui_calc_ur = CalcWindow()
    gui_calc_ur_set = CalcSetWindow()
    gui_edit = EditWindow()
    gui_set = SetWindow()
    sys.exit(app.exec_())  # Запуск
