"""Программа для автоматизации работы ПК RASTRWIN3"""
import logging
import os
import sys
from tkinter import messagebox as mb

import yaml
from PyQt5 import QtWidgets

from calc_model import CalcModel
from edit_model import EditModel
from ini import Ini
from qt.qt_calc_ur import Ui_calc_ur
from qt.qt_calc_ur_set import Ui_calc_ur_set
from qt.qt_choice import Ui_choice
from qt.qt_cor import Ui_cor
from qt.qt_set import Ui_Settings


# from urllib.request import urlopen


class Window:
    """ Класс с общими методами для QT. """
    dict_obj = None
    path_folder = None
    path_file = None
    section = None
    check_status_visibility = None

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
                                                                   filter=filter_)  # 'All Files(*);Text Files(*.txt)'
        if fileName_choose:
            log.info(f'GUI. Выбран файл: {fileName_choose}')
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
            log.info(f'GUI. Для сохранения выбран файл: {fileName_choose}, {_}')
            return fileName_choose

    def choice(self, insert, type_choice: str = 'file', directory=None):
        """
        Функция выбора папки или файла.
        :param type_choice: 'file', 'folder'
        :param insert: объект QT 'QPlainTextEdit' или 'QLineEdit' для вставки пути выбранного файла.
        :param directory: объект QT 'QPlainTextEdit' c начальной папкой для поиска.
        """

        def paste(txt, ins):
            """
            Вставка пути на форму
            :param txt: Значение для вставки в объект.
            :param ins: Объект
            """
            self.ins = ins
            if txt:
                txt = txt.replace('/', '\\')
                if self.ins.__class__.__name__ == 'QPlainTextEdit':
                    self.ins.setPlainText(txt)
                elif self.ins.__class__.__name__ == 'QLineEdit':
                    self.ins.setText(txt)

        if type_choice == 'file':
            self.path_file = self.choice_file(directory=directory.toPlainText().replace('*', ''))
            paste(self.path_file, insert)
        elif type_choice == 'folder':
            self.path_folder = self.choice_folder(directory=directory.toPlainText().replace('*', ''))
            paste(self.path_folder, insert)

        # Сохранить в ini
        ini.write_ini(section='save_form_folder_' + self.section,
                      key='path',
                      value=self.path_rm.toPlainText())

    def task_from_form(self) -> dict:
        """ Возвращает данные с формы QT. """
        task = {}

        def set_obj(obj):
            match obj.__class__.__name__:
                case 'QCheckBox':
                    return obj.isChecked()
                case 'QSpainBox' | 'QSpinBox':
                    return obj.value()
                case 'QComboBox':
                    return obj.currentText()
                case 'QLineEdit':
                    return obj.text()
                case 'QPlainTextEdit':
                    return obj.toPlainText()
                case 'str':
                    return ''
                case _:
                    raise TypeError(f'Добавить {obj.__class__.__name__}')

        for k1 in self.dict_obj:
            if isinstance(self.dict_obj[k1], dict):
                task[k1] = {}
                for k2 in self.dict_obj[k1]:
                    if isinstance(self.dict_obj[k1][k2], dict):
                        task[k1][k2] = {}
                        for k3 in self.dict_obj[k1][k2]:
                            task[k1][k2][k3] = set_obj(self.dict_obj[k1][k2][k3])
                    else:
                        task[k1][k2] = set_obj(self.dict_obj[k1][k2])
            else:
                task[k1] = set_obj(self.dict_obj[k1])
        return task

    def task_from_yaml(self):
        """ Загрузить данные из yaml на форму. """
        _filter = None
        field = self.path_rm
        if self.section == 'calc':
            _filter = 'YAML Files (*.calc);;All files (*.*)'
        elif self.section == 'edit':
            _filter = 'YAML Files (*.cor);;All files (*.*)'
        name_file_load = self.choice_file(directory=field.toPlainText().replace('*', ''),
                                          filter_=_filter)
        if not name_file_load:
            return
        with open(name_file_load) as f:
            task_yaml = yaml.safe_load(f)
        if not task_yaml:
            return
        msg_list = []
        for k1 in self.dict_obj:
            if isinstance(self.dict_obj[k1], dict):
                for k2 in self.dict_obj[k1]:
                    if isinstance(self.dict_obj[k1][k2], dict):
                        for k3 in self.dict_obj[k1][k2]:
                            msg_list.append(self.get_in_form(self.dict_obj[k1][k2][k3], task_yaml, (k1, k2, k3)))
                    else:
                        msg_list.append(self.get_in_form(self.dict_obj[k1][k2], task_yaml, (k1, k2)))
            else:
                msg_list.append(self.get_in_form(self.dict_obj[k1], task_yaml, (k1,)))
        msg_list = [i for i in msg_list if i]
        msg = '\n'.join(msg_list)
        if msg:
            mb.showerror('Ошибка чтения yaml файла', msg)

        self.check_status(self.check_status_visibility)

    @staticmethod
    def get_in_form(obj, task: dict, keys: tuple) -> str | None:
        """
        Присвоить значение объекту формы из задания. Если ключи в задании отсутствует, то выводится предупреждение.
        :param obj: Функция объекта для присвоения значения, например cb_filter.setChecked.
        :param task: Словарь со считываемыми данными.
        :param keys: Картеж ключей для task. Например, keys = (k1, k2) для task[k1][k2].
        :return:  Возвращает сообщение об ошибке или None.
        """
        msg = None
        try:
            for key in keys:
                task = task[key]
            match obj.__class__.__name__:
                case 'QCheckBox':
                    if isinstance(task, bool):
                        obj.setChecked(task)
                    else:
                        obj.setText(task)
                case 'QSpainBox' | 'QSpinBox':
                    obj.setValue(task)
                case 'QComboBox':
                    obj.setCurrentText(task)
                case 'QLineEdit':
                    obj.setText(task)
                case 'QPlainTextEdit':
                    obj.setPlainText(task)
                case 'str':
                    pass
                case _:
                    log.info(f'Добавить {obj.__class__.__name__}')
        except KeyError:
            msg = f'Отсутствует ключ: {keys};'
        except TypeError:
            msg = f'У значения по ключу: {keys} неверный тип данных;'
        finally:
            if msg:
                log.error(msg)
            return msg


class MainChoiceWindow(QtWidgets.QMainWindow, Ui_choice, Window):
    """
    Окно главного меню.
    """

    def __init__(self):
        super(MainChoiceWindow, self).__init__()
        self.setupUi(self)
        self.settings.clicked.connect(lambda: gui_set.show())
        self.correction.clicked.connect(lambda: self.hide_show((gui_choice_window,),
                                                               (gui_edit,)))
        self.calc_ur.clicked.connect(lambda: self.hide_show((gui_choice_window,),
                                                            (gui_calc_ur,)))


class CalcWindow(QtWidgets.QMainWindow, Ui_calc_ur, Window):
    """
    Форма задания и запуска расчетов УР.
    """

    def __init__(self):
        super(CalcWindow, self).__init__()
        self.section = 'calc'
        self.setupUi(self)
        self.b_set.clicked.connect(lambda: gui_calc_ur_set.show())
        self.b_main_choice.clicked.connect(lambda: self.hide_show((gui_calc_ur,),
                                                                  (gui_choice_window,)))
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
        self.b_task_load.clicked.connect(self.task_from_yaml)

        self.b_choice_path_folder.clicked.connect(lambda: self.choice(type_choice='folder',
                                                                      insert=self.path_rm,
                                                                      directory=self.path_rm))
        self.b_choice_path_file.clicked.connect(lambda: self.choice(insert=self.path_rm,
                                                                    directory=self.path_rm))
        self.b_choice_XL.clicked.connect(lambda: self.choice(insert=self.te_XL_path,
                                                             directory=self.path_rm))
        self.b_choice_path_import_folder.clicked.connect(lambda: self.choice(type_choice='folder',
                                                                             insert=self.te_path_import_rg2,
                                                                             directory=self.path_rm))
        self.b_choice_path_import_file.clicked.connect(lambda: self.choice(insert=self.te_path_import_rg2,
                                                                           directory=self.path_rm))

        self.run_calc_rg2.clicked.connect(lambda: self.start())
        self.path_rm.setPlainText(ini.read_ini(section='save_form_folder_calc', key='path'))
        # Подсказки
        self.le_control_sel.setToolTip('Если поле не заполнено, то контролируются все ветви и узлы РМ')
        self.path_rm.setToolTip('Для расчета файлов во всех вложенных папках нужно в конце поставить *')

        self.dict_obj = {
            # Окно запуска расчета.
            'source_path': self.path_rm,
            # Выборка файлов.
            'filter_file': self.cb_filter,
            'max_count_file': self.sb_count_file,
            'criterion': {'years': self.le_condition_file_years,
                          'season': self.le_condition_file_season,
                          'max_min': self.le_condition_file_max_min,
                          'add_name': self.le_condition_file_add_name},
            # Корректировка в txt.
            'cor_rm': {'add': self.cb_cor_txt,
                       'txt': self.te_cor_txt},
            # Импорт ИД для расчетов УР из моделей.
            'CB_Import_Rg2': self.cb_import_model,
            'Import_file': self.te_path_import_rg2,
            'txt_Import_Rg2': self.te_import_rg2,
            # Расчет всех возможных сочетаний. Отключаемые элементы.
            'cb_disable_comb': self.cb_disable_comb,
            'SRS': {'n-1': self.cb_n1,
                    'n-2_abv': self.cb_n2_abv,
                    'n-2_gd': self.cb_n2_gd,
                    'n-3': self.cb_n3},

            'cb_auto_disable': self.cb_auto_disable,
            'auto_disable_choice': self.le_auto_disable_choice,

            'cb_comb_field': self.cb_comb_field,

            'filter_comb': self.cb_filter_comb,
            'filter_comb_val': self.sb_filter_comb_val,
            # Импорт перечня расчетных сочетаний из EXCEL
            'cb_disable_excel': self.cb_disable_excel,
            'srs_XL_path': self.te_XL_path,
            'srs_XL_sheets': self.le_XL_sheets,
            # Расчет всех возможных сочетаний. Контролируемые элементы.
            'cb_control': self.cb_control,
            'cb_control_field': self.cb_control_field,
            'cb_control_sel': self.cb_control_sel,
            'control_sel': self.le_control_sel,
            # Результаты в EXCEL: таблицы контролируемые - отключаемые элементы
            'cb_save_i': self.cb_save_i,
            'cb_tab_KO': self.cb_tab_KO,
            'te_tab_KO_info': self.te_tab_KO_info,
            # Результаты в RG2
            'results_RG2': self.cb_results_pic,
            'pic_overloads': self.cb_pic_overloads,
            'name_pic': self.te_name_pic, }

    def task_save_yaml(self):
        name_file_save = self.save_file(directory=self.path_rm.toPlainText(),
                                        filter_='YAML Files (*.calc);;All files (*.*)')
        if name_file_save:
            with open(name_file_save, 'w') as f:
                yaml.dump(data=self.task_from_form() | ini.to_dict(),
                          stream=f, default_flow_style=False, sort_keys=False, allow_unicode=True)

    def start(self):
        """
        Запуск расчета моделей
        """
        ini.write_ini(section='save_form_folder_calc',
                      key='path',
                      value=self.path_rm.toPlainText())
        config = self.task_from_form() | ini.to_dict()
        cm = CalcModel(config)
        end_info = cm.run()
        cm.save_log(name_file_source=log_file)
        mb.showinfo('Инфо', end_info)


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
                self.cb_gost.setChecked(eval(config['gost']))
                self.cb_skrm.setChecked(eval(config['skrm']))
                self.cb_avr.setChecked(eval(config['avr']))
                self.cb_add_disabling_repair.setChecked(eval(config['add_disabling_repair']))
                self.cb_pa.setChecked(eval(config['pa']))
            except Exception:
                log.error(f'Файл {ini.name} [CalcSetWindow] не читается, перезаписан.')
                self.save_ini_ur()
        else:
            log.info(f'Создан файл {ini.name}.')
            self.save_ini_ur()

    def save_ini_ur(self):
        ini.add(info={'gost': self.cb_gost.isChecked(),
                      'skrm': self.cb_skrm.isChecked(),
                      'avr': self.cb_avr.isChecked(),
                      'add_disabling_repair': self.cb_add_disabling_repair.isChecked(),
                      'pa': self.cb_pa.isChecked()},
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
            config = ini.read_ini(section='Settings')
            try:
                self.LE_shablon_rg2.setText(config['шаблон rg2'])
                self.LE_shablon_rst.setText(config['шаблон rst'])
                self.LE_shablon_sch.setText(config['шаблон sch'])
                self.LE_shablon_trn.setText(config['шаблон trn'])
                self.LE_shablon_anc.setText(config['шаблон anc'])
                self.CB_load_add.setChecked(eval(config['load_add']))
            except Exception:
                log.error(f'Файл {ini.name} [Settings] не читается, перезаписан.')
                self.create_start_value()
                self.save_ini()
        else:
            log.info(f'Создан файл {ini.name}.')
            self.save_ini()

    def create_start_value(self) -> None:
        """
        Записать в настройки путь к шаблонам (если получится его определить)
        """
        documents_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents\RastrWin3\SHABLON')
        if os.path.exists(documents_path):
            self.LE_shablon_rg2.setText(os.path.join(documents_path, 'режим.rg2'))
            self.LE_shablon_rst.setText(os.path.join(documents_path, 'динамика.rst'))
            self.LE_shablon_sch.setText(os.path.join(documents_path, 'сечения.sch'))
            self.LE_shablon_trn.setText(os.path.join(documents_path, 'трансформаторы.trn'))
            self.LE_shablon_anc.setText(os.path.join(documents_path, 'анцапфы.anc'))

    def save_ini(self):
        ini.add(info={'шаблон rg2': self.LE_shablon_rg2.text(),
                      'шаблон rst': self.LE_shablon_rst.text(),
                      'шаблон sch': self.LE_shablon_sch.text(),
                      'шаблон trn': self.LE_shablon_trn.text(),
                      'шаблон anc': self.LE_shablon_anc.text(),
                      'load_add': self.CB_load_add.isChecked()},
                key='Settings')


class EditWindow(QtWidgets.QMainWindow, Ui_cor, Window):
    """
    Окно корректировки моделей.
    """

    def __init__(self):
        super(EditWindow, self).__init__()  # *args, **kwargs
        self.section = 'edit'
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

        # Набор соотношения: проверяемый на отметку элемент, элемент у которого меняется видимость.
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
        self.check_status(self.check_status_visibility)  # Скрыть при старте

        # CB показать / скрыть параметры при переключении.
        for CB, element in self.check_status_visibility:
            CB.clicked.connect(lambda: self.check_status(self.check_status_visibility))
        # CB показать список импортируемых моделей.
        for CB, _ in self.check_import:
            CB.clicked.connect(lambda: self.import_name_table())

        # Функциональные кнопки
        self.task_save.clicked.connect(self.task_save_yaml)
        self.task_load.clicked.connect(self.task_from_yaml)
        self.task_load.clicked.connect(lambda: self.import_name_table())

        self.choice_XL.clicked.connect(lambda: self.choice(insert=self.T_PQN_XL_File,
                                                           directory=self.path_rm))
        self.choice_N.clicked.connect(lambda: self.choice(insert=self.file_N,
                                                          directory=self.path_rm))
        self.choice_V.clicked.connect(lambda: self.choice(insert=self.file_V,
                                                          directory=self.path_rm))
        self.choice_G.clicked.connect(lambda: self.choice(insert=self.file_G,
                                                          directory=self.path_rm))
        self.choice_A.clicked.connect(lambda: self.choice(insert=self.file_A,
                                                          directory=self.path_rm))
        self.choice_A2.clicked.connect(lambda: self.choice(insert=self.file_A2,
                                                           directory=self.path_rm))
        self.choice_D.clicked.connect(lambda: self.choice(insert=self.file_D,
                                                          directory=self.path_rm))
        self.choice_PQ.clicked.connect(lambda: self.choice(insert=self.file_PQ,
                                                           directory=self.path_rm))
        self.choice_IT.clicked.connect(lambda: self.choice(insert=self.file_IT,
                                                           directory=self.path_rm))
        self.choice_from_file.clicked.connect(lambda: self.choice(insert=self.path_rm,
                                                                  directory=self.path_rm))

        self.choice_from_folder.clicked.connect(lambda: self.choice(type_choice='folder',
                                                                    insert=self.path_rm,
                                                                    directory=self.path_rm))
        self.choice_in_folder.clicked.connect(lambda: self.choice(type_choice='folder',
                                                                  insert=self.T_InFolder,
                                                                  directory=self.path_rm))

        self.run_krg2.clicked.connect(lambda: self.start())
        self.b_main_choice.clicked.connect(lambda: self.hide_show((gui_edit,), (gui_choice_window,)))
        # Подсказки
        self.path_rm.setToolTip('Для корректировки файлов во всех вложенных папках нужно в конце поставить *')
        # Загрузить из .ini начальный путь для path_rm
        self.path_rm.setPlainText(ini.read_ini(section='save_form_folder_edit', key='path'))

        self.dict_obj = {
            'source_path': self.path_rm,
            'target_path': self.T_InFolder,
            # Выборка файлов.
            'filter_file': self.CB_KFilter_file,
            'max_count_file': self.D_count_file,
            'criterion': {'years': self.condition_file_years,
                          'season': self.condition_file_season,
                          'max_min': self.condition_file_max_min,
                          'add_name': self.condition_file_add_name},
            # Корректировка в начале.
            'cor_beginning_qt': {'add': self.CB_cor_b,
                                 'txt': self.TE_cor_b},
            # Задание из 'EXCEL'
            'import_val_XL': self.CB_import_val_XL,
            'excel_cor_file': self.T_PQN_XL_File,
            'excel_cor_sheet': self.T_PQN_Sheets,

            # Расчет режима и контроль параметров режима
            'checking_parameters_rg2': self.CB_kontrol_rg2,
            'control_rg2_task': {'node': self.CB_U,
                                 'vetv': self.CB_I,
                                 'Gen': self.CB_gen,
                                 'sel_node': self.kontrol_rg2_Sel},
            # Выводить данные из моделей в XL
            'printXL': self.CB_printXL,
            'set_printXL': {
                'sechen': {'add': self.CB_print_sech},
                'area': {'add': self.CB_print_area},
                'area2': {'add': self.CB_print_area2},
                'darea': {'add': self.CB_print_darea},
                'таблица на выбор': {'tab_name': self.print_tab_log_ar_tab,
                                     'add': self.CB_print_tab_log,
                                     'sel': self.print_tab_log_ar_set,
                                     'par': self.print_tab_log_ar_cols,
                                     'rows': self.print_tab_log_rows,  # поля строк в сводной
                                     'columns': self.print_tab_log_cols,  # поля столбцов в сводной
                                     'values': self.print_tab_log_vals}},  # поля значений в свод
            'print_parameters': {'add': self.CB_print_parametr,
                                 'sel': self.TA_parametr_vibor},
            'print_balance_q': {'add': self.CB_print_balance_Q,
                                'sel': self.balance_Q_vibor},
            # только для UI
            'imp_rg2_name': self.CB_ImpRg2,
            'imp_rg2': self.CB_ImpRg2,
            'Imp_add': {
                'node': {'add': self.CB_N,
                         'import_file_name': self.file_N,
                         'selection': self.CB_Filtr_N,
                         'years': self.Filtr_god_N,
                         'season': self.Filtr_sez_N,
                         'max_min': self.Filtr_max_min_N,
                         'add_name': self.Filtr_dop_name_N,
                         'tables': self.tab_N,
                         'param': self.param_N,
                         'sel': self.sel_N,
                         'calc': self.tip_N, },
                'vetv': {'add': self.CB_V,
                         'import_file_name': self.file_V,
                         'selection': self.CB_Filtr_V,
                         'years': self.Filtr_god_V,
                         'season': self.Filtr_sez_V,
                         'max_min': self.Filtr_max_min_V,
                         'add_name': self.Filtr_dop_name_V,
                         'tables': self.tab_V,
                         'param': self.param_V,
                         'sel': self.sel_V,
                         'calc': self.tip_V, },
                'gen': {'add': self.CB_G,
                        'import_file_name': self.file_G,
                        'selection': self.CB_Filtr_G,
                        'years': self.Filtr_god_G,
                        'season': self.Filtr_sez_G,
                        'max_min': self.Filtr_max_min_G,
                        'add_name': self.Filtr_dop_name_G,
                        'tables': self.tab_G,
                        'param': self.param_G,
                        'sel': self.sel_G,
                        'calc': self.tip_G, },
                'area': {'add': self.CB_A,
                         'import_file_name': self.file_A,
                         'selection': self.CB_Filtr_A,
                         'years': self.Filtr_god_A,
                         'season': self.Filtr_sez_A,
                         'max_min': self.Filtr_max_min_A,
                         'add_name': self.Filtr_dop_name_A,
                         'tables': self.tab_A,
                         'param': self.param_A,
                         'sel': self.sel_A,
                         'calc': self.tip_A, },
                'area2': {'add': self.CB_A2,
                          'import_file_name': self.file_A2,
                          'selection': self.CB_Filtr_A2,
                          'years': self.Filtr_god_A2,
                          'season': self.Filtr_sez_A2,
                          'max_min': self.Filtr_max_min_A2,
                          'add_name': self.Filtr_dop_name_A2,
                          'tables': self.tab_A2,
                          'param': self.param_A2,
                          'sel': self.sel_A2,
                          'calc': self.tip_A2, },
                'darea': {'add': self.CB_D,
                          'import_file_name': self.file_D,
                          'selection': self.CB_Filtr_D,
                          'years': self.Filtr_god_D,
                          'season': self.Filtr_sez_D,
                          'max_min': self.Filtr_max_min_D,
                          'add_name': self.Filtr_dop_name_D,
                          'tables': self.tab_D,
                          'param': self.param_D,
                          'sel': self.sel_D,
                          'calc': self.tip_D, },
                'PQ': {'add': self.CB_PQ,
                       'import_file_name': self.file_PQ,
                       'selection': self.CB_Filtr_PQ,
                       'years': self.Filtr_god_PQ,
                       'season': self.Filtr_sez_PQ,
                       'max_min': self.Filtr_max_min_PQ,
                       'add_name': self.Filtr_dop_name_PQ,
                       'tables': self.tab_PQ,
                       'param': self.param_PQ,
                       'sel': self.sel_PQ,
                       'calc': self.tip_PQ, },
                'IT': {'add': self.CB_IT,
                       'import_file_name': self.file_IT,
                       'selection': self.CB_Filtr_IT,
                       'years': self.Filtr_god_IT,
                       'season': self.Filtr_sez_IT,
                       'max_min': self.Filtr_max_min_IT,
                       'add_name': self.Filtr_dop_name_IT,
                       'tables': self.tab_IT,
                       'param': self.param_IT,
                       'sel': self.sel_IT,
                       'calc': self.tip_IT, },
            }
        }

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
        name_file_save = self.save_file(directory=self.path_rm.toPlainText(),
                                        filter_='YAML Files (*.cor);;All files (*.*)')
        if not name_file_save:
            raise ValueError('Не указан путь к сохраняемому файлу.')
        with open(name_file_save, 'w') as f:
            yaml.dump(data=self.task_from_form() | ini.to_dict(),
                      stream=f, default_flow_style=False, sort_keys=False, allow_unicode=True)

    def start(self):
        """
        Запуск корректировки моделей.
        """
        ini.write_ini(section='save_form_folder_edit',
                      key='path',
                      value=self.path_rm.toPlainText())
        if self.print_tab_log_ar_tab.text() in ['area', 'area2', 'darea', 'sechen']:
            mb.showerror('Ошибка',
                         'В поле таблица на выбор нельзя задавать таблицы: area, area2, darea, sechen.')
            return
        config = self.task_from_form() | ini.to_dict()

        em = EditModel(config)
        end_info = em.run()
        em.save_log(name_file_source=log_file)
        mb.showinfo('Инфо', end_info)


def my_except_hook(func):
    """
    Переназначить функцию для добавления информации об ошибке в диалоговое окно.
    :param func:
    :return:
    """

    def new_func(*args, **kwargs):
        log.error(f'Критическая ошибка: {args[0]}, {args[1]}', exc_info=True)
        mb.showerror('Ошибка', f'Критическая ошибка: {args[0]}, {args[1]}')
        # https://python-scripts.com/python-traceback
        func(*args, **kwargs)

    return new_func


if __name__ == '__main__':
    sys.excepthook = my_except_hook(sys.excepthook)
    log_file = f'log_file.log'
    ini = Ini('settings.ini')
    ini.write_ini(key='path_project', section='other', value=os.getcwd())
    # DEBUG, INFO, WARNING, ERROR и CRITICAL
    # logging.basicConfig(filename='log_file.log', level=logging.DEBUG, filemode='w',
    #                     format='%(asctime)s %(name)s  %(levelname)s:%(message)s')

    log = logging.getLogger(__name__)
    log.setLevel(logging.DEBUG)  # INFO DEBUG
    formatter = logging.Formatter('%(asctime)s %(name)s %(levelname)s:%(message)s')

    file_handler = logging.FileHandler(filename=log_file, mode='w')
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
