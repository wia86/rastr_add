import shutil
from datetime import datetime
import os
import logging
from tkinter import messagebox as mb
import pandas as pd
from rastr_model import RastrModel
from import_rm import ImportFromModel
from cor_xl import CorXL
from print_xl import PrintXL
import yaml

from common import Common
# import loading_sections as ls
log_ed_mod = logging.getLogger(f'__main__.{__name__}')


class EditModel(Common):
    """
    Изменение расчетных моделей.
    """

    def __init__(self, config: dict):
        """
        :param config: Задание и настройки программы
        """
        super(EditModel, self).__init__()
        self.config = config
        RastrModel.config = config['Settings']
        RastrModel.overwrite_new_file = 'question'
        self.cor_xl = None
        self.print_xl = None
        self.rastr_files = None
        self.all_folder = False  # Не перебирать вложенные папки
        RastrModel.all_rm = pd.DataFrame()

        # Добавление импорта данных из РМ с формы.
        self.set_import_model = []
        if self.config['CB_ImpRg2']:
            for tabl in self.config['Imp_add']:
                dict_tabl = self.config['Imp_add'][tabl]
                if dict_tabl['add']:
                    criterion_start = {}
                    if dict_tabl['selection']:
                        criterion_start = {"years": dict_tabl['years'],
                                           "season": dict_tabl['season'],
                                           "max_min": dict_tabl['max_min'],
                                           "add_name": dict_tabl['add_name']}

                    ifm = ImportFromModel(export_rm=RastrModel(full_name=dict_tabl['import_file_name'],
                                                               not_calculated=True),
                                          criterion_start=criterion_start,
                                          tables=dict_tabl['tables'],
                                          param=dict_tabl['param'],
                                          sel=dict_tabl['sel'],
                                          calc=dict_tabl['calc'])
                    self.set_import_model.append(ifm)

    def run(self):
        """
        Запуск корректировки моделей.
        """
        # test_run('edit')
        log_ed_mod.info('\n!!! Запуск корректировки РМ !!!\n')
        self.config["KIzFolder"] = self.config["KIzFolder"].strip()
        if "*" in self.config["KIzFolder"]:
            self.config["KIzFolder"] = self.config["KIzFolder"].replace('*', '')
            self.all_folder = True

        if not os.path.exists(self.config["KIzFolder"]):
            raise ValueError(f'Не найден путь: {self.config["KIzFolder"]}.')

        self.config['folder_result'] = self.config["KIzFolder"] + r"\result"
        if os.path.isfile(self.config["KIzFolder"]):
            self.config['folder_result'] = os.path.dirname(self.config["KIzFolder"]) + r"\result"

        self.config["KInFolder"] = self.config["KInFolder"].strip()
        # папка для сохранения result и KInFolder
        if self.config["KInFolder"] and not os.path.exists(self.config["KInFolder"]):
            if os.path.isdir(self.config["KIzFolder"]):
                log_ed_mod.info("Создана папка: " + self.config["KInFolder"])
                os.makedirs(self.config["KInFolder"])  # создать папку
                self.config['folder_result'] = self.config["KInFolder"] + r"\result"
            else:
                self.config['folder_result'] = os.path.dirname(self.config["KIzFolder"]) + r"\result"

        if not os.path.exists(self.config['folder_result']):
            os.mkdir(self.config['folder_result'])  # создать папку result

        self.config['name_time'] = f"{self.config['folder_result']}\\{datetime.now().strftime('%d-%m-%Y %H-%M-%S')}"

        if "import_val_XL" in self.config:
            # Задать параметры узла по значениям в таблице excel (имя книги, имя листа)
            if self.config["import_val_XL"]:
                self.cor_xl = CorXL(excel_file_name=self.config["excel_cor_file"],
                                    sheets=self.config["excel_cor_sheet"])
                self.cor_xl.init_export_model()

        if os.path.isdir(self.config["KIzFolder"]):  # корр файлы в папке
            if self.all_folder:  # с вложенными папками
                for address, dirs, files in os.walk(self.config["KIzFolder"]):
                    in_dir = ''
                    if self.config["KInFolder"]:
                        in_dir = address.replace(self.config["KIzFolder"], self.config["KInFolder"])
                        if not os.path.exists(in_dir):
                            os.makedirs(in_dir)

                    self.for_file_in_dir(from_dir=address, in_dir=in_dir)

            else:  # без вложенных папок
                self.for_file_in_dir(from_dir=self.config["KIzFolder"], in_dir=self.config["KInFolder"])

        elif os.path.isfile(self.config["KIzFolder"]):  # корр файл
            rm = RastrModel(full_name=self.config["KIzFolder"])
            log_ed_mod.info("\n\n")
            rm.load()

            self.cor_file(rm)
            if self.config["KInFolder"]:
                if os.path.isdir(self.config["KInFolder"]):
                    rm.save(full_name_new=os.path.join(self.config["KInFolder"], rm.name_base))
                else:  # if os.path.isfile(self.config["KInFolder"]):
                    rm.save(full_name_new=self.config["KInFolder"])

        if self.print_xl:
            self.print_xl.finish()

        self.the_end()
        if self.set_info['collapse']:
            t = f',\n'.join(self.set_info['collapse'])
            self.set_info['end_info'] += f"\nВНИМАНИЕ! Развалились модели:\n[{t}]."

        notepad_path = self.config['name_time'] + ' протокол коррекции файлов.log'
        shutil.copyfile(self.config['other']['log_file'], notepad_path)
        with open(self.config['name_time'] + ' задание.cor', 'w') as f:
            yaml.dump(data=self.config, stream=f, default_flow_style=False, sort_keys=False, allow_unicode=True)
        mb.showinfo("Инфо", self.set_info['end_info'])

    def for_file_in_dir(self, from_dir: str, in_dir: str):
        files = os.listdir(from_dir)  # список всех файлов в папке
        self.rastr_files = list(filter(lambda x: x.endswith('.rg2') | x.endswith('.rst'), files))

        for rastr_file in self.rastr_files:  # цикл по файлам .rg2 .rst в папке KIzFolder
            if self.config["KFilter_file"] and self.file_count == self.config["max_file_count"]:
                break  # Если включен фильтр файлов проверяем количество расчетных файлов.
            full_name = os.path.join(from_dir, rastr_file)

            rm = RastrModel(full_name)
            # если включен фильтр файлов и имя стандартизовано
            if self.config["KFilter_file"] and rm.code_name_rg2:
                if not rm.test_name(condition=self.config["cor_criterion_start"], info='Цикл по файлам.'):
                    continue  # пропускаем если не соответствует фильтру
            log_ed_mod.info("\n\n")
            rm.load()
            self.cor_file(rm)
            if self.config["KInFolder"]:
                rm.save(full_name_new=os.path.join(in_dir, rastr_file))

    def cor_file(self, rm):
        """Корректировать файл rm"""
        self.file_count += 1

        # Импорт моделей
        if self.set_import_model:
            for im in self.set_import_model:
                im.import_data_in_rm(rm)

        if self.config['cor_beginning_qt']['add']:
            log_ed_mod.info("\t*** Корректировка моделей в текстовом формате ***")
            rm.cor_rm_from_txt(self.config['cor_beginning_qt']['txt'])
            log_ed_mod.info("\t*** Конец выполнения корректировки моделей в текстовом формате ***")

        # Задать параметры по значениям в таблице excel
        if self.config.get("import_val_XL", False):
            self.cor_xl.run_xl(rm)

        if self.config.get("checking_parameters_rg2", False):
            if not rm.checking_parameters_rg2(self.config['control_rg2_task']):  # Расчет и контроль параметров режима.
                self.set_info['collapse'].append(rm.name_base)

        if self.config.get("printXL", False):
            if not isinstance(self.print_xl, PrintXL):
                self.print_xl = PrintXL(self.config)
            self.print_xl.add_val(rm)
