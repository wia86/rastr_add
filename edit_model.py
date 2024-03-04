import logging
import os

import pandas as pd

from common import Common
from cor_xl import CorXL
from import_rm import ImportFromModel
from print_xl import PrintXL
from rastr_model import RastrModel

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
        self.mark = 'cor'
        self.config = self.config | config
        name_tab = self.config['set_printXL']['таблица на выбор']['tab_name']
        self.config['set_printXL'][name_tab] = self.config['set_printXL']['таблица на выбор']
        del self.config['set_printXL']['таблица на выбор']

        RastrModel.config = config['Settings']
        RastrModel.overwrite_new_file = 'question'
        self.cor_xl = None
        self.print_xl = None
        RastrModel.all_rm = pd.DataFrame()

        # Добавление импорта данных из РМ с формы.
        self.set_import_model = []
        if self.config['imp_rg2']:
            for tabl in self.config['Imp_add']:
                dict_tabl = self.config['Imp_add'][tabl]
                if dict_tabl['add']:
                    criterion_start = {}
                    if dict_tabl['selection']:
                        criterion_start = {'years': dict_tabl['years'],
                                           'season': dict_tabl['season'],
                                           'max_min': dict_tabl['max_min'],
                                           'add_name': dict_tabl['add_name']}

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
        log_ed_mod.info('\n!!! Запуск корректировки РМ !!!\n')
        self.run_common()

        if 'import_val_XL' in self.config:
            # Задать параметры узла по значениям в таблице excel (имя книги, имя листа)
            if self.config['import_val_XL']:
                self.cor_xl = CorXL(excel_file_name=self.config['excel_cor_file'],
                                    sheets=self.config['excel_cor_sheet'])
                self.cor_xl.init_export_model()

        if self.size_date_source == 'nested_folder':
            for address, dirs, files in os.walk(self.source_path):
                in_dir = ''
                if self.target_path:
                    in_dir = address.replace(self.source_path, self.target_path)
                    if not os.path.exists(in_dir):
                        os.makedirs(in_dir)
                self.cycle_rm(path_folder=address, in_dir=in_dir)

        elif self.size_date_source == 'folder':
            self.cycle_rm(path_folder=self.source_path, in_dir=self.target_path)

        elif self.size_date_source == 'file':
            rm = RastrModel(full_name=self.config['init_folder'])
            log_ed_mod.info('\n\n')
            rm.load()

            self.run_file(rm)
            if self.target_path:
                if os.path.isdir(self.target_path):
                    rm.save(full_name_new=os.path.join(self.target_path, rm.name_base))
                else:
                    rm.save(full_name_new=self.target_path)

        if self.print_xl:
            self.print_xl.finish()

        return self.the_end()

    def cycle_rm(self, path_folder: str, in_dir: str):
        """Цикл по файлам"""
        gen_files = (f for f in os.listdir(path_folder) if f.endswith(('.rg2', '.rst')))

        for rastr_file in gen_files:  # цикл по файлам .rg2 .rst в папке source_path
            if self.config['filter_file'] and self.file_count == self.config['max_count_file']:
                break  # Если включен фильтр файлов проверяем количество расчетных файлов.
            full_name = os.path.join(path_folder, rastr_file)

            rm = RastrModel(full_name)
            # если включен фильтр файлов и имя стандартизовано
            if self.config['filter_file'] and rm.code_name_rg2:
                if not rm.test_name(condition=self.config['criterion'], info='Цикл по файлам.'):
                    continue  # Пропустить, если не соответствует фильтру
            log_ed_mod.info('\n\n')
            self.run_file(rm)
            if self.target_path:
                rm.save(full_name_new=os.path.join(in_dir, rastr_file))

    def run_file(self, rm):
        """Корректировать файл rm"""
        rm.load()
        self.file_count += 1

        # Импорт моделей
        if self.set_import_model:
            for im in self.set_import_model:
                im.import_data_in_rm(rm)

        if self.config['cor_beginning_qt']['add']:
            log_ed_mod.info('\t*** Корректировка моделей в текстовом формате ***')
            rm.cor_rm_from_txt(self.config['cor_beginning_qt']['txt'])
            log_ed_mod.info('\t*** Конец выполнения корректировки моделей в текстовом формате ***')

        # Задать параметры по значениям в таблице excel
        if self.config.get('import_val_XL', False):
            self.cor_xl.run_xl(rm)

        if self.config.get('checking_parameters_rg2', False):
            if not rm.checking_parameters_rg2(self.config['control_rg2_task']):  # Расчет и контроль параметров режима.
                self.config['collapse'].append(rm.name_base)

        if self.config.get('printXL', False):
            if not isinstance(self.print_xl, PrintXL):
                self.print_xl = PrintXL(self.config)
            self.print_xl.add_val(rm)
