import logging
import os
import shutil
from abc import ABC
from datetime import datetime
from pathlib import Path
from tkinter import messagebox as mb

import yaml

log_comm = logging.getLogger(f'__main__.{__name__}')


class Common(ABC):
    """
    Содержит общие атрибуты и методы для дочерних классов EditModel и CalcModel.
    """
    size_date_source: str | None = None  # Объем расчетных файлов: 'file', 'folder' или 'nested_folder'
    mark: str | None = None  # Метка класса ('calc' или 'cor')
    source_path = None  # Путь к папке с исходными РМ.
    target_path = None  # Путь к папке для сохранения РМ.
    folder_result = None  # Путь к папке для сохранения результатов работы программы.

    def __init__(self):
        # коллекция для хранения информации о расчете
        self.config = {'collapse': [],
                       'end_info': ''}
        self.file_count = 0  # Счетчик расчетных файлов.
        self.time_start = datetime.now()
        self.time_str_format = '%d-%m-%Y %H_%M_%S'

    def run_common(self):
        """Общие действия подклассов при запуске."""
        self.source_path = self.config['source_path'].strip()

        if '*' in self.source_path:
            self.source_path = self.source_path.replace('*', '')
            self.size_date_source = 'nested_folder'

        if not os.path.exists(self.source_path):
            raise ValueError(f'Не найден путь к данным: {self.source_path}.')

        if os.path.isfile(self.source_path):
            self.size_date_source = 'file'
        else:
            self.size_date_source = 'folder'

        # Создать папку target_path.
        if 'target_path' in self.config:
            self.target_path = self.config['target_path'].strip()
            if self.target_path:
                if self.size_date_source == 'file':
                    dir_ = os.path.dirname(self.target_path)
                    self.folder_result = os.path.join(dir_, self.mark)
                else:
                    Path(self.target_path).mkdir(parents=True, exist_ok=True)
                    self.folder_result = os.path.join(self.target_path, self.mark)

        if not self.folder_result:
            if self.size_date_source == 'file':
                dir_ = os.path.dirname(self.source_path)
                self.folder_result = os.path.join(dir_, self.mark)
            else:
                self.folder_result = os.path.join(self.source_path, self.mark)

        # Создать папку result.
        Path(self.folder_result).mkdir(parents=True, exist_ok=True)

        self.config['name_time'] = os.path.join(self.folder_result,
                                                self.time_start.strftime(self.time_str_format))

    def save_log(self, name_file_source):
        """
        Сохранить файл логирования.
        :param self:
        :param name_file_source: Имя файла с логом.
        """
        path_new_log = f'{self.config["name_time"]} протокол.log'
        shutil.copyfile(name_file_source,
                        path_new_log)

    @staticmethod
    def save_config(config, extension):
        """
        Сохранить файл задания и конфигураций.

        :param config: Словарь для сохранения.
        :param extension: Расширение нового файла
        :return:
        """
        with open(f'{config["name_time"]} задание.{extension}', 'w') as f:
            yaml.dump(data=config,
                      stream=f,
                      default_flow_style=False,
                      sort_keys=False,
                      allow_unicode=True)

    def the_end(self):  # по завершению
        time_end = datetime.now()
        execution_time = time_end - self.time_start
        self.config['end_info'] = (
            f'РАСЧЕТ ЗАКОНЧЕН!'
            f'\nНачало расчета {self.time_start.strftime(self.time_str_format)}.'
            f'\nКонец {time_end.strftime(self.time_str_format)}.'
            f'\nЗатрачено: {execution_time} (файлов: {self.file_count}).')
        if self.config['collapse']:
            t = f',\n'.join(self.config['collapse'])
            self.config['end_info'] += f'\nВНИМАНИЕ! Развалились модели:\n[{t}].'
        log_comm.info(self.config['end_info'])
        self.save_config(self.config, self.mark)
        mb.showinfo('Инфо', self.config['end_info'])

    @staticmethod
    def read_title(txt: str) -> tuple:
        """
        Разделить строку типа 'Рисунок [1] - Южный'.
        :param txt:
        :return: (1, ['Рисунок ', ' - Южный']).
        """
        txt = txt.strip()
        num = txt[txt.find('[') + 1: txt.find(']')]
        txt = txt.split(f'[{num}]')
        num = int(num) if num.isdigit() else 1
        return num, txt
