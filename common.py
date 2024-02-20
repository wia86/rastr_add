import shutil
from abc import ABC
from datetime import datetime
import time
import logging

import yaml

log_g_s = logging.getLogger(f'__main__.{__name__}')


class Common(ABC):
    """
    Содержит общие атрибуты и методы для дочерних классов EditModel и CalcModel.
    """

    def __init__(self):
        # коллекция для хранения информации о расчете
        self.config = {'collapse': [],
                       'end_info': ''}
        self.file_count = 0  # Счетчик расчетных файлов.
        self.now = datetime.now()
        self.time_start = time.time()
        self.now_start = self.now.strftime("%d-%m-%Y %H:%M:%S")

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
        execution_time = time.strftime("%H:%M:%S", time.gmtime(time.time() - self.time_start))
        self.config['end_info'] = (
            f"РАСЧЕТ ЗАКОНЧЕН!"
            f"\nНачало расчета {self.now_start}."
            f"\nКонец {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}."
            f"\nЗатрачено: {execution_time} (файлов: {self.file_count}).")
        log_g_s.info(self.config['end_info'])

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
