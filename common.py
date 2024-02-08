from abc import ABC
from datetime import datetime
import time
import logging

log_g_s = logging.getLogger(f'__main__.{__name__}')


class Common(ABC):
    """
    Содержит общие атрибуты и методы для дочерних классов EditModel и CalcModel.
    """

    def __init__(self):
        # коллекция для хранения информации о расчете
        self.set_info = {'collapse': [],
                         'end_info': ''}
        self.file_count = 0  # Счетчик расчетных файлов.
        self.now = datetime.now()
        self.time_start = time.time()
        self.now_start = self.now.strftime("%d-%m-%Y %H:%M:%S")

    def the_end(self):  # по завершению
        execution_time = time.strftime("%H:%M:%S", time.gmtime(time.time() - self.time_start))
        self.set_info['end_info'] = (
            f"РАСЧЕТ ЗАКОНЧЕН!"
            f"\nНачало расчета {self.now_start}."
            f"\nКонец {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}."
            f"\nЗатрачено: {execution_time} (файлов: {self.file_count}).")
        log_g_s.info(self.set_info['end_info'])

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
