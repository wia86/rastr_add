__all__ = ['ManagerPrintParameters']

from abc import ABC

import pandas as pd


class ManagerPrintParameters(ABC):
    _dict_obj = {}  # строка задания: объект PrintParameters

    @classmethod
    def add(cls, rm, task: str):
        """
        Добавить значения из rm в объект класса PrintParameters.
        Если он отсутствует, то создать его и сохранить в данном классе.
        :param rm:
        :param task: '15177,16095:r;x;b / 15106;15138:pn;qn / ns=2(sechen):psech'
        """
        if not task:
            raise ValueError('Ошибка: не указаны входные параметры.')
        if not isinstance(task, str):
            raise TypeError('Ошибка: неверный тип данных.')

        obj = cls._dict_obj.setdefault(task, PrintParameters(task))
        obj.add_val_parameters(rm)

    @classmethod
    def all_data_in_excel(cls, path_excel: str):
        if not cls._dict_obj:
            return
        if not path_excel:
            raise ValueError('Не указан путь для сохранения.')
        for key in cls._dict_obj:
            cls._dict_obj[key].data_in_excel(path_excel)
        cls._dict_obj = {}


class PrintParameters:
    """ Вывод параметров режима в таблицу excel."""
    _set_output_parameters = set()  # set((key, param), (ny, pn), (10, pn))
    _data_parameters = pd.DataFrame()
    _num_class = 0

    @classmethod
    def _receive_name_sheet(cls):
        cls._num_class += 1
        return f'par{cls._num_class}'

    def __init__(self, task: str):
        """
        :param task: '15177,16095:r;x;b / 15106;15138:pn;qn / ns=2(sechen):psech'
        """
        self._task = task

        task = task.replace(' ', '').split('/')
        for task_i in task:
            keys, parameters = task_i.split(':')  # нр'8;9', 'pn;qn'
            for parameter in parameters.split(';'):  # ['pn','qn']
                for key in keys.split(';'):  # ['15105,15113','15038,15037,4']
                    self._set_output_parameters.add((key, parameter,))

    def add_val_parameters(self, rm):
        """
        Добавить на sheet excel.
        """
        if 'sechen' in self._task:
            if rm.rastr.tables.Find('sechen') < 0:
                rm.downloading_additional_files('sch')

        date = pd.Series(dtype='object')
        for k, p in self._set_output_parameters:
            table, sel = rm.recognize_key(key=k, back='tab sel')
            key = f'{k}_{p}'
            if rm.rastr.tables(table).cols(p).Prop(1) == 2:  # если поле типа строка
                date.loc[key] = rm.txt_field_return(table, sel, p)
            else:
                date.loc[key] = rm.rastr.tables(table).cols.Item(p).Z(rm.index(table_name=table,
                                                                               key_str=sel))

        date = pd.concat([date, pd.Series(rm.info_file)])
        self._data_parameters = pd.concat([self._data_parameters, date], axis=1)

    def data_in_excel(self, path_excel: str):
        """Записать данные в excel"""

        with pd.ExcelWriter(path=path_excel,
                            mode='w',
                            engine='openpyxl') as writer:
            self._data_parameters.T.to_excel(excel_writer=writer,
                                             sheet_name=self._receive_name_sheet(),
                                             header=True,
                                             index=False)
