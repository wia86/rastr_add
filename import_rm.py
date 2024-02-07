__all__ = ['ImportFromModel']
"""Модуль для изменения параметров РМ по заданию в таблице excel."""
import logging
import os
from collections import namedtuple
import pandas as pd

log_im_rm = logging.getLogger(f'__main__.{__name__}')


class ImportFromModel:
    """Клас для импорта из одних моделей в другие"""
    TYPE_IMPORT_RM = {"обновить": 2,
                      "загрузить": 1,
                      "присоединить": 0,
                      "присоединить-обновить": 3,
                      "объединить": 3}

    def __init__(self,
                 export_rm,
                 criterion_start: dict | None = None,
                 tables: str = '',
                 param: str = '',
                 sel: str | None = '',
                 calc: int | str = 2):
        """
        Импорт данных из файлов РМ ('.rg2', '.rst' и др.) в РМ.
        :param export_rm: RastrModel
        :param criterion_start: {"years": "","season": "","max_min": "", "add_name": ""} условие выполнения
        :param tables: таблица для импорта, нр "node, vetv"
        :param param: параметры для импорта: "" все параметры или перечисление, нр 'sel, sta'(ключи необязательно)
        :param sel: выборка нр "sel" или "" - все
        :param calc: число (в формате int | строка) или ключевое слово:
        {"обновить": 2, "загрузить": 1, "присоединить": 0, "присоединить-обновить": 3}
        """
        self.__export_file_name = export_rm.full_name
        log_im_rm.info(f'Экспорт данных из файла "{self.__export_file_name}".')
        if not os.path.exists(self.__export_file_name):
            raise ValueError(f"Ошибка в задании, не найден файл: {self.__export_file_name}")

        self.__criterion_start = criterion_start
        self.__sel = sel if sel else ''

        match calc:
            case int():
                self.__calc = calc
            case str() if calc.isdigit():
                self.__calc = int(calc)
            case str() if calc in self.TYPE_IMPORT_RM:
                self.__calc = self.TYPE_IMPORT_RM[calc]
            case _:
                raise ValueError(f"{self.__class__.__name__}: Ошибка в задании, не распознано задание '{calc=}'.")

        data_import_i = namedtuple('импорт', ['table', 'parameters', 'data'])
        self.__data_import = []
        self.__type_file = export_rm.type_file
        export_rm.load()

        for table in tables.replace(' ', '').split(","):  # разделить на ["таблицы"]
            tab_rm = export_rm.rastr.Tables(table)
            tab_rm.setsel(self.__sel)
            if not tab_rm.count:
                continue
            # Параметры
            if param:  # Добавить к строке параметров ключи текущей таблицы
                param_all = param + ',' + export_rm.rastr.Tables(table).Key
            else:  # если все параметры
                param_all = export_rm.all_cols(table)
            # Данные
            self.__data_import.append(data_import_i(table,
                                                    param_all,
                                                    tab_rm.writesafearray(param_all, "000")))
            log_im_rm.info(f"\tТаблица: {table}, выборка: {self.__sel}, параметры: {param_all!r}.")

    def import_data_in_rm(self, rm) -> None:
        """
        Импорт данных в файлы
        """
        log_im_rm.info(f"\tИмпорт данных в текущую РМ из файла {self.__export_file_name} в РМ.")
        if not rm.code_name_rg2 or rm.test_name(condition=self.__criterion_start,
                                                info=f'\t{self.__class__.__name__} '):
            for i in self.__data_import:
                rm_tab = rm.rastr.Tables(i.table)
                if self.__type_file == rm.type_file:
                    rm_tab.ReadSafeArray(self.__calc, i.parameters, i.data)
                    log_im_rm.info(
                        f"\tТаблица: {i.table}, выборка: {self.__sel}, тип: {self.__calc}, параметры: {i.parameters}.")
                else:
                    set_param_in = set(rm.all_cols(i.table, val_return='list'))
                    set_param_out = set(i.parameters.split(','))
                    common = set_param_out & set_param_in  # пересечение
                    if not common:
                        log_im_rm.warning(f"Таблица: {i.table}, параметры: ОТСУТСТВУЮТ.")
                        return
                    data = pd.DataFrame(data=i.data, columns=i.parameters.split(','))
                    data.drop(columns=list(set_param_out - common), inplace=True)
                    param_all = ','.join(list(data.columns))
                    import_data = tuple(data.itertuples(index=False, name=None))
                    rm_tab.ReadSafeArray(self.__calc, param_all, import_data)
                    log_im_rm.info(f"\tТаблица: {i.table}, выборка: {self.__sel}, тип: {self.__calc}, "
                                   f"параметры: {param_all}, без параметров: {set_param_out - common}.")
