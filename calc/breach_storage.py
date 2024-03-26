__all__ = ['BreachStorage']

import logging
import os
from collections import defaultdict

import pandas as pd

from collection_func import from_list1_only_exists_in_list2, save_to_sqlite
from pivot_table import make_pivot_tables

log_breach_storage = logging.getLogger(f'__main__.{__name__}')


class BreachStorage:
    """Хранение нарушений режимов"""

    collection = None

    def __init__(self):
        self.collection_rm = defaultdict(pd.DataFrame)
        self.collection_all = defaultdict(pd.DataFrame)

    def save_data_active(self,
                         breach,
                         comb_id: int,
                         active_id: int):
        """
        Сохранить перегрузки для текущей комбинации и действия
        :param breach: Объект Breach
        :param comb_id:
        :param active_id:
        """
        if breach.violations:
            for key, val in breach.violations.items():

                if val is not None:
                    val['comb_id'] = comb_id
                    val['active_id'] = active_id

                    self.collection_rm[key] = pd.concat([self.collection_rm[key], val],
                                                        axis=0,
                                                        ignore_index=True)

    def save_data_rm(self):
        """Добавить Контролируемые элементы.
        Перенести данные в общую коллекцию"""
        for key in self.collection_rm:
            self.collection_all[key] = pd.concat([self.collection_all[key], self.collection_rm[key]],
                                                 axis=0,
                                                 ignore_index=True)

        self.collection_rm = defaultdict(pd.DataFrame)

    def save_to_sql(self, path_db: str):

        if not self.collection_all:
            log_breach_storage.debug('Данные для сохранения в db отсутствуют.')
            return

        save_to_sqlite(path_db=path_db,
                       dict_df=self.collection_all)

        for key in self.collection_all:
            log_breach_storage.debug(f'Данные {key} добавлены в db {path_db}.')

    def save_to_xl(self,
                   path_xl_book: str,
                   all_rm: pd.DataFrame,
                   all_comb: pd.DataFrame,
                   all_actions: pd.DataFrame):
        """
        Сохранить в книгу df с нарушениями режима и сформировать сводные.
        :param path_xl_book:
        :param all_rm:
        :param all_comb:
        :param all_actions:
        """
        if not self.collection_all:
            log_breach_storage.debug('Данные для сохранения в excel отсутствуют.')
            return

        dict_columns = {}

        for key in self.collection_all:
            log_breach_storage.debug(f'Данные {key} добавлены в excel {path_xl_book}.')

            self.collection_all[key] = (all_rm.merge(all_comb, how='right', on='rm_id')
                                        .merge(all_actions, how='right', on='comb_id')
                                        .merge(self.collection_all[key], how='right', on=['comb_id', 'active_id']))

            self.collection_all[key].dropna(axis=1, how='all', inplace=True)
            self.collection_all[key].fillna('-', inplace=True)

            mode = 'a' if os.path.exists(path_xl_book) else 'w'
            with pd.ExcelWriter(path=path_xl_book, mode=mode) as writer:
                # Максимальное количество строк листа 1 048 576
                l_max = 1_048_576

                if len(self.collection_all[key]) > l_max:
                    log_breach_storage.error('Число нарушений режима превышает допустимый '
                                             'размер листа xl (1 048 576 строк).')
                    self.collection_all[key] = self.collection_all[key].iloc[:l_max - 3]

                self.collection_all[key].to_excel(excel_writer=writer,
                                                  float_format='%.2f',
                                                  index=False,
                                                  freeze_panes=(1, 1),
                                                  sheet_name=key)
            dict_columns[key] = tuple(self.collection_all[key].columns)

        if dict_columns:
            task_pivot = self.add_pivot_tables(dict_columns)
            make_pivot_tables(book_path=path_xl_book,
                              sheets_info=task_pivot)

    @staticmethod
    def add_pivot_tables(dict_columns: dict) -> dict:
        """
        Формирование задания для сводной.
        :param dict_columns: Словарь с перечнем имен столбцов.

        {'dead': [],
         'i': [],
         'low_u': [],
         'high_u': []}
        :return:  см.sheets_info в make_pivot_tables
        """
        sheets_info = {}

        data_field = {'dead': dict(dead_mode='Режим не моделируется'),
                      'i': dict(i_max='Iрасч.,A',
                                i_dop_r='Iддтн,A',
                                i_zag='Iзагр. ддтн,%',
                                i_dop_r_av='Iадтн,A',
                                i_zag_av='Iзагр. адтн,%'),
                      'low_u': dict(vras='Uр, кВ',
                                    umin='МДН, кВ',
                                    otv_min='Uр, % от МДН',
                                    umin_av='АДН,кВ',
                                    otv_min_av='Uр, % от АДН'),
                      'high_u': dict(vras='Uр, кВ',
                                     umax='Uнр, кВ',
                                     otv_max='Uр, % от Uнр')}

        conditional_formatting = {'i': dict(num_field=[3, 5],
                                            val=100),
                                  'low_u': dict(num_field=[3],
                                                val=0),
                                  'high_u': dict(num_field=[3],
                                                 val=0)}

        for key in dict_columns:
            sheets_info[key] = {}
            columns = dict_columns[key]

            sheet_info = sheets_info[key]
            sheet_info['sheet_name'] = f'Сводная_{key}'
            sheet_info['pt_name'] = f'pt_{key}'

            sheet_info['conditional_formatting'] = True if key != 'dead' else False  # todo ?
            sheet_info['row_fields'] = ['Контролируемые элементы',
                                        'Отключение',
                                        'Ремонт 1',
                                        'Ремонт 2']
            sheet_info['row_fields'] = from_list1_only_exists_in_list2(sheet_info['row_fields'],
                                                                       columns)

            sheet_info['column_fields'] = ['Год', 'Сезон макс/мин'] + [col for col in columns if 'Доп. имя' in col]
            sheet_info['column_fields'] = from_list1_only_exists_in_list2(sheet_info['column_fields'],
                                                                          columns)

            sheet_info['page_fields'] = ['Имя файла',
                                         'Кол. откл. эл.',
                                         'End',
                                         'Наименование СРС',
                                         'Контроль ДТН',
                                         'Темп.(°C)']

            sheet_info['page_fields'] = from_list1_only_exists_in_list2(sheet_info['page_fields'],
                                                                        columns)
            sheet_info['data_field'] = data_field[key]
            sheet_info['conditional_formatting'] = conditional_formatting.get(key)
        return sheets_info
