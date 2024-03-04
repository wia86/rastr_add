import logging

import pandas as pd
from tabulate import tabulate

import collection_func as cf

log_comb_xl = logging.getLogger(f'__main__.{__name__}')


class CombinationXL:
    """
    Класс для создания генератора комбинаций заданных в книге excel
    """

    def __init__(self, path_xl: str, sheet: str):
        # Перечень отключений из excel
        self.srs_xl = pd.read_excel(path_xl,
                                    sheet_name=sheet,
                                    keep_default_na=False)  # keep_default_na: True NaN or  False ''
        self.srs_xl = self.srs_xl[~self.srs_xl['Статус'].str.contains('#')]  # ~ not
        self.srs_xl.dropna(how='all', axis=0, inplace=True)  # Удалить пустые строки.
        if self.srs_xl.empty:
            raise ValueError(f'Таблица отключений из xl пуста.')

    def gen_comb_xl(self, rm) -> pd.DataFrame:
        """
        Генератор комбинаций из XL
        :param rm:
        :return:  комбинацию comb_xl
        """
        for _, row in self.srs_xl.iterrows():
            if row['Условие']:
                if not rm.conditions_test(row['Условие']):
                    continue
            comb_xl = pd.DataFrame(columns=['table',
                                            'index',
                                            'status_repair',
                                            'key',
                                            's_key',
                                            'repair_scheme',
                                            'disable_scheme',
                                            'double_repair_scheme',
                                            'double_repair_scheme_copy'])
            double_repair = True if row['Ключ рем.1'] and row['Ключ рем.2'] else False
            for key_type, scheme_xl_name in (('Ключ откл.', 'Схема при отключении'),
                                             ('Ключ рем.1', 'Ремонтная схема1'),
                                             ('Ключ рем.2', 'Ремонтная схема2')):
                key = row[key_type]
                if key:
                    key = str(key)
                    status_repair = False if key_type == 'Ключ откл.' else True
                    table, s_key = rm.recognize_key(key=key, back='tab s_key')
                    index = rm.index(table_name=table, key_int=s_key)
                    if table and index >= 0:
                        repair_scheme = False
                        disable_scheme = False
                        double_repair_scheme = False
                        double_repair_scheme_copy = False
                        # Если в колонке «Схема при отключении» или «Ремонтная схема» содержится «*», то значение поля
                        # дополняется из соответствующих полей disable_scheme, repair_scheme, double_repair_scheme РМ.
                        scheme_xl = row[scheme_xl_name]
                        add_scheme = []
                        if scheme_xl:
                            scheme_xl = scheme_xl.split('#')[0].replace(' ', '')
                            if '*' in scheme_xl:
                                scheme_xl = scheme_xl.replace('*', '')
                                if status_repair:
                                    add_scheme = rm.dt.t_scheme[table]['repair_scheme'].get(s_key, False)
                                    if double_repair:
                                        double_repair_scheme_copy = \
                                            rm.dt.t_scheme[table]['double_repair_scheme'].get(s_key, False)
                                else:
                                    add_scheme = rm.dt.t_scheme[table]['disable_scheme'].get(s_key, False)
                            scheme_xl = cf.split_task_action(scheme_xl)
                            if add_scheme:
                                if scheme_xl:
                                    scheme_xl.append(add_scheme)
                                else:
                                    scheme_xl = add_scheme
                        if scheme_xl:
                            if status_repair:
                                repair_scheme = scheme_xl
                            else:
                                disable_scheme = scheme_xl

                        comb_xl.loc[len(comb_xl.index)] = [table,
                                                           index,
                                                           status_repair,
                                                           key,
                                                           s_key,
                                                           repair_scheme,
                                                           disable_scheme,
                                                           double_repair_scheme,
                                                           double_repair_scheme_copy]
                    else:
                        log_comb_xl.info(f'Задание комбинаций их XL: в РМ не найден ключ {key!r}')
                        log_comb_xl.info(tabulate(row, headers='keys', tablefmt='psql'))
                        continue
            if not len(comb_xl):
                continue
            yield comb_xl
