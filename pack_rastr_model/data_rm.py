__all__ = ['DataRM']

import logging
from collections import defaultdict

import pandas as pd

log_rm_db = logging.getLogger(f'__main__.{__name__}')


class DataRM:
    """Хранение данных из таблиц РМ в различных форматах и восстановление данных в таблицах."""

    def __init__(self, rm):
        self.rm = rm

        self.data_save = {}
        self.data_columns = None
        self.data_save_sta = {}
        self.data_columns_sta = None
        self.t_sta = {}  # {имя таблицы: {(ip, iq, np): 0 или 1}}
        self.t_name = {}  # {имя таблицы: {ny: имя}}
        self.t_key_i = {}  # {имя таблицы: {(ip, iq, np): индекс}}
        self.t_i_key = {}  # {имя таблицы: {индекс: (ip, iq, np)}}
        self.t_scheme = {}  # {имя таблицы: {тип схемы:{(ip, iq, np): (картеж номеров)}}}
        for tab_name in ['node', 'vetv', 'Generator']:
            self.t_sta[tab_name] = {}
            self.t_key_i[tab_name] = {}
            self.t_i_key[tab_name] = {}
            self.t_scheme[tab_name] = {'repair_scheme': {},
                                       'disable_scheme': {},
                                       'double_repair_scheme': {},
                                       'automation': {}}
            self.t_name[tab_name] = {}
            self.t_name[tab_name][-1] = 'Режим не моделируется'
        self.ny_join_vetv = defaultdict(list)  # {ny: все присоединенные ветви}
        self.ny_unom = {}  # {ny: номинальное напряжение}

        # self.ny_pqng = defaultdict(tuple)  # {ny: (pn, qn, pg, qn)} - все с pn pg > 0 | qn pg > 0 | pg > 0 | qg > 0
        self.v_gr = {}  # {(ip, iq, np): groupid} - все c groupid > 0
        self.v_rxb = {}  # {(ip, iq, np): (r, x, b)} - все

    def save_date_tables(self):
        """
        Сохранить значения в таблицах в исходной схеме сети.
        Сохранить имена ветвей узлов и генераторов в dict и df.
        """
        log_rm_db.debug('Сохранение значений исходных параметров сети.')

        # Запись sta
        self.data_columns_sta = {'vetv': 'ip,iq,np,sta,sel',
                                 'node': 'ny,sta,sel',
                                 'Generator': 'Num,sta'}
        for name_tab in self.data_columns_sta:
            self.data_save_sta[name_tab] = self.rm.rastr.tables(name_tab).writesafearray(
                self.data_columns_sta[name_tab],
                '000')

        # Запись прочих данных которые могут измениться во время расчетов
        self.data_columns = {'vetv': 'ip,iq,np,sta,ktr',
                             'node': 'ny,sta,pn,qn,pg,qg,vzd,bsh',
                             'Generator': 'Num,sta,P'}
        for name_tab in self.data_columns:
            self.data_save[name_tab] = self.rm.rastr.tables(name_tab).writesafearray(self.data_columns[name_tab],
                                                                                     '000')
        # Узлы
        for ny, sta, pn, qn, pg, qg, vzd, bsh in self.data_save['node']:
            self.t_sta['node'][ny] = sta

        t = self.rm.rastr.tables('node').writesafearray('ny,name,dname,uhom', '000')
        for index, (ny, name, dname, uhom) in enumerate(t):
            self.t_key_i['node'][ny] = index
            self.t_i_key['node'][index] = ny
            self.ny_unom[ny] = uhom
            if dname.strip():
                self.t_name['node'][ny] = dname
            else:
                self.t_name['node'][ny] = name if name else f'Узел {ny}'

        # Ветви
        for ip, iq, np_, sta, ktr in self.data_save['vetv']:  # , r, x, b
            s_key = (ip, iq, np_)
            self.t_sta['vetv'][s_key] = sta
            self.ny_join_vetv[ip].append(s_key)
            self.ny_join_vetv[iq].append(s_key)

        t = self.rm.rastr.tables('vetv').writesafearray('ip,iq,np,dname,groupid,r,x,b', '000')
        for index, (ip, iq, np_, dname, groupid, r, x, b) in enumerate(t):
            s_key = (ip, iq, np_)
            self.t_key_i['vetv'][s_key] = index
            self.t_i_key['vetv'][index] = s_key
            if dname.strip():
                self.t_name['vetv'][s_key] = dname
            else:
                self.t_name['vetv'][s_key] = f'{self.t_name["node"][ip]} - {self.t_name["node"][iq]}'

            if groupid:
                self.v_gr[s_key] = groupid
            self.v_rxb[s_key] = (r, x, b)

        # Генераторы
        for Num, sta, P in self.data_save['Generator']:
            self.t_sta['Generator'][Num] = sta

        t = self.rm.rastr.tables('Generator').writesafearray('Num,Name,Node', '000')
        for index, (Num, Name, Node) in enumerate(t):
            self.t_key_i['Generator'][Num] = index
            self.t_i_key['Generator'][index] = Num
            if Name:
                self.t_name['Generator'][Num] = Name
            else:
                self.t_name['Generator'][Num] = f'генератор номер {Num} в узле {self.t_name["node"][Node]}'

    def recover_date_tables(self, restore_only_state: bool) -> bool:
        """
        Восстановить значения в таблицах rastr.
        :param restore_only_state: Истина - только поля sta
        :return:
        """
        if restore_only_state:
            for name_table in self.data_save_sta:
                self.rm.rastr.tables(name_table).ReadSafeArray(2,
                                                               self.data_columns_sta[name_table],
                                                               self.data_save_sta[name_table])
            log_rm_db.debug('Состояние элементов сети восстановлено.')
        else:
            for name_table in self.data_save:
                self.rm.rastr.tables(name_table).ReadSafeArray(2,
                                                               self.data_columns[name_table],
                                                               self.data_save[name_table])
            log_rm_db.debug('Состояние элементов сети и параметров восстановлено.')
        return True
