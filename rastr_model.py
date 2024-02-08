import logging
import pandas as pd
import re
import os
from collections import namedtuple, defaultdict, Counter
from tkinter import messagebox as mb
import win32com.client

from import_rm import ImportFromModel
import collection_func as cf
log_rm = logging.getLogger(f'__main__.{__name__}')


class RastrModel:
    """
    Для хранения параметров файла rg2, rst.
    """
    # Работа с напряжением узлов
    U_NOM = [6, 10, 35, 110, 150, 220, 330, 500, 750]  # номинальные напряжения
    U_MIN_NORM = [5.8, 9.7, 32, 100, 135, 205, 315, 490, 730]  # минимальное нормальное напряжение
    U_LARGEST_WORKING = [7.2, 12, 42, 126, 180, 252, 363, 525, 787]  # наибольшее рабочее напряжения

    # U_LARGEST_WORKING_dict = {6: 7.2, 10: 12, 35: 42, 110: 126,150: 180,220: 252, 330: 363,500: 525,750: 787}
    KEY_TABLES = {'ny': 'node',
                  'ip': 'vetv',
                  'Num': 'Generator',
                  # 'g': 'Generator',
                  'na': 'area',
                  'npa': 'area2',
                  'no': 'darea',
                  'nga': 'ngroup',
                  'ns': 'sechen'}
    rm_id = 0
    all_rm = pd.DataFrame()
    config = None  # todo переместить в __init__
    overwrite_new_file = 'question'

    def __str__(self):
        return self.name_base

    def __repr__(self):
        return f'{self.__class__} {self.name_base}'

    def __init__(self, full_name: str):
        # Информация о файле.
        RastrModel.rm_id += 1
        self.info_file = dict()
        self.info_file['rm_id'] = RastrModel.rm_id
        self.full_name = full_name
        self.dir = os.path.dirname(full_name)
        self.Name = os.path.basename(full_name)  # Вернуть имя с расширением "2020 зим макс.rg2"
        self.name_base, self.type_file = self.Name.rsplit(sep='.', maxsplit=1)
        self.info_file['Имя файла'] = self.name_base

        self.pattern = self.config[f"шаблон {self.type_file}"]
        self.all_auto_shunt = {}
        self.info_file["Темп.(°C)"]: float = 0
        self.rastr = None
        self.additional_name_list = None
        self.add_load = []  # Расширение дополнительных файлов из [trn, anc]

        # Для хранения исходной схемы и параметров сети
        self.v__num_transit = {}  # {(ip, iq, np): номер транзита в тч из 1 ветви}
        self.data_save = None
        self.data_columns = None
        self.data_save_sta = None
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

        # self.ny_pqng = defaultdict(tuple)  # {ny: (pn, qn, pg, qn)} - все с pn pg > 0 | qn pg > 0 | pg > 0 | qg > 0
        self.v_gr = {}  # {(ip, iq, np): groupid} - все c groupid > 0
        self.v_rxb = {}  # {(ip, iq, np): (r, x, b)} - все
        self.vetv_name = None  # df[s_key, 'Контролируемые элементы']
        self.node_name = None  # df[s_key, 'Контролируемые элементы']

        # Произвольный формат названия файла
        self.info_file['Сезон']: str = ''  # 'зимний', 'летний', 'паводок' сезон года
        self.info_file['макс/мин']: str = ''  # 'максимум', 'максимум' потребления в сутках
        self.info_file['Сезон макс/мин']: str = ''  # "Зимний максимум нагрузки"
        self.info_file["Имя режима"]: str = ''  # "Зимний максимум нагрузки 2023 г. Расчетная температура -40 °C."
        name_base_lower = self.name_base.lower()

        # Поиск в имени: сезона
        for test in ['zim', 'зим']:  # , '.12.'
            if test in self.name_base.lower():
                self.info_file['Сезон'] = 'зимний'
        if not self.info_file['Сезон']:
            for test in ['let', 'лет']:  # , '.06.'
                if test in name_base_lower:
                    self.info_file['Сезон'] = 'летний'
        if not self.info_file['Сезон']:
            for test in ['пав', 'flood', 'pav']:
                if test in name_base_lower:
                    self.info_file['Сезон'] = 'паводок'
        self.info_file['Сезон макс/мин'] = f'{self.info_file["Сезон"].capitalize()}.'  # Сделать первую букву заглавной
        self.info_file['Сезон макс/мин'] = self.info_file['Сезон макс/мин'].replace('Паводок', 'Паводок,')
        # Поиск в имени: макс мин
        for test in ['max', 'макс']:
            if test in name_base_lower:
                self.info_file['макс/мин'] = 'максимум'
        if not self.info_file['макс/мин']:
            for test in ['min', 'мин']:
                if test in name_base_lower:
                    self.info_file['макс/мин'] = 'минимум'
        self.info_file['Сезон макс/мин'] = (
            self.info_file['Сезон макс/мин'].replace('.', f' {self.info_file["макс/мин"]} нагрузки.'))

        if self.info_file['Сезон макс/мин']:
            self.info_file["Имя режима"] = self.info_file['Сезон макс/мин']

        # Поиск в имени: год
        match = re.search(re.compile(r"(20[2-9][0-9])"), name_base_lower)
        if match:
            self.info_file['Год'] = match[1]
            self.info_file["Имя режима"] = self.info_file["Имя режима"].replace('.', f' {match[1]} г.')
        else:
            self.info_file['Год'] = ''

        # Температура абв по ГОСТ.
        # 0 не распознан, 1 зим макс, 2 зим мин, 3 зим макс tср(табл), 4 зим мин tср, 5 ПЭВТ (tэкст).
        # Температура гд по ГОСТ.
        #  6 лет макс, 7 лет мин, 8 паводок макс, 9 паводок мин.
        self.code_name_rg2 = 0

        if ("tэкст" in name_base_lower) or ("пэвт" in name_base_lower):
            self.code_name_rg2 = 5
            self.info_file["Имя режима"] = self.info_file["Имя режима"].replace('нагрузки', 'нагрузки (ПЭВТ)')
        else:
            if self.info_file['Сезон макс/мин'] == 'Зимний максимум нагрузки.':
                self.code_name_rg2 = 1
                if "tср" in name_base_lower:
                    self.code_name_rg2 = 3
            if self.info_file['Сезон макс/мин'] == 'Зимний минимум нагрузки.':
                self.code_name_rg2 = 2
                if "tср" in name_base_lower:
                    self.code_name_rg2 = 4
            if self.info_file['Сезон макс/мин'] == 'Летний максимум нагрузки.':
                self.code_name_rg2 = 6
            if self.info_file['Сезон макс/мин'] == 'Летний минимум нагрузки.':
                self.code_name_rg2 = 7
            if self.info_file['Сезон макс/мин'] == 'Паводок, максимум нагрузки.':
                self.code_name_rg2 = 8
            if self.info_file['Сезон макс/мин'] == 'Паводок, минимум нагрузки.':
                self.code_name_rg2 = 9

        self.gost_abv = True if 0 < self.code_name_rg2 < 6 else False
        self.gost_gd = True if self.code_name_rg2 > 5 else False
        # Поиск в строке значения в ()
        match = re.search(re.compile(r"\((.+)\)"), self.name_base)
        if match:
            self.info_file["Доп. имена"] = match[1]
            self.additional_name_list = match[1].split(";")
        else:
            self.info_file["Доп. имена"] = ''

        if "°C" in self.name_base:
            match = re.search(re.compile(r"((-|минус)?\s?\d+([,.]\d*)?)\s?°C"), self.name_base)  # -45.14 °C
            if match:
                self.info_file['Темп.(°C)'] = float(match[1]
                                                    .replace(',', '.')
                                                    .replace('минус', '-')
                                                    .replace(' ', ''))  # число
                self.info_file["Имя режима"] += f' Расчетная температура {self.info_file["Темп.(°C)"]} °C.'

        if self.additional_name_list:
            for i, additional_name_i in enumerate(self.additional_name_list, 1):
                self.info_file['Доп. имя' + str(i)] = additional_name_i
        RastrModel.all_rm = pd.concat([RastrModel.all_rm, pd.Series(self.info_file).to_frame().T],
                                      axis=0, ignore_index=True)
        log_rm.info(self.info_file["Имя режима"])

    def save_value_fields(self):
        """
        Сохранить значения полей в исходной схеме сети (изменяемых в процессе расчетов).
        Сохранить имена ветвей узлов и генераторов в dict и df.
        """
        log_rm.debug('Сохранение значений исходных параметров сети.')

        self.data_save_sta = {'vetv': None, 'node': None, 'Generator': None}
        self.data_columns_sta = {'vetv': 'ip,iq,np,sta,sel',
                                 'node': 'ny,sta,sel',
                                 'Generator': 'Num,sta'}
        for name_tab in self.data_save_sta:
            self.data_save_sta[name_tab] = \
                self.rastr.tables(name_tab).writesafearray(self.data_columns_sta[name_tab], "000")

        self.data_save = {'vetv': None, 'node': None, 'Generator': None}
        self.data_columns = {'vetv': 'ip,iq,np,sta,ktr',  # ,r,x,b
                             'node': 'ny,sta,pn,qn,pg,qg,vzd,bsh',
                             'Generator': 'Num,sta,P'}
        for name_tab in self.data_save:
            self.data_save[name_tab] = \
                self.rastr.tables(name_tab).writesafearray(self.data_columns[name_tab], "000")

        # Узлы
        for ny, sta, pn, qn, pg, qg, vzd, bsh in self.data_save['node']:
            self.t_sta['node'][ny] = sta
            # if pn or qn or pg:
            #     self.ny_pqng[ny] = (pn, qn, pg, qg)
        t = self.rastr.tables('node').writesafearray("ny,name,dname,index", "000")
        for ny, name, dname, index in t:
            self.t_key_i['node'][ny] = index
            self.t_i_key['node'][index] = ny
            if dname:
                self.t_name['node'][ny] = dname
            else:
                self.t_name['node'][ny] = name if name else f'Узел {ny}'

        # Ветви
        for ip, iq, np_, sta, ktr in self.data_save['vetv']:  # , r, x, b
            s_key = (ip, iq, np_)
            self.t_sta['vetv'][s_key] = sta
            self.ny_join_vetv[ip].append(s_key)
            self.ny_join_vetv[iq].append(s_key)

        t = self.rastr.tables('vetv').writesafearray("ip,iq,np,dname,groupid,r,x,b,index", "000")
        for ip, iq, np_, dname, groupid, r, x, b, index in t:
            s_key = (ip, iq, np_)
            self.t_key_i['vetv'][s_key] = index
            self.t_i_key['vetv'][index] = s_key
            # log_rm.info(s_key)
            if dname:
                self.t_name['vetv'][s_key] = dname
            else:
                # log_rm.info(self.t_name["node"][ip])
                # log_rm.info(self.t_name["node"][iq])
                self.t_name['vetv'][s_key] = f'{self.t_name["node"][ip]} - {self.t_name["node"][iq]}'

            if groupid:
                self.v_gr[s_key] = groupid
            self.v_rxb[s_key] = (r, x, b)

        # Генераторы
        for Num, sta, P in self.data_save['Generator']:
            self.t_sta['Generator'][Num] = sta

        t = self.rastr.tables('Generator').writesafearray("Num,index,Name,Node", "000")
        for Num, index, Name, Node in t:
            self.t_key_i['Generator'][Num] = index
            self.t_i_key['Generator'][index] = Num
            if Name:
                self.t_name['Generator'][Num] = Name
            else:
                self.t_name['Generator'][Num] = f'генератор номер {Num} в узле {self.t_name["node"][Node]}'

        # Создать df ['s_key', 'Контролируемые элементы']] для таблиц узлов и ветвей
        self.node_name = pd.DataFrame.from_dict(self.t_name['node'],
                                                orient='index',
                                                columns=['Контролируемые элементы'])
        self.node_name['s_key'] = self.node_name.index
        self.node_name.reset_index(drop=True, inplace=True)

        self.vetv_name = pd.DataFrame.from_dict(self.t_name['vetv'],
                                                orient='index',
                                                columns=['Контролируемые элементы'])
        self.vetv_name['s_key'] = self.vetv_name.index
        self.vetv_name.s_key = self.vetv_name.s_key.apply(lambda xx: str(xx).replace(' ', '')
                                                          .replace('(', '')
                                                          .replace(')', '')
                                                          .replace(',0', ''))
        self.vetv_name.reset_index(drop=True, inplace=True)

    def network_analysis(self, disable_on: bool = True,
                         field: str = 'disable',
                         selection_node_for_disable: str = ''):
        """
        Анализ графа сети.
        :param disable_on: Отметить отключаемые элементы в поле field узлов и ветвей.
        :param field: Поле для отметки отключаемых элементов.
        :param selection_node_for_disable: Выборка в таблице узлы, для выбора отключаемых элементов.
        Например, района, территории, нагрузочной группы для расчета или "" - все узлы.
        """
        log_rm.info('Анализ графа сети с заполнением поля transit в таблице узлов и ветвей.')
        all_ny = [x[0] for x in self.rastr.tables('node').writesafearray("ny", "000")]
        log_rm.debug(f'В РМ {len(all_ny)} узлов.')

        vetv = self.rastr.tables('vetv')
        vetv.setsel('sta=0')
        data_v = vetv.writesafearray('ip,iq,np,groupid,tip,pl_ip,groupid', "000")
        # Все включенные ветви РМ
        all_vetv_sta0 = [(ip, iq, np_) for ip, iq, np_, groupid, tip, pl_ip, groupid in data_v]
        all_ny_in_v = []  # [Все ip iq в таблице ветви, если ny встречается 1 раз, то это тупик.]
        ny_end = set()  # Все тупиковые узлы

        # Поиск узлов с одной отходящей включенной ветвью - это тупик.
        ny_all_vetv = defaultdict(list)  # {ny: все примыкающие включенные ветви}

        for ip, iq, np_ in all_vetv_sta0:
            all_ny_in_v.append(ip)
            all_ny_in_v.append(iq)
            ny_all_vetv[ip].append((ip, iq, np_))
            ny_all_vetv[iq].append((ip, iq, np_))

        for k, v in Counter(all_ny_in_v).items():  # {Номер узла: количество отходящих ветвей}
            if v == 1:
                ny_end.add(k)
        log_rm.debug(f'В РМ {len(ny_end)} тупиков.')

        # Найти остальные узлы тупиковых цепочек.
        all_v_end = set()  # Все тупиковые ветви
        ny_end2 = set()  # Вспомогательный набор
        for ny in ny_end:
            ny_next = ny
            # Поиск в цикле следующего узла цепочки, если его нет, то ny_next равен 0.
            while ny_next > 0:  # ny_next следующий проверяемый узел
                ny_source = ny_next
                v_not_end = []  # [записываем не тупиковые ветви узла]
                for i in ny_all_vetv[ny_source]:
                    if i not in all_v_end:
                        v_not_end.append(i)
                if len(v_not_end) == 1:
                    all_v_end.add(v_not_end[0])
                    ip, iq, np_ = v_not_end[0]
                    ny_next = iq if ip == ny_source else ip
                    ny_end2.add(ny_source)
                else:
                    ny_next = 0
        ny_end = ny_end | ny_end2
        log_rm.debug(f'В РМ {len(ny_end)} тупиковых узлов.')

        # Определить транзитные и узловые узлы
        all_ny_transit = []  # [Все узлы РМ входящие в транзиты]
        all_ny_nodal = {}  # {ny: количество примыкающих не тупиковых ветвей.}
        for ny in all_ny:
            if ny not in ny_end:
                v_not_end = 0  # записываем не тупиковые ветви узла
                for i in ny_all_vetv[ny]:
                    if i not in all_v_end:
                        v_not_end += 1
                if v_not_end > 2:
                    all_ny_nodal[ny] = v_not_end
                else:
                    all_ny_transit.append(ny)
        # Заполнить номера транзитов.
        num_transit = 0
        transit_num_all_ny = defaultdict(list)  # {номер транзита: все входящие узлы}
        transit_num_all_v_end = defaultdict(list)  # {номер транзита: крайние ветви транзита (ip, iq, np)}
        ny_use = set()

        for ny in all_ny_transit:
            if ny in ny_use:
                continue
            ny_use.add(ny)
            num_transit += 1
            transit_num_all_ny[num_transit].append(ny)
            for ip, iq, np_ in ny_all_vetv[ny]:
                v_end_transit = (ip, iq, np_)
                ny_next = iq if ip == ny else ip
                while ny_next:
                    ny_source = ny_next
                    ny_next = 0
                    if ny_source not in all_ny_nodal:
                        ny_use.add(ny_source)
                        transit_num_all_ny[num_transit].append(ny_source)
                        for ip1, iq1, np_1 in ny_all_vetv[ny_source]:
                            if (ip1, iq1, np_1) in all_v_end:
                                continue
                            for i in [ip1, iq1]:
                                if i not in ny_use:
                                    ny_next = i
                                    v_end_transit = (ip1, iq1, np_1)
                                    break
                            if ny_next:
                                break
                    else:
                        transit_num_all_v_end[num_transit].append(v_end_transit)
                        # log_rm.debug((num_transit, v_end_transit))

        log_rm.debug(f'В РМ {num_transit} групп транзитных узлов.')

        # Внести номера транзитов в таблицу узлы растра
        all_ny_transit = []  # [(транзитные узлы, номер транзита,)]
        ny__num_transit = {}  # {номер узла: номер транзита}
        for num in transit_num_all_ny:
            for ny in transit_num_all_ny[num]:
                all_ny_transit.append((ny, num,))
                ny__num_transit[ny] = num
        all_ny_transit = all_ny_transit + [(ny, -(all_ny_nodal[ny]),) for ny in all_ny_nodal]
        self.rastr.tables('node').ReadSafeArray(2, 'ny,transit', all_ny_transit)

        # Внести номера транзитов в таблицу ветви растра
        all_transit_one = []  # [(ip, iq, np_) всех транзитных ветвей состоящих из 1 элемента.]
        all_v_transit = []  # [(ip, iq, np_, num) все транзитные ветви]
        for i in all_vetv_sta0:
            if i in all_v_end:
                continue
            ip, iq, np_ = i
            num = 0
            if ip in ny__num_transit:
                num = ny__num_transit[ip]
            elif iq in ny__num_transit:
                num = ny__num_transit[iq]
            if num:
                all_v_transit.append((ip, iq, np_, num,))
                self.v__num_transit[(ip, iq, np_)] = num
            else:
                # Транзит из 1 ветви
                num_transit += 1
                all_transit_one.append((ip, iq, np_))
                all_v_transit.append((ip, iq, np_, num_transit,))
                self.v__num_transit[(ip, iq, np_)] = num_transit
        vetv.ReadSafeArray(2, 'ip,iq,np,transit', all_v_transit)

        if disable_on:
            # Отключаемые узлы
            node = self.rastr.tables('node')
            node.setsel(selection_node_for_disable + '&transit<-3')  # 4-х и более отходящих транзитов
            node.cols.item(field).calc(1)
            log_rm.info(f'{len(node)} отключаемых узлов')
            # Отключаемы ветви
            node.setsel(selection_node_for_disable)
            sel_ny = node.writesafearray('ny', "000")
            sel_ny = [x[0] for x in sel_ny]
            all_v_disable = []  # Все отключаемые ветви
            transit_use = []  # Уже добавленные в отключения номера транзиты
            v__gr = {(ip, iq, np_): groupid for ip, iq, np_, groupid, tip, pl_ip, groupid in data_v}
            v__pl = {(ip, iq, np_): pl_ip for ip, iq, np_, groupid, tip, pl_ip, groupid in data_v}
            v__tip = {(ip, iq, np_): tip for ip, iq, np_, groupid, tip, pl_ip, groupid in data_v}
            node.setsel('')
            ny__un = {ny: uhom for ny, uhom in self.rastr.tables('node').writesafearray('ny,uhom', "000")}
            for ny in sel_ny:
                if ny not in ny_end:
                    for v in ny_all_vetv[ny]:  # Цикл по прилегающим ветвям
                        if v not in all_v_end:  # Без тупиков
                            if v in all_transit_one:  # todo поверить all_transit_one
                                all_v_disable.append(v)
                            else:
                                ip, iq, np_ = v
                                ny_transit = ip if ip in ny__num_transit else 0
                                if not ny_transit:
                                    ny_transit = iq if iq in ny__num_transit else 0

                                if ny_transit:
                                    num_transit = ny__num_transit[ny_transit]
                                    if num_transit in transit_use or num_transit not in transit_num_all_v_end:
                                        continue
                                    transit_use.append(num_transit)
                                    # Сравнить groupid концов транзита, если одинаковый, то отключаем конец
                                    # с большей суммой напряжений ip и ip
                                    # log_rm.debug(transit_num_all_v_end)
                                    # log_rm.debug(transit_num_all_v_end[num_transit])
                                    # log_rm.debug(num_transit)
                                    ip1, iq1, np_1 = transit_num_all_v_end[num_transit][0]
                                    ip2, iq2, np_2 = transit_num_all_v_end[num_transit][1]
                                    if v__gr[(ip1, iq1, np_1)] == v__gr[(ip2, iq2, np_2)] and v__gr[(ip1, iq1, np_1)]:
                                        # В случае АТ, нужно отключать обмотку ВН
                                        if (ny__un[ip1] + ny__un[iq1]) > (ny__un[ip2] + ny__un[iq2]):
                                            all_v_disable.append((ip1, iq1, np_1))
                                        else:
                                            all_v_disable.append((ip2, iq2, np_2))
                                    else:
                                        # Отключаем оба конца. Если разница P < 1, то любой конец.
                                        # Положительное направление в центр транзита
                                        p1 = v__pl[(ip1, iq1, np_1)]  # Поток от начала к концу со знаком -
                                        if ip1 in all_ny_nodal:
                                            p1 = -p1
                                        p2 = v__pl[(ip2, iq2, np_2)]
                                        if ip2 in all_ny_nodal:
                                            p2 = -p2

                                        if abs(p1 + p2) > 1:
                                            all_v_disable.append((ip2, iq2, np_2))
                                            all_v_disable.append((ip1, iq1, np_1))
                                        else:
                                            if v__tip[(ip1, iq1, np_1)] == 2:  # выключатель
                                                all_v_disable.append((ip2, iq2, np_2))
                                            else:
                                                all_v_disable.append((ip1, iq1, np_1))
            # todo в all_v_disable  есть дубликаты ?
            # todo опционо убрать выключатели
            all_v_disable = tuple(set([(ip, iq, np_, 1) for ip, iq, np_ in all_v_disable]))

            if all_v_disable:
                log_rm.info(f'{len(all_v_disable)} отключаемых ветвей')
                vetv.ReadSafeArray(2, 'ip,iq,np,' + field, all_v_disable)

    @staticmethod
    def name_table_from_key(task_key: str):
        """
        По ключу строки (нр ny=1) определяет имя таблицы.
        :return: table:str 'node' or False если не найдено.
        """
        for key_tables in RastrModel.KEY_TABLES:
            if key_tables in task_key:
                return RastrModel.KEY_TABLES[key_tables]
        return False

    def sta(self, table_name: str, ndx: int = 0, key_int: int | tuple = 0) -> bool:
        """
        Отключить ветвь(группу ветвей, если groupid!=0), узел (с примыкающими ветвями) или генератор.
        Отключаемый элемент определяется по ndx или key_int.
        :param table_name: Имя таблицы: 'node', 'vetv', 'Generator'
        :param ndx:
        :param key_int: Например: узел 10 или ветвь (1, 2, 0).
        :return: False если элемент отключен в исходном состоянии.
        """
        # Проверка ИД
        if table_name not in ['node', 'vetv', 'Generator']:
            raise ValueError(f'При вызове функции sta не правильно указано имя таблицы {table_name}.')
        if not ndx:
            if key_int:
                ndx = self.index(table_name=table_name, key_int=key_int)
            else:
                raise ValueError('При вызове функции sta не указаны входные параметры.')

        rtable = self.rastr.tables(table_name)

        # Проверка состояния отключаемого элемента
        if not key_int:
            key_int = self.t_i_key[table_name].get(ndx, 0)
        if key_int and self.t_sta[table_name]:
            if self.t_sta[table_name].get(key_int):
                return False
        else:
            if rtable.cols.item('sta').Z(ndx) == 1:
                return False

        # Отключение элемента
        if table_name == 'vetv':
            if self.v_gr and key_int:
                groupid = self.v_gr.get(key_int)
            else:
                groupid = rtable.cols.item('groupid').Z(ndx)
            if groupid:
                rtable.setsel(f'groupid={groupid}')
                rtable.cols.item('sta').Calc(1)
                rtable.cols.item('sel').Calc(1)
                return True
        rtable.cols.item('sta').SetZ(ndx, 1)
        rtable.cols.item('sel').SetZ(ndx, 1)
        return True

    def test_name(self, condition: dict, info: str = "") -> bool:
        """
         Проверка имени файла на соответствие условию condition.
        :param condition:
        {"years":"2020,2023...2025","season": "лет, зим, паводок","max_min":"макс","add_name":"-41С;МДП:ТЭ-У"}
        :param info: для вывода в протокол
        :return: True если удовлетворяет
        """
        if not condition:
            return True
        if not (any(condition.values())):  # условие пустое
            return True
        # Проверка года
        if 'years' in condition:
            if condition['years']:
                if not int(self.info_file['Год']) in cf.str_yeas_in_list(str(condition['years'])):
                    log_rm.info(f"{info} {self.Name!r}. Год не проходит по условию: {condition['years']!r}")
                    return False
        # Проверка "зим" "лет" "паводок"
        if 'season' in condition:
            if condition['season']:
                if self.info_file['Сезон'][:3] not in condition['season']:
                    log_rm.info(f'{info} {self.Name!r}. Сезон не проходит по условию: {condition["season"]!r}')
                    return False
        # Проверка "макс" "мин"
        if 'max_min' in condition:
            if condition['max_min']:
                if self.info_file['макс/мин'][:3] not in condition['max_min']:
                    log_rm.info(f'{info} {self.Name!r}. Не проходит по условию: {condition["max_min"]!r}')
                    return False
        # Проверка доп имени, например (-41С;МДП:ТЭ-У)
        if 'add_name' in condition:
            if condition['add_name']:
                if condition['add_name'].strip():
                    for us in condition['add_name'].split(";"):
                        if us not in self.additional_name_list:
                            log_rm.debug(f'{info} {self.Name}. Не проходит по условию: {us!r}')
                            return False
        return True

    def load(self):
        """
        Загрузить модель в Rastr
        """
        if not self.rastr:
            try:
                self.rastr = win32com.client.Dispatch("Astra.Rastr")
            except Exception:
                raise Exception('Com объект Astra.Rastr не найден')

        self.rastr.Load(1, self.full_name, self.pattern)  # Загрузить или перезагрузить
        log_rm.info(f"Загружен файл: {self.full_name}")

        # При загрузке файла rst или rg2 загружать одноименные файлы trn и anc (из той же папки)
        if self.config['load_trn_anc']:
            for type_file_add in ['trn', 'anc']:
                name_file_add = f'{self.dir}\\{self.name_base}.{type_file_add}'
                if os.path.exists(name_file_add):
                    self.add_load.append(type_file_add)
                    self.rastr.Load(1, name_file_add, self.config[f'шаблон {type_file_add}'])
                    log_rm.info(f"Загружен файл: {name_file_add}")

    def downloading_additional_files(self, load_additional: list = None):
        """
        Загрузка в Rastr дополнительных файлов из папки с РМ.
        :param load_additional: ['amt','sch','trn']
        """
        for extension in load_additional:
            files = os.listdir(self.dir)
            names = list(filter(lambda x: x.endswith('.' + extension), files))
            if len(names) > 0:
                self.rastr.Load(1, f'{self.dir}\\{names[0]}', self.config[f"шаблон {extension}"])
                log_rm.info(f"Загружен файл: {names[0]}")
            else:
                raise ValueError(f'Файл с расширением {extension!r} не найден в папке {self.dir}')

    def save(self, full_name_new: str = '', file_name: str = '', folder_name: str = ''):
        """
        Сохранить файл. Указать полное имя или имя файла (без расширения) с папкой.
        """
        if not full_name_new:
            if file_name and folder_name:
                full_name_new = folder_name + '\\' + re.sub(r'[\\/|:?<>.]', '', file_name)
                full_name_new = full_name_new[:252]
                full_name_new += '.' + self.type_file
            else:
                raise ValueError(f'Ошибка в входных аргументах функции.')
        # Запрос о перезаписи файлов.
        if not folder_name:
            folder_name = os.path.dirname(full_name_new)
        if os.path.exists(full_name_new):
            if self.overwrite_new_file == 'question':
                RastrModel.overwrite_new_file = mb.askquestion('Внимание!',
                                                         f'Заменить файлы в папке: {folder_name}')
        # Запись файла.
        if self.overwrite_new_file != 'no':
            self.rastr.Save(full_name_new, self.pattern)
            log_rm.info("Файл сохранен: " + full_name_new)
            if self.add_load:
                for type_file_add in self.add_load:
                    full_name_new_add = f'{full_name_new.rsplit(".", 1)[0]}.{type_file_add}'
                    self.rastr.Save(full_name_new_add,
                                    self.config[f"шаблон {type_file_add}"])
                    log_rm.info("Файл сохранен: " + full_name_new_add)
            return full_name_new

    def checking_parameters_rg2(self, dict_task: dict):
        """  контроль  dict_task = {'node': True, 'vetv': True, 'Gen': True, 'section': True,
             'area': True, 'area2': True, 'darea': True, 'sel_node': "na>0"}  """
        if not self.rgm("checking_parameters_rg2"):
            return False
        # self.fill_field_index('node,vetv,Generator')
        self.rastr.CalcIdop(self.info_file["Темп.(°C)"], 0.0, "")
        log_rm.info(f'Выполнен расчет загрузки ветвей для температуры {self.info_file["Темп.(°C)"]}.')

        node = self.rastr.tables("node")
        branch = self.rastr.tables("vetv")
        generator = self.rastr.tables("Generator")
        # Удаление узлов без ветвей, ветвей без узлов начала или конца, генераторов без узлов.
        all_ny = set([x[0] for x in node.writesafearray("ny", "000")])
        all_ip = set([x[0] for x in branch.writesafearray("ip", "000")])
        all_iq = set([x[0] for x in branch.writesafearray("iq", "000")])
        all_iq_ip = all_ip.union(all_iq)

        # Узлы без ветвей.
        all_ny_not_branches = all_ny - all_iq_ip
        if all_ny_not_branches:
            log_rm.error(f'В таблице node удалены узлы без ветвей: {all_ny_not_branches}')
            for ny_not_branches in all_ny_not_branches:
                self.cor(keys=str(ny_not_branches), values='del', print_log=True)
        # Ветви без узлов.
        all_ip_iq_not_node = all_iq_ip - all_ny
        if all_ip_iq_not_node:
            log_rm.error(f'В таблице vetv есть ссылка на узлы которых нет в таблице node: {all_ip_iq_not_node}')
            for ip, iq, np_ in branch.writesafearray('ip,iq,np', "000"):
                if ip in all_ip_iq_not_node or iq in all_ip_iq_not_node:
                    self.cor(keys=f'{ip},{iq},{np_}', values='del', print_log=True)

        # Генераторы без узлов.
        if generator.size:
            dict_task['Gen'] = False
            all_gen_ny = set([x[0] for x in generator.writesafearray("Node", "000")])
            all_gen_not_node = all_gen_ny - all_ny
            if all_gen_not_node:
                log_rm.error(f'В таблице Generator есть ссылка на узлы которых нет в таблице node: {all_gen_not_node}')
                for Num, Node in generator.writesafearray('Num,Node', "000"):
                    if Node in all_gen_not_node:
                        self.cor(keys=f'Num={Num}', values='del', print_log=True)

        if dict_task["sel_node"]:
            # node.setsel(dict_task["sel_node"])
            # ny_sel = set([x[0] for x in node.writesafearray("ny", "000")])
            # todo додумать
            self.add_fields_in_table(name_tables='node', fields='sel1', type_fields=3)
            node.cols.item("sel1").calc(0)
            node.setsel(dict_task["sel_node"])
            node.cols.item("sel1").calc(1)

        # Напряжения
        if dict_task["node"]:
            log_rm.info("\tКонтроль напряжений.")
            self.voltage_error(choice=dict_task["sel_node"])
            self.voltage_fine(choice=dict_task["sel_node"])

        # Токи
        if dict_task['vetv']:
            # Контроль токовой загрузки

            if dict_task["sel_node"]:
                sel_vetv = "i_zag>=0.1&(ip.sel1|iq.sel1)"
                presence_n_it = {'n_it': "(ip.sel1|iq.sel1)&n_it>0",
                                 'n_it_av': "(ip.sel1|iq.sel1)&n_it_av>0"}
            else:
                sel_vetv = "i_zag>=0.1"
                presence_n_it = {'n_it': "n_it>0",
                                 'n_it_av': "n_it_av>0"}

            log_rm.debug('Контроль токовой загрузки.')
            branch.setsel(sel_vetv)
            if branch.count:  # есть превышения
                j = branch.FindNextSel(-1)
                while j > -1:
                    log_rm.info(f"\t\tВНИМАНИЕ ТОКИ! vetv:{branch.SelString(j)}, "
                                f"{branch.cols.item('name').ZS(j)} - {branch.cols.item('i_zag').ZS(j)} %")
                    j = branch.FindNextSel(j)

            log_rm.debug('Проверка наличия n_it,n_it_av в таблице График_Iдоп_от_Т(graphikIT).')
            graph_it = self.rastr.tables("graphikIT")
            if graph_it.size:
                all_graph_it = set([x[0] for x in graph_it.writesafearray("Num", "000")])
                for field, sel_vetv_n_it in presence_n_it.items():
                    branch.setsel(sel_vetv_n_it)
                    for i in branch.writesafearray(field + ",name,ip,iq,np", "000"):
                        if i[0] > 0 and i[0] not in all_graph_it:
                            log_rm.error(f"\t\tВНИМАНИЕ graphikIT! vetv: {i[1]} [{i[2]},{i[3]},{i[4]}] "
                                         f"{field}={i[0]} не найден в таблице График_Iдоп_от_Т")

        #  ГЕНЕРАТОРЫ
        if dict_task['Gen']:
            log_rm.info("\tКонтроль генераторов")
            sel_gen = "!sta&Node.sel1" if dict_task["sel_node"] else "!sta"
            generator.setsel(sel_gen)
            col = {'Num': 0, 'Node': 1, 'Name': 2, 'Pmin': 3, 'Pmax': 4, 'P': 5, 'NumPQ': 6}
            if generator.count:
                for i in generator.writesafearray(','.join(col), "000"):
                    Pmin = i[col['Pmin']]
                    Pmax = i[col['Pmax']]
                    P = i[col['P']]
                    Name = i[col['Name']]
                    Num = i[col['Num']]
                    Node = i[col['Node']]
                    NumPQ = i[col['NumPQ']]
                    if P < Pmin and Pmin:
                        log_rm.info(f"\t\tВНИМАНИЕ! ГЕНЕРАТОР: {Name}, {Num=},{Node=}, {P=} < {Pmin=}")
                    if P > Pmax and Pmax:
                        log_rm.info(f"\t\tВНИМАНИЕ! ГЕНЕРАТОР: {Name}, {Num=},{Node=}, {P=} > {Pmax=}")
                    if NumPQ and self.rastr.tables("graphik2").size:
                        chart_pq = set([x[0] for x in self.rastr.tables("graphik2").writesafearray("Num", "000")])
                        if NumPQ not in chart_pq:
                            log_rm.info(f"\t\tВНИМАНИЕ! ГЕНЕРАТОР: {Name}, {Num=},{Node=}, "
                                        f"{NumPQ=} не найден в таблице PQ-диаграммы (graphik2)")
        return True

    def cor_rm_from_txt(self, task_txt: str) -> str:
        """
        Корректировать модели по заданию в текстовом формате:
        Имя_функции[действие]{Условие_выполнения}#комментарии\n...
        Имя_функции по умолчанию = изм
        :param task_txt:
        :return: Информация
        """
        info = []

        task_rows = task_txt.split('\n')
        for task_row in task_rows:
            task_row = task_row.split('#')[0]  # удалить текст после '#'
            name_fun = task_row.split('[', 1)[0]  # Имя функции, стоит перед "[".
            name_fun = name_fun.replace(' ', '')
            if not name_fun:
                if '[' in task_row:
                    name_fun = 'изм'
                else:
                    continue  # К следующей строке.
            # Цикличность
            cycle_condition = False
            if '*}' in task_row:
                cycle_condition = True
                task_row = task_row.replace('*}', '}')

            # Параметры функции в квадратных скобках
            sel = ''
            value = ''
            param = ''
            if '[' in task_row and ']' in task_row:
                param = task_row.split('[', 1)[1].rsplit(']', 1)[0]

            if param:
                if ':' in param:
                    sel, value = param.split(':', maxsplit=1)
                else:
                    sel = param

            # Условие выполнения в фигурных скобках
            execution_condition = ''
            if '{' in task_row:
                match = re.search(re.compile(r"\{(.+?)}"), task_row)
                if match:
                    execution_condition = match[1].strip()
                else:
                    raise ValueError(f'Ошибка в условии {task_row}')

                if not self.conditions_test(task_row):
                    log_rm.debug(f'Условие не выполняется: {task_row}')
                    continue  # К следующей строке.
                else:
                    log_rm.debug(f'Условие выполняется: {task_row}')
            if not cycle_condition:
                execution_condition = ''
            info_i = self.txt_task_cor(name=name_fun,
                                       sel=sel,
                                       value=value,
                                       all_task=param,
                                       execution_condition=execution_condition,
                                       cycle_condition=cycle_condition)
            if info_i:
                info.append(info_i)
        all_info = ', '.join(info) if info else ''
        log_rm.debug(all_info)
        return all_info

    def txt_task_cor(self,
                     name: str,
                     sel: str = '',
                     value: str = '',
                     all_task: str = '',
                     execution_condition: str = '',
                     cycle_condition: bool = False) -> str:
        """
        Функция для выполнения задания в текстовом формате
        :param name: Имя функции.
        :param all_task: Всё задание внутри []
        :param sel: Выборка, нр, 15145; 12,13.
        :param value: Значение, нр, name=Промплощадка: изм name; pg=qn*2+10.
        :param cycle_condition: Если истина, то выполнять действие пока условие не станет ложным;
        :param execution_condition: условие выполнения;
        :return: информация
        """
        name = name.lower()
        if 'уд' in name:
            return self.cor(keys=sel, values='del', del_all=('*' in name), print_log=True)
        elif 'изм' in name:
            return self.cor(keys=sel,
                            values=value,
                            print_log=True,
                            execution_condition=execution_condition,
                            cycle_condition=cycle_condition)
        elif 'импорт' in name:
            self.txt_import_rm(type_import=sel, description=value)
        elif 'снять' in name:
            return self.cor(keys='(node); (vetv); (Generator)', values='sel=0', print_log=True)
        elif 'расчет' in name:
            self.rgm(txt='txt_task_cor')
            return 'выполнен расчет режима'
        elif 'добавить' in name:
            self.table_add_row(table=sel, tasks=value)
        elif 'текст' in name:
            self.txt_field_right(tasks=all_task)
            return "\tИсправить пробелы, заменить английские буквы на русские."
        elif 'схн' in name:
            self.shn(choice=sel)
        elif 'сечение' in name:
            raise ValueError('Функция загрузки сечений не реализована.')
            # sel = sel.replace(' ', '')
            # for i in ['ns:', 'psech:', 'выбор:', 'тип:']:
            #     if i not in sel:
            #         raise ValueError(f'В задании "сечение": {sel!r} отсутствует ключ {i!r}')
            # sel = sel.split(';')
            # sd = {}
            # for _ in sel:
            #     key, val = _.split(':')
            #     sd[key] = val
            # ls.loading_section(ns=sd['ns'], p_new=sd['psech'], type_correction=sd['тип'])
        elif 'напряжения' in name:
            self.voltage_error(choice=sel, edit=True)

        elif 'анализ' in name:
            if sel == '-':
                self.network_analysis(disable_on=False)
            else:
                self.network_analysis(selection_node_for_disable=sel)
        elif 'скрм' in name:
            if 'скрм*' in name:
                self.all_auto_shunt = self.auto_shunt_rec(selection=sel)
            else:
                self.all_auto_shunt = self.auto_shunt_rec(selection=sel, only_auto_bsh=True)
            self.auto_shunt_cor(all_auto_shunt=self.all_auto_shunt)
        else:
            raise ValueError(f'Задание {name=} не распознано ({sel=}, {value=})')
        return ''

    def txt_import_rm(self, type_import: str, description: str):
        """
        Импорт данных из РМ.
        :param type_import: Если 'папка', то переносить данные из одноименных файлов в указанной папке,
         'файл' - из указанного файла
        :param description: "(I:\pop);таблица:node; тип:2; поле: pn,qn; выборка:"
        :return:
        """
        description_dict = {}
        path = re.search(re.compile(r"\((.+)\)"), description)[1]
        description_list = description.replace(path, '').replace(' ', '').split(';')
        dict_name = {'таблица': 'tables', 'тип': 'calc', 'поле': 'param', 'выборка': 'sel'}

        for i in description_list:
            if ':' in i:
                key, val = i.split(':')
                for x in dict_name:
                    if x in key:
                        description_dict[dict_name[x]] = val

        if type_import == 'папка':
            file_name = path + '\\' + self.Name
            if os.path.isfile(file_name):
                path = file_name
        if os.path.isfile(path):
            ifm = ImportFromModel(RastrModel(path),
                                  **description_dict)
            ifm.import_data_in_rm(rm=self)
        else:
            if type_import == 'файл':
                raise ValueError(f'Файл для импорта не найден {path}')
            if type_import == 'папка':
                log_rm.error(f'Папка для импорта не найдена {path}')

    def node_include(self) -> str:
        """
        Восстановление питания узлов путем включения выключателей (r<0.011 & x<0.011).
        :return: Информация о включенных узлах
        """
        # self.ny_join_vetv
        log_rm.debug('Восстановление питания отключенных узлов.')
        node_info = ''
        node_all = set()
        node_include = set()

        for ny, sta, pn, qn, pg, qg in self.rastr.tables('node').writesafearray('ny,sta,pn,qn,pg,qg', "000"):
            if not self.t_sta['node'][ny] and sta and (pn or qn or pg or qg):
                node_all.add((ny, self.t_name["node"][ny]))

                for s_key in self.ny_join_vetv[ny]:
                    r, x, _ = self.v_rxb[s_key]
                    if r < 0.011 and x < 0.011:
                        ny_connectivity = s_key[0] if ny != s_key[0] else s_key[1]
                        # log_rm.debug(ny_connectivity)
                        ndx = self.t_key_i['node'][ny_connectivity]
                        if not self.rastr.tables('node').Cols("sta").Z(ndx):  # Питающий узел включен.
                            # Включить узел и ветвь
                            self.rastr.tables('node').Cols("sta").SetZ(self.t_key_i['node'][ny], False)
                            self.rastr.tables('vetv').Cols("sta").SetZ(self.t_key_i['vetv'][s_key], False)
                            node_include.add((ny, self.t_name["node"][ny]))
                            break

        if node_include:
            node_info = "Восстановлено питание узлов:"
            for ny, name in node_include:
                node_info += f' {name} ({ny}),'
            node_info = node_info.strip(',') + ". "

        node_not_include = node_all - node_include
        if node_not_include:
            node_info += "Узлы, оставшиеся без питания:"
            for ny, name in node_not_include:
                node_info += f' {name} ({ny}),'
            node_info = node_info.strip(',') + ". "

        if node_info:
            log_rm.info('\tnode_include: ' + node_info)
        return node_info.strip()

    def cor(self,
            keys: str = '',
            values: str = '',
            print_log: bool = False,
            del_all: bool = False,
            execution_condition: str = '',
            cycle_condition: bool = False) -> str:
        """
        Коррекция значений в таблицах rastrwin.
        В круглых скобках указать имя таблицы (н.р. na=1(node)).
        Если корректировать все строки таблицы, то указать только имя таблицы, н.р. (node).
        Если выборка по ключам, то имя таблицы указывать не нужно (н.р. ny=1 в таблице узлы).
        Краткая форма выборки по узлам: 12;12,13,1;g=12 вместо ny=12;ip=12&iq=13&np=1;Num=12.
        Если np=0, то выборка по ветвям можно записать еще короче: «12,13», вместо «12,13,0».
        При задании краткой формы имя таблицы указывать не нужно.

        :param keys: Если несколько выборок, то указываются через ";"
        "125;ny=25;na=1(node)" для узлов, "Num=25;g=12" для генераторов, "1,2" для ветви,
        "na=2;no=1;npa=1;nga=2" для районов, объединения, территорий и нагрузочных групп;

        :param values:  Удалить строки в таблице 'del'
        Изменить значение параметров: 'pn=10.2;qn=qn*2' ;
        Использование ссылок на другие значения таблиц rastr: 'pn=10.2;qn=qn*2+30:qn+1,2(vetv):ip'

        :param print_log: выводить в лог;
        :param del_all: удалять узлы с генераторами и отходящими ветвями;
        :param execution_condition: условие выполнения;
        :param cycle_condition: Если истина, то выполнять действие пока условие не станет ложным;
        :return: Информация об отключении
        """
        info = []
        if print_log:
            log_rm.info(f"\t\tФункция cor: {keys=},  {values=}")
        keys = keys.replace(' ', '')
        values = values.strip().replace('  ', ' ')
        if not (keys and values.replace(' ', '')):
            raise ValueError(f'{keys=},{values=}')
        for key in keys.split(";"):  # например:['na=11(node)','125', 'g=125', '12,13,0']
            rastr_table, selection_in_table = self.recognize_key(key, 'tab sel')

            for value in values.split(";"):  # разделение задания, например:['pn=10.2', 'qn=5.4']
                param = ''
                if value == 'del':
                    formula = 'del'
                elif '=' in value:
                    param, formula = value.split("=", maxsplit=1)
                    param = param.replace(' ', '')
                    if not param:
                        raise ValueError(f"Задание не распознано, {key=}, {value=}")

                    # В значении есть ссылка и поле не текстовое.
                    if ':' in value and not self.rastr.tables(rastr_table).cols(param).Prop(1) == 2:
                        formula = self.replace_links(formula)
                else:
                    raise ValueError(f"Задание не распознано, {key=}, {value=}")

                info.append(self.group_cor(tabl=rastr_table,
                                           param=param,
                                           selection=selection_in_table,
                                           formula=formula,
                                           del_all=del_all,
                                           execution_condition=execution_condition,
                                           cycle_condition=cycle_condition))
        return ', '.join(info)

    def recognize_key(self, key: str, back: str = 'all'):
        """
        Распознать:
         -имя таблицы;
         -короткий ключ (s_key);
         -выборку в таблице.
        :param key: например:['na=11(node)','125', 'g=125', '12,13,0', '12,13', 'ip=5&iq=3&np=0']
        :param back: тип возвращаемого значения
        :return:'all' (имя таблицы: str, выборка: str, ключ: int|tuple(int,int,int?))
                'tab' имя таблицы: str
                's_key'  ключ: int|tuple(int,int,int)
                'tab s_key'
                'sel' выборка: str
                'tab sel'
        """
        if isinstance(key, str):
            key = key.replace(' ', '')
        else:
            key = str(int(key))
        selection_in_table = key  # выборка в таблице
        if key == '-1':
            rastr_table = 'vetv'
            s_key = -1
        else:
            rastr_table = ''  # имя таблицы
            # проверка наличия явного указания таблицы
            match = re.search(re.compile(r"\((.+?)\)"), key)
            s_key = 0
            if match:  # таблица указана
                rastr_table = match[1]
                selection_in_table = selection_in_table.split('(', 1)[0]

            if selection_in_table:
                # разделение ключей для распознания
                key_comma = selection_in_table.split(",")  # нр для ветви [ , , ], прочее []
                key_equally = selection_in_table.split("=")  # есть = [, ], нет равно []
                if ',' in selection_in_table:  # vetv
                    if len(key_comma) > 3:
                        raise ValueError(f'Ошибка в задании {key=}')
                    rastr_table = 'vetv'
                    if len(key_comma) == 3:
                        selection_in_table = f"ip={key_comma[0]}&iq={key_comma[1]}&np={key_comma[2]}"
                        s_key = (int(key_comma[0]), int(key_comma[1]), int(key_comma[2]),)
                    if len(key_comma) == 2:
                        selection_in_table = f"ip={key_comma[0]}&iq={key_comma[1]}&np=0"
                        s_key = (int(key_comma[0]), int(key_comma[1]), 0)
                else:
                    if selection_in_table.isdigit():
                        rastr_table = 'node'
                        s_key = int(selection_in_table)
                        selection_in_table = "ny=" + selection_in_table
                    elif 'g' == key_equally[0]:
                        rastr_table = 'Generator'
                        s_key = int(key_equally[1])
                        selection_in_table = "Num=" + key_equally[1]
                    elif len(key_equally) == 2:
                        s_key = int(key_equally[1])
                        if not rastr_table:
                            if key_equally[0] in self.KEY_TABLES:
                                rastr_table = self.KEY_TABLES[key_equally[0]]  # вернет имя таблицы
                    elif len(key_equally) > 2:  # "ip = 1&iq = 2&np = 0"
                        rastr_table = 'vetv'
                        ip = int(key_equally[1].split('&')[0])
                        iq = int(key_equally[2].split('&')[0])
                        np_ = 0 if len(key_equally) == 3 else int(key_equally[3])
                        s_key = (ip, iq, np_)  # if np_ else (ip, iq)
            if not rastr_table:
                raise ValueError(f"Таблица не определена: {key=}")

        if back == 's_key':
            return s_key
        elif back == 'sel':
            return selection_in_table
        elif back == 'tab':
            return rastr_table
        elif back == 'tab s_key':
            return rastr_table, s_key
        elif back == 'tab sel':
            return rastr_table, selection_in_table
        return rastr_table, selection_in_table, s_key

    def group_cor(self,
                  tabl: str,
                  param: str = '',
                  selection: str = '',
                  formula: str = '',
                  del_all: bool = False,
                  execution_condition: str = '',
                  cycle_condition: bool = False) -> str:
        """
        Групповая коррекция;
        :param tabl: таблица, нр 'node';
        :param param: параметры, нр 'pn';
        :param selection: выборка, нр 'sel';
        :param formula: 'del' удалить строки, формула для расчета параметра, нр 'pn*2' или значение, нр '10'.
        Меняются все поля в выборке через 'Calc'. А значит formula может быть например 'pn*0.4'
        :param del_all: удалять узлы с генераторами и отходящими ветвями;
        :param execution_condition: условие выполнения;
        :param cycle_condition: Если истина то выполнять действие пока условие не станет ложным;
        :return: Информация об отключении
        """
        if execution_condition:
            if not self.conditions_test(execution_condition):
                return ''
        if cycle_condition and (not execution_condition):
            raise ValueError(f"Ошибка в задании {cycle_condition!r}, {execution_condition!r}.")

        if self.rastr.tables.Find(tabl) < 0:
            raise ValueError(f"В rastrwin не загружена таблица {tabl!r}.")

        table = self.rastr.tables(tabl)
        table.setsel(selection)
        num = table.count
        if not num:
            log_rm.debug(f'В таблице {tabl!r} по выборке {selection!r} не найдено строк.')
            return ''
        index = table.FindNextSel(-1)

        start_value = table.cols.Item(param).ZS(index) if param else ''

        if formula == 'del':
            table.DelRows()
            if del_all and tabl == 'node':  # Удалить ветви и генераторы.
                if 'ny=' in selection:
                    ny = selection.split('=')[1]
                    table_vetv = self.rastr.tables('vetv')
                    table_vetv.setsel(f"ip={ny}|iq={ny}")
                    table_vetv.DelRows()
                    table_gen = self.rastr.tables('Generator')
                    table_gen.setsel("Node=" + ny)
                    table_gen.DelRows()
            return f'удаление {num} строк(и) в таблице {tabl} по выборке {selection}'
        else:
            if table.cols.Find(param) < 0:
                raise ValueError(f"В таблице {tabl!r} нет параметра {param!r}.")

            if table.cols(param).Prop(1) == 2:  # если поле типа строка

                if table.count > 1:
                    new_data = [(*x, formula) for x in table.writesafearray(table.Key, "000")]
                    table.ReadSafeArray(2, table.Key + ',' + param, new_data)
                else:

                    table.cols.Item(param).SetZ(index, formula)

            else:  # если поле типа число
                if isinstance(formula, str):
                    formula = formula.replace(' ', '').replace(',', '.')

                table.cols.item(param).Calc(formula)
                if cycle_condition:
                    for i in range(1000):
                        self.rgm(f'Расчет {i} в цикле')
                        if self.conditions_test(execution_condition):
                            table.cols.item(param).Calc(formula)
                        else:
                            break

            if num > 1:
                return f'изменение {num} строк по выборке {selection} параметра {param}, {formula!r}'
            elif num == 1:
                # [отключение / включение:]    [имя(ВЛ, 1СШ, ЮЖНАЯ) / выборка]
                # [имя(ВЛ, 1СШ, ЮЖНАЯ) / выборка]    [: изменение]    [нагрузки / генерации][с до]
                # [имя(ВЛ, 1СШ, ЮЖНАЯ) / выборка]    [: изменение]    [парам(vzd)][с до]
                info = ''
                if param == 'sta':
                    if str(formula) == '1':
                        info = 'отключение: '
                    elif str(formula) == '0':
                        info = 'включение: '

                name = selection
                for n in ['dname', 'name', 'Name']:
                    if table.cols.Find(n) > -1:
                        name1 = table.cols(n).Z(index).strip()
                        if name1:
                            name = name1
                            break

                info += name
                if param == 'sta':
                    return info

                info += ': изменение '
                if param in ['pg', 'P']:
                    info += 'генерации '
                elif param in ['pn', 'qn']:
                    info += 'нагрузки '
                else:
                    info += f'{param} '
                info += f' c {start_value} до {table.cols(param).ZS(index)}'
                return info

    def txt_field_return(self, table_name: str, selection: str, field_name: str):
        """
        Считать поле в таблице rastrwin.
        :param table_name:
        :param field_name:
        :param selection: Выборка в таблице
        :return: Значение поля
        """
        table = self.rastr.tables(table_name)
        table.setsel(selection)
        return table.cols.Item(field_name).Z(table.FindNextSel(-1))

    def voltage_fix_frame(self):
        """
        В таблице узлы задать поля umin(uhom*1.15*0.7), umin_av(uhom*1.1*0.7), если uhom>100
        и и_umax
        """
        log_rm.info(f"Заполнение поля umax и umin таблицы node.")
        node = self.rastr.tables("node")
        data = []
        field = 'ny,uhom,umax,umin,umin_av'
        for ny, uhom, umax, umin, umin_av in node.writesafearray(field, "000"):
            if umin == 0 and uhom > 100:
                umin = uhom * 1.15 * 0.7
            if umin_av == 0 and uhom > 100:
                umin_av = uhom * 1.1 * 0.7
            if not umax:
                if uhom < 50:
                    umax = uhom * 1.2
                if 50 < uhom < 300:
                    umax = uhom * 1.1455
                if uhom == 330:
                    umax = uhom * 1.1
                if uhom > 400:
                    umax = uhom * 1.05
                if uhom == 750:
                    umax = uhom * 1.0493

            data.append((ny, uhom, umax, umin, umin_av,))
        node.ReadSafeArray(2, field, data)

    def voltage_fine(self, choice: str = ''):
        """
        Проверка расчетного напряжения:
        меньше наибольшего рабочего,
        минимального рабочего напряжения,
        больше минимально-допустимого.
        :param choice: Выборка в таблице узлы
        """
        log_rm.info('Проверка расчетного напряжения.')
        node = self.rastr.tables("node")
        for i in range(len(self.U_NOM)):
            sel_node = f"!sta&uhom={self.U_NOM[i]}"
            if choice:
                sel_node += f"&{choice}"
            node.setsel(sel_node)
            if node.count:
                for name, ny, vras in node.writesafearray('name,ny,vras', "000"):
                    vras = round(vras, 1)
                    if not self.U_MIN_NORM[i] < vras < self.U_LARGEST_WORKING[i]:
                        if self.U_MIN_NORM[i] > vras:
                            log_rm.warning(f"\tНизкое напряжение: {ny=}, {name=}, {vras=}, uhom={self.U_NOM[i]}")
                        if vras > self.U_LARGEST_WORKING[i]:
                            log_rm.warning(f"\tПревышение наибольшего рабочего напряжения: "
                                           f"{ny=}, {name=}, {vras=}, uhom={self.U_NOM[i]}")

        sel_node = "vras>0&vras<umin"  # Отклонение напряжения от umin минимально допустимого, в %
        if choice:
            sel_node += "&" + choice
        node.setsel(sel_node)
        if node.count:
            for name, ny, vras, umin in node.writesafearray('name,ny,vras,umin', "000"):
                log_rm.warning(f"\tНапряжение ниже минимально-допустимого: {ny=}, {name=}, {vras=}, {umin=}")

    def voltage_error(self, choice: str = '', edit: bool = False):
        """
        Проверка номинального напряжения на соответствие ряду [6, 10, 35, 110, 220, 330, 500, 750].
        Если umax<uhom, то umax удаляется;
        Если umin>uhom, umin_av>uhom, то umin, umin_av удаляется.
        :param choice: Выборка в таблице узлы
        :param edit: Испраить значения в РМ
        """
        node = self.rastr.tables("node")
        if edit:
            self.fill_field_index('node')
        else:
            self.add_fields_in_table(name_tables='node', fields='index', type_fields=0)
        data = []
        fild_set = ''
        node.setsel(choice)
        if node.count:
            rst_on = False
            fild_set = 'name,ny,uhom,index,umax,umin,umin_av'
            if node.cols.Find("umin_av") < 0:
                fild_set = 'name,ny,uhom,index,na,npa,nga'
                rst_on = True
            data_b = node.writesafearray(fild_set, "000")
            for name, ny, uhom, index, umax, umin, umin_av in data_b:
                add = False
                # Номинальное напряжение.
                if uhom not in self.U_NOM:
                    for x in range(len(self.U_NOM)):
                        if self.U_MIN_NORM[x] < uhom < self.U_LARGEST_WORKING[x]:
                            log_rm.warning(f"\tНесоответствие номинального напряжения: "
                                           f"{ny=}, {name=}, {uhom=}->{self.U_NOM[x]}.")
                            uhom = self.U_NOM[x]
                            add = True
                            break
                # Ошибки
                if not rst_on and umax and umax < uhom:
                    log_rm.warning(f"\tОшибка:{ny=},{name=}, {umax=}<{uhom=}.")
                    umax = 0
                    add = True
                if not rst_on and umin > uhom:
                    log_rm.warning(f"\tОшибка: {ny=},{name=}, {umin=}>{uhom=}.")
                    umin = 0
                    add = True
                if not rst_on and umin_av > uhom:
                    log_rm.warning(f"\tОшибка: {ny=},{name=}, {umin_av=}>{uhom=}.")
                    umin_av = 0
                    add = True

                if edit and add:
                    data.append((name, ny, uhom, index, umax, umin, umin_av,))

        if edit and data:
            log_rm.warning(f"\tОшибки исправлены.")
            node.ReadSafeArray(2, fild_set, data)

    def rgm(self, txt: str = "", param: str = '') -> bool:
        """
        Расчет режима
        :param txt: Для вывода в лог
        :param param: Параметр функции rastr.rgm(param)
        Параметр функции rastr.rgm():
        "" – c параметрами по умолчанию; 40 мс
        "p" – расчет с плоского старта; 93 мс
        "z" – отключение стартового алгоритма; 39 мс
        "c" – отключение контроля данных; 38 мс
        "r" – отключение подготовки данных (можно использовать
        при повторном расчете режима с измененными значениями узловых мощностей и модулей напряжения). 34 мс
        "zcr" - 19 мс
        :return: False если развалился.
        """
        if txt:
            log_rm.debug(f"Расчет режима: {txt}")
        for i in (param, '', '', 'p', 'p', 'p'):
            kod_rgm = self.rastr.rgm(i)  # 0 сошелся, 1 развалился
            if not kod_rgm:  # 0 сошелся
                return True
        # развалился
        log_rm.info(f"Расчет режима: {txt} !!!РАЗВАЛИЛСЯ!!!")
        return False

    def all_cols(self, tab: str, val_return: str = 'str'):
        """
        Возвращает все поля таблицы, кроме начинающихся с '_'.
        :param tab:
        :param val_return: Варианты: 'str' или 'list'
        :return:
        """
        cls = self.rastr.Tables(tab).Cols
        cols_list = []
        for col in range(cls.Count):
            if cls(col).Name[0] != '_':
                # print(cls(col).Name)
                cols_list.append(cls(col).Name)
        if val_return == 'str':
            return ','.join(cols_list)
        elif val_return == 'list':
            return cols_list

    def table_add_row(self, table: str = '', tasks: str = '') -> int:
        """
        Добавить запись в таблицу и вернуть index.
        :param table: таблица, например "vetv";
        :param tasks: параметры в формате "ip=1;iq=2; np=10; i_dop=100.5";
        :return: index;
        """
        table = table.strip()
        if not all([table, tasks]):
            raise ValueError(f'\tОшибка при добавлении в таблицу <{table=}> строки <{tasks=}>')

        r_table = self.rastr.tables(table)
        r_table.AddRow()  # добавить строку в конце таблицы
        index = r_table.size - 1
        for task_i in tasks.split(";"):
            if task_i:
                if task_i.count('=') == 1:
                    parameters, value = task_i.split("=")
                    parameters = parameters.replace(' ', '')
                    if all([parameters, value]):
                        if r_table.cols.Find(parameters) < 0:
                            raise ValueError(f"В таблице {r_table!r} нет параметра {parameters!r}.")
                        if r_table.cols(parameters).Prop(1) == 2:  # если поле типа строка
                            r_table.cols.Item(parameters).SetZ(index, value)
                        else:
                            r_table.cols.item(parameters).SetZ(index, value.replace(' ', '').replace('.', ','))

                    else:
                        raise ValueError(f'\tОшибка при добавлении в таблицу <{table=}> строки <{task_i=}>')
                else:
                    raise ValueError(f'\tОшибка при добавлении в таблицу <{table=}> строки <{task_i=}>(проблемы с = )')

        log_rm.info(f'\tВ таблицу <{table}> добавлена строка <{tasks}>, индекс <{index}>')
        return index

    def txt_field_right(self, tasks: str = 'node:name,dname;vetv:dname;Generator:Name'):
        """Исправить пробелы, заменить английские буквы на русские."""
        for task in tasks.replace(' ', '').split(';'):
            name_table, field_table = task.split(':')
            for field_table_i in field_table.split(','):
                self.cor_letter_space(table=name_table, field=field_table_i)

    def cor_letter_space(self, table: str, field: str):
        """
        Изменить текстовые значения в таблице.
        Английские буквы менять на русские.
        Удалить пробел в начале и в конце.
        Заменить 2 пробела на 1.
        :param table: имя таблицы
        :param field: имя поля
        :return:
        """
        matching_letter = {
            "E": "Е",
            "T": "Т",
            "O": "О",
            "P": "Р",
            "A": "А",
            "H": "Н",
            "K": "К",
            "X": "Х",
            "C": "С",
            "B": "В",
            "M": "М",
            "e": "е",
            "o": "о",
            "p": "р",
            "a": "а",
            "x": "х",
            "c": "с",
            "b": "в"}
        r_table = self.rastr.tables(table)
        data = []
        fields = f'{field},{r_table.Key}'
        for i in r_table.writesafearray(fields, "000"):
            i = list(i)
            new = i[0]
            # заменить буквы
            for key in matching_letter:
                while key in new:
                    new = new.replace(key, matching_letter[key])
            while '  ' in new:
                new = new.replace('  ', ' ')
            while ' -' in new and ' - ' not in new:
                new = new.replace(' -', ' - ')
            while '- ' in new and ' - ' not in new:
                new = new.replace('- ', ' - ')
            new = new.strip()
            if not i[0] == new:
                log_rm.info(f'\t\tИсправление текстового поля: {table, field} <{i[0]}> на <{new}>')
                i[0] = new
                data.append(i)
        if data:
            r_table.ReadSafeArray(2, fields, data)

    def shn(self, choice: str = ''):
        """
        Добавить зависимости СХН в таблицу узлы (uhom>100 nsx=1, uhom<100 nsx=2)
        :param choice: выборка, нр na=100
        """
        log_rm.info("\tДобавлены зависимости СХН в таблицу узлы (uhom>100 nsx=1, uhom<100 nsx=2)")
        choice = choice + '&' if choice else ''
        self.group_cor(tabl="node", param="nsx", selection=choice + "uhom>100", formula="1")
        self.group_cor(tabl="node", param="nsx", selection=choice + "uhom<100", formula="2")

    def cor_pop(self, zone: str, new_pop: int | float) -> bool:
        """
        Изменить потребление.
        :param zone: Например, "na=3", "npa=2" или "no=1"
        :param new_pop: значение потребления
        :return:
        """
        eps = 0.003 if new_pop < 50 else 0.0003  # точность расчета, *100=%
        react_cor = True  # менять реактивное потребление пропорционально
        if '=' not in str(zone):
            raise ValueError(f"Ошибка в задании, cor_pop: {zone=}, {new_pop=}")
        zone_id = zone.partition('=')[0]
        name_zone = {"na": "area", "npa": "area2", "no": "darea",
                     "name_na": "район", "name_npa": "территория", "name_no": "объединение",
                     "p_na": "pop", "p_npa": "pop", "p_no": "pp"}
        t_node = self.rastr.tables("node")
        t_zone = self.rastr.tables(name_zone[zone_id])
        t_zone.setsel(zone)
        ndx_z = t_zone.FindNextSel(-1)
        t_area = self.rastr.tables("area")
        if zone_id == "no":
            t_area.setsel(zone)
        if t_zone.cols.Find("set_pop") > 0:
            t_zone.cols.Item("set_pop").SetZ(ndx_z, new_pop)
        name_z = t_zone.cols.item('name').ZS(ndx_z)
        pop_beginning = self.rastr.Calc("val", name_zone[zone_id], name_zone[f'p_{zone_id}'], zone)
        for i in range(10):  # максимальное число итераций
            self.rgm('cor_pop')
            pop = self.rastr.Calc("val", name_zone[zone_id], name_zone[f'p_{zone_id}'], zone)
            # нр("val", "darea", "pp", "no=1")
            ratio = new_pop / pop  # 50/55=0.9
            if abs(ratio - 1) > eps:
                if zone_id == "no":
                    ndx_na = t_area.FindNextSel(-1)
                    while ndx_na != -1:
                        i_na = t_area.cols.item("na").Z(ndx_na)
                        t_node.setsel("na=" + str(i_na))
                        t_node.cols.item("pn").Calc(f"pn*{ratio}")
                        if react_cor:
                            t_node.cols.item("qn").Calc(f"qn*{ratio}")
                        ndx_na = t_area.FindNextSel(ndx_na)

                elif zone_id in ["npa", "na"]:
                    t_node.setsel(zone)
                    t_node.cols.item("pn").Calc("pn*" + str(ratio))
                    if react_cor:
                        t_node.cols.item("qn").Calc("qn*" + str(ratio))

                if not self.rgm('cor_pop'):
                    log_rm.error(f"Аварийное завершение расчета, cor_pop: {zone=}, {new_pop=}")
                    return False
            else:
                log_rm.info(f"\t\tПотребление {name_z!r}({zone}) = {pop_beginning:.1f} МВт изменено на {pop:.1f} МВт"
                            f" (задано {new_pop}, отклонение {abs(new_pop - pop):.1f} МВт, {i + 1} ит.)")
                return True

    def auto_shunt_rec(self, selection: str = '', only_auto_bsh: bool = False) -> dict:
        """
        Функция формирует словарь all_auto_shunt с объектами класса AutoShunt для записи СКРМ.
        :param selection: Выборка в таблице узлы
        :param only_auto_bsh: True узлы только с заданным значением в поле AutoBsh. False все узлы с СКРМ
        :return словарь[ny] = namedtuple('СКРМ')
        """
        log_rm.debug(f'Поиск узлов с СКРМ {selection=}')
        all_auto_shunt = {}
        KU = namedtuple('СКРМ', ['ny', 'name', 'ny_adjacency', 'ny_control', 'umin', 'umax',
                                 'type', ])  # KU компенсирующее устройство
        have_AutoBsh = True
        node = self.rastr.tables('node')
        vetv = self.rastr.tables('vetv')
        if node.cols.Find('AutoBsh') < 0:
            have_AutoBsh = False
            if only_auto_bsh:
                raise ValueError(f"В таблице node нет параметра AutoBsh.")
        selection_result = selection + '&pn=0&qn=0&pg=0&qg=0&bsh!=0' if selection else 'pn=0&qn=0&pg=0&qg=0&bsh!=0'
        node.setsel(selection_result)
        i = node.FindNextSel(-1)
        while i > -1:
            AutoBsh = ''
            if have_AutoBsh:
                AutoBsh = node.cols.item("AutoBsh").ZS(i)
                AutoBsh = AutoBsh.replace(' ', '')
                if not AutoBsh and only_auto_bsh:
                    i = node.FindNextSel(i)
                    continue  # если только по полю AutoBsh и оно не заполнено, то к следующему узлу
            ny = node.cols.item("ny").Z(i)
            name = node.cols.item("name").Z(i)
            type_ = 'ШР' if node.cols.item("bsh").Z(i) > 0 else 'БСК'
            vetv.setsel(f'ip={ny}|iq={ny}')
            if not vetv.count == 1:
                i = node.FindNextSel(i)
                continue  # если не 1 ветвь, то к следующему узлу
            iv = vetv.FindNextSel(-1)
            ip = vetv.cols.item("ip").Z(iv)
            iq = vetv.cols.item("iq").Z(iv)

            ny_adjacency = ip if ny == iq else iq
            ny_control = ''

            if AutoBsh:  # 105-126.5;ny=100
                if '(' in AutoBsh:
                    log_rm.error(f'Ошибка в задании {AutoBsh=}')
                    i = node.FindNextSel(i)
                    continue
                if ';' in AutoBsh:
                    try:
                        u, ny_control = AutoBsh.split(';')
                        ny_control = int(ny_control.replace('ny=', ''))
                    except Exception:
                        raise ValueError(f'Ошибка в задании {AutoBsh=}')
                    AutoBsh = u
                umin, umax = AutoBsh.split('-')
                if not (umin and umax):
                    log_rm.error(f'Ошибка в задании {AutoBsh=}')
                    i = node.FindNextSel(i)
                    continue
            else:
                uhom = node.cols.item("uhom").Z(i)
                if uhom > 300:
                    umin = round(uhom * 0.95, 1)
                    umax = round(uhom * 1.05, 1)
                else:
                    umin = round(uhom * 0.95, 1)
                    umax = round(uhom * 1.14, 1)

            all_auto_shunt[ny] = KU(ny, name, ny_adjacency, ny_control, int(umin), int(umax), type_)
            log_rm.debug(f'Обнаружено СКРМ: {ny=} {name=} {ny_adjacency=} {ny_control=} {umin=} {umax=}')
            i = node.FindNextSel(i)
        return all_auto_shunt

    def auto_shunt_cor(self, all_auto_shunt: dict) -> str:
        """
        Функция включает или отключает узлы с СКРМ в соответствии с уставкой по напряжению.
        :param all_auto_shunt: Словарь с namedtuple('СКРМ')
        """
        changes_in_rm = ''
        if not all_auto_shunt:
            return ''
        for ny in all_auto_shunt:
            ku = all_auto_shunt[ny]
            ny_test = ku.ny_control if ku.ny_control else ku.ny_adjacency
            i = self.index(table_name='node', key_str=f'ny={ny}')
            i_test = self.index(table_name='node', key_str=f'ny={ny_test}')
            node = self.rastr.tables('node')
            sta = node.cols.item("sta").Z(i)  # 1 откл, 0 вкл
            volt_test = round(node.cols.item("vras").Z(i_test), 1)
            if volt_test:
                if volt_test < ku.umin:
                    if ku.type == 'БСК':  # включить
                        if sta:
                            self.sta_node_with_branches(ny=ny, sta=0)
                            self.rgm('auto_shunt_cor')
                            volt_result = round(node.cols.item("vras").Z(i_test), 1)
                            changes_in_rm += (f'\nВключена БСК {ny=} {ku.name!r},'
                                              f' напряжение увеличилось с {volt_test} до {volt_result}.')
                    elif ku.type == 'ШР':  # отключить
                        if not sta:
                            node.cols.item("sta").SetZ(i, 1)
                            self.rgm('auto_shunt_cor')
                            volt_result = round(node.cols.item("vras").Z(i_test), 1)
                            changes_in_rm += (f'\nОтключен ШР {ny=} {ku.name!r},'
                                              f' напряжение увеличилось с {volt_test} до {volt_result}.')
                elif volt_test > ku.umax:
                    if ku.type == 'БСК':  # отключить
                        if not sta:
                            node.cols.item("sta").SetZ(i, 1)
                            self.rgm('auto_shunt_cor')
                            volt_result = round(node.cols.item("vras").Z(i_test), 1)
                            changes_in_rm += (f'\nОтключена БСК {ny=} {ku.name!r},'
                                              f' напряжение снизилось с {volt_test} до {volt_result}.')
                    elif ku.type == 'ШР':  # включить
                        if sta:
                            self.sta_node_with_branches(ny=ny, sta=0)
                            self.rgm('auto_shunt_cor')
                            volt_result = round(node.cols.item("vras").Z(i_test), 1)
                            changes_in_rm += (f'\nВключен ШР {ny=} {ku.name!r},'
                                              f' напряжение снизилось с {volt_test} до {volt_result}.')
        if changes_in_rm:
            log_rm.info(changes_in_rm)
        return changes_in_rm

    def table_index_setsel(self, table_name: str, setsel: str):
        """
        Вернуть list из индексов строк таблице в соответствии с выборкой.
        :param table_name: Имя таблицы
        :param setsel: Выборка в таблице
        :return:
        """
        # todo удалить
        table = self.rastr.tables(table_name)
        if table.cols.Find("index") < 0:
            self.fill_field_index(table_name)
        table = self.rastr.tables(table_name)
        table.setsel(setsel)
        return [x[0] for x in table.writesafearray("index", "000")]

    def add_fields_in_table(self, name_tables: str, fields: str, type_fields: int, prop=(), replace=False):
        """
        Добавить поля в таблицу, если они отсутствуют.
        :param name_tables: Можно несколько через запятую.
        :param fields: Можно несколько через запятую.
        :param type_fields: Тип поля: 0 целый, 1 вещ, 2 строка, 3 переключатель(sta sel), 4 перечисление, 6 цвет
        :param prop: ((0-12, значение), ()) prop=((8, '2'), (0, 'yes')) или ((8, '2'),)
        0 Имя, 1 Тип, 2 Ширина, 3 Точность, 4 Заголовок
        5 Формула "str(ip.name)+"+"+str(iq.name)+"_"+str(ip.uhom)"
        6-, 7-, 8 Перечисление – ссылка, 9 Описание, 10 Минимум, 11 Максимум, 12 Масштаб
        :param replace: True предварительно удалить поле если оно существует
        """
        for name_table in name_tables.replace(' ', '').split(','):
            table = self.rastr.tables(name_table)
            for field in fields.replace(' ', '').split(','):
                if table.cols.Find(field) > -1 and replace:
                    table.Cols.Remove(field)
                if table.cols.Find(field) < 0:
                    table.Cols.Add(field, type_fields)
                    if prop != ():
                        for property_number, val in prop:
                            table.Cols(field).SetProp(property_number, val)  # (номер свойства, новое значение)
                            log_rm.debug(f'В таблицу {name_table} добавлено поле {field}.')
                            # table.Cols(field).Prop(5)  # Получить значение
                else:
                    log_rm.debug(f'В таблицу {name_table} поле {field} уже имеется.')

    def df_from_table(self, table_name: str, fields: str = '', setsel: str = ''):
        """
        Возвращает DataFrame из таблицы.
        :param table_name:
        :param fields: Если не указывать, то все поля.
        :param setsel: Выборка в таблице
        :return:
        """
        table = self.rastr.tables(table_name)
        table.setsel(setsel)
        # if not table.count:
        #     return False
        if not fields:
            fields = self.all_cols(table_name)
        fields = fields.replace(' ', '').replace(',,', ',').strip(',')
        part_table = table.writesafearray(fields, "000")
        return pd.DataFrame(data=part_table, columns=fields.split(','))

    def table_from_df(self, df: pd.DataFrame, table_name: str, fields: str = '', type_import: int = 2):
        """
        Записать в таблицу растр DataFrame.
        :param table_name:
        :param df:
        :param fields: Если не указывать, то все колонки.
        :param type_import: Обновить: 2, загрузить: 1, дополнить: 0, обновить-добавить: 3.
        :return:
        """
        table = self.rastr.tables(table_name)

        if not fields:
            fields = ','.join(df.columns)

        table.ReadSafeArray(type_import, fields, tuple(df.itertuples(index=False, name=None)))

    def fill_field_index(self, name_tables: str):
        """
        Создать и заполнить поле index таблиц
        :param name_tables:
        """
        for name_table in name_tables.replace(' ', '').split(','):
            self.add_fields_in_table(name_tables=name_table, fields='index', type_fields=0)
            table = self.rastr.tables(name_table)
            keys = [(*x, i) for i, x in enumerate(table.writesafearray(table.Key, "000"), 0)]
            table.ReadSafeArray(2, table.Key + ',index', keys)
            log_rm.debug(f'В таблице {name_table} заполнено поле index.')

    def sta_node_with_branches(self, ny: int, sta: int):
        """Включить/отключить узел с ветвями."""
        if not ny:
            raise ValueError(f'Ошибка в задании {ny=}')
        self.cor(keys=str(ny), values='sta=' + str(sta))
        vetv = self.rastr.tables('vetv')
        vetv.setsel(f'ip={ny}|iq={ny}')
        vetv.cols.item("sta").calc(sta)

    def replace_links(self, formula: str) -> str:
        """
        Функция заменяет в формуле ссылки на значения в таблицах rastr, на соответствующие значения.
        :param formula: '(10.5+15,16,2:r)*ip.uhom'
        :return: formula: '(10.5+z)*ip.uhom'
        """
        # formula = formula.replace(' ', '')
        formula_list = re.split('\*|/|\^|\+|-|\(|\)|==|!=|&|\||not|>|<|<=|=<|>=|=>', formula)
        for formula_i in formula_list:
            if ':' in formula_i:
                if any([txt in formula_i for txt in ['years', 'season', 'max_min', 'add_name']]):
                    continue
                sel_all, field = formula_i.split(':')
                name_table, sel = self.recognize_key(sel_all, 'tab sel')
                self.rgm(f'для определения значения {formula}')
                index = self.index(table_name=name_table, key_str=sel)
                if index > -1:
                    new_val = self.rastr.tables(name_table).cols.Item(field).ZS(index)
                    formula = formula.replace(formula_i, new_val)
                else:
                    raise ValueError(f'В таблице {name_table} отсутствует {sel}')
        return formula

    def index(self, table_name: str, key_int: int | tuple = 0, key_str: str = '') -> int:
        """
        Возвращает номер строки в таблице по ключу в одном из форматов.
        :param table_name: Имя таблицы:'vetv' ...
        :param key_int: Например: узел 10 или ветвь (1, 2, 0). При наличии t_key_i индекс берется из них.
        :param key_str: Например: 'ny=10' или 'ip=1&iq=2&np=3'
        :return: index
        """
        if not table_name:
            raise ValueError(f'Ошибка в задании {table_name=}.')
        if key_int:
            if table_name in ['node', 'vetv', 'Generator'] and key_int in self.t_key_i[table_name]:
                return self.t_key_i[table_name][key_int]
            else:
                t = self.rastr.tables(table_name)

                if table_name == 'vetv':
                    np_ = key_int[2] if len(key_int) == 3 else 0
                    t.setsel(f'ip={key_int[0]}&iq={key_int[1]}&np={np_}')
                else:
                    t.setsel(f'{t.Key}={key_int}')
                i = t.FindNextSel(-1)
                if i == -1:
                    raise ValueError(f'index: В таблице {table_name} не найдена строка по ключу {key_int} ')
                    # log_rm.warning
                return i
        if key_str:
            t = self.rastr.tables(table_name)
            t.setsel(key_str)
            i = t.FindNextSel(-1)
            if i == -1:
                raise ValueError(f'index: В таблице {table_name} не найдена строка по ключу {key_str} ')
                # log_rm.warning
            return i

    def conditions_test(self, conditions: str) -> bool:
        """
        В строке типа "years : 2026...2029& ny=1: vras>125|(not ny=1: na==2)" проверяет выполнение условий.
        Если в conditions имеются {}, то значения берутся внутри скобок
        :param conditions:
        :return:
        """
        log_rm.debug(f'Проверка условия: {conditions}')
        if '{' in conditions:
            match = re.search(re.compile(r"\{(.+?)}"), conditions)
            if match:
                conditions = match[1].strip()
            else:
                raise ValueError(f'Ошибка в условии {conditions}')
        conditions = conditions.strip('*')  # если в условии предусмотрен цикл
        conditions_s = conditions
        conditions = self.replace_links(conditions)
        conditions_list = re.split('\*|/|\^|\+|-|\(|\)|==|!=|&|\||not|>|<|<=|=<|>=|=>', conditions)
        for condition in conditions_list:
            if ':' in condition:
                for key_txt in ['years', 'season', 'max_min', 'add_name']:
                    if not self.code_name_rg2:  # Если имя не стандартное, то True.
                        conditions = conditions.replace(condition, 'True')
                        continue
                    else:
                        if key_txt in condition:
                            par, value = condition.split(':')

                            if self.test_name(condition={par.replace(' ', ''): value.strip()},
                                              info=condition):
                                conditions = conditions.replace(condition, 'True')
                            else:
                                conditions = conditions.replace(condition, 'False')
        if ':' in conditions:
            raise ValueError("Ошибка в условии: " + conditions)
        try:
            conditions_test = bool(eval(conditions))
            log_rm.debug(f'{conditions_test=}: {conditions}')
            return conditions_test
        except Exception:
            raise ValueError(f'Ошибка у условии: {conditions_s!r}.')

    @staticmethod
    def key_to_str(ob):
        """Преобразовать переменную в строку. Для удаления np=0 key vetv.
        Убрать последний ноль в tuple или list"""
        if isinstance(ob, tuple | list):
            if ob[-1] == 0:
                ob = ob[:-1]
            return ','.join(map(str, ob))
        else:
            return str(ob)
