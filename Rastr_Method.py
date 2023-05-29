import logging
import re
from collections import namedtuple
from typing import Union
import pandas as pd
log_r_m = logging.getLogger('__main__.' + __name__)


class RastrMethod:
    """
    Сборник методов работы с rastr
    """
    # Работа с напряжением узлов
    U_NOM = [6, 10, 35, 110, 150, 220, 330, 500, 750]  # номинальные напряжения
    U_MIN_NORM = [5.8, 9.7, 32, 100, 135,  205, 315, 490, 730]  # минимальное нормальное напряжение
    U_LARGEST_WORKING = [7.2, 12, 42, 126, 180,  252, 363, 525, 787]  # наибольшее рабочее напряжения

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

    def __init__(self):
        self.rastr = None

    def cor(self, keys: str = '', values: str = '', print_log: bool = False, del_all: bool = False) -> str:
        """
        Коррекция значений в таблицах rastrwin.

        В круглых скобках указать имя таблицы (н.р. na=1(node)).
        Если корректировать все строки таблицы, то указать только имя таблицы, н.р. (node).
        Если выборка по ключам, то имя таблицы указывать не нужно (н.р. ny=1 в таблице узлы).
        Краткая форма выборки по узлам: 12;12,13,1;g=12 вместо ny=12;ip=12&iq=13&np=1;Num=12.
        Если np=0, то выборка по ветвям можно записать еще короче: «12,13», вместо «12,13,0».
        При задании краткой формы имя таблицы указывать не нужно.

        :param keys: "125;ny=25;na=1(node)" для узлов, "Num=25;g=12" для генераторов, "1,2" для ветви,
        "na=2;no=1;npa=1;nga=2" для районов, объединения, территорий и нагрузочных групп;
        Если несколько выборок, то указываются через ";".

        :param values:  Удалить строки в таблице 'del'
        Изменить значение параметров: 'pn=10.2;qn=qn*2' ;
        Использование ссылок на другие значения таблиц rastr: 'pn=10.2;qn=qn*2+30:qn+1,2(vetv):ip'
        :param print_log: выводить в лог;
        :param del_all: удалять узлы с генераторами и отходящими ветвями;
        :return: Информация об отключении
        """
        info = []
        if print_log:
            log_r_m.info(f"\t\tФункция cor: {keys=},  {values=}")
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

                info.append(self.group_cor(tabl=rastr_table, param=param, selection=selection_in_table,
                                           formula=formula, del_all=del_all))
        return ', '.join(info)

    def recognize_key(self, key: str, back: str = 'all'):
        """
        Распознать имя таблицы и выборку в таблице по короткой записи.
        :param key: например:['na=11(node)','125', 'g=125', '12,13,0', '12,13']
        :param back: тип возвращаемого значения
        :return:'all' (имя таблицы: str, выборка: str, ключ: int|tuple(int,int,int?))
                'tab' имя таблицы: str
                's_key'  ключ: int|tuple(int,int,int)
                'tab s_key'
                'sel' выборка: str
                'tab sel'
        """
        key = key.replace(' ', '')
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
                key_comma = selection_in_table.split(",")  # нр для ветви [,,], прочее []
                key_equally = selection_in_table.split("=")  # есть = [,], нет равно []
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
                  del_all: bool = False) -> str:
        """
        Групповая коррекция;
        :param tabl: таблица, нр 'node';
        :param param: параметры, нр 'pn';
        :param selection: выборка, нр 'sel';
        :param formula: 'del' удалить строки, формула для расчета параметра, нр 'pn*2' или значение, нр '10'.
        Меняются все поля в выборке через 'Calc'. А значит formula может быть например 'pn*0.4'
        :param del_all: удалять узлы с генераторами и отходящими ветвями;
        :return: Информация об отключении
        """
        if self.rastr.tables.Find(tabl) < 0:
            raise ValueError(f"В rastrwin не загружена таблица {tabl!r}.")

        table = self.rastr.tables(tabl)
        table.setsel(selection)
        num = table.count
        if not num:
            log_r_m.debug(f'В таблице {tabl!r} по выборке {selection!r} не найдено строк.')
            return ''
        i = table.FindNextSel(-1)

        start_value = table.cols.Item(param).ZS(i) if param else ''

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
                while i > -1:
                    table.cols.Item(param).SetZ(i, formula)
                    i = table.FindNextSel(i)
            else:
                if type(formula) == str:
                    formula = formula.replace(' ', '').replace(',', '.')
                table.cols.item(param).Calc(formula)

            if num > 1:
                return f'изменение {selection} параметра {param}, {formula!r}'
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
                        name1 = table.cols(n).Z(i).strip()
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
                info += f' c {start_value} до {table.cols(param).ZS(i)}'
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
        log_r_m.info(f"\t Заполнение поля umax и umin таблицы node.")
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
        log_r_m.info('Проверка расчетного напряжения.')
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
                            log_r_m.warning(f"\tНизкое напряжение: {ny=}, {name=}, {vras=}, uhom={self.U_NOM[i]}")
                        if vras > self.U_LARGEST_WORKING[i]:
                            log_r_m.warning(f"\tПревышение наибольшего рабочего напряжения: "
                                            f"{ny=}, {name=}, {vras=}, uhom={self.U_NOM[i]}")

        sel_node = "vras>0&vras<umin"  # Отклонение напряжения от umin минимально допустимого, в %
        if choice:
            sel_node += "&" + choice
        node.setsel(sel_node)
        if node.count:
            for name, ny, vras, umin in node.writesafearray('name,ny,vras,umin', "000"):
                log_r_m.warning(f"\tНапряжение ниже минимально-допустимого: {ny=}, {name=}, {vras=}, {umin=}")

    def voltage_error(self, choice: str = '', edit: bool = False):
        """
        Проверка номинального напряжения на соответствие ряду [6, 10, 35, 110, 220, 330, 500, 750].
        Если umax<uhom, то umax удаляется;
        Если umin>uhom, umin_av>uhom, то umin, umin_av удаляется.
        :param choice: выборка в таблице узлы
        :param edit:
        """
        node = self.rastr.tables("node")
        if edit:
            self.fill_field_index('node')
        else:
            self.add_fields_in_table(name_tables='node', fields='index', type_fields=0)
        data = []
        node.setsel(choice)
        if node.count:
            data_b = node.writesafearray('name,ny,uhom,index,umax,umin,umin_av', "000")
            for name, ny, uhom, index, umax, umin, umin_av in data_b:
                add = False
                # Номинальное напряжение.
                if uhom not in self.U_NOM:
                    for x in range(len(self.U_NOM)):
                        if self.U_MIN_NORM[x] < uhom < self.U_LARGEST_WORKING[x]:
                            log_r_m.warning(f"\tНесоответствие номинального напряжения: "
                                            f"{ny=}, {name=}, {uhom=}->{self.U_NOM[x]}.")
                            uhom = self.U_NOM[x]
                            add = True
                            break
                # Ошибки
                if umax and umax < uhom:
                    log_r_m.warning(f"\tОшибка:{ny=},{name=}, {umax=}<{uhom=}.")
                    umax = 0
                    add = True
                if umin > uhom:
                    log_r_m.warning(f"\tОшибка: {ny=},{name=}, {umin=}>{uhom=}.")
                    umin = 0
                    add = True
                if umin_av > uhom:
                    log_r_m.warning(f"\tОшибка: {ny=},{name=}, {umin_av=}>{uhom=}.")
                    umin_av = 0
                    add = True

                if edit and add:
                    data.append((name, ny, uhom, index, umax, umin, umin_av,))

        if edit:
            log_r_m.warning(f"\tОшибки исправлены.")
            node.ReadSafeArray(2, 'name,ny,uhom,index,umax,umin,umin_av', data)

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
            log_r_m.debug(f"Расчет режима: {txt}")
        for i in (param, '', '', 'p', 'p', 'p'):
            kod_rgm = self.rastr.rgm(i)  # 0 сошелся, 1 развалился
            if not kod_rgm:  # 0 сошелся
                return True
        # развалился
        log_r_m.info(f"Расчет режима: {txt} !!!РАЗВАЛИЛСЯ!!!")
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

        log_r_m.info(f'\tВ таблицу <{table}> добавлена строка <{tasks}>, индекс <{index}>')
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
            if not i[0] == new:
                log_r_m.info(f'\t\tИсправление текстового поля: {table, field} <{i[0]}> на <{new}>')
                i[0] = new
                data.append(i)
        if data:
            r_table.ReadSafeArray(2, fields, data)

    def shn(self, choice: str = ''):
        """
        Добавить зависимости СХН в таблицу узлы (uhom>100 nsx=1, uhom<100 nsx=2)
        :param choice: выборка, нр na=100
        """
        log_r_m.info("\tДобавлены зависимости СХН в таблицу узлы (uhom>100 nsx=1, uhom<100 nsx=2)")
        all_choice = '' if choice == '' else choice + '&'
        self.group_cor(tabl="node", param="nsx", selection=all_choice + "uhom>100", formula="1")
        self.group_cor(tabl="node", param="nsx", selection=all_choice + "uhom<100", formula="2")

    def cor_pop(self, zone: str, new_pop: Union[int, float]) -> bool:
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
                    log_r_m.error(f"Аварийное завершение расчета, cor_pop: {zone=}, {new_pop=}")
                    return False
            else:
                log_r_m.info(f"\t\tПотребление {name_z!r}({zone}) = {pop_beginning:.1f} МВт изменено на {pop:.1f} МВт"
                             f" (задано {new_pop}, отклонение {abs(new_pop - pop):.1f} МВт, {i + 1} ит.)")
                return True

    def auto_shunt_rec(self, selection: str = '', only_auto_bsh: bool = False) -> dict:
        """
        Функция формирует словарь all_auto_shunt с объектами класса AutoShunt для записи СКРМ.
        :param selection: Выборка в таблице узлы
        :param only_auto_bsh: True узлы только с заданным значением в поле AutoBsh. False все узлы с СКРМ
        :return словарь[ny] = namedtuple('СКРМ')
        """
        log_r_m.debug(f'Поиск узлов с СКРМ {selection=}')
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
                    log_r_m.error(f'Ошибка в задании {AutoBsh=}')
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
                    log_r_m.error(f'Ошибка в задании {AutoBsh=}')
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
            log_r_m.debug(f'Обнаружено СКРМ: {ny=} {name=} {ny_adjacency=} {ny_control=} {umin=} {umax=}')
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
            i = self.index(name_table='node', key_str=f'ny={ny}')
            i_test = self.index(name_table='node', key_str=f'ny={ny_test}')
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
            log_r_m.info(changes_in_rm)
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
        :param prop: ((0-12, значение),()) prop=((8, '2'), (0, 'yes')) или ((8, '2'), )
        0 Имя, 1 Тип, 2 Ширина, 3 Точность, 4 Заголовок
        5 Формула   "str(ip.name)+"+"+str(iq.name)+"_"+str(ip.uhom)"
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
                            table.Cols(field).SetProp(property_number, val)  # (номер свойства,новое значение)
                            log_r_m.info(f'В таблицу {name_table} добавлено поле {field}.')
                            # table.Cols(field).Prop(5)  # Получить значение
                else:
                    log_r_m.info(f'В таблицу {name_table} поле {field} уже имеется.')

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
            keys = [(*x, i) for i, x in enumerate(table.writesafearray(table.Key, "000"), 1)]
            table.ReadSafeArray(2, table.Key + ',index', keys)
            log_r_m.debug(f'В таблице {name_table} заполнено поле index.')

    def sta_node_with_branches(self, ny: int, sta: int):
        """Включить/отключить узел с ветвями."""
        if not ny:
            raise ValueError(f'Ошибка в задании {ny=}')
        self.cor(keys=str(ny), values='sta='+str(sta))
        vetv = self.rastr.tables('vetv')
        vetv.setsel(f'ip={ny}|iq={ny}')
        vetv.cols.item("sta").calc(sta)
