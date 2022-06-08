import logging
import re
from collections import namedtuple
from typing import Union

log_r_m = logging.getLogger('__main__.' + __name__)


class RastrMethod:
    """
    Сборник методов работы с rastr
    """
    # Работа с напряжением узлов
    U_NOM = [6, 10, 35, 110, 220, 330, 500, 750]  # номинальные напряжения
    U_MIN_NORM = [5.8, 9.7, 32, 100, 205, 315, 490, 730]  # минимальное нормальное напряжение
    U_LARGEST_WORKING = [7.2, 12, 42, 126, 252, 363, 525, 787]  # наибольшее рабочее напряжения
    KEY_TABLES = {'ny': 'node',
                  'Num': 'Generator',
                  # 'g': 'Generator',
                  'na': 'area',
                  'npa': 'area2',
                  'no': 'darea',
                  'nga': 'ngroup',
                  'ns': 'sechen'}

    def __init__(self):
        self.rastr = None

    def cor(self, keys: str = '', tasks: str = '', print_log: bool = True, del_all: bool = False):
        """
        Коррекция значений в таблицах rastrwin;
        Если несколько выборок, то указываются через ; (н.р. na=1|na=2;no=1; npa=1; nga=2; Num=25; g=12).
        В круглых скобках указать имя таблицы (н.р. na=1(node)).
        Если корректировать все строки таблицы, то указать только имя таблицы, н.р. (node).
        Если выборка по ключам, то имя таблицы указывать не нужно (н.р. ny=1 в таблице узлы).
        Краткая форма выборки по узлам: 12;21 вместо ny=12;ny=21.
        Краткая форма выборки по узлам: 12,13,0 вместо ip=12&iq=13&np=0.
        Краткая форма выборки по генераторам: g=12 вместо Num=12.
        При задании краткой формы имя таблицы указывать не нужно.
        :param keys: "125;ny=25;na=1(node)" для узла, "Num=25;g=12" для генераторов, "1,2,0" для ветви,
        "na=2;no=1;npa=1;nga=2" для районов, объединения, территорий и нагрузочных групп;
        :param tasks: 'del' удалить строки в таблице, 'pn=10.2;qn=qn*2' изменить значение параметров;
        :param print_log: выводить в лог;
        :param del_all: удалять узлы с генераторами и отходящими ветвями;
        """
        if print_log:
            log_r_m.info(f"\t\tФункция cor: {keys=},  {tasks=}")
        keys = keys.replace(' ', '')
        tasks = tasks.strip().replace('  ', ' ')
        if not (keys and tasks.replace(' ', '')):
            raise ValueError(f'{keys=},{tasks=}')
        for key in keys.split(";"):  # например:['na=11(node)','125', 'g=125', '12,13,0']

            rastr_table, selection_in_table = self.recognize_key(key)

            for task in tasks.split(";"):  # разделение задания например:['pn=10.2', 'qn=5.4']
                param = ''
                if task == 'del':
                    formula = 'del'
                elif task.count('=') == 1:
                    param, formula = task.split("=")
                    param = param.replace(' ', '')
                    if not (formula and param):
                        raise ValueError(f"Задание не распознано, {key=}, {task=}")
                else:
                    raise ValueError(f"Задание не распознано, {key=}, {task=}")

                self.group_cor(tabl=rastr_table, param=param, selection=selection_in_table,
                               formula=formula, del_all=del_all)

    def recognize_key(self, key: str) -> tuple:
        """
        Распознать по имя таблицы и выбору в таблице по короткой записи.
        :param key: например:['na=11(node)','125', 'g=125', '12,13,0']
        :return: (имя таблицы: str, выборка: str)
        """
        key = key.replace(' ', '')
        selection_in_table = key  # выборка в таблице
        rastr_table = ''  # имя таблицы
        # проверка наличия явного указания таблицы
        match = re.search(re.compile(r"\((.+?)\)"), key)

        if match:  # таблица указана
            rastr_table = match[1]
            selection_in_table = selection_in_table.split('(', 1)[0]

        if selection_in_table:
            # разделение ключей для распознания
            key_comma = selection_in_table.split(",")  # нр для ветви [,,], прочее []
            key_equally = selection_in_table.split("=")  # есть = [,], нет равно []
            if ',' in selection_in_table:  # vetv
                if len(key_comma) != 3:
                    raise ValueError(f'Ошибка в задании {key=}')
                rastr_table = 'vetv'
                selection_in_table = f"ip={key_comma[0]}&iq={key_comma[1]}&np={key_comma[2]}"
            else:
                if selection_in_table.isdigit():
                    rastr_table = 'node'
                    selection_in_table = "ny=" + selection_in_table
                elif 'g' == key_equally[0]:
                    rastr_table = 'Generator'
                    selection_in_table = "Num=" + key_equally[1]
                else:
                    if not rastr_table:
                        if key_equally[0] in self.KEY_TABLES:
                            rastr_table = self.KEY_TABLES[key_equally[0]]  # вернет имя таблицы
        if not rastr_table:
            raise ValueError(f"Таблица не определена: {key}")
        return rastr_table, selection_in_table

    def group_cor(self, tabl: str, param: str, selection: str, formula: str, del_all: bool = False):
        """
        Групповая коррекция;
        :param tabl: таблица, нр 'node';
        :param param: параметры, нр 'pn';
        :param selection: выборка, нр 'sel';
        :param formula: 'del' удалить строки, формула для расчета параметра, нр 'pn*2' или значение, нр '10'.
        Меняются все поля в выборке через 'Calc'. А значит formula может быть например 'pn*0.4'
        :param del_all: удалять узлы с генераторами и отходящими ветвями;
        """
        if self.rastr.tables.Find(tabl) < 0:
            raise ValueError(f"В rastrwin не загружена таблица {tabl!r}.")

        table = self.rastr.tables(tabl)
        table.setsel(selection)
        if not table.count:
            log_r_m.error(f'В таблице {tabl} по выборке {selection} не найдено строк.')

        if formula == 'del':
            table.DelRows()
            if del_all and tabl == 'node':
                if 'ny=' in selection:
                    ny = selection.split('=')[1]
                    table_vetv = self.rastr.tables('vetv')
                    table_vetv.setsel(f"ip={ny}|iq={ny}")
                    table_vetv.DelRows()
                    table_gen = self.rastr.tables('Generator')
                    table_gen.setsel("Node=" + ny)
                    table_gen.DelRows()
        else:
            if table.cols.Find(param) < 0:
                raise ValueError(f"В таблице {tabl!r} нет параметра {param!r}.")

            if table.cols(param).Prop(1) == 2:  # если поле типа строка
                i = table.FindNextSel(-1)
                while i > -1:
                    table.cols.Item(param).SetZ(i, formula)
                    i = table.FindNextSel(i)
            else:
                cor_param = table.cols.item(param)
                cor_param.Calc(formula.replace(' ', ''))

    def voltage_nominal(self, choice: str = 'uhom>30', edit: bool = False):
        """
        Проверка номинального напряжения на соответствие ряду [6, 10, 35, 110, 220, 330, 500, 750].
        :param choice: выборка в таблице 'узлы'
        :param edit: Исправить номинальные напряжения в узлах
        """
        node = self.rastr.tables("node")
        node.setsel(choice)
        j = node.FindNextSel(-1)
        while j > -1:
            uhom = node.cols.item("uhom").Z(j)

            if uhom not in self.U_NOM:
                ny = node.cols.item('ny').ZS(j)
                name = node.cols.item('name').ZS(j)
                log_r_m.warning(f"\tНесоответствие номинального напряжения! {ny=}, {name=}, {uhom=}.")
                if edit:
                    for x in range(len(self.U_NOM)):
                        if self.U_MIN_NORM[x] < uhom < self.U_LARGEST_WORKING[x]:
                            node.cols.item("uhom").SetZ(j, self.U_NOM[x])
                            log_r_m.info(f"\tВнесены изменения! {ny=}, {name=}, uhom={self.U_NOM[x]}")
                            break
                # Если напряжение не исправилось
                if node.cols.item('uhom').Z(j) not in self.U_NOM:
                    log_r_m.error(f"\tНоминальное напряжение не исправлено! {ny=}, {name=}, {uhom=}")

            j = node.FindNextSel(j)

    def voltage_normal(self, choice: str = ''):
        """
        Проверка расчетного напряжения: меньше наибольшего рабочего, больше минимального рабочего напряжения.
        :param choice: выборка в таблице узлы
        """
        node = self.rastr.tables("node")
        for i in range(len(self.U_NOM)):
            sel_node = "!sta&uhom=" + str(self.U_NOM[i])
            if choice:
                sel_node += "&" + choice
            node.setsel(sel_node)
            j = node.FindNextSel(-1)
            while j > -1:
                ny = node.cols.item('ny').Z(j)
                if not self.U_MIN_NORM[i] < node.cols.item("vras").Z(j) < self.U_LARGEST_WORKING[i]:
                    ny = node.cols.item('ny').ZS(j)
                    name = node.cols.item('name').ZS(j)
                    vras = node.cols.item('vras').ZS(j)
                    if self.U_MIN_NORM[i] > node.cols.item("vras").Z(j):
                        log_r_m.warning(f"\tНизкое напряжение! ny={ny}, имя: {name}, vras={vras}, uhom={self.U_NOM[i]}")
                    if node.cols.item("vras").Z(j) > self.U_LARGEST_WORKING[i]:
                        log_r_m.warning(
                            f"\tВысокое напряжение! ny={ny}, имя: {name}, vras={vras}, uhom={self.U_NOM[i]}")
                j = node.FindNextSel(j)

    def voltage_deviation(self, choice: str = ''):
        """
        Проверка расчетного напряжения: больше минимально-допустимого.
        :param choice: выборка в таблице узлы
        """
        node = self.rastr.tables("node")
        sel_node = "otv_min<0"  # Отклонение напряжения от umin минимально допустимого, в %
        if choice:
            sel_node += "&" + choice
        node.setsel(sel_node)
        j = node.FindNextSel(-1)
        while j > -1:
            ny = node.cols.item('ny').ZS(j)
            name = node.cols.item('name').ZS(j)
            vras = node.cols.item('vras').ZS(j)
            umin = node.cols.item('umin').ZS(j)
            log_r_m.warning(f"\tНапряжение ниже минимально-допустимого! ny={ny}, имя: {name}, vras={vras},umin={umin}")
            j = node.FindNextSel(j)

    def voltage_error(self, choice: str = ''):
        """
        - если umax<uhom, то umax удаляется;
        - если umin>uhom, umin_av>uhom, то umin, umin_av удаляется.
        :param choice: выборка в таблице узлы
        """
        node = self.rastr.tables("node")
        sel_node = "umax<uhom&umax!=0"
        if choice:
            sel_node += "&" + choice
        node.setsel(sel_node)
        j = node.FindNextSel(-1)
        while j > -1:
            ny = node.cols.item('ny').ZS(j)
            name = node.cols.item('name').ZS(j)
            umax = node.cols.item('umax').ZS(j)
            uhom = node.cols.item('uhom').ZS(j)
            log_r_m.warning(f"\tumax<uhom! {ny=},{name=}, {umax=},{uhom=}. umax удалено.")
            node.cols.item('umax').SetZ(j, 0)
            j = node.FindNextSel(j)

        sel_node = "umin>uhom"
        if choice:
            sel_node += "&" + choice
        node.setsel(sel_node)
        j = node.FindNextSel(-1)
        while j > -1:
            ny = node.cols.item('ny').ZS(j)
            name = node.cols.item('name').ZS(j)
            umin = node.cols.item('umin').ZS(j)
            uhom = node.cols.item('uhom').ZS(j)
            log_r_m.warning(f"\tumax<uhom! {ny=},{name=}, {umin=},{uhom=}. umin удалено.")
            node.cols.item('umin').SetZ(j, 0)
            j = node.FindNextSel(j)

        sel_node = "umin_av>uhom"
        if choice:
            sel_node += "&" + choice
        node.setsel(sel_node)
        j = node.FindNextSel(-1)
        while j > -1:
            ny = node.cols.item('ny').ZS(j)
            name = node.cols.item('name').ZS(j)
            umin_av = node.cols.item('umin_av').ZS(j)
            uhom = node.cols.item('uhom').ZS(j)
            log_r_m.warning(f"\tumax<uhom! {ny=},{name=}, {umin_av=},{uhom=}. umin_av удалено.")
            node.cols.item('umin_av').SetZ(j, 0)
            j = node.FindNextSel(j)

    def rgm(self, txt: str = "") -> bool:
        """Расчет режима"""
        for i in ['', '', '', 'p', 'p', 'p']:
            kod_rgm = self.rastr.rgm(i)  # 0 сошелся, 1 развалился
            if not kod_rgm:  # 0 сошелся
                if txt:
                    log_r_m.debug(f"\tРасчет режима: {txt}")
                return True
        # развалился
        log_r_m.error(f"Расчет режима: {txt} !!!РАЗВАЛИЛСЯ!!!")
        return False

    def sel0(self, txt=''):
        """ Снять отметку узлов, ветвей и генераторов"""
        self.rastr.Tables("node").cols.item("sel").Calc("0")
        self.rastr.Tables("vetv").cols.item("sel").Calc("0")
        self.rastr.Tables("Generator").cols.item("sel").Calc("0")
        if txt:
            log_r_m.info("\tСнять отметку узлов, ветвей и генераторов")

    def all_cols(self, tab: str):
        """Возвращает все поля таблицы: 'ny,pn....', кроме начинающихся с '_'. """
        cls = self.rastr.Tables(tab).Cols
        cols_list = []
        for col in range(cls.Count):
            if cls(col).Name[0] != '_':
                # if cls(col).Name not in ["kkluch", "txt_zag", "txt_adtn_zag", "txt_ddtn", "txt_adtn", "txt_ddtn_zag"]:
                # print(cls(col).Name)
                cols_list.append(cls(col).Name)
        return ','.join(cols_list)

    def table_add_row(self, table: str = '', tasks: str = '') -> int:
        """
        Добавить запись в таблицу и вернуть index.
        :param table: таблица, например "vetv";
        :param tasks: параметры в формате "ip=1;iq=2; np=10; i_dop=100";
        :return: index;
        """
        if not all([table, tasks]):
            raise ValueError(f'\tОшибка при добавлении в таблицу <{table=}> строки <{tasks=}>')

        r_table = self.rastr.tables(table)
        r_table.AddRow()  # добавить строку в конце таблицы
        index = r_table.size - 1
        for task_i in tasks.replace(' ', '').split(";"):
            if task_i:
                if task_i.count('=') == 1:
                    parameters, value = task_i.split("=")
                    if all([parameters, value]):
                        if r_table.cols.Find(parameters) < 0:
                            raise ValueError(f"В таблице {r_table!r} нет параметра {parameters!r}.")
                        if r_table.cols(parameters).Prop(1) == 2:  # если поле типа строка
                            r_table.cols.Item(parameters).SetZ(index, value.replace('_', ' '))
                        else:
                            r_table.cols.item(parameters).SetZ(index, value)

                    else:
                        raise ValueError(f'\tОшибка при добавлении в таблицу <{table=}> строки <{task_i=}>')
                else:
                    raise ValueError(f'\tОшибка при добавлении в таблицу <{table=}> строки <{task_i=}>(проблемы с = )')

        log_r_m.info(f'\tВ таблицу <{table}> добавлена строка <{tasks}>, индекс <{index}>')
        return index

    def cor_txt_field(self, table_field: str = 'node:name,dname;vetv:dname;Generator:Name'):
        """        Исправить пробелы, заменить английские буквы на русские.        """
        log_r_m.info("\tИсправить пробелы, заменить английские буквы на русские.")
        for task in table_field.replace(' ', '').split(';'):
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
        if self.rastr.Tables.Find(table) == -1:
            raise ValueError(f'В rastr не найдена таблица <{table}>.')
        r_table = self.rastr.tables(table)
        if r_table.cols.Find(field) == -1:
            raise ValueError(f'В rastr таблице не найдено поле <{field}>.')

        r_field = r_table.cols.item(field)
        r_table.setsel("")
        index = r_table.FindNextSel(-1)
        while index >= 0:
            val1 = r_field.ZS(index)
            # заменить буквы
            for key in matching_letter:
                while key in r_field.ZS(index):
                    r_field.SetZ(index, r_field.ZS(index).replace(key, matching_letter[key]))
            # пробелы
            while '  ' in r_field.ZS(index):
                r_field.SetZ(index, r_field.ZS(index).replace('  ', ' '))
            r_field.SetZ(index, r_field.ZS(index).strip())
            while ' -' in r_field.ZS(index) and ' - ' not in r_field.ZS(index):
                r_field.SetZ(index, r_field.ZS(index).replace(' -', ' - '))
            while '- ' in r_field.ZS(index) and ' - ' not in r_field.ZS(index):
                r_field.SetZ(index, r_field.ZS(index).replace('- ', ' - '))
            if not val1 == r_field.ZS(index):
                log_r_m.info(f'\t\tИсправление текстового поля: {table, field} <{val1}> на <{r_field.ZS(index)}>')
            index = r_table.FindNextSel(index)

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
        :param zone: "na=3", "npa=2" или "no=1"
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
        if t_zone.cols.Find("pop_zad") > 0:
            t_zone.cols.Item("pop_zad").SetZ(ndx_z, new_pop)
        name_z = t_zone.cols.item('name').ZS(ndx_z)
        pop_beginning = self.rastr.Calc("val", name_zone[zone_id], name_zone[f'p_{zone_id}'], zone)
        for i in range(10):  # максимальное число итераций
            self.rgm()
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

                if not self.rgm():
                    log_r_m.error(f"Аварийное завершение расчета, cor_pop: {zone=}, {new_pop=}")
                    return False
            else:
                log_r_m.info(f"\t\tПотребление {name_z!r}({zone}) = {pop_beginning:.1f} МВт изменено на {pop:.1f} МВт"
                             f" (задано {new_pop}, отклонение {abs(new_pop - pop):.1f} МВт, {i + 1} ит.)")
                return True

    def txt_task_cor(self, name: str, sel: str = '', value: str = ''):
        """
        Функция для выполнения задания в текстовом формате
        :param name: имя функции
        :param sel: выборка
        :param value: значение
        :return:
        """
        name = name.lower()
        if 'уд' in name:
            self.cor(keys=sel, tasks='del', del_all=('*' in name))
        elif 'изм' in name:
            self.cor(keys=sel, tasks=value)
        elif 'снять' in name:
            self.cor(keys='(node) (vetv) (Generator)', tasks='sel=0')
        elif 'расчет' in name:
            self.rgm(txt='txt_task_cor')
        elif 'добавить' in name:
            self.table_add_row(table=sel, tasks=value)
        elif 'текст' in name:
            self.cor_txt_field(table_field=sel)
        elif 'схн' in name:
            self.shn(choice=sel)
        elif 'ном' in name:  # номинальные напряжения
            self.voltage_nominal(choice=sel, edit=True)
        elif 'скрм' in name:
            if 'скрм*' in name:
                self.all_auto_shunt = self.auto_shunt_rec(selection=sel, only_AutoBsh=False)
            else:
                self.all_auto_shunt = self.auto_shunt_rec(selection=sel, only_AutoBsh=True)
            self.auto_shunt_cor(all_auto_shunt=self.all_auto_shunt)
        else:
            raise ValueError(f'Задание {name=} не распознано ({sel=}, {value=})')

    def test_parameter_rm_all(self, statement_all: str) -> bool:
        """
        Проверяет все утверждения и возвращает истина если все истина.
        :param statement_all: Например, 'ny=15302: vras>510|ny=15302: vras<525.5'
        :return:
        """
        statement_list = statement_all.split('|')
        for statement_i in statement_list:
            if statement_i:
                if ':' not in statement_i:
                    raise ValueError(f"Ошибка в  утверждении (нет ':'): {statement_i=}")
                sel, statement = statement_i.replace(' ', '').split(':')
                if not self.test_parameter_rm(sel, statement):
                    log_r_m.debug(f'Условие {statement_i!r} не выполняется')
                    return False
        return True

    def test_parameter_rm(self, sel: str, statement: str) -> bool:
        """
        Проверяет верность утверждения.
        :param sel: 'ny=1'
        :param statement: 'vras>125'
        :return: true или false
        """

        if not (statement and sel):
            raise ValueError(f"Ошибка в  утверждении (нет условия или выборки): {sel=}, {statement=}")
        tabl_name, tabl_choice = self.recognize_key(sel)
        parameter = ''
        value = ''
        for operator in ['=<', '<=', '=>', '>=', '=', '<', '>', '']:
            if operator in statement:
                parameter, value = statement.split(operator)
                break
        if not (parameter and value):
            raise ValueError("Ошибка в  утверждении (оператор сравнения не распознан): " + statement)

        try:
            value = float(value)
        except ValueError:
            raise ValueError("Ошибка в  утверждении (значение не число): " + statement)

        rm_val = self.rastr.Calc("val", tabl_name, parameter, tabl_choice)
        if operator in ['=<', '<=']:
            if rm_val <= value:
                return True
        elif operator in ['=>', '>=']:
            if rm_val >= value:
                return True
        elif operator == '>':
            if rm_val > value:
                return True
        elif operator == '<':
            if rm_val < value:
                return True
        elif operator == '=':
            if rm_val == value:
                return True
        return False

    def auto_shunt_rec(self, selection: str = '', only_AutoBsh: bool = True) -> dict:
        """
        Функция формирует словарь all_auto_shunt с объектами класса AutoShunt для записи СКРМ.
        :param selection: выборка в таблице узлы
        :param only_AutoBsh: True узлы только с заданным значением в поле AutoBsh. False все узлы с СКРМ
        :return словарь[ny] = namedtuple('СКРМ')
        """
        all_auto_shunt = {}
        KU = namedtuple('СКРМ', ['ny', 'name', 'ny_adjacency', 'ny_control', 'umin', 'umax',
                                 'type', ])  # KU компенсирующее устройство
        have_AutoBsh = True
        node = self.rastr.tables('node')
        vetv = self.rastr.tables('vetv')
        if node.cols.Find('AutoBsh') < 0:
            have_AutoBsh = False
            if only_AutoBsh:
                raise ValueError(f"В таблице node нет параметра AutoBsh.")
        selection_result = selection + '&pn=0&qn=0&pg=0&qg=0&bsh!=0' if selection else 'pn=0&qn=0&pg=0&qg=0&bsh!=0'
        node.setsel(selection_result)
        i = node.FindNextSel(-1)
        while i > -1:
            AutoBsh = ''
            if have_AutoBsh:
                AutoBsh = node.cols.item("AutoBsh").ZS(i)
                AutoBsh = AutoBsh.replace(' ', '')
                if not AutoBsh and only_AutoBsh:
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
            log_r_m.debug(f'Найдено СКРМ {ny=} {name=} {ny_adjacency=} {ny_control=} {umin=} {umax=}')
            i = node.FindNextSel(i)
        return all_auto_shunt

    def auto_shunt_cor(self, all_auto_shunt: dict) -> str:
        """
        Функция включает или отключает узлы с СКРМ в соответствии с уставкой по напряжению.
        :param all_auto_shunt: словарь с namedtuple('СКРМ')
        """
        changes_in_rm = ''
        if not all_auto_shunt:
            return ''
        for ny in all_auto_shunt:
            ku = all_auto_shunt[ny]
            ny_test = ku.ny_control if ku.ny_control else ku.ny_adjacency
            i = self.index_in_table(name_table='node', key=f'ny={ny}')
            i_test = self.index_in_table(name_table='node', key=f'ny={ny_test}')
            node = self.rastr.tables('node')
            sta = node.cols.item("sta").Z(i)  # 1 откл, 0 вкл
            volt_test = round(node.cols.item("vras").Z(i_test), 1)
            if volt_test:
                if volt_test < ku.umin:
                    if ku.type == 'БСК':  # включить
                        if sta:
                            self.enable_node_with_branches(ny)
                            self.rgm()
                            volt_result = round(node.cols.item("vras").Z(i_test), 1)
                            changes_in_rm += (f'\nВключена БСК {ny=} {ku.name!r},'
                                              f' напряжение увеличилось с {volt_test} до {volt_result}.')
                    elif ku.type == 'ШР':  # отключить
                        if not sta:
                            node.cols.item("sta").SetZ(i, 1)
                            self.rgm()
                            volt_result = round(node.cols.item("vras").Z(i_test), 1)
                            changes_in_rm += (f'\nОтключен ШР {ny=} {ku.name!r},'
                                              f' напряжение увеличилось с {volt_test} до {volt_result}.')
                elif volt_test > ku.umax:
                    if ku.type == 'БСК':  # отключить
                        if not sta:
                            node.cols.item("sta").SetZ(i, 1)
                            self.rgm()
                            volt_result = round(node.cols.item("vras").Z(i_test), 1)
                            changes_in_rm += (f'\nОтключена БСК {ny=} {ku.name!r},'
                                              f' напряжение снизилось с {volt_test} до {volt_result}.')
                    elif ku.type == 'ШР':  # включить
                        if sta:
                            self.enable_node_with_branches(ny)
                            self.rgm()
                            volt_result = round(node.cols.item("vras").Z(i_test), 1)
                            changes_in_rm += (f'\nВключен ШР {ny=} {ku.name!r},'
                                              f' напряжение снизилось с {volt_test} до {volt_result}.')
        log_r_m.info(changes_in_rm)
        return changes_in_rm

    def index_in_table(self, name_table: str, key: str) -> int:
        """
        Функция по ключу и имени таблицы возвращает индекс строки.
        :param name_table: например node
        :param key: например ny=100
        :return: Индекс строки в таблице. Если не найден key вернет -1. Если не найдена таблица вернет -2.
        """
        if not (name_table and key):
            raise ValueError(f'Ошибка в задании {name_table=} {key=}')
        if self.rastr.Tables.Find(name_table) == -1:
            raise ValueError (f'Таблица {name_table=} не найдена в rastr')

        table = self.rastr.tables(name_table)
        table.setsel(key)
        index = table.FindNextSel(-1)
        if index < 0:
            log_r_m.error(f'В таблице {name_table=} не найден {key=}')
        return index

    def enable_node_with_branches(self, ny: int):
        """Включить узел с ветвями."""
        if not ny:
            raise ValueError(f'Ошибка в задании {ny=}')
        self.cor(keys=str(ny), tasks='sta=0')
        vetv = self.rastr.tables('vetv')
        vetv.setsel(f'ip={ny}|iq={ny}')
        vetv.cols.item("sta").calc('0')

    def change_loading_section(self, ns: int, new_loading: float, way: str = 'pg'):
        """
        # TODO Изменить переток мощности в сечении номер ns до величины new_loading путем (way) изменения нагрузки
         ('pn') или генерации ('qn')
        :param ns:
        :param new_loading:
        :param way:
        """
        table = self.rastr.tables('sechen')