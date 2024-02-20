"""Модуль загрузки сечения """
import logging

log_ls = logging.getLogger('__main__.' + __name__)


def loading_section(rm,
                    ns: int,
                    p_new: float | str,
                    method: str = 'pg',
                    ignore_pmin: bool = True,
                    max_cycle: int = 30,
                    accuracy: float = 0.5,
                    dr_p_set: float = 0.01) -> bool:
    """
    Изменить переток мощности в сечении номер ns до величины p_new за счет изменения нагрузки('pn') или
    генерации ('qn') в отмеченных узлах и генераторах.
    :param rm: RastrModel
    :param ns: Номер сечения
    :param p_new: Требуемая мощность в сечении
    :param method: Изменения нагрузки 'pn' или генерации 'pg'.
    :param ignore_pmin: Поле Pmin генераторов и pg_min узлов не учитывается если True.
    :param max_cycle: Максимальное количество циклов.
    :param accuracy: Точность задания мощности сечения, но не превышает заданную.
    :param dr_p_set: Начальная величина реакции в узле.
    :return: True если заданный переток в сечении достигнут.
    """
    choice_node = 'sel&!sta'
    log_ls.info(f'\tИзменить переток мощности в сечении {ns=}, {p_new=}, {method=}.')
    if rm.rastr.tables.Find("sechen") == -1:
        rm.downloading_additional_files('sch')
    try:
        i_ns = rm.index(table_name='sechen', key_int=ns)
    except ValueError:
        raise ValueError(f'Сечение {ns=} отсутствует в файле сечений.')

    grline = rm.rastr.Tables("grline")
    section_tab = rm.rastr.tables('sechen')
    name_ns = section_tab.cols('name').ZS(i_ns)
    if p_new in ['pmax', 'pmin']:
        p_new = section_tab.cols(p_new).Z(i_ns)

    try:
        p_new = float(p_new)
    except ValueError:
        raise ValueError(f'Заданная величина перетока мощности не распознано {p_new!r}')

    if not p_new:
        p_new = 0.01
    rm.rastr.sensiv_start("")

    grline.SetSel(f'ns={ns}')
    data_grline = grline.writesafearray('ip,iq', "000")
    for ip, iq in data_grline:
        rm.rastr.sensiv_back(4, 1., ip, iq, 0)

    rm.rastr.sensiv_write("")
    rm.rastr.sensiv_end()

    node = rm.rastr.tables("node")
    node.SetSel(choice_node)
    select_add = "tip>1 &" if method == 'pg' else ''
    # Сумма реакций в узле.
    db = abs(rm.rastr.Calc("sum", "node", "dr_p", f"{select_add}!sta & (abs(dr_p)>{dr_p_set}) & dr_p>0"))
    db += abs(rm.rastr.Calc("sum", "node", "dr_p", f"{select_add}!sta & (abs(dr_p)>{dr_p_set}) & dr_p<0"))
    if db < dr_p_set:
        log_ls.error("Невозможно изменить переток мощности в сечении.")
        return False

    set_gen = []
    if method == 'pg':
        if rm.rastr.Tables("Generator").cols.Find("sel") < 0:
            rm.rastr.Tables("Generator").Cols.Add('sel', 3)

        # Анализ генераторов и узлов РМ.
        gen_sect = []  # Выборка отмеченных генераторов и генераторов в отмеченных узлах.
        node_sel = set([x[0] for x in node.writesafearray('ny', "000")])
        gen_all = rm.rastr.Tables("Generator").writesafearray('Node,Num,sel', "000")
        node_with_gen = set()
        for index, (node_i, num, sel) in enumerate(gen_all):
            node_with_gen.add(node_i)
            if sel or (node_i in node_sel):  # отмечен узел или генератор
                gen_sect.append((num, index))
        # Выборка отмеченных узлов без генераторов.
        node_sect = [node_i for node_i in node_sel if node_i not in node_with_gen]

        for g, i in gen_sect:
            set_gen.append(Gen(rm, key=g, i=i, ignore_pmin=ignore_pmin))
        for n in node_sect:
            set_gen.append(NodeGen(rm, key=n,
                                   i=rm.index(table_name='node', key_int=n),
                                   ignore_pmin=ignore_pmin))

    p_section = section_tab.cols('psech')
    for cycle in range(1, max_cycle):
        p_sect_cur = p_section.Z(i_ns)
        if cycle == 1:
            log_ls.info(f'\tНачальный переток мощности в сечении {name_ns}: {p_sect_cur:.2f}.')
        change_p = p_new - p_sect_cur
        log_ls.debug(f'\t{cycle=}, {p_sect_cur=:.2f}, {p_new=}, {change_p=:.2f} МВт.')

        if abs(change_p) < accuracy:
            log_ls.info(f'\tЗаданная точность достигнута psech={p_sect_cur:.2f},'
                        f' отклонение {change_p:.2f} МВт, количество итераций: {cycle - 1}.')
            return True

        if method == 'pn':
            p_sum = sum((x[0] for x in node.writesafearray('pn', "000")))
            dr_p_sum = sum((x[0] for x in node.writesafearray('dr_p', "000")))
            if not p_sum:
                log_ls.error('Изменение мощности сечения: сумма нагрузки узлов равна 0')
                return False
            if dr_p_sum < 0:
                coefficient = 1 + (1 - (p_sum - change_p) / p_sum)
            else:
                coefficient = (p_sum - change_p) / p_sum
            node.cols("pn").Calc(f"pn*({coefficient})")
            node.cols("qn").Calc(f"qn*({coefficient})")

        elif method == 'pg':
            section_up_sum = 0
            section_down_sum = 0
            for gen in set_gen:
                gen.reserve_p()
                gen.direction = ''
                if change_p * gen.info['dr_p'] > 0:
                    if gen.reserve_p_up:
                        section_up_sum += gen.reserve_p_up
                        gen.direction = 'up'
                elif change_p * gen.info['dr_p'] < 0:
                    if gen.reserve_p_down:
                        section_down_sum += gen.reserve_p_down
                        gen.direction = 'down'

            sum_reserve = section_up_sum + section_down_sum

            if not sum_reserve:
                log_ls.error(f'Исчерпан резерв мощности.')
                return False

            # На сколько МВт нужно снизить Р.
            reduce_p = min(abs(section_down_sum / sum_reserve * change_p), section_down_sum)
            # На сколько МВт нужно увеличить Р.
            increase_p = min(abs(abs(change_p) - reduce_p), section_up_sum)

            if (section_down_sum + section_up_sum) < change_p:
                log_ls.info("Генерации не хватает")
            # На сколько нужно умножить резерв Рген и прибавить к резерву Рген, для снижения генерации.
            koef_p_down = (1 - (section_down_sum - reduce_p) / section_down_sum) if section_down_sum else 0
            # На сколько нужно умножить резерв Рген и прибавить его к резерву Рген, для увеличения генерации.
            koef_p_up = (1 - (section_up_sum - increase_p) / section_up_sum) if section_up_sum else 0

            for gen in set_gen:
                match gen.direction:
                    case "up":
                        gen.ratio = koef_p_up
                    case "down":
                        gen.ratio = koef_p_down
                gen.change_generation()
        rm.rgm('loading_section')

    p_sect_cur = p_section.Z(i_ns)
    log_ls.info(f'Заданная точность не достигнута P={p_sect_cur:.2f}, '
                f'отклонение {p_new - p_sect_cur:.2f}.')
    return False


class ObjectGeneration:
    """Для описания общих свойств генераторов и генерации в узле."""
    ignore_pmin = None  # Не учитывать Pmin
    # dr_p_koeff = 0 # если 1, то умножаем дополнительно на dr_p в этом случае больше загружаются
    # # генераторы которые меньше влияют на изменение мощности в сечении
    reserve_p_up = 0  # Резерв Р на увеличение генерации.
    reserve_p_down = 0  # Резерв Р на снижение генерации.
    pg = None  # Текущее значение генерации.
    direction = ''  # Нужно увеличивать 'up' или уменьшать генерацию 'down'.
    sta = 0  # 1 отключен, 0 включен.
    ratio = 0  # Соотношение для получения требуемой мощности генерации.
    info = None
    _rm = None
    _pg_name = None  # 'pg' или 'P'
    table_name = None  # 'node' или 'Generator'
    _i = None  # Индекс в таблице rastr
    _key_name = None  # 'Num' or 'ny'
    key = None  # Значение Num или ny

    def __init__(self, rm, key: int, i: int, ignore_pmin: bool):
        self.ignore_pmin = ignore_pmin
        self.key = key
        self._i = i
        self._rm = rm

    def reserve_p(self):
        self.sta = self._rm.rastr.Calc("val",
                                       self.table_name,
                                       "sta",
                                       f"{self._key_name}={self.key}")
        if self.sta:  # Отключен > 0.
            self.pg = 0
            self.reserve_p_up = self.info['pg_max']
            self.reserve_p_down = 0
        else:  # Включен 0.
            self.pg = self._rm.rastr.Calc("val",
                                          self.table_name,
                                          self._pg_name,
                                          f"{self._key_name}={self.key}")
            self.reserve_p_up = self.info['pg_max'] - self.pg
            self.reserve_p_down = self.pg

    def change_generation(self):

        deviation_pg = None  # На сколько нужно изменить генерацию в узле
        pg_new = None
        match self.direction:
            case 'up':
                deviation_pg = self.ratio * self.reserve_p_up
                pg_new = self.pg + deviation_pg
            case 'down':
                deviation_pg = self.ratio * self.pg
                pg_new = self.pg - deviation_pg

        if not deviation_pg:
            return False

        pg_new = self._pg_correction(pg_new)
        # Включить если отключен.
        if pg_new and self.sta:
            self.sta = 0
            self._rm.rastr.tables(self.table_name).cols.Item('sta').SetZ(self._i, self.sta)
        self._rm.rastr.tables(self.table_name).cols.Item(self._pg_name).SetZ(self._i, pg_new)

    def _pg_correction(self, pg_new: float) -> float:
        if pg_new > self.info['pg_max']:
            pg_new = self.info['pg_max']
        if pg_new < 0:
            pg_new = 0
        if 0 < pg_new < self.info['pg_min']:
            pg_new = self.pg
        return pg_new


class Gen(ObjectGeneration):
    table_name = 'Generator'
    _pg_name = 'P'
    _key_name = 'Num'
    """Генератор для изменения мощности в сечении."""

    def __init__(self, rm, key: int, i: int, ignore_pmin: bool):
        super().__init__(rm, key, i, ignore_pmin)
        self.info = self._rm.df_from_table(table_name=self.table_name,
                                           fields='Name,Pmin,Pmax,Num,Node',
                                           setsel=f'{self._key_name}={self.key}')
        self.info.rename(columns={'Name': 'name',
                                  'Node': 'ny',
                                  'Pmin': 'pg_min',
                                  'Pmax': 'pg_max'},
                         inplace=True)
        self.info = self.info.loc[0].to_dict()

        self.info['dr_p'] = self._rm.rastr.Calc("val", "node", "dr_p", f"ny={self.info['ny']}")
        if not self.info['pg_max']:
            raise ValueError(f"В генераторе [{self.info['Num']}] {self.info['name']} не задано поле [Pmax].")

    def __str__(self):
        return f'g={self.key} {self.info["name"]}'


class NodeGen(ObjectGeneration):
    """Узел для изменения мощности в сечении."""
    table_name = 'node'
    _pg_name = 'pg'
    _key_name = 'ny'

    def __init__(self, rm, key: int, i: int, ignore_pmin: bool):
        super().__init__(rm, key, i, ignore_pmin)

        self.info = self._rm.df_from_table(table_name=self.table_name,
                                           fields='name,pg_min,pg_max,ny,dr_p',
                                           setsel=f'{self._key_name}={self.key}'
                                           ).loc[0].to_dict()
        self.info['dr_p'] = self._rm.rastr.Calc("val", "node", "dr_p", f"ny={self.info['ny']}")
        if not self.info['pg_max']:
            raise ValueError(f"В узле [{self.info['ny']}] {self.info['name']} не задано поле [pg_max].")

    def __str__(self):
        return f'ny={self.key} {self.info["name"]}'
