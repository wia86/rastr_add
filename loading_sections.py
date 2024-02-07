"""Модуль загрузки сечения """
import logging
from typing import Union

log_ls = logging.getLogger('__main__.' + __name__)

def loading_section(self, ns: int, p_new: Union[float, str], type_correction: str = 'pg'):
    """
    Изменить переток мощности в сечении номер ns до величины p_new за счет изменения нагрузки('pn') или
    генерации ('qn') в отмеченных узлах и генераторах
    :param ns: номер сечения
    :param p_new:
    :param type_correction:  'pn' изменения нагрузки или 'pg' генерации
    """
    # --------------настройки----------
    choice = 'sel&!sta'
    max_cycle = 30  # максимальное количество циклов
    accuracy = 0.05  # процент, точность задания мощности сечения, но не превышает заданную
    dr_p_zad = 0.01  # величина реакции начальная

    log_ls.info(f'\tИзменить переток мощности в сечении {ns=}: P={p_new}, выборка: {choice}, тип: {type_correction}.')
    if self.rastr.tables.Find("sechen") == -1:
        self.downloading_additional_files(['sch'])

    index_ns = self.index(table_name='sechen', key_int=ns)
    if index_ns == -1:
        raise ValueError(f'Сечение {ns=} отсутствует в файле сечений.')
    grline = self.rastr.Tables("grline")
    sechen = self.rastr.tables('sechen')
    name_ns = sechen.cols('name').ZS(index_ns)
    if p_new in ['pmax', 'pmin']:
        p_new = sechen.cols(p_new).Z(index_ns)
    try:
        p_new = float(p_new)
    except ValueError:
        raise ValueError(f'Заданная величина перетока мощности не распознано {p_new!r}')
    if not p_new:
        p_new = 0.01
    p_current = round(self.rastr.Calc("sum", "sechen", "psech", f"ns={ns}"), 2)
    log_ls.info(f'\tТекущий переток мощности в сечении {name_ns!r}: {p_current}.')

    self.rastr.sensiv_start("")
    grline.SetSel(f'ns={ns}')
    index_grline = grline.FindNextSel(-1)
    while not index_grline == -1:
        self.rastr.sensiv_back(4, 1., grline.Cols("ip").Z(index_grline), grline.Cols("iq").Z(index_grline), 0)
        index_grline = grline.FindNextSel(index_grline)

    self.rastr.sensiv_write("")
    self.rastr.sensiv_end()

    node = self.rastr.tables("node")
    change_p = round(p_new - p_current, 2)
    db = 0
    # 'pn'
    p_sum = 0
    dr_p_sum = 0
    # 'pg'
    node_all = {}

    if type_correction == 'pn':  # изменение нагрузки

        choice_dr_p = f"!sta & abs(dr_p) > {dr_p_zad}"  # !sta вкл
        db = abs(self.rastr.Calc("sum", "node", "dr_p", choice_dr_p + "&dr_p>0"))
        db += abs(self.rastr.Calc("sum", "node", "dr_p", choice_dr_p + "&dr_p<0"))

    elif type_correction == 'pg':

        if self.rastr.Tables("Generator").cols.Find("sel") < 0:
            log_ls.info('В таблицу Generator добавляется отсутствующее поле sel')
            self.rastr.Tables("Generator").Cols.Add('sel', 3)

        # Доотметить узлы и генераторы которые нужно корректировать
        # отметить генераторы у отмеченных узлов
        node.SetSel("sel")
        i = node.FindNextSel(-1)
        while i >= 0:
            self.group_cor(tabl="Generator", param="sel", selection=f"Node={node.cols('ny').ZS(i)}", formula="1")
            i = node.FindNextSel(i)
        # отметить узлы у отмеченных генераторов
        generators = self.rastr.tables("Generator")
        generators.SetSel("sel")
        i = generators.FindNextSel(-1)
        while i >= 0:
            self.group_cor(tabl="node", param="sel", selection=f"ny={generators.cols('Node').ZS(i)}", formula="1")
            i = generators.FindNextSel(i)
        choice_dr_p = f"tip>1 &!sta & abs(dr_p) > {dr_p_zad}"  # tip>1 ген   !sta вкл
        db = abs(self.rastr.Calc("sum", "node", "dr_p", choice_dr_p + "&dr_p>0"))
        db += abs(self.rastr.Calc("sum", "node", "dr_p", choice_dr_p + "&dr_p<0"))

        node.SetSel(choice)
        i = node.FindNextSel(-1)

        while i >= 0:
            nd = NodeGeneration(rastr=self.rastr, i=i)
            node_all[node.cols("ny").Z(i)] = nd
            i = node.FindNextSel(i)

    if db < dr_p_zad:
        log_ls.error("Невозможно изменить мощность по сечению (с учетом отмеченных узлов и/или генераторов)")
        return False

    for cycle in range(max_cycle):

        p_current = round(self.rastr.Calc("sum", "sechen", "psech", f"ns={ns}"), 2)
        change_p = round(p_new - p_current, 2)
        log_ls.debug(f'\t{cycle=}, {p_current=}, {p_new=}, {change_p=} МВт ({round(abs(change_p / p_new) * 100)} %)')

        if abs(change_p / p_new) * 100 < accuracy:
            if (p_current < p_new and p_new > 0) or (p_current > p_new and p_new < 0):
                log_ls.info(f'\tЗаданная точность достигнута P={p_current},'
                         f' отклонение {change_p}. {cycle + 1} итераций')
                break

        # изменение нагрузки
        if type_correction == 'pn':
            node.SetSel(choice)
            i = node.FindNextSel(-1)
            while not i == -1:
                p_sum += node.cols("pn").Z(i)
                dr_p_sum += node.cols("dr_p").Z(i)
                i = node.FindNextSel(i)
            if not p_sum:
                log_ls.error('Изменение мощности сечения: сумма нагрузки узлов равна 0')
                break
            if dr_p_sum < 0:
                coefficient = 1 + (1 - (p_sum - change_p) / p_sum)
            else:
                coefficient = (p_sum - change_p) / p_sum
            node.cols("pn").Calc(f"pn*({coefficient})")
            node.cols("qn").Calc(f"qn*({coefficient})")

        # изменение генерации
        elif type_correction == 'pg':
            NodeGeneration.change_p = change_p
            section_up_sum = 0
            section_down_sum = 0
            for nd in node_all:
                if nd.use:
                    nd.reserve_p()
                    if change_p * nd.dr_p > 0:
                        if nd.reserve_p_up:
                            section_up_sum += nd.reserve_p_up
                            nd.up_pgen = True
                    elif change_p * nd.dr_p < 0:
                        if nd.reserve_p_down:
                            section_down_sum += nd.reserve_p_down
                            nd.up_pgen = False
            log_ls.debug(f'')
            if not (section_up_sum and section_down_sum):
                log_ls.error(f'Не удалось добиться заданной точности в сечении')

            # на сколько МВт нужно снизить Р
            reduce_p = abs(section_down_sum / (section_down_sum + section_up_sum) * change_p)
            if section_down_sum < reduce_p:
                reduce_p = section_down_sum
            # на сколько МВт нужно увеличить Р
            increase_p = abs(abs(change_p) - reduce_p)
            if section_up_sum < increase_p:
                increase_p = section_up_sum

            if (section_down_sum + section_up_sum) < change_p:
                log_ls.info("Генерации не хватает")
            # Коэффициент на сколько нужно умножить резерв Рген и прибавить к резерву Рген, для снижения генерации
            koef_p_down = 0
            # Коэффициент: на сколько нужно умножить резерв Рген и прибавить его к резерву Рген,
            # для увеличения генерации
            koef_p_up = 0
            if section_down_sum:
                koef_p_down = 1 - (section_down_sum - reduce_p) / section_down_sum
            if section_up_sum:
                koef_p_up = 1 - (section_up_sum - increase_p) / section_up_sum

            for nd in node_all:
                if nd.use:
                    nd.change(koef_p_down=koef_p_down, koef_p_up=koef_p_up)

        self.rgm('loading_section')
    else:
        log_ls.info(f'Заданная точность не достигнута P={p_current}, отклонение {change_p}.')


class NodeGeneration:
    """Класс для хранения информации об узле для изменения мощности в сечении."""
    dr_p_koeff = 0  # если 1, то умножаем дополнительно на dr_p в этом случае больше загружаются
    # генераторы которые меньше влияют на изменение мощности в сечении

    no_pmin = True  # ' не учитывать Pmin
    abs_change_p = None  # todo что это?
    change_p = 0
    unbalance_p = 0

    def __init__(self, i: int, rastr):
        """
        :param i: Индекс в таблице узлы
        :param rastr:
        """
        self.gen_available = False  # Узел с генераторами
        self.use = True
        self.up_pgen = True
        self.reserve_p_up = 0
        self.reserve_p_down = 0
        self.rastr = rastr
        self.i = i
        self.node_t = self.rastr.tables("node")
        gen_t = self.rastr.tables("Generator")
        self.ny = self.node_t.Cols("ny").Z(self.i)
        self.dr_p = self.node_t.Cols("dr_p").Z(self.i)
        self.gen_all = {}
        dr_p = self.node_t.Cols("dr_p").Z(self.i)
        self.name = self.node_t.Cols("name").ZS(self.i)
        txt = f'\t\tУзел {self.ny}: {self.name}'
        gen_t.SetSel(f"Node={self.ny}")
        if gen_t.count:
            gen_t.SetSel(f"Node={self.ny}&sel")  # все генераторы дб отмечены, если не отмечен то не используем
            i = gen_t.FindNextSel(-1)
            while i >= 0:  # ЦИКЛ ген
                self.gen_available = True  # узел с генераторами
                gen = Gen(rastr=self.rastr, i=i)
                self.gen_all[gen.Num] = gen
                i = gen_t.FindNextSel(i)
        else:
            self.pg_max = self.node_t.Cols("pg_max").Z(self.i)
            self.pg_min = self.node_t.Cols("pg_min").Z(self.i)

    def reserve_p(self):
        self.reserve_p_up = 0
        self.reserve_p_down = 0
        if self.gen_available:
            for gen in self.gen_all:
                if gen.use:
                    gen.reserve_p()
                    self.reserve_p_up += gen.reserve_p_up
                    self.reserve_p_down += gen.reserve_p_down
        else:
            if self.pg_max:
                self.reserve_p_up = self.pg_max - self.node_t.Cols("pg").Z(self.i)
            else:
                log_ls.info(f"в узле {self.ny} {self.name} не задано поле pg_max")
            self.reserve_p_down = self.node_t.Cols("pg").Z(self.i)

    def change(self, koef_p_down: float = 0, koef_p_up: float = 0):
        # --------------настройки----------
        change_p = abs(NodeGeneration.abs_change_p)
        unbalance_p = NodeGeneration.unbalance_p
        pg_node = self.node_t.Cols("pg").Z(self.i)

        if self.up_pgen:
            deviation_pg = koef_p_up * self.reserve_p_up  # На сколько нужно изменить генерацию в узле
        else:
            deviation_pg = pg_node * koef_p_down

        if not deviation_pg:
            return False

        if unbalance_p > 0:
            if unbalance_p > deviation_pg:
                unbalance_p = unbalance_p - deviation_pg
                deviation_pg = 0
            if unbalance_p < deviation_pg:
                deviation_pg = deviation_pg - unbalance_p
                unbalance_p = 0

        if not self.gen_available:  # нет генераторов
            if self.up_pgen:  # увеличиваем генерацию узла, koef_p_up
                if self.pg_max and self.pg_max > pg_node:
                    if self.pg_min > pg_node + deviation_pg:  # (от 0 до pg_min)
                        if self.pg_min and not self.no_pmin:  # если есть Рмин и учитываем Рмин то
                            if change_p > self.pg_min:
                                self.node_t.cols.Item("pg").SetZ(self.i, self.pg_min)
                                # unbalance_p = unbalance_p + (self.pg_min - deviation_pg)
                                change_p = change_p - self.pg_min
                        else:  # нет Рмин или не учитываем Рмин
                            self.node_t.cols.Item("pg").SetZ(self.i, pg_node + deviation_pg)
                            change_p = change_p - deviation_pg
                    elif self.pg_max > pg_node + deviation_pg and (
                            self.pg_min < pg_node + deviation_pg or self.pg_min == pg_node + deviation_pg):
                        # (от pg_min (включительно) до pg_max)v
                        self.node_t.cols.Item("pg").SetZ(self.i, pg_node + deviation_pg)
                        change_p = change_p - deviation_pg
                    elif self.pg_max < pg_node + deviation_pg or self.pg_max == pg_node + deviation_pg:
                        # (больше или равно pg_max)
                        self.node_t.cols.Item("pg").SetZ(self.i, self.pg_max)
                        change_p = change_p - (self.pg_max - pg_node)

            else:  # снижаем генерацию узла,KefPG_Down
                if self.pg_min < pg_node - deviation_pg or self.pg_min == pg_node - deviation_pg:
                    # (от pg_min (включительно) до pg_node)
                    self.node_t.cols.Item("pg").SetZ(self.i, pg_node - deviation_pg)
                    change_p = change_p - deviation_pg

                elif self.pg_min > pg_node - deviation_pg and (pg_node - deviation_pg) > 0:  # (от 0 до pg_min)
                    if self.pg_min > 0 and not self.no_pmin:  # если есть Рмин и учитываем Рмин то
                        self.node_t.cols.Item("pg").SetZ(self.i, self.pg_min)
                        change_p = change_p - (pg_node - self.pg_min)
                        deviation_pg = deviation_pg - (pg_node - self.pg_min)
                        if change_p > self.pg_min:
                            self.node_t.cols.Item("sta").SetZ(self.i, True)
                            # unbalance_p = unbalance_p + (self.pg_min - deviation_pg)
                            change_p = change_p - self.pg_min

                    else:  # если Рмин не учитываем
                        self.node_t.cols.Item("pg").SetZ(self.i, pg_node - deviation_pg)
                        change_p = change_p - deviation_pg

                elif pg_node - deviation_pg < 0 or pg_node == deviation_pg:  # (меньше или равно 0)
                    self.node_t.cols.Item("pg").SetZ(self.i, 0)
                    change_p = change_p - pg_node



class Gen:
    """Класс для хранения информации о генераторах в узле для изменения мощности в сечении."""

    def __init__(self, i: int, rastr):
        self.reserve_p_up = 0
        self.reserve_p_down = 0
        self.use = True
        self.rastr = rastr
        self.i = i
        self.gen_t = self.rastr.tables("Generator")
        self.Num = self.gen_t.Cols("Num").Z(self.i)
        self.gen_name = self.gen_t.Cols("Name").Z(self.i)
        self.Pmax = self.gen_t.Cols("Pmax").Z(self.i)
        if not self.Pmax:
            log_ls.debug(f"У генератора {self.Num!r}: {self.gen_name!r}  не задано Pmax")
        self.Pmin = self.gen_t.Cols("Pmin").Z(self.i)

    def reserve_p(self):
        if self.gen_t.Cols("sta").Z(self.i):
            self.reserve_p_up = self.Pmax
        else:  # если генератор включен
            self.reserve_p_down = self.gen_t.Cols("P").Z(self.i)
            if self.Pmax:
                self.reserve_p_up = self.Pmax - self.gen_t.Cols("P").Z(self.i)

