import logging
import pandas as pd
from tabulate import tabulate

import collection_func as cf
log_auto = logging.getLogger(f'__main__.{__name__}')


class Automation:
    """
    Моделирование действия ПА
    """

    def __init__(self, rm):
        """
        Выполняется один раз при загрузки РМ
        :param rm: RastrModel
        """
        self.current_time = 0  # Отсчет времени с момента начала СРС, увеличивается по ходу срабатывания ПА
        log_auto.debug('Инициализация автоматики')
        self.n_action = {}
        self.df_pa = pd.DataFrame()
        self.exist = False  # Наличие автоматики
        self.all_num = set()  # Все номера в таблице automation
        self.all_num_device = {}  # {Номер ПА: устройство ПА}
        self.all_num_auto = set()  # Используемые номера automation в таблице automation
        self.all_num_test = set()  # Номера в таблице automation отмеченные test - проверять всегда
        self.num_activation = set()  # Номера активированных ПА.
        # Активируется если:
        # - недопустимое отклонение UI в таблице узлы, ветви
        # - ПА step=1, test=1, условие выполняется.
        # В противном случае номер ПА исключается (если ранее был активирован).

        # загрузка таблицы automation
        if rm.rastr.tables.Find('automation') > -1:
            if rm.rastr.tables('automation').count:
                self.exist = True
                self.df_pa = rm.df_from_table(table_name='automation')
                if rm.rastr.tables.Find('automation_pattern') > -1:
                    df_automation_pattern = rm.df_from_table(table_name='automation_pattern')
                    df_automation_pattern['name'] = df_automation_pattern['name'].str.strip()

                    df_automation_pattern.set_index('name', inplace=True)
                    dict_name_action = df_automation_pattern.to_dict()['pattern']

                    self.df_pa.replace({"action": dict_name_action}, inplace=True)
                    self.df_pa.replace({"condition": dict_name_action}, inplace=True)
                    self.df_pa = self.df_pa[(self.df_pa['sta'] == 0)]
                    self.df_pa.loc[self.df_pa['step'] == 0, 'step'] = 1
                    self.all_num = set(self.df_pa['n'].unique())
                    self.all_num_test = set(self.df_pa.loc[self.df_pa['test'] == 1, 'n'].unique())
                    self.df_pa['active_time'] = 0
            else:
                log_auto.info('Таблица automation пуста')
                # return
        else:
            log_auto.info('Таблица automation не найдена')
            # return

            # Анализ задания автоматики из таблиц node и vetv
        df = pd.DataFrame(columns=['n',
                                   'test',
                                   'name',
                                   'step',
                                   'time',
                                   'set_point',
                                   'action',
                                   'condition',
                                   'sta'])
        n_new_pa = max(self.all_num) if self.all_num else 0
        fields = 'automation,repair_scheme,disable_scheme,double_repair_scheme'
        for name_t in ('node', 'vetv', 'Generator'):
            t = rm.rastr.tables(name_t)
            d = t.writesafearray(f'{fields},{t.Key}', "000")
            for automation, repair_scheme, disable_scheme, double_repair_scheme, *s_key in d:
                if len(s_key) == 1:
                    s_key = s_key[0]
                else:
                    s_key = tuple(s_key)

                for name_fields, z in (('repair_scheme', repair_scheme),
                                       ('disable_scheme', disable_scheme),
                                       ('automation', automation),
                                       ('double_repair_scheme', double_repair_scheme)):
                    z = z.split('#')[0]
                    if z:
                        z_list = cf.split_task_action(z)
                        for i, z_list_i in enumerate(z_list):
                            if not z_list_i.replace('.', '').isdigit():
                                if '[' in z_list_i:
                                    name = z_list_i.split('[')[0]
                                    condition = z_list_i.split('{')[1].split('}')[0] if '{' in z_list_i else ''
                                    action = z_list_i.split('[')[1].split(']')[0]
                                    n_new_pa += 1
                                    self.all_num.add(n_new_pa)
                                    z_list[i] = str(n_new_pa)
                                    df.loc[len(df.index)] = [n_new_pa,  # 'n',
                                                             0,  # 'test',
                                                             name,
                                                             1,  # 'step',
                                                             0,  # 'time',
                                                             '',  # 'set_point',
                                                             action.replace(' ', ''),
                                                             condition.replace(' ', ''),
                                                             0]  # 'sta']
                        rm.t_scheme[name_t][name_fields][s_key] = tuple(z_list)
                        if name_fields == 'automation':
                            self.all_num_auto = self.all_num_auto | set(z_list)

        if len(df):
            if len(self.df_pa):
                self.df_pa = pd.concat([self.df_pa, df])
            else:
                self.df_pa = df

        if len(self.df_pa):
            for n in self.all_num:
                self.all_num_device[n] = AutoDevice(n=n, data=self.df_pa[self.df_pa['n'] == n])
            log_auto.debug(tabulate(self.df_pa, headers='keys', tablefmt='psql'))
        # self.df_pa['active'] = False

    def reset(self):
        """
        Сброс активации ПА.
        Выполняется в случае окончания рассмотрения СРС с действием ПА.
        """
        self.current_time = 0
        self.num_activation = set()

    def scheme_description(self, number: str) -> tuple:
        """
        По номеру n в таблице automation возвращает строки задания в текстовом виде
        :param number: "Номер_ПА.номер_ступени"
        :return: (list(название из таблицы automation), list(задание из той же таблицы))
        """
        if number in self.n_action:
            return self.n_action[number][0], self.n_action[number][1]

        names = []
        tasks = []
        if '.' in number:
            n, step = number.split('.')
        else:
            n = number
            step = -1
        n = int(n)
        step = int(step)
        cut = self.df_pa[self.df_pa['n'] == n]
        if step > -1:
            cut = cut[cut['step'] == step]

        if not len(cut):
            raise ValueError(f'В таблице automation отсутствует запись с номером {number!r}')

        if cut['action'].all():

            for i in cut.index:
                task = f"[{cut.loc[i, 'action']}]"
                if cut.loc[i, 'condition']:
                    task += f"{{{cut.loc[i, 'condition']}}}"

                tasks.append(task)
                names.append(cut.loc[i, 'name'])

            self.n_action[number] = (names, tasks)
            return names, tasks
        else:
            raise ValueError(f'В таблице automation в записи с номером {number!r} отсутствует описание действия.')

    def execute_action_pa(self, rm, df_init: pd.DataFrame) -> str:
        # не забыть про restore_only_state
        pass

    def active(self, rm, overloads):
        """
        Отметка и снятие отметки колонки df_pa[active] активной автоматики
        в промежутке между циклами действия ПА.
        :param rm:
        :param overloads:
        :return: True если есть ПА для действия
        """
        # Проверка таблиц узлы и ветви
        all_found_automatics = set()  # все номера активируемых автоматик
        if len(overloads):
            for r in overloads.itertuple():
                if r.i_max:
                    # Ветвь
                    s_key = tuple(r.s_key.split(','))
                    auto = rm.t_scheme['vetv']['automation'].get(s_key, False)
                    if auto:
                        for a in auto:
                            # если "a" уже запущена, то проверяем текущую ступень, если нет, то минимальную
                            all_found_automatics.add(a)
                            # if a not in self.num_activation:
                            self.num_activation.add(a)
                else:
                    auto = rm.t_scheme['node']['automation'].get(r.s_key, False)
                    if auto:
                        for a in auto:
                            all_found_automatics.add(a)
                            # if a not in self.num_activation:
                            self.num_activation.add(a)

        # Проверка всегда тестируемой ПА: если condition мин ступени истина, то активируется вся ПА с тем же номером
        if self.all_num_test:
            for num_test in self.all_num_test:
                # self.num_activation
                device = self.all_num_device[num_test]
                test_add = False
                for t_c in device.test_condition:
                    if rm.conditions_test(t_c):
                        all_found_automatics.add(num_test)
                        self.num_activation.add(num_test)
                        test_add = True
                        break
                if not test_add:
                    self.num_activation.remove(num_test)
                    device.time_active = 0
                    # todo проверить что device в all_num_device меняется

        # отсеять ПА которая не актуальна
        for n in self.num_activation:
            if n not in all_found_automatics:
                self.num_activation.remove(n)

        return True if self.num_activation else False


class AutoDevice:
    """Моделирование устройства ПА"""

    def __init__(self, n: int,
                 data: pd.DataFrame):
        self.n = n
        self.data = data
        self.data.index = range(len(self.data))
        # log_auto.debug(tabulate(data, headers='keys', tablefmt='psql'))
        self.time_active = 0

        self.test = True if data.loc[0, 'test'] else False  # тестируется всегда

        self.all_step = sorted(list(self.data['step'].unique()))
        self.step_active = min(self.all_step)  # увеличивается по ходу действия ПА
        self.all_time = sorted(list(self.data['time'].unique()))

    def test_condition(self, rm, overload):
        """Проверка условия выполнения активной ступени"""
        # res = True  # все условия выполняются
        test_condition = self.data.loc[self.data['step'] == self.step_active, 'condition'].to_list()
        for cond in test_condition:
            if cond:
                if not rm.conditions_test(cond):
                    return False
        # проверим уставку
        set_point = self.data.loc[self.data['step'] == self.step_active, 'set_point'].to_list()
        for s_p in set_point:
            if s_p:
                if overload.i_max:
                    pass

                if overload.umin:
                    pass
                    # Снижение напряжения в узле

                elif overload.umax:
                    pass
                    # Повышение напряжения в узле

    def reset(self):
        """
        Сброс настроек устройства на начальные
        """
        self.step_active = min(self.all_step)
        self.time_active = 0
