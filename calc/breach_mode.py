__all__ = ['Breach']

import logging
from abc import ABC

import pandas as pd

from collection_func import convert_s_key

log_breach = logging.getLogger(f'__main__.{__name__}')


class TypeViolationDead(ABC):
    violation = None

    def add(self):
        pass


class TypeViolation(ABC):
    violation = None

    def add(self, rm, selection: str):
        """
        Добавить нарушение режима
        :param rm:
        :param selection:
        """
        pass

    @staticmethod
    def add_dname(rm, df, table_name):
        """
        Вставить в df столбец с dname
        :param rm:
        :param df:
        :param table_name:
        """
        dname = rm.dt.t_name[table_name]
        df.insert(0,
                  'Контролируемые элементы',
                  df.s_key.apply(lambda x: dname[convert_s_key(x)]))


class Dead(TypeViolationDead):

    def __init__(self, value, name):
        """

        :param value: Значение
        :param name: Имя столбца
        """
        self._value = value
        self._name = name

    def add(self):
        self.violation = pd.DataFrame([self._value],
                                      columns=[self._name])


class OverloadsI(TypeViolation):

    def add(self,
            rm,
            selection: str):
        self.violation = rm.df_from_table(table_name='vetv',
                                          fields='s_key,'  # 'Ключ контроль,'
                                                 'txt_zag,'  # 'txt_zag,' 
                                                 'i_max,'  # 'Iрасч.(A),'
                                                 'i_dop_r,'  # 'Iддтн(A),'
                                                 'i_zag,'  # 'Iзагр.ддтн(%),'
                                                 'i_dop_r_av,'  # 'Iадтн(A),'
                                                 'i_zag_av',  # 'Iзагр.адтн(%),'
                                          setsel=selection)
        if self.violation is not None:
            self.add_dname(rm, self.violation, 'vetv')
            log_breach.info(f'Выявлено {len(self.violation)} превышений ДТН.')


class LowVoltages(TypeViolation):

    def add(self,
            rm,
            selection: str):
        self.violation = rm.df_from_table(table_name='node',
                                          fields='ny,'  # 'Ключ контроль,'
                                          # 'txt_zag,'  # 'txt_zag,'
                                          # todo сделать что бы в txt_zag были значения узлов?
                                                 'vras,'  # 'Uрасч.(кВ),'
                                                 'umin,'  # 'Uмин.доп.(кВ),'
                                                 'umin_av,'  # 'U ав.доп.(кВ),'
                                                 'otv_min,'
                                          # отклонение vras от 'Uмин.доп.' (%)
                                                 'otv_min_av',
                                          # отклонение vras от 'U ав.доп.' (%)
                                          setsel=selection)

        if self.violation is not None:
            self.violation.rename(columns={'ny': 's_key'}, inplace=True)

            self.add_dname(rm, self.violation, 'node')
            log_breach.info(f'Выявлено {len(self.violation)} точек недопустимого снижения напряжения.')


class HighVoltages(TypeViolation):

    def add(self,
            rm,
            selection: str):
        self.violation = rm.df_from_table(table_name='node',
                                          fields='ny,'  # 'Ключ контроль,'
                                                 'vras,'  # 'Uрасч.(кВ),'
                                                 'umax,'  # 'Uнаиб.раб.(кВ)'
                                                 'otv_max',  # 'Uнаиб.раб.(кВ)'
                                          setsel=selection)

        if self.violation is not None:
            self.violation.rename(columns={'ny': 's_key'}, inplace=True)
            self.add_dname(rm, self.violation, 'node')
            log_breach.info(f'Выявлено {len(self.violation)} точек недопустимого превышения напряжения.')


class Breach:
    """Нарушения режима"""

    def __init__(self):
        self.violations = {}

    def add(self, name: str, obj: TypeViolationDead | TypeViolation):
        """
        Добавить данные о нарушении режима.
        :param name:
        :param obj:
        """
        if obj.violation is not None:
            self.violations[name] = obj.violation

    def yes(self):
        """Имеет место нарушение режима """

        return True if self.violations else False
