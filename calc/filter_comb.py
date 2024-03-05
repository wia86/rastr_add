__all__ = ['FilterCombination']

import logging

import pandas as pd

log_filter_comb = logging.getLogger(f'__main__.{__name__}')


class FilterCombination:
    """ Класс для отсеивания сочетаний н-2 и н-3 из анализа н-1"""

    # dict[((ip, iq, np), disable_scheme, comb.repair_scheme)] = set((ip, iq, np), ...)
    # отключаемая ветвь: (ветви из списка отключаемых ветвей, загрузка которых изменяется)
    _disable_effect = {}

    def __init__(self, diff: int | float, df_ib_norm: pd.DataFrame):
        """
        :param diff: Значение в % выше которого считается, что ветви оказывают взаимное влияние.
        :param df_ib_norm: Загрузка контролируемых элементов сети в нормальном режиме
        col: [key - 'ip= &iq= np= ', ib - ток начала ветви, i_dop - допустимый ток].
        """
        self._diff = diff
        self._df_ib_norm = df_ib_norm
        self._df_ib_norm.rename(columns={'ib': 'ib_norm'}, inplace=True)
        self.count_false_comb = 0

    def add_n1(self, rm, comb, df_ip_n1: pd.DataFrame):
        """
        Проверить изменение загрузки отключаемых элементов при нормативном возмущении (НВ).
        :param rm:
        :param comb:
        :param df_ip_n1: Токовая загрузка отключаемых элементов сети при НВ.
        """
        df_diff_i = self._df_ib_norm.merge(df_ip_n1)
        df_diff_i['change_loading'] = df_diff_i.apply(lambda x:
                                                      self._calc_change(x['ib_norm'], x['ib'], x['i_dop']), axis=1)

        change_loading_branch = df_diff_i.loc[(df_diff_i.change_loading > self._diff) & (df_diff_i.ib > 0)]
        change_loading_branch = set(change_loading_branch['s_key'])
        key_branch = (comb.s_key[0],
                      comb.disable_scheme[0],
                      comb.repair_scheme[0])

        disable_num_transit = rm.v__num_transit.get(comb.s_key[0])

        # Если нет анализа транзитов или отключаемая ветвь в ненумерованном транзите.
        if (not len(rm.v__num_transit)) or (not disable_num_transit):
            self._disable_effect[key_branch] = change_loading_branch
            return

        # Добавить если принадлежат разным транзитам
        self._disable_effect[key_branch] = set()
        for branch in change_loading_branch:
            if rm.v__num_transit.get(branch) != disable_num_transit:
                self._disable_effect[key_branch].add(branch)

    @staticmethod
    def _calc_change(x1: float, x2: float, x_dop: float) -> float:
        """ Функция выполняет расчет разности двух значений относительно допустимого значения в %."""
        if not x_dop:
            x_dop = 100
        return abs(x1 - x2) / x_dop * 100

    def test_comb(self, comb: pd.DataFrame) -> bool:
        """
        Проверка взаимного влияния элементов в комбинации.
        :param comb:
        :return: True если все элементы взаимно влияют на загрузку друг друга.
        """

        if len(comb.loc[(comb.table == 'node') | comb.double_repair_scheme]):
            return True

        all_key = tuple(comb['s_key'])
        for i in comb[['s_key', 'disable_scheme', 'repair_scheme']].itertuples(index=False, name=None):
            for key in all_key:
                if i[0] != key:  # Отбираем ключи прочих ветвей.
                    if key not in self._disable_effect.get(i, []):
                        self.count_false_comb += 1
                        return False
        return True
