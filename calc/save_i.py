__all__ = ['SaveI']

import os
import sqlite3

import pandas as pd

from collection_func import s_key_vetv_in_tuple


class SaveI:
    """Класс для хранения токовой загрузки"""
    # Для хранения токовой загрузки контролируемых элементов в пределах одной РМ
    _save_i_rm = None
    path_db = None

    def init_for_rm(self):
        self._save_i_rm = pd.DataFrame()

    def add_data(self,
                 rm,
                 comb_id: int,
                 active_id: int,
                 setsel: str):
        data = rm.df_from_table(table_name='vetv',
                                fields='s_key,'  # 'Ключ контроль,'
                                       'i_max,'  # 'Iрасч.(A),'
                                       'i_dop_r,'  # 'Iддтн(A),'
                                       'i_zag,'  # 'Iзагр.ддтн(%),'
                                       'i_dop_r_av,'  # 'Iадтн(A),'
                                       'i_zag_av',  # 'Iзагр.адтн(%),'
                                setsel=setsel)
        data['comb_id'] = comb_id
        data['active_id'] = active_id
        self._save_i_rm = pd.concat([self._save_i_rm, data],
                                    axis=0)

    def end_for_rm(self, rm,
                   path_db: str):
        """Сохранить в db токовую загрузку элементов сети для текущей РМ"""
        self.path_db = path_db

        self._save_i_rm.insert(0,
                               'Контролируемые элементы',
                               self._save_i_rm.s_key.apply(
                                   lambda x: rm.dt.t_name['vetv'][s_key_vetv_in_tuple(x)]))

        con = sqlite3.connect(self.path_db)
        self._save_i_rm.to_sql('save_i',
                               con,
                               if_exists='append')
        con.commit()
        con.close()

    def max_i_to_xl(self, path_xl: str):
        """Данные токовой загрузки из db сгруппировать по элементам и сохранить в xl"""

        con = sqlite3.connect(self.path_db)

        group_i_max = pd.read_sql_query("""
        SELECT s_key, 
        "Контролируемые элементы", 
        "Год", 
        "Сезон макс/мин", 
        "Темп.(°C)", 
        "Кол. откл. эл.", 
        count(*) AS "Кол.СРС", 
        "Наименование СРС", 
        max(i_max) AS "Iрасч.,A", 
        i_dop_r AS "Iддтн,А", 
        i_zag AS "Iзагр. ддтн,%", 
        i_dop_r_av AS "Iадтн,А", 
        i_zag_av AS "Iзагр. адтн,%"
        FROM (
        SELECT *
        FROM save_i AS si
           INNER JOIN all_actions AS aa
              ON si.comb_id = aa.comb_id AND si.active_id = aa.active_id
           INNER JOIN all_comb AS ac
              ON ac.comb_id = aa.comb_id
           INNER JOIN all_rm AS ar
              ON ar.rm_id = ac.rm_id
        )
        GROUP BY s_key, "Год", "Сезон макс/мин", "Темп.(°C)", "Кол. откл. эл.";
        """, con)

        con.commit()
        con.close()

        # Запись в xl данных о максимальной токовой нагрузке

        mode = 'a' if os.path.exists(path_xl) else 'w'
        with pd.ExcelWriter(path=path_xl,
                            mode=mode) as writer:
            group_i_max.to_excel(excel_writer=writer,
                                 float_format='%.2f',
                                 index=False,
                                 freeze_panes=(1, 1),
                                 sheet_name='Максимальные токи')
