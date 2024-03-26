__all__ = ['DataCalc']

import pandas as pd

from collection_func import save_to_sqlite


class DataCalc:

    def __init__(self):
        self.all_comb = pd.DataFrame()
        self.all_actions = pd.DataFrame()  # Действие оперативного персонала или ПА

    def add_all_comb(self, info_srs: dict):
        self.all_comb = pd.concat([self.all_comb,
                                   pd.Series(info_srs).to_frame().T],
                                  axis=0,
                                  ignore_index=True)

    def add_all_actions(self, info_action: dict):
        self.all_actions = pd.concat([self.all_actions,
                                      pd.Series(info_action).to_frame().T],
                                     axis=0,
                                     ignore_index=True)

    def save_to_sql(self, path_db: str):

        dict_df = {'all_comb': self.all_comb,
                   'all_actions': self.all_actions
                   }

        save_to_sqlite(path_db=path_db,
                       dict_df=dict_df)

