__all__ = ['BreachStorage']

import logging
import os
from collections import defaultdict, namedtuple

import pandas as pd
import win32com.client

log_breach_storage = logging.getLogger(f'__main__.{__name__}')


class BreachStorage:
    """Хранение нарушений режимов"""

    collection = None

    def __init__(self):
        self.collection_rm = defaultdict(pd.DataFrame)
        self.collection_all = defaultdict(pd.DataFrame)

    def save_data_active(self,
                         breach,
                         comb_id: int,
                         active_id: int):
        """
        Сохранить перегрузки для текущей комбинации и действия
        :param breach: Объект Breach
        :param comb_id:
        :param active_id:
        """
        if breach.violations:
            for key, val in breach.violations.items():

                if val is not None:
                    val['comb_id'] = comb_id
                    val['active_id'] = active_id

                    self.collection_rm[key] = pd.concat([self.collection_rm[key], val],
                                                        axis=0,
                                                        ignore_index=True)

    def save_data_rm(self):
        """Добавить Контролируемые элементы.
        Перенести данные в общую коллекцию"""
        for key in self.collection_rm:
            self.collection_all[key] = pd.concat([self.collection_all[key], self.collection_rm[key]],
                                                 axis=0,
                                                 ignore_index=True)

        self.collection_rm = defaultdict(pd.DataFrame)

    def save_to_sql(self, con):
        for key in self.collection_all:
            self.collection_all[key].to_sql(key,
                                            con,
                                            if_exists='replace')

    def save_to_xl(self,
                   book_path: str,
                   all_rm: pd.DataFrame,
                   all_comb: pd.DataFrame,
                   all_actions: pd.DataFrame):
        log_breach_storage.debug(f'Запись параметров режима в excel.')

        for key in self.collection_all:
            self.collection_all[key] = (all_rm.merge(all_comb)
                                        .merge(all_actions)
                                        .merge(self.collection_all[key]))

            # for col in ['Отключение', 'Ремонт 1', 'Ремонт 2', 'Доп. имя']:
            #     for col_df in self.collection_all[key].columns:
            #         if col in col_df:
            #             self.collection_all[key].fillna(value={col_df: 0},
            #                                             inplace=True)
            #             self.collection_all[key].loc[self.collection_all[key][col_df] == 0, col_df] = '-'

            mode = 'a' if os.path.exists(book_path) else 'w'
            with pd.ExcelWriter(path=book_path, mode=mode) as writer:
                # todo если много элементов то будет ошибка
                self.collection_all[key].to_excel(excel_writer=writer,
                                                  float_format='%.2f',
                                                  index=False,
                                                  freeze_panes=(1, 1),
                                                  sheet_name=key)

        # self.add_pivot_tables(book_path=book_path,
        #                       sheets_name=list(self.collection_all.keys()))

    @staticmethod
    def add_pivot_tables(book_path, sheets_name: list):
        # Сводная.
        if not sheets_name:
            log_breach_storage.info('Отклонений параметров режима от допустимых значений не выявлено.')
            return

        log_breach_storage.info(f'Формирование сводных таблиц.')
        excel = win32com.client.Dispatch('Excel.Application')
        excel.ScreenUpdating = False  # Обновление экрана
        try:
            book = excel.Workbooks.Open(book_path)
        except Exception:
            raise Exception(f'Ошибка при открытии файла {book_path=}')

        task_pt = {
            'dead': {},
            'i': {},
            'low_u': {},
            'high_u': {}
        }

        for sheet_name in sheets_name:
            try:
                sheet = book.sheets[sheet_name]
            except Exception:
                raise Exception(f'Не найден лист: {sheet_name}')

            # Создать объект таблица из всего диапазона листа.
            obj_table = sheet.ListObjects.Add(SourceType=1,
                                              Source=sheet.Range(sheet.UsedRange.address))
            obj_table.Name = f'Таблица_{sheet_name}'
            obj_table_columns = get_list_columns(obj_table)

            # Создать КЭШ xlDatabase, ListObjects
            pt_cache = book.PivotCaches().add(1, obj_table)
            task_pivot = namedtuple('task_pivot',
                                    ['sheet_name', 'pivot_table_name', 'data_field'])
            task_pivot_cur = None
            if sheet_name == 'i':
                task_pivot_cur = task_pivot('Сводная_I', 'Свод_I',
                                            dict(i_max='Iрасч.,A',
                                                 i_dop_r='Iддтн,A',
                                                 i_zag='Iзагр. ддтн,%',
                                                 i_dop_r_av='Iадтн,A',
                                                 i_zag_av='Iзагр. адтн,%'))
            elif sheet_name == 'low_u':
                task_pivot_cur = task_pivot('Сводная_Umin', 'Свод_Umin',
                                            dict(vras='Uр, кВ',
                                                 umin='МДН, кВ',
                                                 otv_min='Uр, % от МДН',
                                                 umin_av='АДН,кВ',
                                                 otv_min_av='Uр, % от АДН'))
            elif sheet_name == 'high_u':
                task_pivot_cur = task_pivot('Сводная_Umax', 'Свод_Umax',
                                            dict(vras='Uр, кВ',
                                                 umax='Uнр, кВ',
                                                 otv_max='Uр, % от Uнр'))
            elif sheet_name == 'crash':
                task_pivot_cur = task_pivot('Сводная_не_сходятся', 'Свод_crash',
                                            dict(alive='Режим не сошелся'))
        #
        #     row_fields = [col for col in ['Контролируемые элементы', 'Отключение', 'Ремонт 1', 'Ремонт 2'] if col in full_breach[sheet_name].columns]
        #
        #     column_fields = (['Год', 'Сезон макс/мин'] + [col for col in full_breach[sheet_name].columns if 'Доп. имя' in col])
        #
        #     sheet_pivot = book.Sheets.Add()
        #     sheet_pivot.Name = task_pivot_cur.sheet_name
        #
        #     pt = pt_cache.CreatePivotTable(TableDestination=task_pivot_cur.sheet_name + '!R1C1',
        #                                    TableName=task_pivot_cur.pivot_table_name)
        #     pt.ManualUpdate = True  # True не обновить сводную
        #     pt.AddFields(row_fields=row_fields,
        #                  column_fields=column_fields,
        #                  PageFields=['Имя файла', 'Кол. откл. эл.', 'End', 'Наименование СРС', 'Контроль ДТН',
        #                              'Темп.(°C)'],
        #                  AddToTable=False)
        #     for field_df, field_pt in task_pivot_cur.data_field.items():
        #         pt.AddDataField(Field=pt.PivotFields(field_df),
        #                         Caption=field_pt,
        #                         Function=-4136)  # xlMax -4136 xlSum -4157
        #         pt.PivotFields(field_pt).NumberFormat = '0'
        #
        #     if len(task_pivot_cur.data_field) > 1:
        #         pt.PivotFields('Контролируемые элементы').ShowDetail = True  # группировка
        #     pt.RowAxisLayout(1)  # 1 xlTabularRow показывать в табличной форме!!!!
        #     if len(task_pivot_cur.data_field) > 1:
        #         pt.DataPivotField.Orientation = 1  # xlRowField = 1 'Значения' в столбцах или строках xlColumnField
        #     pt.RowGrand = False  # Удалить строку общих итогов
        #     pt.ColumnGrand = False  # Удалить столбец общих итогов
        #     pt.MergeLabels = True  # Объединять одинаковые ячейки
        #     pt.HasAutoFormat = False  # Не обновлять ширину при обновлении
        #     pt.NullString = '--'  # Заменять пустые ячейки
        #     pt.PreserveFormatting = False  # Сохранять формат ячеек при обновлении
        #     pt.ShowDrillIndicators = False  # Показывать кнопки свертывания
        #     for row in row_fields + column_fields:
        #         pt.PivotFields(row).Subtotals = [False, False, False, False, False, False, False, False, False,
        #                                          False, False, False]  # промежуточные итоги и фильтры
        #     if len(task_pivot_cur.data_field) > 1:
        #         field = list(task_pivot_cur.data_field)[2]
        #         pt.PivotFields(field).Orientation = 3  # xlPageField = 3
        #         pt.PivotFields(field).CurrentPage = '(All)'  #
        #     pt.TableStyle2 = ""  # стиль
        #     pt.ColumnRange.ColumnWidth = 10  # ширина строк
        #     pt.RowRange.ColumnWidth = 20
        #     pt.DataBodyRange.HorizontalAlignment = -4108  # xlCenter = -4108
        #     pt.TableRange1.WrapText = True  # перенос текста в ячейке
        #     for i in range(7, 13):
        #         pt.TableRange1.Borders(i).LineStyle = 1  # лево
        #     # Условное форматирование
        #     if sheet_name != 'crash':
        #         for i in range(3, len(task_pivot_cur.data_field) + 1, 2):
        #             dpz = pt.DataBodyRange.Rows(i).Cells(1)
        #             dpz.FormatConditions.AddColorScale(2)  # ColorScaleType:=2
        #             dpz.FormatConditions(dpz.FormatConditions.count).SetFirstPriority()
        #             dpz.FormatConditions(1).ColorScaleCriteria(1).Type = 0  # xlConditionValueNumber = 0
        #             if list(task_pivot_cur.data_field)[2] == 'i_zag':
        #                 dpz.FormatConditions(1).ColorScaleCriteria(1).Value = 100
        #             else:
        #                 dpz.FormatConditions(1).ColorScaleCriteria(1).Value = 0
        #             dpz.FormatConditions(1).ColorScaleCriteria(1).FormatColor.ThemeColor = 1  # xlThemeColorDark1=1
        #             dpz.FormatConditions(1).ColorScaleCriteria(1).FormatColor.TintAndShade = 0
        #             dpz.FormatConditions(1).ColorScaleCriteria(2).Type = 2  # xlConditionValueHighestValue = 2
        #             dpz.FormatConditions(1).ColorScaleCriteria(2).FormatColor.ThemeColor = 3 + i  # номер темы
        #             dpz.FormatConditions(1).ColorScaleCriteria(2).FormatColor.TintAndShade = -0.249977111117893
        #             dpz.FormatConditions(1).ScopeType = 2  # xlDataFieldScope = 2 применить ко всем значениям поля
        #     pt.ManualUpdate = False  # обновить сводную
        # book.Save()
        # book.Close()
        # excel.Quit()


def get_list_columns(obj_table):
    count_columns = obj_table.ListColumns.Count
    list_columns = []
    for i in count_columns:
        list_columns.append(obj_table.ListColumns(i).Name)
