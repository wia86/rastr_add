import logging
import os
import sqlite3
from collections import namedtuple
from datetime import datetime
from itertools import combinations

import pandas as pd
import win32com.client
from tabulate import tabulate

from calc import Automation
from calc import CombinationXL
from calc import FillTable
from calc import FilterCombination
from calc import Drawings
from calc import SaveI
from common import Common
from rastr_model import RastrModel

log_calc = logging.getLogger(f'__main__.{__name__}')


class CalcModel(Common):
    """
    Расчет нормативных возмущений.
    """
    fill_table = None
    drawings = None
    save_i = None
    mark = 'calc'

    def __init__(self, config: dict):
        """
        :param config: Задание и настройки программы
        """
        super(CalcModel, self).__init__()
        self.config = self.config | config
        RastrModel.config = config['Settings']

        RastrModel.overwrite_new_file = 'question'
        self.info_srs = None  # СРС
        self.comb_id = 1
        self.set_comb = None  # {количество отключений: контроль ДТН, 1:'ДДТН',2:'АДТН'}
        # self.auto_shunt = {}

        self.restore_only_state = True

        self.book_path: str = ''  # Путь к файлу excel.
        self.book_db: str = ''  # Путь к файлу db.
        self.pa = None  # Объект Automation
        self.task_full_name = ''  # Путь к файлу задания rg2.

        self.disable_df_vetv = pd.DataFrame()

        if self.config['results_RG2']:
            self.drawings = Drawings(name_drawing=self.config['name_pic'])

        self.info_action = None
        RastrModel.all_rm = pd.DataFrame()
        self.all_comb = pd.DataFrame()
        self.all_actions = pd.DataFrame()  # действие оперативного персонала или ПА

        self.breach = {'i': pd.DataFrame(), 'high_u': pd.DataFrame(), 'low_u': pd.DataFrame()}

        # Для хранения токовой загрузки контролируемых элементов.
        if self.config['cb_save_i']:
            self.save_i = SaveI()

        self.combination_xl = None  # Для создания объекта класса CombinationXL.
        self.filter_comb = None  # Для создания объекта класса FilterCombination.

    def run(self):
        """
        Запуск расчета нормативных возмущений (НВ) в РМ.
        """
        log_calc.info('Запуск расчета нормативных возмущений (НВ) в расчетной модели (РМ).\n')
        self.run_common()

        if self.config['cb_disable_excel']:  # Отключаемые элементы сети по excel.
            self.combination_xl = CombinationXL(path_xl=self.config['srs_XL_path'],
                                                sheet=self.config['srs_XL_sheets'])
        # Цикл, если несколько файлов задания.
        if self.config['CB_Import_Rg2'] and os.path.isdir(self.config['Import_file']):
            task_files = os.listdir(self.config['Import_file'])
            task_files = list(filter(lambda x: x.endswith('.rg2'), task_files))
            for task_file in task_files:  # цикл по файлам '.rg2' в папке
                self.task_full_name = os.path.join(self.config['Import_file'], task_file)
                log_calc.info(f'Текущий файл задания: {self.task_full_name}\n')
                self.run_task()
                self.config['name_time'] = os.path.join(self.folder_result,
                                                        datetime.now().strftime(self.time_str_format))
        else:
            if self.config['CB_Import_Rg2']:
                self.task_full_name = self.config['Import_file']
                self.run_task()
            else:
                self.run_task()
        if self.filter_comb:
            self.config['filter_comb_info'] = (f'Рассчитано комбинаций: {self.comb_id}, '
                                               f'отфильтровано {self.filter_comb.count_false_comb} ')

        return self.the_end()

    def run_task(self):
        """
        Запуск расчета с текущим файлом импорта задания или без него.
        """

        # todo ссылки ниже убрать от сюда
        self.book_path = self.config['name_time'] + ' результаты расчетов.xlsx'
        self.book_db = self.config['name_time'] + ' данные.db'

        # Папка с вложенными папками
        if self.size_date_source == 'nested_folder':
            for address, dir_, file_ in os.walk(self.source_path):
                self.cycle_rm(path_folder=address)

        # Папка без вложенных папок
        elif self.size_date_source == 'folder':
            self.cycle_rm(path_folder=self.source_path)

        # один файл
        elif self.size_date_source == 'file':
            rm = RastrModel(self.source_path)
            if not rm.code_name_rg2:
                raise ValueError(f'Имя файла {self.source_path!r} не подходит.')
            self.run_file(rm=rm)

        self.processing_results()

    def processing_results(self):

        con = sqlite3.connect(self.book_db)

        # Записать данные о перегрузках в SQL.
        for key in self.breach:
            self.breach[key].to_sql(key, con, if_exists='replace')

        # Записать данные о выполненных расчетах в SQL.
        name_df = {'all_rm': RastrModel.all_rm,
                   'all_comb': self.all_comb,
                   'all_actions': self.all_actions,
                   }
        if self.drawings:
            name_df['all_drawings'] = self.drawings.df_drawing

        for key in name_df:
            name_df[key].to_sql(key, con, if_exists='replace')

        con.commit()
        con.close()

        # Считать из SQL данные токовой загрузки и запись в xl.
        if self.save_i:
            self.save_i.max_i_to_xl(book_path=f'{self.config["name_time"]} Imax.xlsx')

        # Вставить таблицы К-О в word.
        if self.fill_table:
            self.fill_table.insert_word()

        # Вставить таблицы c перечнем рисунков в xl.
        if self.drawings:
            self.drawings.add_to_xl(book_path=f'{self.config["name_time"]} рисунки.xlsx')
            self.drawings.add_macro(macro_path=f'{self.config["other"]["path_project"]}'
                                               f'\help\Сделать рисунки в word.rbs')

        log_calc.debug(f'Запись параметров режима в excel.')
        # Запись данных о перегрузке в xl
        full_breach = {}
        for key in self.breach:
            if len(self.breach[key]):
                full_breach[key] = (RastrModel.all_rm.merge(self.all_comb)
                                                     .merge(self.all_actions)
                                                     .merge(self.breach[key]))

                for col in ['Отключение', 'Ремонт 1', 'Ремонт 2', 'Доп. имя']:
                    for col_df in full_breach[key].columns:
                        if col in col_df:
                            full_breach[key].fillna(value={col_df: 0},
                                                    inplace=True)
                            full_breach[key].loc[full_breach[key][col_df] == 0, col_df] = '-'

                mode = 'a' if os.path.exists(self.book_path) else 'w'
                with pd.ExcelWriter(path=self.book_path, mode=mode) as writer:
                    # todo если много элементов в full_breach то будет ошибка
                    full_breach[key].to_excel(excel_writer=writer,
                                              float_format='%.2f',
                                              index=False,
                                              freeze_panes=(1, 1),
                                              sheet_name=key)
        # Запись в xl данных о СРС в которых режим развалился
        crash = self.all_actions[self.all_actions.alive == 0]
        if len(crash):
            full_breach['crash'] = (RastrModel.all_rm.merge(self.all_comb)
                                                     .merge(self.all_actions[self.all_actions.alive == 0]))
            mode = 'a' if os.path.exists(self.book_path) else 'w'
            with (pd.ExcelWriter(path=self.book_path, mode=mode) as writer):
                full_breach['crash'].to_excel(excel_writer=writer,
                                              float_format='%.2f',
                                              index=False,
                                              freeze_panes=(1, 1),
                                              sheet_name='crash')

        # Сводная.
        if len(full_breach):
            log_calc.info(f'Формирование сводных таблиц ({self.book_path}).')
            excel = win32com.client.Dispatch('Excel.Application')
            excel.ScreenUpdating = False  # Обновление экрана
            try:
                book = excel.Workbooks.Open(self.book_path)
            except Exception:
                raise Exception(f'Ошибка при открытии файла {self.book_path=}')

            for key in full_breach:
                try:
                    sheet = book.sheets[key]
                except Exception:
                    raise Exception(f'Не найден лист: {key}')
                # Создать объект таблица из всего диапазона листа.

                tabl_overload = sheet.ListObjects.Add(SourceType=1, Source=sheet.Range(sheet.UsedRange.address))
                tabl_overload.Name = f'Таблица_{key}'
                pt_cache = book.PivotCaches().add(1, tabl_overload)  # Создать КЭШ xlDatabase, ListObjects

                task_pivot = namedtuple('task_pivot',
                                        ['sheet_name', 'pivot_table_name', 'data_field'])
                task_pivot_cur = None
                if key == 'i':
                    task_pivot_cur = task_pivot('Сводная_I', 'Свод_I',
                                                dict(i_max='Iрасч.,A',
                                                     i_dop_r='Iддтн,A',
                                                     i_zag='Iзагр. ддтн,%',
                                                     i_dop_r_av='Iадтн,A',
                                                     i_zag_av='Iзагр. адтн,%'))
                elif key == 'low_u':
                    task_pivot_cur = task_pivot('Сводная_Umin', 'Свод_Umin',
                                                dict(vras='Uр, кВ',
                                                     umin='МДН, кВ',
                                                     otv_min='Uр, % от МДН',
                                                     umin_av='АДН,кВ',
                                                     otv_min_av='Uр, % от АДН'))
                elif key == 'high_u':
                    task_pivot_cur = task_pivot('Сводная_Umax', 'Свод_Umax',
                                                dict(vras='Uр, кВ',
                                                     umax='Uнр, кВ',
                                                     otv_max='Uр, % от Uнр'))
                elif key == 'crash':
                    task_pivot_cur = task_pivot('Сводная_не_сходятся', 'Свод_crash',
                                                dict(alive='Режим не сошелся'))

                RowFields = [col for col in ['Контролируемые элементы', 'Отключение', 'Ремонт 1', 'Ремонт 2']
                             if col in full_breach[key].columns]

                ColumnFields = (['Год', 'Сезон макс/мин'] +
                                [col for col in full_breach[key].columns if 'Доп. имя' in col])

                sheet_pivot = book.Sheets.Add()
                sheet_pivot.Name = task_pivot_cur.sheet_name

                pt = pt_cache.CreatePivotTable(TableDestination=task_pivot_cur.sheet_name + '!R1C1',
                                               TableName=task_pivot_cur.pivot_table_name)
                pt.ManualUpdate = True  # True не обновить сводную
                pt.AddFields(RowFields=RowFields,
                             ColumnFields=ColumnFields,
                             PageFields=['Имя файла', 'Кол. откл. эл.', 'End', 'Наименование СРС', 'Контроль ДТН',
                                         'Темп.(°C)'],
                             AddToTable=False)
                for field_df, field_pt in task_pivot_cur.data_field.items():
                    pt.AddDataField(Field=pt.PivotFields(field_df),
                                    Caption=field_pt,
                                    Function=-4136)  # xlMax -4136 xlSum -4157
                    pt.PivotFields(field_pt).NumberFormat = '0'

                if len(task_pivot_cur.data_field) > 1:
                    pt.PivotFields('Контролируемые элементы').ShowDetail = True  # группировка
                pt.RowAxisLayout(1)  # 1 xlTabularRow показывать в табличной форме!!!!
                if len(task_pivot_cur.data_field) > 1:
                    pt.DataPivotField.Orientation = 1  # xlRowField = 1 'Значения' в столбцах или строках xlColumnField
                pt.RowGrand = False  # Удалить строку общих итогов
                pt.ColumnGrand = False  # Удалить столбец общих итогов
                pt.MergeLabels = True  # Объединять одинаковые ячейки
                pt.HasAutoFormat = False  # Не обновлять ширину при обновлении
                pt.NullString = '--'  # Заменять пустые ячейки
                pt.PreserveFormatting = False  # Сохранять формат ячеек при обновлении
                pt.ShowDrillIndicators = False  # Показывать кнопки свертывания
                for row in RowFields + ColumnFields:
                    pt.PivotFields(row).Subtotals = [False, False, False, False, False, False, False, False, False,
                                                     False, False, False]  # промежуточные итоги и фильтры
                if len(task_pivot_cur.data_field) > 1:
                    field = list(task_pivot_cur.data_field)[2]
                    pt.PivotFields(field).Orientation = 3  # xlPageField = 3
                    pt.PivotFields(field).CurrentPage = '(All)'  #
                pt.TableStyle2 = ""  # стиль
                pt.ColumnRange.ColumnWidth = 10  # ширина строк
                pt.RowRange.ColumnWidth = 20
                pt.DataBodyRange.HorizontalAlignment = -4108  # xlCenter = -4108
                pt.TableRange1.WrapText = True  # перенос текста в ячейке
                for i in range(7, 13):
                    pt.TableRange1.Borders(i).LineStyle = 1  # лево
                # Условное форматирование
                if key != 'crash':
                    for i in range(3, len(task_pivot_cur.data_field) + 1, 2):
                        dpz = pt.DataBodyRange.Rows(i).Cells(1)
                        dpz.FormatConditions.AddColorScale(2)  # ColorScaleType:=2
                        dpz.FormatConditions(dpz.FormatConditions.count).SetFirstPriority()
                        dpz.FormatConditions(1).ColorScaleCriteria(1).Type = 0  # xlConditionValueNumber = 0
                        if list(task_pivot_cur.data_field)[2] == 'i_zag':
                            dpz.FormatConditions(1).ColorScaleCriteria(1).Value = 100
                        else:
                            dpz.FormatConditions(1).ColorScaleCriteria(1).Value = 0
                        dpz.FormatConditions(1).ColorScaleCriteria(1).FormatColor.ThemeColor = 1  # xlThemeColorDark1=1
                        dpz.FormatConditions(1).ColorScaleCriteria(1).FormatColor.TintAndShade = 0
                        dpz.FormatConditions(1).ColorScaleCriteria(2).Type = 2  # xlConditionValueHighestValue = 2
                        dpz.FormatConditions(1).ColorScaleCriteria(2).FormatColor.ThemeColor = 3 + i  # номер темы
                        dpz.FormatConditions(1).ColorScaleCriteria(2).FormatColor.TintAndShade = -0.249977111117893
                        dpz.FormatConditions(1).ScopeType = 2  # xlDataFieldScope = 2 применить ко всем значениям поля
                pt.ManualUpdate = False  # обновить сводную
            book.Save()
            book.Close()
            excel.Quit()
        else:
            log_calc.info('Отклонений параметров режима от допустимых значений не выявлено.')


    def cycle_rm(self, path_folder: str):
        """Цикл по файлам"""

        gen_files = (f for f in os.listdir(path_folder) if f.endswith('.rg2'))

        for rastr_file in gen_files:  # цикл по файлам '.rg2' в папке

            if self.config['filter_file'] and self.file_count == self.config['max_count_file']:
                break  # Если включен фильтр файлов проверяем количество расчетных файлов.
            full_name = os.path.join(path_folder, rastr_file)
            rm = RastrModel(full_name)
            # Если включен фильтр файлов и имя стандартизовано
            if not rm.code_name_rg2:
                log_calc.info(f'Имя файл {full_name} не распознано.')
                continue
            if self.config['filter_file']:
                if not rm.test_name(condition=self.config['criterion'],
                                    info=f'Имя файла {full_name} не подходит.'):
                    continue  # Пропустить, если не соответствует фильтру
            log_calc.info('\n\n')
            self.run_file(rm)

    def run_file(self, rm):
        """
        Рассчитать РМ.
        """
        if self.save_i:
            self.save_i.init_for_rm()

        self.set_comb = {}  # {количество отключений: контроль ДТН, 1:'ДДТН',2:'АДТН'}
        self.file_count += 1

        rm.load()

        if self.config['cor_rm']['add']:
            rm.cor_rm_from_txt(self.config['cor_rm']['txt'])

        # Импорт из РМ c ИД.
        if self.task_full_name:
            # task_full_name: полный путь к текущему файлу задания (rg2)
            # 'таблица: node, vetv; тип: 2; поле: disable_scheme, automation; выборка: sel'
            for row in self.config['txt_Import_Rg2'].split('\n'):
                row = row.replace(' ', '').split('#')[0]  # удалить текст после '#'
                if row:
                    rm.txt_import_rm(type_import='файл', description=f'({self.task_full_name});{row}')

        rm.rastr.CalcIdop(rm.info_file['Темп.(°C)'], 0.0, "")
        log_calc.info(f'Выполнен расчет ДТН для температуры: {rm.info_file["Темп.(°C)"]} °C.')

        # Подготовка.
        rm.voltage_fix_frame()
        # if self.config['CalcSetWindow']['skrm']:
        #     self.auto_shunt = rm.auto_shunt_rec(selection='')

        # В поле all_disable складываем элементы авто отмеченные и отмеченные в поле comb_field
        rm.add_fields_in_table(name_tables='vetv,node,Generator', fields='all_disable', type_fields=3)

        if self.config['CalcSetWindow']['pa']:
            self.pa = Automation(rm)
            if not self.pa.exist:
                self.config['CalcSetWindow']['pa'] = False

        # Сохранить текущее состояние РМ
        rm.dt.save_date_tables()

        # Контролируемые элементы сети.
        if self.config['cb_control']:
            log_calc.debug('Инициализация контролируемых элементов сети.')
            # all_control для отметки всех контролируемых узлов и ветвей (авто и field)

            if self.config['cb_control_field']:
                log_calc.debug(f'Отмеченные в поле control элементы добавлены в контролируемые.')
                for table_name in ['vetv', 'node']:
                    rm.group_cor(tabl=table_name,
                                 param='all_control',
                                 selection='control',
                                 formula='1')

            if self.config['cb_control_sel']:
                control_sel = self.config['control_sel'].replace(' ', '')
                log_calc.debug(f'Элементы по выборке {control_sel}  добавлены в контролируемые.')
                if control_sel:
                    table = rm.rastr.tables('node')
                    table.setsel(control_sel)
                    if not table.count:
                        raise ValueError(f'По выборке {control_sel} не найдены узлы в РМ.')
                    ny_sel = [x[0] for x in table.writesafearray('ny', '000')]
                    sel_v = set()
                    for ny in ny_sel:
                        for ip, iq, np_ in rm.dt.ny_join_vetv[ny]:
                            sel_v.add((ip, iq, np_, 1))
                    rm.rastr.tables('vetv').ReadSafeArray(2, 'ip,iq,np,all_control', list(sel_v))
                    rm.rastr.tables('node').ReadSafeArray(2, 'ny,all_control', [(ny, 1) for ny in ny_sel])

                else:  # Контролировать все узлы и ветви.
                    rm.rastr.Tables('node').cols.item('all_control').Calc('1')
                    rm.rastr.Tables('vetv').cols.item('all_control').Calc('1')

            # Поля для сортировки ветвей и др.
            rm.add_fields_in_table(name_tables='vetv', fields='temp,temp1', type_fields=1)
            rm.rastr.tables('vetv').cols.item('temp').calc('ip.uhom')
            rm.rastr.tables('vetv').cols.item('temp1').calc('iq.uhom')

            if self.config['cb_tab_KO']:
                self.fill_table = FillTable(rm=rm,
                                            setsel='all_control')
            else:  # Для отметки всех контролируемых ветвей и ветвей с теми же groupid
                log_calc.info('Добавление в контролируемые элементы ветвей по groupid.')
                table = rm.rastr.tables('vetv')
                table.setsel('all_control & groupid>0')
                if table.count:
                    for gr in set(table.writesafearray('groupid', '000')):
                        rm.group_cor(tabl='vetv',
                                     param='all_control',
                                     selection=f'groupid={gr[0]}',
                                     formula=1)

        # Нормальная схема сети
        self.info_srs = dict()  # СРС
        self.info_srs['Наименование СРС'] = 'Нормальная схема сети.'
        self.info_srs['Наименование СРС без()'] = 'Нормальная схема сети'
        self.info_srs['comb_id'] = self.comb_id
        self.info_srs['Кол. откл. эл.'] = 0
        self.info_srs['Контроль ДТН'] = 'ДДТН'
        self.info_srs['rm_id'] = RastrModel.rm_id
        log_calc.info(f"Сочетание {self.comb_id}: {self.info_srs['Наименование СРС']}")
        self.do_action(rm)

        # Отключаемые элементы сети. Расчет всех возможных сочетаний.
        if self.config['cb_disable_comb']:

            # Выбор количества одновременно отключаемых элементов
            # н-1
            if self.config['SRS']['n-1']:
                self.set_comb[1] = 'ДДТН'
            # н-2
            if self.config['CalcSetWindow']['gost']:
                if self.config['SRS']['n-2_abv'] and rm.gost_abv:
                    self.set_comb[2] = 'AДТН'
                if self.config['SRS']['n-2_gd'] and rm.gost_gd:
                    self.set_comb[2] = 'ДДТН'
            else:
                if self.config['SRS']['n-2_abv'] or self.config['SRS']['n-2_gd']:
                    self.set_comb[2] = 'ДДТН'
            # н-3
            if self.config['SRS']['n-3']:
                if self.config['CalcSetWindow']['gost']:
                    if rm.gost_gd:
                        self.set_comb[3] = 'АДТН'
                else:
                    self.set_comb[3] = 'ДДТН'
            log_calc.info(f'Расчетные СРС: {self.set_comb}.')

            # Выбор отключаемых элементов автоматически по выборке в таблице узлы
            if self.config['cb_auto_disable']:
                rm.network_analysis(field='all_disable',
                                    selection_node_for_disable=self.config['auto_disable_choice'])
            else:
                rm.network_analysis(disable_on=False)

            # Выбор отключаемых элементов из отмеченных в поле disable
            if self.config['cb_comb_field']:
                for table_name in ['vetv', 'node', 'Generator']:
                    rm.group_cor(tabl=table_name,
                                 param='all_disable',
                                 selection='disable',
                                 formula='1')

            # Создать df отключаемых узлов и ветвей и генераторов. Сортировка.
            # Генераторы
            disable_df_gen = rm.df_from_table(table_name='Generator',
                                              fields='key,Num',
                                              setsel='all_disable')
            disable_df_gen['table'] = 'Generator'
            disable_df_gen.rename(columns={'Num': 's_key'}, inplace=True)
            disable_df_gen['index'] = disable_df_gen['s_key'].apply(lambda x: rm.dt.t_key_i['Generator'][x])
            # Узлы
            disable_df_node = rm.df_from_table(table_name='node',
                                               fields='name,uhom,key,ny',
                                               setsel='all_disable')

            disable_df_node['table'] = 'node'
            disable_df_node.sort_values(by=['uhom', 'name'],  # столбцы сортировки
                                        ascending=(False, True),  # обратный порядок
                                        inplace=True)
            disable_df_node.drop(['name'], axis=1, inplace=True)
            disable_df_node.rename(columns={'ny': 's_key'}, inplace=True)
            disable_df_node['index'] = disable_df_node['s_key'].apply(lambda x: rm.dt.t_key_i['node'][x])
            # Ветви
            self.disable_df_vetv = rm.df_from_table(table_name='vetv',
                                                    fields='name,key,temp,temp1,tip,ip,iq,np,ib,i_dop',
                                                    setsel='all_disable')
            self.disable_df_vetv['table'] = 'vetv'
            self.disable_df_vetv['uhom'] = (self.disable_df_vetv[['temp', 'temp1']].max(axis=1) * 10000 +
                                            self.disable_df_vetv[['temp', 'temp1']].min(axis=1))
            self.disable_df_vetv.sort_values(by=['tip', 'uhom', 'name'],  # столбцы сортировки
                                             ascending=(False, False, True),  # обратный порядок
                                             inplace=True)  # изменить df
            # todo использовать s_key_vetv_in_tuple
            self.disable_df_vetv['s_key'] = None
            for i in self.disable_df_vetv.index:
                self.disable_df_vetv.at[i, 's_key'] = (self.disable_df_vetv.at[i, 'ip'],
                                                       self.disable_df_vetv.at[i, 'iq'],
                                                       self.disable_df_vetv.at[i, 'np'],)

            self.disable_df_vetv.drop(['temp', 'temp1', 'tip', 'name', 'ip', 'iq', 'np'], axis=1, inplace=True)

            self.disable_df_vetv['index'] = self.disable_df_vetv['s_key'].apply(lambda x: rm.dt.t_key_i['vetv'][x])
            log_calc.info(f'Количество отключаемых элементов сети:'

                          f' ветвей - {len(self.disable_df_vetv.axes[0])},'
                          f' узлов - {len(disable_df_node.axes[0])},'
                          f' генераторов - {len(disable_df_gen.axes[0])}.')

            disable_df_all = pd.concat([self.disable_df_vetv.drop(['ib', 'i_dop'], axis=1),
                                        disable_df_node,
                                        disable_df_gen])
            # Фильтр комбинаций
            if not self.config['SRS']['n-1'] or not (self.config['SRS']['n-2_abv']
                                                     or self.config['SRS']['n-2_gd']
                                                     or self.config['SRS']['n-3']):
                self.config['filter_comb'] = False

            if self.config['filter_comb']:
                self.disable_df_vetv.set_index('index', inplace=True)
                self.disable_df_vetv.drop(['table', 'uhom'], axis=1, inplace=True)
                self.filter_comb = FilterCombination(diff=self.config['filter_comb_val'],
                                                     df_ib_norm=self.disable_df_vetv)

            # Цикл по всем возможным сочетаниям отключений
            for n_, self.info_srs['Контроль ДТН'] in self.set_comb.items():  # Цикл н-1 н-2 н-3.
                if n_ > len(disable_df_all):
                    break
                log_calc.info(f"Количество отключаемых элементов в комбинации: {n_} ({self.info_srs['Контроль ДТН']}).")
                if n_ == 1:
                    disable_all = disable_df_all.copy()
                else:
                    disable_all = \
                        disable_df_all[(disable_df_all['uhom'] > 300) | (disable_df_all['table'] != 'node')]
                disable_all.drop(['uhom'], axis=1, inplace=True)
                name_columns = list(disable_all.columns)
                disable_all = tuple(disable_all.itertuples(index=False, name=None))

                for comb in combinations(disable_all, r=n_):  # Цикл по комбинациям.
                    # log_calc.debug(f'Комбинация элементов {comb}')
                    comb_df = pd.DataFrame(data=comb, columns=name_columns)
                    unique_set_actions = []

                    comb_df['repair_scheme'] = False
                    comb_df['disable_scheme'] = False
                    comb_df['double_repair_scheme'] = False
                    for index in comb_df.index:
                        t = comb_df.loc[index, 'table']
                        k = comb_df.loc[index, 's_key']
                        for nm in ('repair_scheme', 'disable_scheme', 'double_repair_scheme'):
                            value = rm.dt.t_scheme[t][nm].get(k, False)
                            if value:
                                comb_df.at[index, nm] = 1  # долбаный глюк at
                                comb_df.at[index, nm] = value

                    # Если нет дополнительных изменений сети, то всего 1 сочетание.
                    if not comb_df[['disable_scheme', 'repair_scheme', 'double_repair_scheme']].any().any():
                        comb_df['status_repair'] = True
                        comb_df.loc[0, 'status_repair'] = False
                        self.calc_comb(rm, comb_df)
                        continue
                    comb_df['double_repair_scheme_copy'] = comb_df['double_repair_scheme']
                    # Цикл по всем возможным комбинациям внутри сочетания, вызванные ремонтами и отключениями.
                    # Под i понимаем номер отключаемого элемента, остальные в ремонте.
                    # Если -1, то ремонт всех элементов.

                    i_min = 0 if len(comb_df) == 3 else -1
                    for i in range(n_ - 1, i_min - 1, -1):  # От последнего до первого элемента или -1.

                        # Если в ремонте 2 элемента.
                        double_repair = True if (n_ == 2 and i == -1) or (n_ == 3) else False
                        if self.info_srs['Контроль ДТН'] == 'AДТН' and double_repair and n_ == 2:
                            continue  # Не расчетный случай по ГОСТ.

                        comb_df['status_repair'] = True  # Истина, если элемент в ремонте. Ложь отключен.
                        if i != -1:
                            comb_df.loc[i, 'status_repair'] = False

                        comb_df['double_repair_scheme'] = False
                        double_repair_scheme = []
                        if double_repair:
                            double_repair_scheme = self.find_double_repair_scheme(comb_df)

                        # Суммировать текущий набор изменений сети в set и проверить на уникальность.
                        set_actions = set()
                        for _, row in comb_df.iterrows():
                            if row['status_repair']:
                                if double_repair_scheme:
                                    set_actions.add(tuple(double_repair_scheme))
                                else:
                                    if row['repair_scheme']:
                                        set_actions.add(tuple(row['repair_scheme']))
                            else:
                                if row['disable_scheme']:
                                    set_actions.add(tuple(row['disable_scheme']))

                        if set_actions not in unique_set_actions:
                            unique_set_actions.append(set_actions)
                            self.calc_comb(rm, comb_df)

        # Отключаемые элементы сети по excel.
        if self.config['cb_disable_excel']:
            gen_comb_xl = self.combination_xl.gen_comb_xl(rm)
            for comb in gen_comb_xl:
                if comb['double_repair_scheme'].any():
                    self.find_double_repair_scheme(comb)
                self.info_srs['Контроль ДТН'] = 'ДДТН'
                if self.config['CalcSetWindow']['gost']:
                    if comb.shape[0] == 3 or (comb.shape[0] == 2 and rm.gost_abv):
                        self.info_srs['Контроль ДТН'] = 'АДТН'
                    if rm.gost_abv and (comb.shape[0] == 3 or (comb.shape[0] == 2 and all(comb['status_repair']))):
                        log_calc.info(f'Сочетание отклонено по ГОСТ: ')
                        log_calc.info(tabulate(comb, headers='keys', tablefmt='psql'))
                        continue
                self.calc_comb(rm, comb, source='xl')

        # Доработка перечня перегрузок РМ
        comb_min = min(self.all_comb[self.all_comb.rm_id == RastrModel.rm_id].comb_id.to_list())

        for key in self.breach:
            if len(self.breach[key]):
                sel = self.breach[key].comb_id >= comb_min
                if len(self.breach[key][sel]):
                    tabl = rm.dt.vetv_name if key == 'i' else rm.dt.node_name
                    self.breach[key].loc[sel, 'Контролируемые элементы'] = \
                        self.breach[key].loc[sel, ['s_key']] \
                            .merge(tabl, how='left') \
                            .set_index(self.breach[key].loc[sel].index)['Контролируемые элементы']

        if self.save_i:
            self.save_i.end_for_rm(rm, path_db=self.book_db)

        # Вывод таблиц К-О в excel
        if self.fill_table:
            self.fill_table.insert_tables_to_xl(name_rm=rm.info_file["Имя режима"],
                                                file_name=rm.info_file["Имя файла"],
                                                file_count=self.file_count,
                                                name_table=self.config['te_tab_KO_info'],
                                                path_xl=f'{self.config["name_time"]} таблицы К-О.xlsx')

    def calc_comb(self, rm, comb: pd.DataFrame, source: str = 'combinatorics'):
        """
        Смоделировать отключение элементов в комбинации.
        :param rm:
        :param comb:
        :param source: Варианты: 'combinatorics' или 'xl'
        :return:
        """
        # Фильтр н-2-3
        if self.config['filter_comb'] and len(comb) > 1 and source == 'combinatorics':
            if not self.filter_comb.test_comb(comb):
                return False

        # Восстановление схемы
        self.restore_only_state = rm.dt.recover_date_tables(self.restore_only_state)
        # log_calc.debug(tabulate(comb, headers='keys', tablefmt='psql'))
        comb.sort_values(by='status_repair', inplace=True)

        # Для добавления в 'Наименование СРС' данных о disable_scheme, double_repair_scheme и repair_scheme
        comb['scheme_info'] = ''
        log_calc.debug(tabulate(comb, headers='keys', tablefmt='psql'))

        # Отключение элементов
        repair2_one = True  # Для выполнения действия с двойным отключением на 2-м элементе.
        for i in comb.index:
            if not rm.sta(table_name=comb.loc[i, 'table'],
                          ndx=comb.loc[i, 'index']):  # отключить элемент
                log_calc.info(
                    f'Комбинация отклонена: элемент {rm.dt.t_name[comb.loc[i, "table"]][comb.loc[i, "s_key"]]!r}'
                    f' отключен в исходной РМ.')
                return False
            scheme_info = ''

            # Ремонтная схема
            if comb.loc[i, 'status_repair']:
                if comb.loc[i, 'double_repair_scheme']:
                    if repair2_one:
                        repair2_one = False
                    else:
                        scheme_info = self.perform_action(rm, comb.loc[i, 'double_repair_scheme'])
                else:
                    if comb.loc[i, 'repair_scheme']:
                        scheme_info = self.perform_action(rm, comb.loc[i, 'repair_scheme'])

            # Схема при отключении
            if (not comb.loc[i, 'status_repair']) and comb.loc[i, 'disable_scheme']:
                scheme_info = self.perform_action(rm, comb.loc[i, 'disable_scheme'])

            if scheme_info:
                comb.loc[i, 'scheme_info'] = f' ({scheme_info})'
        log_calc.debug('Элементы сети из сочетания отключены.')

        # Имя сочетания
        for k in ['Отключение', 'Ключ откл.', 'Ремонт 1', 'Ключ рем.1', 'Ремонт 2', 'Ключ рем.2']:
            self.info_srs.pop(k, None)  # Очистить

        dname = rm.dt.t_name[comb['table'].iloc[0]][comb['s_key'].iloc[0]]
        if comb.iloc[0]['status_repair']:
            all_name_srs = 'Ремонт '
            self.info_srs['Ремонт 1'] = dname + comb['scheme_info'].iloc[0]
            self.info_srs['Ключ рем.1'] = rm.key_to_str(comb['s_key'].iloc[0])
        else:
            all_name_srs = 'Отключение '
            self.info_srs['Отключение'] = dname + comb['scheme_info'].iloc[0]
            self.info_srs['Ключ откл.'] = rm.key_to_str(comb['s_key'].iloc[0])

        name_srs_base = all_name_srs + dname
        all_name_srs += dname + comb['scheme_info'].iloc[0]
        if len(comb) > 1:
            dname = rm.dt.t_name[comb['table'].iloc[1]][comb['s_key'].iloc[1]]
            all_name_srs += ' при ремонте' if 'Откл' in all_name_srs else ' и'
            all_name_srs += f' {dname}{comb["scheme_info"].iloc[1]}'
            name_srs_base += ' при ремонте' if 'Откл' in all_name_srs else ' и'
            name_srs_base += f' {dname}'
            if comb.iloc[0]['status_repair']:
                self.info_srs['Ремонт 2'] = dname + comb['scheme_info'].iloc[1]
                self.info_srs['Ключ рем.2'] = rm.key_to_str(comb['s_key'].iloc[1])
            else:
                self.info_srs['Ремонт 1'] = dname + comb['scheme_info'].iloc[1]
                self.info_srs['Ключ рем.1'] = rm.key_to_str(comb['s_key'].iloc[1])
        if len(comb) == 3:
            dname = rm.dt.t_name[comb['table'].iloc[2]][comb['s_key'].iloc[2]]
            all_name_srs += f', {dname}{comb["scheme_info"].iloc[2]}'
            name_srs_base += f', {dname}'
            self.info_srs['Ремонт 2'] = dname + comb['scheme_info'].iloc[2]
            self.info_srs['Ключ рем.2'] = rm.key_to_str(comb['s_key'].iloc[2])

        self.info_srs['Наименование СРС без()'] = name_srs_base  # re.sub(r'\(.+\)', '', all_name_srs).strip()
        all_name_srs += '.'

        self.info_srs['Наименование СРС'] = all_name_srs.strip()
        self.info_srs['comb_id'] = self.comb_id
        self.info_srs['Кол. откл. эл.'] = comb.shape[0]
        log_calc.info(f'Сочетание {self.comb_id}: {all_name_srs}')

        self.do_action(rm, comb)

    def perform_action(self, rm, task_action: list) -> str:
        """
        Выполнить действия, записанные в поле repair_scheme, disable_scheme.
        :param task_action: list('10', '2')
        :param rm:
        :return: Наименование внесенных изменений в расчетное НВ.
        """
        info = []
        # if not type(task_action) == tuple:
        #     task_action = tuple(task_action)
        for task_action_i in task_action:
            names, actions = self.pa.scheme_description(number=task_action_i)
            for i, action in enumerate(actions):
                name = rm.cor_rm_from_txt(action)
                if name:
                    if self.restore_only_state:
                        self.test_not_only_sta(name)
                    if names[i]:
                        name = names[i]
                    info.append(name)

        all_info = ', '.join(info) if info else ''
        return all_info

    def test_not_only_sta(self, txt):
        """
        Проверка на наличие изменений в сети кроме состояния.
        :param txt: Строка сформированная group_cor
        """
        for i in ['нагрузки', 'генерации', 'ktr', 'pn', 'qn', 'pg', 'qg', 'vzd', 'bsh', 'P']:
            # список параметров сверять с функцией rm.td.group_cor, rm.td.data_columns
            if i in txt:
                self.restore_only_state = False
                break

    def do_action(self, rm, comb=pd.DataFrame()):
        """
        Цикл по действиям ПА для ввода режима в область допустимых значений.
        :param rm:
        :param comb:
        """
        self.info_srs['comb_id'] = self.comb_id
        self.all_comb = pd.concat([self.all_comb, pd.Series(self.info_srs).to_frame().T],
                                  axis=0, ignore_index=True)
        self.info_action = dict()
        self.info_action['comb_id'] = self.comb_id
        self.info_action['active_id'] = 1
        self.info_action['End'] = False
        self.info_action['alive'] = 1
        # Если False - значит есть ПА, True - конец расчета сочетания (перегрузку нечем ликвидировать или отсутствует).
        # Цикл по действиям (ПА или ОП)
        while True:
            overloads = self.do_control(rm, comb)
            if self.config['CalcSetWindow']['pa'] and self.pa.active(overloads):  # TODO overloads не задано
                self.info_action['Action'] += self.pa.execute_action_pa(rm)
            else:
                if self.config['CalcSetWindow']['pa']:
                    self.pa.reset()
                self.info_action['End'] = True

            self.all_actions = pd.concat([self.all_actions, pd.Series(self.info_action).to_frame().T],
                                         axis=0, ignore_index=True)
            if self.info_action['End']:
                self.comb_id += 1  # код комбинации
                return

            self.info_action['active_id'] += 1
            for k in ['АРВ', 'СКРМ', 'Action']:
                self.info_action.pop(k, None)

    def do_control(self, rm, comb=pd.DataFrame()):
        """
        Проверка параметров режима.
        Заполняет таблицу Контроль - Отключение.
        Наполняет self.breach['i', 'low_u', 'high_u'].
        return overloads
        """
        log_calc.debug(f'Проверка параметров УР.')
        violation = False
        test_rgm = rm.rgm('do_control')
        if self.config['CalcSetWindow']['avr'] and len(comb):
            self.info_action['АРВ'] = rm.node_include()
            if 'Восстановлено' in self.info_action['АРВ']:
                test_rgm = rm.rgm('Перерасчет после действия АВР.')
        # if self.config['CalcSetWindow']['skrm']:
        #     self.info_action['СКРМ'] = rm.auto_shunt_cor(all_auto_shunt=self.auto_shunt)
        #     if self.info_action['СКРМ']:
        #         test_rgm = rm.rgm('do_control')

        if not test_rgm:
            self.info_action['alive'] = 0
            log_calc.debug(f'Режим развалился.')
            # return None
        else:
            # Сохранение загрузки отключаемых элементов в н-1 для фильтра
            if self.config['filter_comb'] and len(comb) == 1 and comb.table[0] == 'vetv':
                df_ip_n1 = rm.df_from_table(table_name='vetv',
                                            fields='key,ib',
                                            setsel='all_disable')
                self.filter_comb.add_n1(rm=rm, comb=comb, df_ip_n1=df_ip_n1)

            # проверка на наличие перегрузок ветвей (ЛЭП, трансформаторов, выключателей)
            if self.info_srs['Контроль ДТН'] == 'АДТН':
                selection_v = 'all_control & i_zag_av > 0.1004'
                selection_n = 'all_control & vras<umin_av & !sta'
            else:
                selection_v = 'all_control & i_zag > 0.1004'
                selection_n = 'all_control & vras<umin & !sta'

            tv = rm.rastr.tables('vetv')
            tv.SetSel(selection_v)
            # overloads_i = None
            # high_voltage = None
            # low_voltages = None
            if tv.count:
                overloads_i = rm.df_from_table(table_name='vetv',
                                               fields='s_key,'  # 'Ключ контроль,'
                                                      'txt_zag,'  # 'txt_zag,' 
                                                      'i_max,'  # 'Iрасч.(A),'
                                                      'i_dop_r,'  # 'Iддтн(A),'
                                                      'i_zag,'  # 'Iзагр.ддтн(%),'
                                                      'i_dop_r_av,'  # 'Iадтн(A),'
                                                      'i_zag_av',  # 'Iзагр.адтн(%),'
                                               setsel=selection_v)
                overloads_i['comb_id'] = self.comb_id
                overloads_i['active_id'] = self.info_action['active_id']
                self.breach['i'] = pd.concat([self.breach['i'], overloads_i], axis=0, ignore_index=True)
                violation = True
                log_calc.info(f'Выявлено {len(overloads_i)} превышений ДТН.')
            # проверка на наличие недопустимого снижение напряжения
            tn = rm.rastr.tables('node')
            tn.SetSel(selection_n)
            if tn.count:
                low_voltages = rm.df_from_table(table_name='node',
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
                                                setsel=selection_n)
                low_voltages.rename(columns={'ny': 's_key'}, inplace=True)
                low_voltages['comb_id'] = self.comb_id
                low_voltages['active_id'] = self.info_action['active_id']
                self.breach['low_u'] = pd.concat([self.breach['low_u'], low_voltages], axis=0, ignore_index=True)
                violation = True
                log_calc.info(f'Выявлено {len(low_voltages)} точек недопустимого снижения напряжения.')

            # проверка на наличие недопустимого повышения напряжения
            tn.SetSel('all_control & umax<vras & umax>0 & !sta')
            if tn.count:
                high_voltage = rm.df_from_table(table_name='node',
                                                fields='ny,'  # 'Ключ контроль,'
                                                       'vras,'  # 'Uрасч.(кВ),'
                                                       'umax,'  # 'Uнаиб.раб.(кВ)'
                                                       'otv_max',  # 'Uнаиб.раб.(кВ)'
                                                setsel='all_control & umax<vras & umax>0 & !sta')
                high_voltage.rename(columns={'ny': 's_key'}, inplace=True)
                high_voltage['comb_id'] = self.comb_id
                high_voltage['active_id'] = self.info_action['active_id']
                self.breach['high_u'] = pd.concat([self.breach['high_u'], high_voltage], axis=0, ignore_index=True)
                violation = True
                log_calc.info(f'Выявлено {len(high_voltage)} точек недопустимого превышения напряжения.')

            if self.save_i:
                self.save_i.add_data(rm,
                                     comb_id=self.comb_id,
                                     active_id=self.info_action['active_id'])

            # Таблица КОНТРОЛЬ - ОТКЛЮЧЕНИЕ
            if self.fill_table:
                self.fill_table.add_value(rm,
                                          name_srs=self.info_srs['Наименование СРС'],
                                          comb_id=self.comb_id,
                                          active_id=self.info_action["active_id"])
        # Добавить рисунки.
        if self.config['results_RG2'] and (not self.config['pic_overloads'] or
                                           (self.config['pic_overloads'] and violation)):
            self.drawings.draw(rm,
                               folder_path=self.config['name_time'],
                               comb_id=self.comb_id,
                               active_id=self.info_action["active_id"],
                               name_srs=self.info_srs["Наименование СРС без()"])
        return None

    @staticmethod
    def find_double_repair_scheme(comb_df):
        """
        Функция поиска общего действия double_repair_scheme в ремонтируемых элементах comb.
        Добавляет в колонку double_repair_scheme общее действие из колонки double_repair_scheme_copy и возвращает его.
        :param comb_df:
        """
        double_repair_scheme = []
        if comb_df.loc[comb_df['status_repair'], 'double_repair_scheme_copy'].all():
            double_repair_scheme = comb_df.loc[comb_df['status_repair'], 'double_repair_scheme_copy'].to_list()
            double_repair_scheme = list(set(double_repair_scheme[0]) & set(double_repair_scheme[1]))
            for i in comb_df.index:
                if comb_df['status_repair'].iloc[i]:
                    comb_df['double_repair_scheme'].iloc[i] = double_repair_scheme
                else:
                    comb_df['double_repair_scheme'].iloc[i] = False
        return double_repair_scheme
