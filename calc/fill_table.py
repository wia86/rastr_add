__all__ = ['FillTable']

"""Модуль для заполнения таблиц контролируемые - отключаемые элементы в excel."""
import logging
import os
from collections import defaultdict

import pandas as pd
import win32com.client
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

from common import Common

log_fill_bt = logging.getLogger(f'__main__.{__name__}')


class FillTable:
    # DF для хранения токовых перегрузок и недопустимого снижения U
    control_I = None
    control_U = None

    def __init__(self, rm):
        self.rm = rm
        log_fill_bt.debug('Инициализация таблицы "контролируемые - отключаемые" элементы.')

        self.control_I = rm.df_from_table(table_name='vetv',
                                          fields='dname,name,temp,temp1,i_dop_r,i_dop_r_av,groupid'
                                                 ',key,tip',  # ip, iq, np
                                          setsel='all_control')
        dname_list = []
        for dname, name in zip(list(self.control_I.dname), list(self.control_I.name)):
            if dname.strip():
                dname_list.append(dname)
            else:
                dname_list.append(name)
        self.control_I.dname = dname_list
        self.control_I.drop(['name'], axis=1, inplace=True)

        if len(self.control_I):
            # Сортировка
            self.control_I['uhom'] = (self.control_I[['temp', 'temp1']].max(axis=1) * 10000 +
                                      self.control_I[['temp', 'temp1']].min(axis=1))
            self.control_I.sort_values(by=['tip', 'uhom', 'dname'],  # столбцы сортировки
                                       ascending=(False, False, True),  # обратный порядок
                                       inplace=True)  # изменить df
            self.control_I.drop(['temp', 'temp1', 'uhom', 'tip'], axis=1, inplace=True)
            self.control_I['i_dop_r'] = self.control_I['i_dop_r'].round(0).astype(int)
            self.control_I['i_dop_r_av'] = self.control_I['i_dop_r_av'].round(0).astype(int)
            self.control_I.rename(columns={'i_dop_r': 'ДДТН, А',
                                           'i_dop_r_av': 'АДТН, А',
                                           'dname': 'Контролируемый элемент'}, inplace=True)
            self.control_I.set_index('index', inplace=True)
            self.control_I = self.control_I.T
            self.control_I.index = pd.MultiIndex.from_product([['-'], ['-'], self.control_I.index])

        self.control_U = rm.df_from_table(table_name='node',
                                          fields='index,dname,name,umin,umin_av,uhom',  # ny,umax
                                          setsel='all_control')

        dname_list = []
        for dname, name in zip(list(self.control_U.dname), list(self.control_U.name)):
            if dname.strip():
                dname_list.append(dname)
            else:
                dname_list.append(name)
        self.control_U.dname = dname_list
        self.control_U.drop(['name'], axis=1, inplace=True)

        if len(self.control_U):
            self.control_U.sort_values(by=['uhom', 'dname'],  # столбцы сортировки
                                       ascending=(False, True),  # обратный порядок
                                       inplace=True)  # изменить df
            self.control_U.drop(['uhom'], axis=1, inplace=True)
            self.control_U['umin'] = self.control_U['umin'].round(1)
            self.control_U['umin_av'] = self.control_U['umin_av'].round(1)
            self.control_U.rename(columns={'umin': 'МДН, кВ',
                                           'umin_av': 'АДН, кВ',
                                           'dname': 'Контролируемый элемент'}, inplace=True)
            self.control_U.set_index('index', inplace=True)
            self.control_U = self.control_U.T
            self.control_U.index = pd.MultiIndex.from_product([['-'], ['-'], self.control_U.index])

    def add_groupid(self):
        log_fill_bt.info('Добавление в контролируемые элементы ветвей по groupid.')
        table = self.rm.rastr.tables('vetv')
        table.setsel('all_control & groupid>0')
        if table.count:
            for gr in set(table.writesafearray('groupid', '000')):
                self.rm.group_cor(tabl='vetv',
                                  param='all_control',
                                  selection=f'groupid={gr[0]}',
                                  formula=1)

    def add_value(self):

        log_fill_bt.debug('Запись параметров УР в таблицу КО.')
        if len(self.control_I):
            ci = self.rm.df_from_table(table_name='vetv',
                                       fields='index,i_max,i_zag,i_zag_av',
                                       setsel='all_control')
            ci.set_index('index', inplace=True)
            ci['i_max'] = ci['i_max'].round(0)
            ci['i_zag'] = ci['i_zag'].round(0)
            ci['i_zag_av'] = ci['i_zag_av'].round(0)
            ci = ci.T
            ci.index = pd.MultiIndex.from_product([[self.info_srs['Наименование СРС']],
                                                   [f'{self.comb_id}.'
                                                    f'{self.info_action["active_id"]}'],
                                                   ['I, А', 'I, % от ДДТН', 'I, % от АДТН']])
            self.control_I = pd.concat([self.control_I, ci], axis=0)

        if len(self.control_U):
            cu = self.rm.df_from_table(table_name='node',
                                       fields='index,vras,otv_min,otv_min_av',
                                       setsel='all_control')
            cu.set_index('index', inplace=True)
            cu['vras'] = cu['vras'].round(1)
            cu['otv_min'] = cu['otv_min'].round(2)
            cu['otv_min_av'] = cu['otv_min_av'].round(2)
            cu = cu.T
            cu.index = pd.MultiIndex.from_product([[self.info_srs['Наименование СРС']],
                                                   [f'{self.comb_id}.'
                                                    f'{self.info_action["active_id"]}'],
                                                   ['U, кВ', 'U, % от МДН', 'U, % от АДН']])
            self.control_U = pd.concat([self.control_U, cu], axis=0)

    def insert_word(self):

        log_fill_bt.info('Вставить таблицы К-О в word.')

        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = False
        book = excel.Workbooks.Open(self.book_path)

        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        word.ScreenUpdating = False
        doc = word.Documents.Add()  # doc = word.Documents.Open(r'I:\file.docx')

        doc.PageSetup.PageWidth = 29.7 * 28.35  # CentimetersToPoints( format_list_i (2) ) 1 см = 28,35
        doc.PageSetup.PageHeight = 42.0 * 28.35  # CentimetersToPoints( format_list_i (1) )
        doc.PageSetup.Orientation = 1  # 1 книжная или 0 альбомная

        cursor = word.Selection
        cursor.Font.Size = 12
        cursor.Font.Name = 'Times New Roman'
        cursor.EndKey(Unit=6)  # перейти в конец текста

        for i in range(1, book.Worksheets.Count + 1):
            if book.Worksheets(i).name[:1].isnumeric():
                sheet = book.Worksheets(i)
                cursor.TypeText(Text=sheet.Cells(1, 1).value)
                cursor.TypeParagraph()

                sheet.Range(sheet.UsedRange.address.replace('$', '').replace('A1', 'A2')).Copy()
                # cursor.PasteExcelTable(LinkedToExcel=False, WordFormatting=False, RTF=False)
                cursor.PasteAndFormat(Type=13)  # 13 Вставить в виде рисунка.
                cursor.InsertBreak(Type=0)
                # разрыв:7 страницы с новой строки, 0-в той же строке,1 и 8 колонки,
                # 2-5 раздела со след стр,6 и 9-11 перенос на новую стр

        word.ScreenUpdating = True
        doc.SaveAs2(FileName=self.config['name_time'] + ' таблицы К-О.docx')  # FileFormat=16 .docx
        doc.Close()

    def finish(self):
        if not (len(self.control_I) or len(self.control_U)):
            return

        name_sheet = f'{self.file_count}_{self.rm.info_file["Имя файла"]}'.replace('[', '').replace(']', '')[:28]
        control_df_dict = {}
        if len(self.control_I):
            control_df_dict[name_sheet + '{I}'] = self.control_I
            self.control_I = None
        if len(self.control_U):
            control_df_dict[name_sheet + '{U}'] = self.control_U
            self.control_U = None
        # https://www.geeksforgeeks.org/how-to-write-pandas-dataframes-to-multiple-excel-sheets/

        # mode = 'a' if os.path.exists(self.book_path) else 'w'
        if not os.path.exists(self.book_path):
            Workbook().save(self.book_path)

        with pd.ExcelWriter(path=self.book_path, mode='a', engine='openpyxl') as writer:
            for name_sheet, df_control in control_df_dict.items():
                if '{I}' in name_sheet:
                    # Поиск столбцов с одинаковыми dname; ДДТН, А; АДТН, А; groupid
                    # https/www.geeksforgeeks.org/how-to-find-drop-duplicate-columns-in-a-pandas-dataframe/
                    df_control_head = df_control.iloc[:4].T  # включая groupid
                    duplicated_true = df_control_head.duplicated(keep=False)
                    groupid_true = df_control.loc['-', '-', 'groupid'] > 0
                    selection_columns = duplicated_true & groupid_true  # выборка в столбцах df_control для проверки

                    dict_equals = defaultdict(list)  # {номер:[перечень индексов столбцов с одинаковыми колонками]}
                    if selection_columns.any():
                        df_control_head = df_control_head[selection_columns]
                        duplicated_unique = df_control_head.drop_duplicates()
                        for i in range(len(duplicated_unique)):
                            col_unique = duplicated_unique.iloc[i, :]
                            for ii in range(len(df_control_head)):
                                control_col = df_control_head.iloc[ii, :]
                                if col_unique.equals(control_col):
                                    dict_equals[str(i)].append(int(control_col.name))
                    # Объединить столбцы с одинаковыми dname; ДДТН, А; АДТН, А; groupid
                    if dict_equals:
                        for cols in dict_equals.values():
                            df_control[cols[0]] = df_control[cols].max(axis=1)
                            df_control.drop(columns=cols[1:], inplace=True)

                df_control.to_excel(excel_writer=writer,
                                    sheet_name=name_sheet,
                                    header=False,
                                    startrow=1,
                                    freeze_panes=(2, 3),
                                    index=True)

        # Форматирование таблиц Отключение - Контроль
        wb = load_workbook(self.book_path)
        for name_sheet in control_df_dict:
            ws = wb[name_sheet]
            num_tab, name_tab = Common.read_title(self.config['te_tab_KO_info'])
            ws['A1'] = f'{name_tab[0]}{num_tab + self.file_count - 1}{name_tab[1]} {self.rm.info_file["Имя режима"]}'
            ws['A2'] = 'Наименование режима'
            ws['B2'] = 'Номер режима'
            ws['C2'] = 'Наименование параметра'
            # ws.merge_cells('A2:B4')
            thins = Side(border_style='thin', color='000000')
            max_column_lit = get_column_letter(ws.max_column)
            ws.merge_cells(f'A1:{max_column_lit}1')

            # Данные
            for row in range(3, ws.max_row + 1):
                for col in range(4, ws.max_column + 1):
                    ws.cell(row, col).border = Border(thins, thins, thins, thins)
                    if ws.cell(row, 3).value in ['I, А', 'U, кВ']:
                        ws.cell(row, col).font = Font(bold=True)
                    if 'I, %' in ws.cell(row, 3).value:
                        if ws.cell(row, col).value >= 100:
                            ws.cell(row, col).fill = PatternFill(fill_type='solid', fgColor='00FF9900')
                    if 'U, %' in ws.cell(row, 3).value:
                        if ws.cell(row, col).value > 0:
                            ws.cell(row, col).fill = PatternFill(fill_type='solid', fgColor='00FF9900')
            # Колонки
            for litter, L in {'A': 35, 'B': 6, 'C': 17}.items():
                ws.column_dimensions[litter].width = L
            for n in range(4, ws.max_column + 1):
                ws[f'{get_column_letter(n)}2'].alignment = Alignment(textRotation=90, wrap_text=True,
                                                                     horizontal='center', vertical='center')
                ws.column_dimensions[get_column_letter(n)].width = 9
                ws[f'{get_column_letter(n)}2'].font = Font(bold=True)
                ws[f'{get_column_letter(n)}2'].border = Border(thins, thins, thins, thins)
            # Строки
            if '{I}' in name_sheet:
                ws.row_dimensions[5].hidden = True  # Скрыть
                ws.row_dimensions[6].hidden = True
            for n in range(1, ws.max_row + 1):
                ws[f'A{n}'].alignment = Alignment(wrap_text=True, horizontal='left', vertical='center')
                ws[f'B{n}'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
                ws[f'C{n}'].alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            ws.row_dimensions[2].height = 145
        wb.save(self.book_path)
