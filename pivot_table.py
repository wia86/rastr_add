import itertools

import win32com.client


def make_pivot_tables(book_path: str,
                      sheets_info: dict):
    """
    Сделать сводную таблицу в файле path_xl_book
    :param book_path:
    :param sheets_info: Должен содержать
        {sheet_name_id:
         {'sheet_name': 'Сводная_i',
          'pt_name': 'pt_i',

          'conditional_formatting': {
              'num_field': [3, 5],
              'conditional_field': 'i_zag',
              'val': 100,}

          'row_fields': ['Контролируемые элементы',
                         'Отключение'],
          'page_fields': ['Имя файла',
                          'Кол. откл. эл.'],
          'column_fields' ['Год',
                           'Сезон макс/мин']
          }
    """

    assert bool(book_path)
    assert bool(sheets_info)

    excel = win32com.client.Dispatch('Excel.Application')
    excel.ScreenUpdating = False  # Обновление экрана
    try:
        book = excel.Workbooks.Open(book_path)
    except Exception:
        raise Exception(f'Ошибка при открытии файла {book_path=}')

    for sheet_name in sheets_info.keys():
        task = sheets_info[sheet_name]
        try:
            sheet = book.sheets[sheet_name]
        except Exception:
            raise Exception(f'Не найден лист: {sheet_name}')

        # Создать объект таблица из всего диапазона листа.
        obj_table = sheet.ListObjects.Add(SourceType=1,
                                          Source=sheet.Range(sheet.UsedRange.address))
        obj_table.Name = f'Таблица_{sheet_name}'
        obj_table_columns = get_list_columns(obj_table)

        assert set(task['row_fields']).issubset(set(obj_table_columns))
        assert set(task['column_fields']).issubset(set(obj_table_columns))
        assert set(task['page_fields']).issubset(set(obj_table_columns))

        # Создать КЭШ xlDatabase, ListObjects
        pt_cache = book.PivotCaches().add(1, obj_table)

        sheet_pivot = book.Sheets.Add()
        sheet_pivot.Name = task['sheet_name']

        pt = pt_cache.CreatePivotTable(TableDestination=task['sheet_name'] + '!R1C1',
                                       TableName=task['pt_name'])
        pt.ManualUpdate = True  # True не обновить сводную

        row_fields = task['row_fields']
        column_fields = task['column_fields']
        page_fields = task['page_fields']
        pt.AddFields(RowFields=row_fields,
                     ColumnFields=column_fields,
                     PageFields=page_fields,
                     AddToTable=False)

        # Поля значений
        for field_df, field_pt in task['data_field'].items():
            pt.AddDataField(Field=pt.PivotFields(field_df),
                            Caption=field_pt,
                            Function=-4136)  # xlMax -4136 xlSum -4157
            pt.PivotFields(field_pt).NumberFormat = '0'

        # if 'Контролируемые элементы' in row_fields:
        #     pt.PivotFields('Контролируемые элементы').ShowDetail = True  # группировка

        pt.RowAxisLayout(1)  # Показывать в табличной форме.

        if len(task['data_field']) > 1:
            pt.DataPivotField.Orientation = 1
            # xlRowField = 1 'Значения' в столбцах или строках xlColumnField = 2.

        pt.RowGrand = False  # Удалить строку общих итогов.
        pt.ColumnGrand = False  # Удалить столбец общих итогов.
        pt.MergeLabels = True  # Объединять одинаковые ячейки.
        pt.HasAutoFormat = False  # Не обновлять ширину при обновлении.
        pt.NullString = '--'  # Заменять пустые ячейки.
        pt.PreserveFormatting = False  # Сохранять формат ячеек при обновлении.
        pt.ShowDrillIndicators = False  # Показывать кнопки свертывания.

        for row in itertools.chain(row_fields, column_fields):
            # Промежуточные итоги и фильтры.
            pt.PivotFields(row).Subtotals = [False] * 12

        pt.TableStyle2 = ""
        pt.ColumnRange.ColumnWidth = 10  # Ширина строк.
        pt.RowRange.ColumnWidth = 20
        pt.DataBodyRange.HorizontalAlignment = -4108  # xlCenter = -4108
        pt.TableRange1.WrapText = True  # Перенос текста в ячейке

        # Условное форматирование
        if task['conditional_formatting']:
            cf = task['conditional_formatting']
            for i in cf['num_field']:
                dpz = pt.DataBodyRange.Rows(i).Cells(1)
                dpz_fc = dpz.FormatConditions
                dpz_fc.AddColorScale(2)
                dpz_fc(dpz_fc.count).SetFirstPriority()

                dpz_fc_csc = dpz_fc(1).ColorScaleCriteria(1)
                dpz_fc_csc.Type = 0  # xlConditionValueNumber = 0

                dpz_fc_csc.Value = cf['val']

                dpz_fc_csc.FormatColor.ThemeColor = 1  # xlThemeColorDark1=1
                dpz_fc_csc.FormatColor.TintAndShade = 0

                dpz_fc_csc2 = dpz_fc(1).ColorScaleCriteria(2)
                dpz_fc_csc2.Type = 2  # xlConditionValueHighestValue = 2
                dpz_fc_csc2.FormatColor.ThemeColor = 3 + i  # Номер темы
                dpz_fc_csc2.FormatColor.TintAndShade = -0.249977111117893

                dpz_fc(1).ScopeType = 2  # xlDataFieldScope = 2 применить ко всем значениям поля
        pt.ManualUpdate = False  # Обновить сводную.

    book.Save()
    book.Close()
    excel.Quit()


def get_list_columns(obj_table) -> list:
    """
    Вернуть список полей в таблице (Excel Application).
    :param obj_table:
    :return:
    """
    count_columns = obj_table.ListColumns.Count
    list_columns = []
    for i in range(1, count_columns):
        list_columns.append(obj_table.ListColumns(i).Name)
    return list_columns
