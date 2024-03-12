__all__ = ['ManagerPrintParameters']

import os.path
from abc import ABC

import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator
import openpyxl
import pandas as pd


class ManagerPrintParameters(ABC):
    _dict_obj = {}  # строка задания: объект PrintParameters

    @classmethod
    def add(cls, rm, task: str):
        """
        Добавить значения из rm в объект класса PrintParameters.
        Если он отсутствует, то создать его и сохранить в данном классе.
        :param rm:
        :param task: '15177,16095:r;x;b / 15106;15138:pn;qn / ns=2(sechen):psech'
        """
        if not task:
            raise ValueError('Ошибка: не указаны входные параметры.')
        if not isinstance(task, str):
            raise TypeError('Ошибка: неверный тип данных.')

        obj = cls._dict_obj.setdefault(task, PrintParameters(task))
        obj.add_val_parameters(rm)

    @classmethod
    def all_data_in_excel(cls, path_excel: str):
        if not cls._dict_obj:
            return
        if not path_excel:
            raise ValueError('Не указан путь для сохранения.')
        for key in cls._dict_obj:
            cls._dict_obj[key].data_in_excel(path_excel)
        cls._dict_obj = {}


class PrintParameters:
    """ Вывод параметров режима в таблицу excel."""
    _set_output_parameters = set()  # set((key, param), (ny, pn), (10, pn))
    _data_parameters = pd.DataFrame()
    _num_class = 0

    @classmethod
    def _receive_name_sheet(cls):
        cls._num_class += 1
        return f'par{cls._num_class}'

    def __init__(self, task: str):
        """
        :param task: '15177,16095:r;x;b / 15106;15138:pn;qn / ns=2(sechen):psech'
        """
        self._task = task

        task = task.replace(' ', '').split('/')
        for task_i in task:
            keys, parameters = task_i.split(':')  # нр'8;9', 'pn;qn'
            for parameter in parameters.split(';'):  # ['pn','qn']
                for key in keys.split(';'):  # ['15105,15113','15038,15037,4']
                    self._set_output_parameters.add((key, parameter,))

    def add_val_parameters(self, rm):
        """
        Добавить на sheet excel.
        """
        if 'sechen' in self._task:
            if rm.rastr.tables.Find('sechen') < 0:
                rm.downloading_additional_files('sch')

        date = pd.Series(dtype='object')
        for k, p in self._set_output_parameters:
            table, sel = rm.recognize_key(key=k, back='tab sel')
            key = f'{k}_{p}'
            if rm.rastr.tables(table).Cols.Find(p) == -1:
                raise ValueError(f'В таблице {table} отсутствует поле {p}')
            for name in ['dname', 'name', 'Name']:
                if rm.rastr.tables(table).Cols.Find(name) != -1:
                    name_cur = rm.txt_field_return(table, sel, name)
                    if name_cur:
                        key = f'{name_cur} ({key})'
                        break
            if rm.rastr.tables(table).cols(p).Prop(1) == 2:  # если поле типа строка
                date.loc[key] = rm.txt_field_return(table, sel, p)
            else:
                date.loc[key] = rm.rastr.tables(table).cols.Item(p).Z(rm.index(table_name=table,
                                                                               key_str=sel))

        date = pd.concat([date, pd.Series(rm.info_file)])
        self._data_parameters = pd.concat([self._data_parameters, date], axis=1)

    def data_in_excel(self, path_excel: str):
        """Записать данные в excel"""
        self._data_parameters = self._data_parameters.T

        with pd.ExcelWriter(path=path_excel,
                            mode='w',
                            engine='openpyxl') as writer:
            self._data_parameters.to_excel(excel_writer=writer,
                                           sheet_name=self._receive_name_sheet(),
                                           header=True,
                                           index=False)
        self._data_processing(path_excel)

    def _data_processing(self, path_excel):
        # Салехард без графиков
        file_names = self._data_parameters['Имя файла'].unique().tolist()
        for file_name in file_names:
            data = self._data_parameters[self._data_parameters['Имя файла'] == file_name]

            data = data[['узел салехард (nga=1_dp)',
                         'ПС 220 кВ Салехард : 1СШ-220 (10901541_vras)',
                         'ПС 220 кВ Салехард : 1СШ-110 (10901542_vras)',
                         'ПС 220 кВ Салехард : 2СШ-220 (10901145_qg)',
                         'ПС 220 кВ Салехард : 2СШ-110 (10901034_qsh)']]
            data.rename(columns={'узел салехард (nga=1_dp)': 'Потери в Салехардском узле',
                                 'ПС 220 кВ Салехард : 1СШ-220 (10901541_vras)': 'Напряжение на шинах 220 кВ ПС Салехард',
                                 'ПС 220 кВ Салехард : 1СШ-110 (10901542_vras)': 'Напряжение на шинах 110 кВ ПС Салехард',
                                 'ПС 220 кВ Салехард : 2СШ-220 (10901145_qg)': 'Мощность УШР (Салехард)',
                                 'ПС 220 кВ Салехард : 2СШ-110 (10901034_qsh)': 'Мощность БСК (Салехард)',
                                 }, inplace=True)

            data['Мощность УШР (Салехард)'] = - data['Мощность УШР (Салехард)']

            sheet_name = file_name.replace('ПРМ с ВЭ Тюменской ЭС ', '')

            with pd.ExcelWriter(path=path_excel,
                                mode='a',
                                engine='openpyxl') as writer:
                data.to_excel(excel_writer=writer,
                              sheet_name=sheet_name,
                              header=True,
                              index=False)

            bsk_not = data[data['Мощность БСК (Салехард)'] == 0]
            x0 = bsk_not['Мощность УШР (Салехард)'].values
            y0dp = bsk_not['Потери в Салехардском узле'].values
            y0u110 = bsk_not['Напряжение на шинах 110 кВ ПС Салехард'].values
            y0u220 = bsk_not['Напряжение на шинах 220 кВ ПС Салехард'].values
            bsk_yes = data[data['Мощность БСК (Салехард)'] != 0]
            x1 = bsk_yes['Мощность УШР (Салехард)'].values
            y1dp = bsk_yes['Потери в Салехардском узле'].values
            y1u110 = bsk_yes['Напряжение на шинах 110 кВ ПС Салехард'].values
            y1u220 = bsk_yes['Напряжение на шинах 220 кВ ПС Салехард'].values

            f = plt.figure(figsize=(16, 5))

            ax1 = f.add_subplot(1, 3, 1)
            ax1.plot(x0, y0dp, '--o', label='БСК отключены')
            ax1.plot(x1, y1dp, ':s', label='Включена одна БСК')

            ax2 = f.add_subplot(1, 3, 2)
            ax2.plot(x0, y0u110, '--o', label='БСК отключены')
            ax2.plot(x1, y1u110, ':s', label='Включена одна БСК')

            ax3 = f.add_subplot(1, 3, 3)
            ax3.plot(x0, y0u220, '--o', label='БСК отключены')
            ax3.plot(x1, y1u220, ':s', label='Включена одна БСК')

            ax1.grid()
            ax1.set_ylabel('Потери мощности в узле Салехард')
            ax1.set_xlabel('Мощность реактора')
            ax1.xaxis.set_major_locator(MultipleLocator(10))
            ax1.legend()
            ax2.legend()
            ax3.legend()

            ax2.grid()
            ax2.set_ylabel('Напряжение на шинах 220 кВ Салехард')
            ax2.set_xlabel('Мощность реактора')
            ax2.xaxis.set_major_locator(MultipleLocator(10))

            ax3.grid()
            ax3.set_ylabel('Напряжение на шинах 110 кВ Салехард')
            ax3.set_xlabel('Мощность реактора')
            ax3.xaxis.set_major_locator(MultipleLocator(10))

            # plt.show()
            # plt.show()
            path_pic = os.path.join(os.path.dirname(path_excel), 'plot.png')
            plt.savefig(path_pic)

            # Вставить графики в word
            wb = openpyxl.load_workbook(path_excel)
            ws = wb[sheet_name]
            for c in 'ABCDEF':
                ws.column_dimensions[c].width = 40
            img = openpyxl.drawing.image.Image(path_pic)
            img.anchor = 'A14'
            ws.add_image(img)
            wb.save(path_excel)

    def _data_processing2(self, path_excel):
        # Надым c графиками
        file_names = self._data_parameters['Имя файла'].unique().tolist()
        for file_name in file_names:
            data = self._data_parameters[self._data_parameters['Имя файла'] == file_name]

            data = data[['узел салехард (nga=1_dp)',
                         'ПС 220 кВ Надым : 1СШ-220 (10901530_vras)',
                         'ПС 220 кВ Надым : 1С-110 (10901511_vras)',
                         'ПС 220 кВ Надым : 2СШ-220 (10901139_qg)',
                         'ПС 220 кВ Надым : ОСШ-110 (10901026_qsh)']]
            data['Сумма qsh'] = (data['ПС 220 кВ Надым : ОСШ-110 (10901026_qsh)']
                                 - data['ПС 220 кВ Надым : 2СШ-220 (10901139_qg)'])
            sheet_name = file_name.replace('ПРМ с ВЭ Тюменской ЭС ', '')

            with pd.ExcelWriter(path=path_excel,
                                mode='a',
                                engine='openpyxl') as writer:
                data.to_excel(excel_writer=writer,
                              sheet_name=sheet_name,
                              header=True,
                              index=False)
            # Графики
            x = data['Сумма qsh'].values
            f = plt.figure(figsize=(16, 5))
            ax1 = f.add_subplot(1, 3, 1)
            ax1.plot(x, data['узел салехард (nga=1_dp)'].values)
            ax2 = f.add_subplot(1, 3, 2)
            ax2.plot(x, data['ПС 220 кВ Надым : 1СШ-220 (10901530_vras)'].values)
            ax3 = f.add_subplot(1, 3, 3)
            ax3.plot(x, data['ПС 220 кВ Надым : 1С-110 (10901511_vras)'].values)

            ax1.grid()
            ax1.set_ylabel('Потери мощности в узле Салехард')
            ax1.set_xlabel('Мощность реактора')
            ax2.grid()
            ax2.set_ylabel('Напряжение на шинах 220 кВ Надым')
            ax2.set_xlabel('Мощность реактора')
            ax3.grid()
            ax3.set_ylabel('Напряжение на шинах 110 кВ Надым')
            ax3.set_xlabel('Мощность реактора')

            # plt.show()
            path_pic = os.path.join(os.path.dirname(path_excel), 'plot.png')
            plt.savefig(path_pic)

            # Вставить графики в word
            wb = openpyxl.load_workbook(path_excel)
            ws = wb[sheet_name]
            for c in 'ABCDEF':
                ws.column_dimensions[c].width = 40
            img = openpyxl.drawing.image.Image(path_pic)
            img.anchor = 'A14'
            ws.add_image(img)
            wb.save(path_excel)
