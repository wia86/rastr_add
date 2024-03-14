import shutil

import pytest
import yaml

from calc_model import CalcModel


@pytest.fixture(scope='module')
def clear_dir():
    shutil.rmtree(r'I:\rastr_add2\test_rm\РМ v1\calc', ignore_errors=True)


def test_calc_model_xl_draw(clear_dir, name_file=r'I:\rastr_add2\test_rm\test comb xl and drawings.calc'):
    """Отключения по excel."""
    with open(name_file) as f:
        print(CalcModel(yaml.safe_load(f)).run())


def test_calc_model_n1(clear_dir, name_file=r'I:\rastr_add2\test_rm\test n1.calc'):
    """Все возможные отключения н-1."""
    with open(name_file) as f:
        print(CalcModel(yaml.safe_load(f)).run())


def test_calc_model_n3(clear_dir, name_file=r'I:\rastr_add2\test_rm\test n3.calc'):
    """Все возможные отключения н-3."""
    with open(name_file) as f:
        print(CalcModel(yaml.safe_load(f)).run())


def test_calc_model_fill_tabl(clear_dir, name_file=r'I:\rastr_add2\test_rm\test fill tabl.calc'):
    """Наполнение таблиц контролируемые - отключаемые элементы в excel."""
    with open(name_file) as f:
        print(CalcModel(yaml.safe_load(f)).run())
