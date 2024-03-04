import shutil

import pytest
import yaml

from calc_model import CalcModel


@pytest.fixture(scope='module')
def clear_dir():
    shutil.rmtree(r'I:\rastr_add2\test_rm\РМ v1\calc', ignore_errors=True)


def test_calc_models_xl(clear_dir, name_file=r'I:\rastr_add2\test_rm\test comb xl.calc'):
    """Отключения по excel."""
    with open(name_file) as f:
        print(CalcModel(yaml.safe_load(f)).run())


def test_calc_models_n1(clear_dir, name_file=r'I:\rastr_add2\test_rm\test n1.calc'):
    """Все возможные отключения н-1."""
    with open(name_file) as f:
        print(CalcModel(yaml.safe_load(f)).run())


def test_calc_models_n3(clear_dir, name_file=r'I:\rastr_add2\test_rm\test n3.calc'):
    """Все возможные отключения н-3."""
    with open(name_file) as f:
        print(CalcModel(yaml.safe_load(f)).run())
