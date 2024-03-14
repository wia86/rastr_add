"""Модуль для отладки при написании кода."""
import pytest
import yaml

from calc_model import CalcModel
from edit_model import EditModel


@pytest.mark.skip()
@pytest.mark.parametrize('name_file', (r'I:\rastr_add2\test_rm\test fill tabl del big.calc',))
def test_current(name_file):
    with open(name_file) as f:
        print(CalcModel(yaml.safe_load(f)).run())

