import pytest
import yaml

from calc_model import CalcModel


@pytest.mark.parametrize('name_file', (r'I:\rastr_add2\test_rm\test comb xl.calc',))
def test_edit_models(name_file):
    # Быстрый пуск из yaml файла EditModel CalcModel
    with open(name_file) as f:
        CalcModel(yaml.safe_load(f)).run()
