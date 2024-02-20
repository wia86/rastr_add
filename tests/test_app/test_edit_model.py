import pytest
import yaml

from edit_model import EditModel


@pytest.mark.parametrize('name_file', (r'I:\rastr_add\test_rm\test cor rm.cor',))
def test_edit_models(name_file):

    # Быстрый пуск из yaml файла EditModel CalcModel
    with open(name_file) as f:
        EditModel(yaml.safe_load(f)).run()

