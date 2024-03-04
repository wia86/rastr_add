import pytest
import yaml

from calc_model import CalcModel
# from edit_model import EditModel


@pytest.mark.skip()
@pytest.mark.parametrize('name_file', (r'I:\ОЭС Северо-Запада\КПР 2023\Модели РДУ СЗ\ПРМ УР+КЗ\ЭС Архангельск – Коми\откл ВЛ Гриба.calc',))
def test_current(name_file):
    with open(name_file) as f:
        print(CalcModel(yaml.safe_load(f)).run())

