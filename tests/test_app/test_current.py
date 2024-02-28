import pytest
import yaml

from calc_model import CalcModel
# from edit_model import EditModel


@pytest.mark.skip()
@pytest.mark.parametrize('name_file', (r'I:\ОЭС Урала\Тюм_ЭС\КПР Тюменьэнерго до 2030 года\ИД СО\Элегаз по xl.calc',))
def test_current(name_file):
    with open(name_file) as f:
        CalcModel(yaml.safe_load(f)).run()

