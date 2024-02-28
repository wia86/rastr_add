import pytest

from collection_func import str_yeas_in_list


@pytest.mark.parametrize("t, result", [('2021,2028', [2021, 2028]),
                                       ('2021...2023', [2021, 2022, 2023]),
                                       ('2021...2023, 2029', [2021, 2022, 2023, 2029])])
def test_str_yeas_in_list(t, result):
    assert str_yeas_in_list(t) == result


@pytest.mark.parametrize("t, s, result", [('2021-2023, 2029', ('-',), [2021, 2022, 2023, 2029])])
def test_str_yeas_in_list(t, s, result):
    assert str_yeas_in_list(str_init=t, sep=s) == result


@pytest.mark.parametrize("t, expected_exception", [('2021..2023', ValueError),
                                                   ((2021, 2023,), AttributeError)])
def test_str_yeas_in_list_error(t, expected_exception):
    with pytest.raises(expected_exception):
        str_yeas_in_list(t)


@pytest.mark.parametrize("t, expected_exception", [('2023...2021', ValueError)])
def test_str_yeas_in_list_error(t, expected_exception):
    with pytest.raises(expected_exception) as excep:
        str_yeas_in_list(t)
    assert excep.value.args[0] == f'Неверный формат: {t}'
