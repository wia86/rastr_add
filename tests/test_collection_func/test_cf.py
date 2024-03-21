import pytest

from collection_func import s_key_vetv_in_tuple, convert_s_key


@pytest.mark.parametrize("s_key, result", [('2021,2028', (2021, 2028, 0)),
                                           ('2021,2023,0', (2021, 2023, 0)),
                                           ('1,2,3', (1, 2, 3))])
def test_s_key_vetv_in_tuple(s_key, result):
    assert s_key_vetv_in_tuple(s_key) == result


@pytest.mark.parametrize("t, expected_exception", [(10, TypeError),
                                                   ('10', ValueError),
                                                   ('10, 10,10,10', ValueError)])
def test_s_key_vetv_in_tuple_error(t, expected_exception):
    with pytest.raises(expected_exception):
        s_key_vetv_in_tuple(t)


@pytest.mark.parametrize("s_key, result", [('2021,2028', (2021, 2028, 0)),
                                           (10, 10)])
def test_convert_s_key(s_key, result):
    assert convert_s_key(s_key) == result


@pytest.mark.parametrize("t, expected_exception", [([1, 2], TypeError),
                                                   ((10, 20, 0), TypeError)])
def test_convert_s_key_error(t, expected_exception):
    with pytest.raises(expected_exception):
        convert_s_key(t)
