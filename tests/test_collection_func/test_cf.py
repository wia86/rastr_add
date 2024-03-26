import pytest

from collection_func import s_key_vetv_in_tuple, convert_s_key, from_list1_only_exists_in_list2


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


@pytest.mark.parametrize("list1, list2, res_list", [([1, 2, 3], [2, 3, 4], [2, 3]),
                                                    ([1, 2, 3], (20, 30), [])])
def test_from_list1_only_exists_in_list2(list1, list2, res_list):
    assert from_list1_only_exists_in_list2(list1, list2) == res_list


@pytest.mark.parametrize("list1, list2, expected_exception", [([1, 2], '1, 2', TypeError),
                                                              ([1, 2], 1, TypeError),
                                                              ('1', '12', TypeError)])
def test_from_list1_only_exists_in_list2_error(list1, list2, expected_exception):
    with pytest.raises(expected_exception):
        from_list1_only_exists_in_list2(list1, list2)
