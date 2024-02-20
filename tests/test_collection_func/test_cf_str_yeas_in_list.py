from collection_func import str_yeas_in_list

import pytest
import numpy as np


@pytest.mark.parametrize("t, result", [('2021,2028', np.array([2021, 2028], int)),
                                       ('2021...2023', np.array([2021, 2022, 2023], int)),
                                       ('2021...2023, 2029', np.array([2021, 2022, 2023, 2029], int))])
def test_str_yeas_in_list(t, result):
    assert all(str_yeas_in_list(t) == result)


@pytest.mark.parametrize("id_str, expected_exception", [('2021..2023', ValueError),
                                                        ((2021, 2023,), AttributeError)])
def test_str_yeas_in_list_error(id_str, expected_exception):
    with pytest.raises(expected_exception):
        str_yeas_in_list(id_str)