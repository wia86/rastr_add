from collection_func import split_task_action

import pytest


@pytest.mark.parametrize("t, result",
                         [('[1,2,0:sta=1],[2,3:sta=0]{5,7:sta==1},[9,8:sta=1],6',
                           ['[1,2,0:sta=1]', '[2,3:sta=0]{5,7:sta==1}', '[9,8:sta=1]', '6']),
                          ('[1,2,0:sta=1]', ['[1,2,0:sta=1]']),
                          ('', False), ])
def test_split_task_action(t, result):
    assert split_task_action(t) == result


@pytest.mark.parametrize("t, expected_exception", [(['[1,2,0:sta=1]', 6], TypeError)])
def test_split_task_action_error(t, expected_exception):
    with pytest.raises(expected_exception):
        split_task_action(t)
