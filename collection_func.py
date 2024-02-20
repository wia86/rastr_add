import numpy as np
import re


def str_yeas_in_list(id_str: str):
    """
    Преобразует перечень годов.
    :param id_str: "2021,2023...2025"
    :return: [2021,2023,2024,2025] или []
    """
    years_list = id_str.replace(" ", "").split(',')
    if years_list:
        years_list_new = np.array([], int)
        for it in years_list:
            if "..." in it:
                i_years = it.split('...')
                years_list_new = np.hstack(
                    [years_list_new, np.array(np.arange(int(i_years[0]), int(i_years[1]) + 1), int)])
            elif "…" in it:
                i_years = it.split('…')
                years_list_new = np.hstack(
                    [years_list_new, np.array(np.arange(int(i_years[0]), int(i_years[1]) + 1), int)])
            else:
                years_list_new = np.hstack([years_list_new, int(it)])
        return np.sort(years_list_new)
    else:
        return []


def split_task_action(txt: str) -> list | bool:
    """
    Разделить строку по запятым, если запятая не внутри [] {}
    :param txt: '[1,2,0:sta=1],[2,3:sta=0]{5,7:sta==1},[9,8:sta=1],6'
    :return: ['[1,2,0:sta=1]', '[2,3:sta=0]{5,7:sta==1}', '[9,8:sta=1]', '6'] или  False
    """
    if not txt:
        return False
    # Вычленить значения в [] и {}.
    actions = re.findall(re.compile(r"\[(.+?)]"), txt)
    conditions = re.findall(re.compile(r"\{(.+?)}"), txt)

    # Заменить значения в [ ] и { } на act_cond_{n}
    dict_key = {}  # замена, действие
    for n, action in enumerate(actions + conditions):
        dict_key[f'act_cond_{n}'] = action
        txt = txt.replace(action, f'act_cond_{n}')

    #  Заменить act_cond_{n} на значения в [ ] и { }.
    result = []
    for part in txt.split(','):
        for key in dict_key:
            if key in part:
                part = part.replace(key, dict_key[key])
        result.append(part)
    return result


if __name__ == '__main__':
    # print(str_yeas_in_list('2021...2025'))
    print(split_task_action('[1,2,0:sta=1],[2,3:sta=0]{5,7:sta==1},[9,8:sta=1],6'))
