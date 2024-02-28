import re


def str_yeas_in_list(str_init: str, sep: tuple = ('...', '…')) -> list:
    """
    Преобразует перечень годов с диапазонами в отсортированный массив.
    :param str_init: '2021,2023...2025'
    :param sep: Картеж разделителей.
    :return: [2021,2023,2024,2025] или []
    """
    list_init = str_init.replace(' ', '').split(',')
    if not list_init:
        return []
    years_list_new = []
    for list_init_i in list_init:
        for sep_i in sep:
            if sep_i in list_init_i:
                min_val, max_val = list_init_i.split(sep_i)
                min_val = int(min_val)
                max_val = int(max_val)
                if min_val > max_val:
                    raise ValueError(f'Неверный формат: {str_init}')
                min_max = list(range(min_val, max_val + 1))
                years_list_new.extend(min_max)
                break
        else:
            years_list_new.append(int(list_init_i))
    return sorted(years_list_new)


def split_task_action(txt: str) -> list | bool:
    """
    Разделить строку по запятым, если запятая не внутри [] {}
    :param txt: '[1,2,0:sta=1],[2,3:sta=0]{5,7:sta==1},[9,8:sta=1],6'
    :return: ['[1,2,0:sta=1]', '[2,3:sta=0]{5,7:sta==1}', '[9,8:sta=1]', '6'] или  False
    """
    if not txt:
        return False
    # Вычленить значения в [] и {}.
    actions = re.findall(re.compile(r'\[(.+?)]'), txt)
    conditions = re.findall(re.compile(r'\{(.+?)}'), txt)

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
