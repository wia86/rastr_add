import numpy as np

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


if __name__ == '__main__':
    to_str((1, 2))
