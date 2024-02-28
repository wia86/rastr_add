import configparser
import logging
import os

log_ini = logging.getLogger(f'__main__.{__name__}')


class Ini:
    """Работа с ini файлами"""

    def __init__(self, name):
        self.name = name  # Имя файла.ini

    def __str__(self):
        return self.name

    def exists(self):
        return True if os.path.exists(self.name) else False

    def add(self, info: dict, key: str):
        """Записать info в файле ini по ключу"""
        config = configparser.ConfigParser()
        config.read(self.name)
        config[key] = info
        with open(self.name, 'w') as configfile:
            config.write(configfile)

    def read_ini(self, section: str = '', key: str = ''):
        """
        Вернуть значение из файла .ini.
        """
        if self.exists:
            config = configparser.ConfigParser()
            config.read(self.name)
            try:
                if section and key:
                    return config[section][key]
                else:
                    if section:
                        return config[section]
                    else:
                        return config
            except LookupError:
                log_ini.error(f'Ошибка чтения файла {self} {section} {key}')
                return ''

    def write_ini(self, section: str, key: str, value):
        """
        Записать в файл .ini значение value в раздел section по ключу key.
        """
        config = configparser.ConfigParser()
        config.read(self.name)
        config[section] = {key: value}
        with open(self.name, 'w') as configfile:
            config.write(configfile)

    def to_dict(self) -> dict:
        """Вернуть .ini в виде словаря."""
        if self.exists():
            config = configparser.ConfigParser()
            config.read(self.name)
            pars = {}
            try:
                for section in config:
                    pars[section] = {}
                    for key in config[section]:
                        if config[section][key] in ['True', 'False']:
                            pars[section][key] = eval(config[section][key])
                        else:
                            pars[section][key] = config[section][key]
            except LookupError:
                raise LookupError(f'Ошибка чтения {self}.')
            return pars
        else:
            raise LookupError(f'Отсутствует файл {self}')
