class ExceptionTask(Exception):
    """Обработка исключений связанных с неправильным заданием исходных данных.
    Данные можно исправить и повторно запустить задание без перезапуска программы."""

    def __init__(self, *args):
        self.message = args[0] if args else None

    def __str__(self):
        return f"Ошибка: {self.message}"

