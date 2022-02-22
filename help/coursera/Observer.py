from abc import ABC, abstractmethod


class NotificationManager:  # Наблюдаемая система
    def __init__(self):
        self.__subscribers = set()  # При инициализации множество подписчиков задается пустым

    def subscribe(self, subscriber):
        self.__subscribers.add(
            subscriber)  # Для того чтобы подписать пользователя, он добавляется во множество подписчиков

    def unsubcribe(self, subscriber):
        self.__subscribers.remove(subscriber)  # Удаление подписчика из списка

    def notify(self, message):
        for subscriber in self.__subscribers:
            subscriber.update(message)  # Отправка уведомления всем подписчикам


class AbstractObserver(ABC):
    @abstractmethod
    def update(self, message):  # Абстрактный наблюдатель задает метод update
        pass


class MessageNotifier(AbstractObserver):
    def __init__(self, name):
        self.__name = name

    def update(self, message):  # Конкретная реализация метода update
        print(f'{self.__name} recieved message!')


class MessagePrinter(AbstractObserver):
    def __init__(self, name):
        self.__name = name

    def update(self, message):  # Конкретная реализация метода update
        print(f'{self.__name} recieved message: {message}')

notifier1 = MessageNotifier("Notifier1")
printer1 = MessagePrinter("Printer1")
printer2 = MessagePrinter("Printer2")

manager = NotificationManager()

manager.subscribe(notifier1)
manager.subscribe(printer1)
manager.subscribe(printer2)

manager.notify("Hi!")