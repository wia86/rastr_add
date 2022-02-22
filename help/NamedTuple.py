from collections import namedtuple

#создать класс
car = namedtuple('авто', ['пробег', 'цвет'])
car = namedtuple('авто', 'пробег цвет')
uaz = namedtuple('грузовик', car._fields + ('прицеп',))
#создать класс и экземпляр
uaz2 = namedtuple('грузовик', car._fields + ('прицеп',))(100, 'red', 'yes')
print('uaz2:', uaz2)  # uaz2: грузовик(пробег=100, цвет='red', прицеп='yes')
#создать экземпляр
almera = car(200, 'зеленый')
print('uaz.__class__: ', uaz.__class__)# <class 'type'>
print(uaz.__dict__)
# доступ к полям
print(uaz._fields)#('пробег', 'цвет')
print(almera.пробег)
print(almera[0])
print(tuple(almera))  # использовать как картеж
print('обычный картеж ', (10,))
print(almera._asdict())  # возвращает содержимое именованного кортежа в виде словаря:

almera2 = almera._make([999,'красный', ])
# метод класса _make() может использоваться для создания новых экземпляров класса namedtuple из (итерируемой) последовательности:
print('almera2:', almera2)  #

print('_replace:', almera2._replace(пробег=100))  # создает (мелкую) копию кортежа и позволяет вам выборочно заменять
# некоторые его поля

from types import SimpleNamespace
car1 = SimpleNamespace(цвет='красный',пробег=3812.4,автомат=True)
print('car1:', car1)  #car1: namespace(цвет='красный', пробег=3812.4, автомат=True)
car1.пробег = 12
car1.лобовое_стекло = 'треснутое'
del car1.автомат
print('car1:', car1)