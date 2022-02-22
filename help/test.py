conditions='years = 2026| season=лет| max_min=min| add_name=0°C'
e = conditions.split('|')
condition_dict = dict(*e)


class Car:
    def __init__(self, color, mileage):
        self.color = color
        self.mileage = mileage

    def __repr__(self):
        return (f'{self.__class__.__name__}('
            f'{self.color!r}, {self.mileage!r})')

    # def __str__(self):
    #     return '__str__ для объекта Car'


c= Car('зеленый', '454')
i=1