
class Student(object):
    def __init__(self, name, gender):
        self.name = name
        self.__gender = gender

    def get_gender(self):
        print(self.__gender)

    def set_gender(self, gender):
        if gender != 'male' and gender != 'female':
            raise ValueError('wrong gender')
        else:
            self.__gender = gender


lisa = Student('Lisa', 'female')
lisa.get_gender()
lisa.set_gender('male')
lisa.get_gender()


